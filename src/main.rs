#![allow(non_camel_case_types)]
#![allow(non_snake_case)]
#![allow(unused)]

use regex::Regex;
use serde::Serialize;
use serde::Deserialize;
use wmi::{COMLibrary, Variant, WMIConnection, WMIDateTime};
use std::collections::HashMap;

#[derive(Deserialize, Debug)]
struct Win32_PnPEntity {
    DeviceID: Option<String>,
    Name: Option<String>
}

#[derive(Deserialize, Debug)]
struct Win32_PortResource {
    StartingAddress: u64,
    EndingAddress: u64,
    CSName: String,
    Status: String,
}

#[derive(Serialize, Deserialize, Debug)]
struct LPTPort {
    name: String,
    device_id: Option<String>,
    cs_name: String,
    status: String,
    address_start: u64,
    address_end: u64
}

fn extract_lpt_name(whole_name: &str) -> Option<String> {
    let re = Regex::new(r"\bLPT\d+\b").unwrap();
    if let Some(mat) = re.find(&whole_name) {
        Some(mat.as_str().to_string())
    } else {
        None
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let com_con = COMLibrary::new()?;
    let wmi_con = WMIConnection::new(com_con.into())?;

    let results: Vec<Win32_PnPEntity> = wmi_con.raw_query("SELECT DeviceID, Name FROM Win32_PnPEntity WHERE ClassGuid='{4D36E978-E325-11CE-BFC1-08002BE10318}'")?;

    if results.len() == 0 {
        println!("No LPT ports found");
    } else {
        for os in results {
            if let Some(name) = os.Name {
                if let Some(lpt_id) = extract_lpt_name(name.as_str()) {
                    let query: String = format!("ASSOCIATORS OF {{Win32_PnPEntity.DeviceID=\"{}\"}} WHERE ResultClass = Win32_PortResource", &os.DeviceID.as_ref().unwrap()).replace("\\", "\\\\");
                    let results: Vec<Win32_PortResource> = wmi_con.raw_query(query)?;
                    let mut lpt_ports: Vec<LPTPort> = vec![];
                    for a in results {
                        lpt_ports.push(LPTPort { name: lpt_id.clone(), device_id: os.DeviceID.clone(), cs_name: a.CSName.clone(), status: a.Status.clone(), address_start: a.StartingAddress, address_end: a.EndingAddress });
                        println!("{:?}", lpt_ports.last().unwrap());
                    }
                } 
            }
        }
    }

    Ok(())
}
