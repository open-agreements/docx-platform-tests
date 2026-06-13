use std::{env, fs, process};

const PROTOCOL_VERSION: &str = "1";
const DOCX_RS_VERSION: &str = "0.4.20";

fn main() {
    if env::args().nth(1).as_deref() == Some("--print-library-version") {
        let _ = std::any::type_name::<docx_rs::Docx>();
        println!("docx-rs {DOCX_RS_VERSION}");
        return;
    }

    let protocol_version = arg_value("--protocol-version");
    if protocol_version.as_deref() != Some(PROTOCOL_VERSION) {
        println!(
            "docx-rs adapter speaks protocol v{}, got {}",
            PROTOCOL_VERSION,
            protocol_version.unwrap_or_default()
        );
        process::exit(3);
    }

    let Some(operation_path) = arg_value("--operation") else {
        eprintln!("missing required adapter protocol arguments");
        process::exit(1);
    };
    let operation_text = fs::read_to_string(operation_path).unwrap_or_else(|err| {
        eprintln!("failed to read operation descriptor: {err}");
        process::exit(1);
    });
    let operation: serde_json::Value = serde_json::from_str(&operation_text).unwrap_or_else(|err| {
        eprintln!("failed to parse operation descriptor: {err}");
        process::exit(1);
    });
    let operation_name = operation
        .get("operationName")
        .and_then(|value| value.as_str())
        .unwrap_or_default();

    match operation_name {
        "acceptAllTrackedChanges" | "rejectAllTrackedChanges" => {
            unsupported(
                "docx-rs can generate revision markup but exposes no API to accept or reject existing tracked changes",
            );
        }
        "replaceFirstTextOccurrence" => {
            unsupported(
                "docx-rs exposes no public existing-document text replacement API; implementing this would require adapter-side XML surgery",
            );
        }
        _ => unsupported(&format!(
            "docx-rs adapter does not implement operation '{operation_name}'"
        )),
    }
}

fn arg_value(flag: &str) -> Option<String> {
    let args: Vec<String> = env::args().collect();
    args.windows(2)
        .find(|window| window[0] == flag)
        .map(|window| window[1].clone())
}

fn unsupported(reason: &str) -> ! {
    println!("{reason}");
    process::exit(2);
}
