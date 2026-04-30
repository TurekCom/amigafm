use std::env;
use std::fs;
use std::path::PathBuf;

fn main() {
    println!("cargo:rerun-if-changed=x64/nvdaControllerClient.dll");
    println!("cargo:rerun-if-changed=x64");

    let manifest_dir = PathBuf::from(env::var("CARGO_MANIFEST_DIR").expect("missing manifest dir"));
    let x64_dir = manifest_dir.join("x64");
    let source_dll = x64_dir.join("nvdaControllerClient.dll");
    if !source_dll.exists() {
        return;
    }

    let out_dir = PathBuf::from(env::var("OUT_DIR").expect("missing out dir"));
    let profile_dir = out_dir
        .ancestors()
        .nth(3)
        .expect("failed to locate target profile directory");
    let target_dll = profile_dir.join("nvdaControllerClient.dll");

    let _ = fs::copy(source_dll, target_dll);

    if let Ok(entries) = fs::read_dir(&x64_dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            let Some(name) = path.file_name().and_then(|name| name.to_str()) else {
                continue;
            };
            let lower = name.to_ascii_lowercase();
            if lower.starts_with("bass") && lower.ends_with(".dll") {
                let _ = fs::copy(&path, profile_dir.join(name));
            }
        }
    }
}
