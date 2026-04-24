use std::env;
use std::fs;
use std::path::PathBuf;

fn main() {
    println!("cargo:rerun-if-changed=x64/nvdaControllerClient.dll");

    let manifest_dir = PathBuf::from(env::var("CARGO_MANIFEST_DIR").expect("missing manifest dir"));
    let source_dll = manifest_dir.join("x64").join("nvdaControllerClient.dll");
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
}
