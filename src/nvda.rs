use std::env;
use std::path::{Path, PathBuf};

use libloading::Library;

type NvdaTestIfRunning = unsafe extern "system" fn() -> u32;
type NvdaSpeakText = unsafe extern "system" fn(*const u16) -> u32;
type NvdaCancelSpeech = unsafe extern "system" fn() -> u32;

pub struct NvdaController {
    _library: Option<Library>,
    speak_text: Option<NvdaSpeakText>,
    cancel_speech: Option<NvdaCancelSpeech>,
    available: bool,
}

impl NvdaController {
    pub fn new() -> Self {
        let mut unique_paths = Vec::new();
        for path in Self::candidate_paths() {
            if !unique_paths.contains(&path) {
                unique_paths.push(path);
            }
        }

        for dll_path in unique_paths {
            if let Some(controller) = Self::load_from_path(&dll_path) {
                return controller;
            }
        }

        Self {
            _library: None,
            speak_text: None,
            cancel_speech: None,
            available: false,
        }
    }

    pub fn is_available(&self) -> bool {
        self.available
    }

    pub fn speak(&self, text: &str) {
        if !self.available {
            return;
        }

        let Some(speak_text) = self.speak_text else {
            return;
        };

        let wide_text: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();

        unsafe {
            if let Some(cancel_speech) = self.cancel_speech {
                let _ = cancel_speech();
            }
            let _ = speak_text(wide_text.as_ptr());
        }
    }

    pub fn speak_non_interrupting(&self, text: &str) {
        if !self.available {
            return;
        }

        let Some(speak_text) = self.speak_text else {
            return;
        };

        let wide_text: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();

        unsafe {
            let _ = speak_text(wide_text.as_ptr());
        }
    }

    fn load_from_path(dll_path: &Path) -> Option<Self> {
        if !dll_path.exists() {
            return None;
        }

        let library = unsafe { Library::new(dll_path) }.ok()?;
        let test_if_running: NvdaTestIfRunning =
            unsafe { *library.get(b"nvdaController_testIfRunning\0").ok()? };
        let speak_text: NvdaSpeakText =
            unsafe { *library.get(b"nvdaController_speakText\0").ok()? };
        let cancel_speech: NvdaCancelSpeech =
            unsafe { *library.get(b"nvdaController_cancelSpeech\0").ok()? };

        let available = unsafe { test_if_running() == 0 };

        Some(Self {
            _library: Some(library),
            speak_text: Some(speak_text),
            cancel_speech: Some(cancel_speech),
            available,
        })
    }

    fn candidate_paths() -> Vec<PathBuf> {
        let mut paths = Vec::new();

        if let Ok(exe_path) = env::current_exe() {
            if let Some(exe_dir) = exe_path.parent() {
                paths.push(exe_dir.join("nvdaControllerClient.dll"));
                paths.push(exe_dir.join("x64").join("nvdaControllerClient.dll"));
            }
        }

        if let Ok(current_dir) = env::current_dir() {
            paths.push(current_dir.join("nvdaControllerClient.dll"));
            paths.push(current_dir.join("x64").join("nvdaControllerClient.dll"));
        }

        paths.push(
            PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                .join("x64")
                .join("nvdaControllerClient.dll"),
        );
        paths
    }
}
