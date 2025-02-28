use anyhow::Result;
use tracing::debug;
use windows::Win32::System::Com::{CoInitialize, CoUninitialize};

pub struct Com;

impl Drop for Com {
    fn drop(&mut self) {
        debug!("Dropping Com instance");
        unsafe {
            CoUninitialize();
            debug!("CoUninitialize called");
        };
    }
}

impl Com {
    pub fn new() -> Result<Self> {
        unsafe {
            let _ = CoInitialize(None);
        };
        Ok(Com)
    }
}
