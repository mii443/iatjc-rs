use std::sync::mpsc::Sender;

use windows::Win32::UI::TextServices::{ITfEditSession_Impl, ITfEditSession};
use windows_core::{implement, HRESULT};

#[implement(ITfEditSession)]
pub struct EditSession {
    sender: Sender<u32>
}

impl EditSession {
    pub fn new(sender: Sender<u32>) -> EditSession {
        EditSession {
            sender
        }
    }
}

impl ITfEditSession_Impl for EditSession {
    fn DoEditSession(&self, ec: u32) -> windows_core::Result<()> {
        if let Err(_) = self.sender.send(ec) {
            return Err(windows_core::Error::new(HRESULT::from_win32(0), "Failed to send message"));
        }

        Ok(())
    }
}
