use std::sync::{atomic::{AtomicI32, Ordering}, Mutex, RwLock};

use windows::{Win32::{Foundation::{HWND, POINT, RECT, E_INVALIDARG, E_NOINTERFACE, E_NOTIMPL, E_UNEXPECTED, S_OK, BOOL}, System::{Com::{IDataObject, FORMATETC}, Ole::CONNECT_E_ADVISELIMIT}, UI::TextServices::{ITextStoreACP, ITextStoreACPSink, ITextStoreACP_Impl, TEXT_STORE_LOCK_FLAGS, TEXT_STORE_TEXT_CHANGE_FLAGS, TS_AS_TEXT_CHANGE, TS_ATTRVAL, TS_E_NOLOCK, TS_E_SYNCHRONOUS, TS_LF_READ, TS_LF_READWRITE, TS_LF_SYNC, TS_RT_PLAIN, TS_SD_LOADING, TS_SD_READONLY, TS_SELECTION_ACP, TS_SS_REGIONS, TS_STATUS, TS_ST_NONE, TS_TEXTCHANGE}}};
use windows_core::{IUnknown, IUnknownImpl, Interface, HRESULT};

fn flag_check(value: u32, flag: u32) -> bool {
    (value & flag) == flag
}

struct AdviceSink {
    text_store_sink: Option<ITextStoreACPSink>,
    mask: u32
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum LockType {
    None,
    Read,
    ReadWrite,
}

impl From<u32> for LockType {
    fn from(flags: u32) -> Self {
        if flag_check(flags, TS_LF_READWRITE.0) {
            LockType::ReadWrite
        } else if flag_check(flags, TS_LF_READ.0) {
            LockType::Read
        } else {
            LockType::None
        }
    }
}

pub struct TfTextStore {
    ref_count: AtomicI32,
    advice_sink: Mutex<AdviceSink>,
    input_text: RwLock<String>,
    lock_state: RwLock<(LockType, u32)>
}

impl TfTextStore {
    pub fn new() -> Self {
        Self {
            ref_count: AtomicI32::new(1),
            advice_sink: Mutex::new(AdviceSink {
                text_store_sink: None,
                mask: 0
            }),
            input_text: RwLock::new(String::new()),
            lock_state: RwLock::new((LockType::None, 0))
        }
    }

    pub fn is_locked(&self, flags: u32) -> bool {
        let lock_state = self.lock_state.read().unwrap();
        lock_state.0 != LockType::None && flag_check(lock_state.1, flags)
    }

    pub fn try_lock(&self, flags: u32) -> Result<LockGuard, ()> {
        let mut lock_state = self.lock_state.write().unwrap();

        if lock_state.0 == LockType::None {
            let lock_type = LockType::from(flags);
            *lock_state = (lock_type, flags);

            Ok(LockGuard { text_store: self })
        } else {
            Err(())
        }
    }

    pub fn set_string(&self, text: &str) -> bool {
        if let Ok(_lock) = self.try_lock(TS_LF_READWRITE.0) {
            let old_len = self.input_text.read().unwrap().len() as i32;

            let mut input_text = self.input_text.write().unwrap();
            *input_text = text.to_string();
            let new_len = input_text.len() as i32;

            let text_change = TS_TEXTCHANGE {
                acpStart: 0,
                acpOldEnd: old_len,
                acpNewEnd: new_len
            };

            drop(input_text);

            let advice_sink = self.advice_sink.lock().unwrap();
            if flag_check(advice_sink.mask, TS_AS_TEXT_CHANGE) {
                if let Some(sink) = &advice_sink.text_store_sink {
                    unsafe {
                        sink.OnTextChange(TS_ST_NONE, &text_change).ok();
                    }
                }
            }

            true
        } else {
            false
        }
    }
}

pub struct LockGuard<'a> {
    text_store: &'a TfTextStore
}

impl <'a> Drop for LockGuard<'a> {
    fn drop(&mut self) {
        let mut lock_state = self.text_store.lock_state.write().unwrap();
        *lock_state = (LockType::None, 0);
    }
}

impl IUnknownImpl for TfTextStore {
    type Impl = Self;

    fn AddRef(&self) -> u32 {
        self.ref_count.fetch_add(1, Ordering::SeqCst) as u32
    }

    unsafe fn Release(&self) -> u32 {
        let count = self.ref_count.fetch_sub(1, Ordering::SeqCst) - 1;
        count as u32
    }

    unsafe fn GetTrustLevel(&self, value: *mut i32) -> windows_core::HRESULT {
        if !value.is_null() {
            unsafe {
                *value = 0;
            }

             S_OK
        } else {
            E_INVALIDARG
        }
    }

    unsafe fn QueryInterface(&self, iid: *const windows_core::GUID, interface: *mut *mut std::ffi::c_void) -> HRESULT {
        let iid = unsafe { &*iid };

        if iid == &<IUnknown as Interface>::IID || iid == &<ITextStoreACP as Interface>::IID {
            unsafe {
                *interface = self as *const _ as *mut std::ffi::c_void;
            }
            self.AddRef();

            S_OK
        } else {
            unsafe {
                *interface = std::ptr::null_mut();
            }

            E_NOINTERFACE
        }
    }

    fn get_impl(&self) -> &Self::Impl {
        self
    }
}

impl ITextStoreACP_Impl for TfTextStore {
    fn AdviseSink(&self, riid: *const windows_core::GUID, punk: Option<&windows_core::IUnknown>, mask: u32) -> windows_core::Result<()> {
        let punk = match punk {
            Some(punk) => punk,
            None => return Err(E_INVALIDARG.into())
        };

        let mut advice_sink = self.advice_sink.lock().unwrap();

        if let Some(existing_sink) = &advice_sink.text_store_sink {
            advice_sink.mask = mask;
            
            Ok(())
        } else if advice_sink.text_store_sink.is_some() {
            Err(CONNECT_E_ADVISELIMIT.into())
        } else {
            let mut sink: Option<ITextStoreACPSink> = None;
            let hr = unsafe { punk.query(&<ITextStoreACPSink as Interface>::IID, &mut sink as *mut _ as *mut _) };

            if hr.is_ok() {
                advice_sink.text_store_sink = sink;
                advice_sink.mask = mask;

                return Ok(());
            }

            Err(hr.into())
        }
    }

    fn UnadviseSink(&self, punk: Option<&windows_core::IUnknown>) -> windows_core::Result<()> {
        let mut advice_sink = self.advice_sink.lock().unwrap();

        if let Some(_existing_sink) = &advice_sink.text_store_sink {
            advice_sink.text_store_sink = None;
            advice_sink.mask = 0;

            Ok(())
        } else {
            Err(E_INVALIDARG.into())
        }
    }

    fn RequestLock(&self, dwlockflags: u32) -> windows_core::Result<windows_core::HRESULT> {
        let advice_sink = self.advice_sink.lock().unwrap();

        if advice_sink.text_store_sink.is_none() {
            return Ok(E_UNEXPECTED);
        }

        let is_currently_locked = {
            let lock_state = self.lock_state.read().unwrap();
            lock_state.0 != LockType::None
        };

        if is_currently_locked {
            if flag_check(dwlockflags, TS_LF_SYNC) {
                return Ok(TS_E_SYNCHRONOUS)
            } else {
                return Ok(E_NOTIMPL)
            }
        } else {
            if let Ok(_guard) = self.try_lock(dwlockflags) {
                if let Some(sink) = &advice_sink.text_store_sink {
                    let hr = unsafe { sink.OnLockGranted(TEXT_STORE_LOCK_FLAGS(dwlockflags)) };

                    return match hr {
                        Ok(_) => Ok(S_OK),
                        Err(e) => Err(e)
                    };
                }
            }

            Ok(S_OK)
        }
    }

    fn GetStatus(&self) -> windows_core::Result<windows::Win32::UI::TextServices::TS_STATUS> {
        let status = TS_STATUS {
            dwDynamicFlags: TS_SD_READONLY | TS_SD_LOADING,
            dwStaticFlags: TS_SS_REGIONS
        };

        Ok(status)
    }

    fn GetText(&self, acpstart: i32, acpend: i32, pchplain: windows_core::PWSTR, cchplainreq: u32, pcchplainret: *mut u32, prgruninfo: *mut windows::Win32::UI::TextServices::TS_RUNINFO, cruninforeq: u32, pcruninforet: *mut u32, pacpnext: *mut i32) -> windows_core::Result<()> {
        if !self.is_locked(TS_LF_READ.0) {
            return Err(TS_E_NOLOCK.into());
        }

        let input_text = self.input_text.read().unwrap();
        let text_len = input_text.len();
        let copy_len = std::cmp::min(text_len as u32, cchplainreq);

        if copy_len > 0 && !pchplain.is_null() {
            let src_slice = input_text.as_bytes();
            let dest_slice = unsafe { std::slice::from_raw_parts_mut(pchplain.0 as *mut u8, copy_len as usize) };
            dest_slice.copy_from_slice(&src_slice[0..copy_len as usize]);
        }

        if !pcchplainret.is_null() {
            unsafe {
                *pcchplainret = copy_len;
            }
        }

        if !prgruninfo.is_null() && cruninforeq > 0 {
            unsafe {
                (*prgruninfo).r#type = TS_RT_PLAIN;
                (*prgruninfo).uCount = text_len as u32;
            }
        }

        if !pcruninforet.is_null() {
            unsafe {
                *pcruninforet = 1;
            }
        }

        if !pacpnext.is_null() {
            unsafe {
                *pacpnext = acpstart + text_len as i32;
            }
        }

        Ok(())
    }

    fn QueryInsert(&self, _acpteststart: i32, _acptestend: i32, _cch: u32, _pacpresultstart: *mut i32, _pacpresultend: *mut i32) -> windows_core::Result<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn GetSelection(&self, _ulindex: u32, _ulcount: u32, _pselection: *mut TS_SELECTION_ACP, _pcfetched: *mut u32) -> windows_core::Result<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn SetSelection(&self, _ulcount: u32, _pselection: *const TS_SELECTION_ACP) -> windows_core::Result<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn SetText(&self, _dwflags: u32, _acpstart: i32, _acpend: i32, _pchtext: &windows_core::PCWSTR, _cch: u32) -> windows_core::Result<TS_TEXTCHANGE> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn GetFormattedText(&self, _acpstart: i32, _acpend: i32) -> windows_core::Result<IDataObject> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn GetEmbedded(&self, _acppos: i32, _rguidservice: *const windows_core::GUID, _riid: *const windows_core::GUID) -> windows_core::Result<windows_core::IUnknown> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn QueryInsertEmbedded(&self, _pguidservice: *const windows_core::GUID, _pformatetc: *const FORMATETC) -> windows_core::Result<BOOL> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn InsertEmbedded(&self, _dwflags: u32, _acpstart: i32, _acpend: i32, _pdataobject: Option<&IDataObject>) -> windows_core::Result<TS_TEXTCHANGE> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn InsertTextAtSelection(&self, _dwflags: u32, _pchtext: &windows_core::PCWSTR, _cch: u32, _pacpstart: *mut i32, _pacpend: *mut i32, _pchange: *mut TS_TEXTCHANGE) -> windows_core::Result<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn InsertEmbeddedAtSelection(&self, _dwflags: u32, _pdataobject: Option<&IDataObject>, _pacpstart: *mut i32, _pacpend: *mut i32, _pchange: *mut TS_TEXTCHANGE) -> windows_core::Result<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn RequestSupportedAttrs(&self, _dwflags: u32, _cfilterattrs: u32, _pafilterattrs: *const windows_core::GUID) -> windows_core::Result<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn RequestAttrsAtPosition(&self, _acppos: i32, _cfilterattrs: u32, _pafilterattrs: *const windows_core::GUID, _dwflags: u32) -> windows_core::Result<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn RequestAttrsTransitioningAtPosition(&self, _acppos: i32, _cfilterattrs: u32, _pafilterattrs: *const windows_core::GUID, _dwflags: u32) -> windows_core::Result<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn FindNextAttrTransition(&self, _acpstart: i32, _acphalt: i32, _cfilterattrs: u32, _pafilterattrs: *const windows_core::GUID, _dwflags: u32, _pacpnext: *mut i32, _pffound: *mut BOOL, _plfoundoffset: *mut i32) -> windows_core::Result<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn RetrieveRequestedAttrs(&self, _ulcount: u32, _paattrvals: *mut TS_ATTRVAL, _pcfetched: *mut u32) -> windows_core::Result<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn GetEndACP(&self) -> windows_core::Result<i32> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn GetActiveView(&self) -> windows_core::Result<u32> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn GetACPFromPoint(&self, _vcview: u32, _ptscreen: *const POINT, _dwflags: u32) -> windows_core::Result<i32> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn GetTextExt(&self, _vcview: u32, _acpstart: i32, _acpend: i32, _prc: *mut RECT, _pfclipped: *mut BOOL) -> windows_core::Result<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn GetScreenExt(&self, _vcview: u32) -> windows_core::Result<RECT> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    
    fn GetWnd(&self, _vcview: u32) -> windows_core::Result<HWND> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
}
