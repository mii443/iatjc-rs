use std::{ops::Deref, rc::Rc};

use anyhow::Result;

use windows::Win32::UI::TextServices::{ITfContext, ITfDocumentMgr, ITfFnReconversion, ITfFunctionProvider, GUID_SYSTEM_FUNCTIONPROVIDER};
use windows_core::{IUnknown, Interface};
use tracing::{debug, error, info, instrument, warn, span, Level};

use crate::{text_store::TfTextStore, thread_mgr::ThreadMgr};

pub struct TSF {
    client_id: u32,
    thread_mgr: Option<ThreadMgr>,
    doc_mgr: Option<ITfDocumentMgr>,
    text_store: Option<Rc<TfTextStore>>,
    context: Option<ITfContext>,
    edit_cookie: u32,
    func_prov: Option<ITfFunctionProvider>,
    reconvert: Option<ITfFnReconversion>
}

impl TSF {
    #[instrument(name = "tsf_new", level = "debug", skip_all)]
    pub fn new() -> Self {
        info!("Creating new TSF instance");
        Self {
            client_id: 0,
            thread_mgr: None,
            doc_mgr: None,
            text_store: None,
            context: None,
            edit_cookie: 0,
            func_prov: None,
            reconvert: None
        }
    }

    #[instrument(name = "tsf_initialize", level = "debug", skip_all, err)]
    pub fn initialize(&mut self) -> Result<()> {
        let span = span!(Level::INFO, "initialize_tsf");
        let _enter = span.enter();
        
        info!("Initializing TSF");
        
        debug!("Creating thread manager");
        self.thread_mgr = Some(ThreadMgr::new()?);
        let thread_mgr = self.thread_mgr.as_ref().unwrap();
        debug!("Thread manager created successfully");
        
        debug!("Creating document manager");
        let doc_mgr = thread_mgr.create_document_manager()?;
        self.doc_mgr = Some(doc_mgr);
        debug!("Document manager created successfully");

        debug!("Activating thread manager");
        self.client_id = thread_mgr.activate()?;
        debug!("Thread manager activated with client_id: {}", self.client_id);

        debug!("Creating text store");
        self.text_store = Some(Rc::new(TfTextStore::new()));
        let text_store = self.text_store.as_ref().unwrap();
        debug!("Text store created successfully");

        let doc_mgr = self.doc_mgr.as_ref().unwrap();

        debug!("Creating context with client_id: {}", self.client_id);
        let (context, edit_cookie) = unsafe {
            let mut context = None;
            let mut edit_cookie = 0;
            debug!("Casting text store to IUnknown");
            let text_store = text_store.deref() as *const _ as *mut IUnknown;
            debug!("Creating context");
            let result = doc_mgr.CreateContext(self.client_id, 0, Some(&*text_store), &mut context, &mut edit_cookie);
            if result.is_err() {
                error!("Failed to create context: {:?}", result);
                return Err(anyhow::anyhow!("Failed to create context: {:?}", result));
            }

            (context.unwrap(), edit_cookie)
        };
        debug!("Context created successfully with edit_cookie: {}", edit_cookie);

        self.context = Some(context);
        self.edit_cookie = edit_cookie;

        debug!("Pushing context to document manager");
        unsafe {
            match doc_mgr.Push(self.context.as_ref().unwrap()) {
                Ok(_) => debug!("Context pushed successfully"),
                Err(e) => {
                    error!("Failed to push context: {:?}", e);
                    return Err(e.into());
                }
            }
        }

        debug!("Getting function provider");
        let func_prov = match thread_mgr.get_function_provider(&GUID_SYSTEM_FUNCTIONPROVIDER) {
            Ok(fp) => {
                debug!("Function provider retrieved successfully");
                fp
            },
            Err(e) => {
                error!("Failed to get function provider: {:?}", e);
                return Err(e);
            }
        };
        self.func_prov = Some(func_prov);

        if let Some(func_prov) = &self.func_prov {
            debug!("Getting reconversion function");
            let reconvert: ITfFnReconversion = unsafe {
                match func_prov.GetFunction(&windows_core::GUID::zeroed(), &ITfFnReconversion::IID) {
                    Ok(func) => {
                        match func.cast() {
                            Ok(reconv) => {
                                debug!("Reconversion function retrieved and cast successfully");
                                reconv
                            },
                            Err(e) => {
                                error!("Failed to cast function to ITfFnReconversion: {:?}", e);
                                return Err(e.into());
                            }
                        }
                    },
                    Err(e) => {
                        error!("Failed to get reconversion function: {:?}", e);
                        return Err(e.into());
                    }
                }
            };

            self.reconvert = Some(reconvert);
        }

        debug!("Setting focus to document manager");
        unsafe {
            match thread_mgr.thread_mgr.SetFocus(Some(self.doc_mgr.as_ref().unwrap())) {
                Ok(_) => debug!("Focus set successfully"),
                Err(e) => {
                    error!("Failed to set focus: {:?}", e);
                    return Err(e.into());
                }
            }
        }
        
        info!("TSF initialized successfully");
        Ok(())
    }

    #[instrument(name = "tsf_uninitialize", level = "debug", skip_all)]
    pub fn uninitialize(&mut self) {
        info!("Uninitializing TSF");
        
        if let Some(thread_mgr) = &self.thread_mgr {
            debug!("Deactivating thread manager");
            unsafe {
                match thread_mgr.thread_mgr.Deactivate() {
                    Ok(_) => debug!("Thread manager deactivated successfully"),
                    Err(e) => warn!("Failed to deactivate thread manager: {:?}", e)
                }
            }
        }

        debug!("Clearing resources");
        self.reconvert = None;
        debug!("Reconversion function cleared");
        
        self.func_prov = None;
        debug!("Function provider cleared");
        
        self.edit_cookie = 0;
        debug!("Edit cookie reset");
        
        self.context = None;
        debug!("Context cleared");
        
        self.text_store = None;
        debug!("Text store cleared");
        
        self.doc_mgr = None;
        debug!("Document manager cleared");
        
        self.thread_mgr = None;
        debug!("Thread manager cleared");
        
        info!("TSF uninitialized successfully");
    }
}