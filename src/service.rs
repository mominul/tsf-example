use std::{cell::RefCell, mem::transmute};

use windows::{
    core::{implement, ComInterface, Result},
    Win32::{
        Foundation::{E_FAIL, S_OK},
        UI::TextServices::{
            ITfContext, ITfDocumentMgr, ITfSource, ITfTextInputProcessor,
            ITfTextInputProcessor_Impl, ITfThreadMgr, ITfThreadMgrEventSink,
            ITfThreadMgrEventSink_Impl, TF_INVALID_COOKIE,
        },
    },
};

#[implement(ITfTextInputProcessor, ITfThreadMgrEventSink)]
pub struct TextService {
    thread_mgr: RefCell<Option<ITfThreadMgr>>,
    event_sink_cookie: RefCell<u32>,
}

impl TextService {
    pub fn new() -> Self {
        TextService {
            thread_mgr: RefCell::new(None),
            event_sink_cookie: RefCell::new(TF_INVALID_COOKIE),
        }
    }
}

impl ITfTextInputProcessor_Impl for TextService {
    fn Activate(&self, ptim: Option<&ITfThreadMgr>, _tid: u32) -> Result<()> {
        log::trace!("TextService::Activate");
        let thread_mgr = ptim.map(|v| v.clone());
        self.thread_mgr.replace(thread_mgr);

        let source: ITfSource = self.thread_mgr.borrow().as_ref().unwrap().cast()?;
        let sink: ITfThreadMgrEventSink = unsafe { self.cast().unwrap() };
        let res = unsafe { source.AdviseSink(&ITfThreadMgrEventSink::IID, &sink) };

        if let Ok(cookie) = res {
            *self.event_sink_cookie.borrow_mut() = cookie;
            log::trace!("TextService::Activate: Cookie set!");
        } else {
            *self.event_sink_cookie.borrow_mut() = TF_INVALID_COOKIE;
            _ = self.Deactivate(); // cleanup any half-finished init
            log::trace!("TextService::Activate: Fail to set Cookie!");
            return E_FAIL.ok();
        }

        S_OK.ok()
    }

    fn Deactivate(&self) -> Result<()> {
        log::trace!("TextService::Deactivate");

        if *self.event_sink_cookie.borrow() == TF_INVALID_COOKIE {
            log::trace!("TextService::Deactivate: Never advised");
            return S_OK.ok(); // never Advised
        }

        if let Ok(source) = self
            .thread_mgr
            .borrow()
            .as_ref()
            .unwrap()
            .cast::<ITfSource>()
        {
            unsafe {
                _ = source.UnadviseSink(*self.event_sink_cookie.borrow());
                log::trace!("TextService::Deactivate: Unadvised!");
            }
        }

        *self.event_sink_cookie.borrow_mut() = TF_INVALID_COOKIE;

        // We release the reference of the ITfThreadMgr
        self.thread_mgr.replace(None);

        S_OK.ok()
    }
}

impl ITfThreadMgrEventSink_Impl for TextService {
    fn OnInitDocumentMgr(&self, _pdim: Option<&ITfDocumentMgr>) -> Result<()> {
        log::trace!("TextService::OnInitDocumentMgr");
        S_OK.ok()
    }

    fn OnUninitDocumentMgr(&self, _pdim: Option<&ITfDocumentMgr>) -> Result<()> {
        log::trace!("TextService::OnUninitDocumentMgr");
        S_OK.ok()
    }

    fn OnSetFocus(
        &self,
        _pdimfocus: Option<&ITfDocumentMgr>,
        _pdimprevfocus: Option<&ITfDocumentMgr>,
    ) -> Result<()> {
        log::trace!("TextService::OnSetFocus");
        S_OK.ok()
    }

    fn OnPushContext(&self, _pic: Option<&ITfContext>) -> Result<()> {
        log::trace!("TextService::OnPushContext");
        S_OK.ok()
    }

    fn OnPopContext(&self, _pic: Option<&ITfContext>) -> Result<()> {
        log::trace!("TextService::OnPopContext");
        S_OK.ok()
    }
}
