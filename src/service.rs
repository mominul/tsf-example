use std::{cell::RefCell, ptr::null_mut};

use windows::{
    core::{implement, ComInterface, Result},
    Win32::{
        Foundation::{E_FAIL, S_OK},
        UI::TextServices::{
            ITfContext, ITfDocumentMgr, ITfEditRecord, ITfLangBarItem, ITfLangBarItemMgr,
            ITfSource, ITfTextEditSink, ITfTextEditSink_Impl, ITfTextInputProcessor,
            ITfTextInputProcessor_Impl, ITfThreadMgr, ITfThreadMgrEventSink,
            ITfThreadMgrEventSink_Impl, TF_GTP_INCL_TEXT, TF_INVALID_COOKIE,
        },
    },
};

use crate::languagebar::LangBarItemButton;

#[implement(ITfTextInputProcessor, ITfThreadMgrEventSink, ITfTextEditSink)]
pub struct TextService {
    thread_mgr: RefCell<Option<ITfThreadMgr>>,
    event_sink_cookie: RefCell<u32>,
    edit_sink_context: RefCell<Option<ITfContext>>,
    edit_sink_cookie: RefCell<u32>,
    langbar_item: RefCell<Option<ITfLangBarItem>>,
}

impl TextService {
    pub fn new() -> Self {
        TextService {
            thread_mgr: RefCell::new(None),
            event_sink_cookie: RefCell::new(TF_INVALID_COOKIE),
            edit_sink_context: RefCell::new(None),
            edit_sink_cookie: RefCell::new(TF_INVALID_COOKIE),
            langbar_item: RefCell::new(None),
        }
    }

    fn init_text_edit_sink(&self, doc_mgr: &ITfDocumentMgr) {
        log::trace!("TextService::init_text_edit_sink");
        // clear out any previous sink first
        self.uninit_text_edit_sink();

        // setup a new sink advised to the topmost context of the document
        let context = unsafe { doc_mgr.GetTop() };
        let Ok(context) = context else {
            return;
        };

        if let Ok(source) = context.cast::<ITfSource>() {
            let sink: ITfTextEditSink = unsafe { self.cast().unwrap() };
            if let Ok(cookie) = unsafe { source.AdviseSink(&ITfTextEditSink::IID, &sink) } {
                self.edit_sink_cookie.replace(cookie);
                self.edit_sink_context.replace(Some(context));
            }
        }
    }

    fn uninit_text_edit_sink(&self) {
        log::trace!("TextService::uninit_text_edit_sink");
        if *self.edit_sink_cookie.borrow() != TF_INVALID_COOKIE {
            if let Ok(source) = self
                .edit_sink_context
                .borrow()
                .as_ref()
                .unwrap()
                .cast::<ITfSource>()
            {
                unsafe {
                    _ = source.UnadviseSink(*self.edit_sink_cookie.borrow());
                }
            }

            self.edit_sink_context.replace(None);
            self.edit_sink_cookie.replace(TF_INVALID_COOKIE);
        }
    }

    fn init_language_bar(&self) {
        log::trace!("TextService::init_language_bar");
        let Ok(mgr) = self.thread_mgr.borrow().as_ref().unwrap().cast::<ITfLangBarItemMgr>() else {
            return;
        };

        let item = LangBarItemButton::new();
        let item: ITfLangBarItem = item.into();

        unsafe {
            if mgr.AddItem(&item).is_ok() {
                self.langbar_item.replace(Some(item));
            }
        }
    }

    fn uninit_lang_bar(&self) {
        log::trace!("TextService::uninit_lang_bar");
        let Some(item) = self.langbar_item.replace(None) else {
            return;
        };

        if let Ok(mgr) = self
            .thread_mgr
            .borrow()
            .as_ref()
            .unwrap()
            .cast::<ITfLangBarItemMgr>()
        {
            unsafe {
                _ = mgr.RemoveItem(&item);
            }
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
            self.event_sink_cookie.replace(cookie);
            log::trace!("TextService::Activate: Cookie set!");
        } else {
            self.event_sink_cookie.replace(TF_INVALID_COOKIE);
            _ = self.Deactivate(); // cleanup any half-finished init
            log::trace!("TextService::Activate: Fail to set Cookie!");
            return E_FAIL.ok();
        }

        //  If there is the focus document manager already,
        //  we advise the TextEditSink.
        let doc_mgr = unsafe { self.thread_mgr.borrow().as_ref().unwrap().GetFocus() };
        if let Ok(doc_mgr) = doc_mgr {
            self.init_text_edit_sink(&doc_mgr);
        }

        // Initialize Language Bar.
        self.init_language_bar();

        S_OK.ok()
    }

    fn Deactivate(&self) -> Result<()> {
        log::trace!("TextService::Deactivate");

        // Unadvise TextEditSink if it is advised.
        self.uninit_text_edit_sink();

        // Uninitialize ThreadMgrEventSink.
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

        self.event_sink_cookie.replace(TF_INVALID_COOKIE);

        // Uninitialize Language Bar.
        self.uninit_lang_bar();

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
        pdimfocus: Option<&ITfDocumentMgr>,
        _pdimprevfocus: Option<&ITfDocumentMgr>,
    ) -> Result<()> {
        log::trace!("TextService::OnSetFocus");
        // Whenever focus is changed, we initialize the TextEditSink.
        if let Some(doc_mgr) = pdimfocus {
            self.init_text_edit_sink(doc_mgr);
        } else {
            self.uninit_text_edit_sink();
        }

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

impl ITfTextEditSink_Impl for TextService {
    fn OnEndEdit(
        &self,
        _pic: Option<&ITfContext>,
        _ecreadonly: u32,
        peditrecord: Option<&ITfEditRecord>,
    ) -> Result<()> {
        log::trace!("TextService::OnEndEdit");
        let record = peditrecord.unwrap();

        // did the selection change?
        // The selection change includes the movement of caret as well.
        // The caret position is represent as the empty selection range when
        // there is no selection.
        if let Ok(selection_changed) = unsafe { record.GetSelectionStatus() } {
            if selection_changed.into() {
                log::trace!("TextService::OnEndEdit: Selection changed");
            }
        }

        // text modification?
        if let Ok(text_changes) = unsafe { record.GetTextAndPropertyUpdates(TF_GTP_INCL_TEXT, &[]) }
        {
            let mut ranges = vec![None];
            if unsafe { text_changes.Next(&mut ranges, null_mut()).is_ok() } {
                log::trace!("TextService::OnEndEdit: Updated range found");
            }
        }

        S_OK.ok()
    }
}
