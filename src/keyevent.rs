use windows::{
    core::{Result, GUID, ComInterface},
    Win32::{
        Foundation::{BOOL, LPARAM, S_OK, WPARAM},
        UI::{
            Input::KeyboardAndMouse::{VK_F6, VK_KANJI},
            TextServices::{
                ITfContext, ITfKeyEventSink_Impl, TF_MOD_ALT, TF_MOD_IGNORE_ALL_MODIFIER,
                TF_MOD_ON_KEYUP, TF_PRESERVEDKEY, ITfKeyEventSink, ITfKeystrokeMgr,
            },
        },
    },
};

use crate::service::TextService;

const GUID_PRESERVEDKEY_ONOFF: GUID = GUID::from_u128(0x6a0bde41_6adf_11d7_a6ea_00065b84435c);
const GUID_PRESERVEDKEY_F6: GUID = GUID::from_u128(0x6a0bde42_6adf_11d7_a6ea_00065b84435c);

// the preserved keys declaration
//
// VK_KANJI is the virtual key for Kanji key, which is available in 106
// Japanese keyboard.
const KEY_ON_OFF0: TF_PRESERVEDKEY = TF_PRESERVEDKEY {
    uVKey: 0xC0,
    uModifiers: TF_MOD_ALT,
};
const KEY_ON_OFF1: TF_PRESERVEDKEY = TF_PRESERVEDKEY {
    uVKey: VK_KANJI.0 as _,
    uModifiers: TF_MOD_IGNORE_ALL_MODIFIER,
};
const KEY_F6: TF_PRESERVEDKEY = TF_PRESERVEDKEY {
    uVKey: VK_F6.0 as _,
    uModifiers: TF_MOD_ON_KEYUP,
};

// the description for the preserved keys
const KEY_ON_OFF_DESC: &str = "OnOff";
const KEY_F6_DESC: &str = "Function 6";

impl TextService {
    pub fn init_key_event_sink(&self) -> Result<()> {
        if let Ok(mgr) = self
            .thread_mgr
            .borrow()
            .as_ref()
            .unwrap()
            .cast::<ITfKeystrokeMgr>()
        {
            unsafe {
                let sink: ITfKeyEventSink = self.cast_to().unwrap();
                mgr.AdviseKeyEventSink(*self.client_id.borrow(), &sink, true)?;
            }
        }

        Ok(())
    }

    pub fn uninit_key_event_sink(&self) {
        if let Ok(mgr) = self
            .thread_mgr
            .borrow()
            .as_ref()
            .unwrap()
            .cast::<ITfKeystrokeMgr>()
        {
            unsafe {
                _ = mgr.UnadviseKeyEventSink(*self.client_id.borrow());
            }
        }
    }
}

impl ITfKeyEventSink_Impl for TextService {
    // Called by the system whenever this service gets the keystroke device focus.
    fn OnSetFocus(&self, _fforeground: BOOL) -> Result<()> {
        S_OK.ok()
    }

    // Called by the system to query this service wants a potential keystroke.
    fn OnTestKeyDown(
        &self,
        _pic: Option<&ITfContext>,
        _wparam: WPARAM,
        _lparam: LPARAM,
    ) -> Result<BOOL> {
        todo!()
    }

    // Called by the system to query this service wants a potential keystroke.
    fn OnTestKeyUp(
        &self,
        _pic: Option<&ITfContext>,
        _wparam: WPARAM,
        _lparam: LPARAM,
    ) -> Result<BOOL> {
        todo!()
    }

    // Called by the system to offer this service a keystroke.  If TRUE is returned,
    // the application will not handle the keystroke.
    fn OnKeyDown(&self, _pic: Option<&ITfContext>, _wparam: WPARAM, _lparam: LPARAM) -> Result<BOOL> {
        todo!()
    }

    // Called by the system to offer this service a keystroke.  If TRUE is returned,
    // the application will not handle the keystroke.
    fn OnKeyUp(&self, _pic: Option<&ITfContext>, _wparam: WPARAM, _lparam: LPARAM) -> Result<BOOL> {
        todo!()
    }

    // Called when a hotkey (registered by us, or by the system) is typed.
    fn OnPreservedKey(&self, _pic: Option<&ITfContext>, _rguid: *const GUID) -> Result<BOOL> {
        todo!()
    }
}
