use std::iter::once;

use windows::{
    core::{Interface, Result, GUID, VARIANT},
    Win32::{
        Foundation::{BOOL, E_FAIL, LPARAM, S_OK, WPARAM},
        UI::{
            Input::KeyboardAndMouse::{VK_F6, VK_KANJI},
            TextServices::{
                ITfCompartmentMgr, ITfContext, ITfKeyEventSink, ITfKeyEventSink_Impl,
                ITfKeystrokeMgr, GUID_COMPARTMENT_EMPTYCONTEXT, GUID_COMPARTMENT_KEYBOARD_DISABLED,
                GUID_COMPARTMENT_KEYBOARD_OPENCLOSE, TF_MOD_ALT, TF_MOD_IGNORE_ALL_MODIFIER,
                TF_MOD_ON_KEYUP, TF_PRESERVEDKEY,
            },
        },
    },
};

use crate::service::{TextService, TextService_Impl};

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
        log::trace!("TextService::init_key_event_sink");
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
        log::trace!("TextService::uninit_key_event_sink");
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

    // Register a hot key.
    pub fn init_preserved_key(&self) {
        log::trace!("TextService::init_preserved_key");
        let Ok(mgr) = self
            .thread_mgr
            .borrow()
            .as_ref()
            .unwrap()
            .cast::<ITfKeystrokeMgr>()
        else {
            return;
        };

        let desc_onoff: Vec<u16> = KEY_ON_OFF_DESC.encode_utf16().chain(once(0)).collect();
        let desc_f6: Vec<u16> = KEY_F6_DESC.encode_utf16().chain(once(0)).collect();

        unsafe {
            // register Alt+~ key
            _ = mgr.PreserveKey(
                *self.client_id.borrow(),
                &GUID_PRESERVEDKEY_ONOFF,
                &KEY_ON_OFF0,
                &desc_onoff,
            );
            // register KANJI key
            _ = mgr.PreserveKey(
                *self.client_id.borrow(),
                &GUID_PRESERVEDKEY_ONOFF,
                &KEY_ON_OFF1,
                &desc_onoff,
            );
            // register F6 key
            _ = mgr.PreserveKey(
                *self.client_id.borrow(),
                &GUID_PRESERVEDKEY_F6,
                &KEY_F6,
                &desc_f6,
            );
        }
    }

    // Uninit a hot key.
    pub fn uninit_preserved_key(&self) {
        log::trace!("TextService::uninit_preserved_key");
        let Ok(mgr) = self
            .thread_mgr
            .borrow()
            .as_ref()
            .unwrap()
            .cast::<ITfKeystrokeMgr>()
        else {
            return;
        };

        unsafe {
            _ = mgr.UnpreserveKey(&GUID_PRESERVEDKEY_ONOFF, &KEY_ON_OFF0);
            _ = mgr.UnpreserveKey(&GUID_PRESERVEDKEY_ONOFF, &KEY_ON_OFF1);
            _ = mgr.UnpreserveKey(&GUID_PRESERVEDKEY_F6, &KEY_F6);
        }
    }

    /// GUID_COMPARTMENT_KEYBOARD_DISABLED is the compartment in the context
    /// object.
    pub fn is_keyboard_disabled(&self) -> bool {
        log::trace!("TextService::is_keyboard_disabled");
        unsafe {
            let Ok(doc_focus) = self.thread_mgr.borrow().as_ref().unwrap().GetFocus() else {
                // if there is no focus document manager object, the keyboard
                // is disabled.
                return true;
            };

            let Ok(context) = doc_focus.GetTop() else {
                // if there is no context object, the keyboard is disabled.
                return true;
            };

            // Check GUID_COMPARTMENT_KEYBOARD_DISABLED.
            if let Ok(comp_mgr) = context.cast::<ITfCompartmentMgr>() {
                if let Ok(disabled) = comp_mgr.GetCompartment(&GUID_COMPARTMENT_KEYBOARD_DISABLED) {
                    if let Ok(var) = disabled.GetValue() {
                        if let Ok(val) = i32::try_from(&var) {
                            log::trace!("Got value CompartmentDisabled: {val}");
                            return val != 0;
                        }
                    }
                }
            }

            // Check GUID_COMPARTMENT_EMPTYCONTEXT.
            if let Ok(comp_mgr) = context.cast::<ITfCompartmentMgr>() {
                if let Ok(context) = comp_mgr.GetCompartment(&GUID_COMPARTMENT_EMPTYCONTEXT) {
                    if let Ok(var) = context.GetValue() {
                        if let Ok(val) = i32::try_from(&var) {
                            log::trace!("Got value CompartmentEmptyContext: {val}");
                            return val != 0;
                        }
                    }
                }
            }
        }
        false
    }

    // GUID_COMPARTMENT_KEYBOARD_OPENCLOSE is the compartment in the thread manager
    // object.
    pub fn is_keyboard_open(&self) -> bool {
        log::trace!("TextService::is_keyboard_open");
        unsafe {
            if let Ok(comp_mgr) = self
                .thread_mgr
                .borrow()
                .as_ref()
                .unwrap()
                .cast::<ITfCompartmentMgr>()
            {
                if let Ok(compartment) =
                    comp_mgr.GetCompartment(&GUID_COMPARTMENT_KEYBOARD_OPENCLOSE)
                {
                    if let Ok(var) = compartment.GetValue() {
                        if let Ok(val) = i32::try_from(&var) {
                            log::trace!("Got value Compartment_KEYBOARD_OPENCLOSE: {val}");
                            return val != 0;
                        }
                    }
                }
            }
        }
        false
    }

    // GUID_COMPARTMENT_KEYBOARD_OPENCLOSE is the compartment in the thread manager
    // object.
    pub fn set_keyboard_open(&self, open: bool) -> Result<()> {
        log::trace!("TextService::set_keyboard_open");
        unsafe {
            if let Ok(comp_mgr) = self
                .thread_mgr
                .borrow()
                .as_ref()
                .unwrap()
                .cast::<ITfCompartmentMgr>()
            {
                if let Ok(compartment) =
                    comp_mgr.GetCompartment(&GUID_COMPARTMENT_KEYBOARD_OPENCLOSE)
                {
                    let open: i32 = open.into();
                    let var = VARIANT::from(open);
                    compartment.SetValue(*self.client_id.borrow(), &var)?;
                    return S_OK.ok();
                }
            }
        }

        E_FAIL.ok()
    }

    fn is_key_eaten(&self, param: WPARAM) -> bool {
        log::trace!("TextService::is_key_eaten");
        // if the keyboard is disabled, we don't eat keys.
        if self.is_keyboard_disabled() {
            return false;
        }

        if !self.is_keyboard_open() {
            return false;
        }

        // we're only interested in VK_A - VK_Z, when this is open.
        // is on
        (param.0 >= b'A'.into()) && (param.0 <= b'Z'.into())
    }
}

impl ITfKeyEventSink_Impl for TextService_Impl {
    // Called by the system whenever this service gets the keystroke device focus.
    fn OnSetFocus(&self, _fforeground: BOOL) -> Result<()> {
        log::trace!("TextService::OnSetFocus");
        S_OK.ok()
    }

    // Called by the system to query this service wants a potential keystroke.
    fn OnTestKeyDown(
        &self,
        _pic: Option<&ITfContext>,
        wparam: WPARAM,
        _lparam: LPARAM,
    ) -> Result<BOOL> {
        log::trace!("TextService::OnTestKeyDown");
        Ok(self.is_key_eaten(wparam).into())
    }

    // Called by the system to query this service wants a potential keystroke.
    fn OnTestKeyUp(
        &self,
        _pic: Option<&ITfContext>,
        wparam: WPARAM,
        _lparam: LPARAM,
    ) -> Result<BOOL> {
        log::trace!("TextService::OnTestKeyUp");
        Ok(self.is_key_eaten(wparam).into())
    }

    // Called by the system to offer this service a keystroke.  If TRUE is returned,
    // the application will not handle the keystroke.
    fn OnKeyDown(
        &self,
        _pic: Option<&ITfContext>,
        wparam: WPARAM,
        _lparam: LPARAM,
    ) -> Result<BOOL> {
        log::trace!("TextService::OnKeyDown");
        Ok(self.is_key_eaten(wparam).into())
    }

    // Called by the system to offer this service a keystroke.  If TRUE is returned,
    // the application will not handle the keystroke.
    fn OnKeyUp(&self, _pic: Option<&ITfContext>, wparam: WPARAM, _lparam: LPARAM) -> Result<BOOL> {
        log::trace!("TextService::OnKeyUp");
        Ok(self.is_key_eaten(wparam).into())
    }

    // Called when a hotkey (registered by us, or by the system) is typed.
    fn OnPreservedKey(&self, _pic: Option<&ITfContext>, rguid: *const GUID) -> Result<BOOL> {
        log::trace!("TextService::OnPreservedKey");
        if unsafe { *rguid } == GUID_PRESERVEDKEY_ONOFF {
            let open = self.is_keyboard_open();
            _ = self.set_keyboard_open(!open);
            Ok(true.into())
        } else {
            Ok(false.into())
        }
    }
}
