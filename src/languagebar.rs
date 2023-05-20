use std::{
    cell::RefCell,
    iter::{once, repeat},
    ptr::{null_mut, write},
};

use windows::{
    core::{implement, ComInterface, Error, IUnknown, Result, BSTR, GUID},
    w,
    Win32::{
        Foundation::{BOOL, E_NOINTERFACE, E_NOTIMPL, POINT, RECT, S_OK},
        System::Ole::{CONNECT_E_ADVISELIMIT, CONNECT_E_CANNOTCONNECT},
        UI::{
            TextServices::{
                ITfLangBarItem, ITfLangBarItemButton, ITfLangBarItemButton_Impl,
                ITfLangBarItemSink, ITfLangBarItem_Impl, ITfMenu, ITfSource, ITfSource_Impl,
                TfLBIClick, TF_LANGBARITEMINFO, TF_LBI_STYLE_BTN_MENU,
            },
            WindowsAndMessaging::{LoadImageW, HICON, IMAGE_FLAGS, IMAGE_ICON},
        },
    },
};

use crate::globals::{
    CLSID_TEXT_SERVICE, DLL_INSTANCE, GUID_LANGBAR_ITEM_BUTTON, LANGBAR_ITEM_DESC,
};

// The cookie for the sink to CLangBarItemButton.
pub const TEXTSERVICE_LANGBARITEMSINK_COOKIE: u32 = 0x0fab0fab;

// The ids of the menu item of the language bar button.
const MENUITEM_INDEX_0: u32 = 0;
const MENUITEM_INDEX_1: u32 = 1;

// The descriptions of the menu item of the language bar button.
const MENU_ITEM_DESCRIPTION0: &str = "Menu Item Description 0";
const MENU_ITEM_DESCRIPTION1: &str = "Menu Item Description 1";

#[implement(ITfLangBarItem, ITfLangBarItemButton, ITfSource)]
pub struct LangBarItemButton {
    sink: RefCell<Option<ITfLangBarItemSink>>,
    info: TF_LANGBARITEMINFO,
}

impl LangBarItemButton {
    pub fn new() -> Self {
        let desc: Vec<u16> = LANGBAR_ITEM_DESC
            .encode_utf16()
            .chain(repeat(0))
            .take(32)
            .collect();

        let info = TF_LANGBARITEMINFO {
            clsidService: CLSID_TEXT_SERVICE, // This LangBarItem belongs to this TextService.
            guidItem: GUID_LANGBAR_ITEM_BUTTON, // GUID of this LangBarItem.
            dwStyle: TF_LBI_STYLE_BTN_MENU,   // This LangBar is a button type with a menu
            ulSort: 0,                        // The position of this LangBar Item is not specified.
            szDescription: desc.try_into().unwrap(), // Set the description of this LangBar Item.
        };

        Self {
            sink: RefCell::new(None),
            info,
        }
    }
}

impl ITfLangBarItem_Impl for LangBarItemButton {
    fn GetInfo(&self, pinfo: *mut TF_LANGBARITEMINFO) -> Result<()> {
        unsafe {
            write(pinfo, self.info);
        }

        S_OK.ok()
    }

    fn GetStatus(&self) -> Result<u32> {
        Ok(0)
    }

    fn Show(&self, _fshow: BOOL) -> Result<()> {
        E_NOTIMPL.ok()
    }

    fn GetTooltipString(&self) -> Result<BSTR> {
        let string: Vec<u16> = LANGBAR_ITEM_DESC.encode_utf16().chain(once(0)).collect();

        BSTR::from_wide(&string)
    }
}

impl ITfLangBarItemButton_Impl for LangBarItemButton {
    fn OnClick(&self, _click: TfLBIClick, _pt: &POINT, _prcarea: *const RECT) -> Result<()> {
        S_OK.ok()
    }

    fn InitMenu(&self, pmenu: Option<&ITfMenu>) -> Result<()> {
        let menu = pmenu.unwrap();
        let desc0: Vec<u16> = MENU_ITEM_DESCRIPTION0
            .encode_utf16()
            .chain(once(0))
            .collect();
        let desc1: Vec<u16> = MENU_ITEM_DESCRIPTION1
            .encode_utf16()
            .chain(once(0))
            .collect();

        // Add the menu items.
        unsafe {
            _ = menu.AddMenuItem(MENUITEM_INDEX_0, 0, None, None, &desc0, null_mut());
            _ = menu.AddMenuItem(MENUITEM_INDEX_1, 0, None, None, &desc1, null_mut());
        }

        S_OK.ok()
    }

    fn OnMenuSelect(&self, wid: u32) -> Result<()> {
        // This is callback when the menu item is selected.

        match wid {
            MENUITEM_INDEX_0 => (),
            MENUITEM_INDEX_1 => (),
            _ => (),
        }

        S_OK.ok()
    }

    fn GetIcon(&self) -> Result<HICON> {
        let icon = unsafe {
            LoadImageW(
                DLL_INSTANCE,
                w!("IDI_TEXTSERVICE"),
                IMAGE_ICON,
                16,
                16,
                IMAGE_FLAGS(0),
            )
        };

        icon.map(|i| HICON(i.0))
    }

    fn GetText(&self) -> Result<BSTR> {
        let string: Vec<u16> = LANGBAR_ITEM_DESC.encode_utf16().chain(once(0)).collect();

        BSTR::from_wide(&string)
    }
}

impl ITfSource_Impl for LangBarItemButton {
    fn AdviseSink(&self, riid: *const GUID, punk: Option<&IUnknown>) -> Result<u32> {
        // We allow only ITfLangBarItemSink interface.
        let iid = unsafe { *riid };

        if iid != ITfLangBarItemSink::IID {
            return Err(Error::from(CONNECT_E_CANNOTCONNECT));
        }

        // We support only one sink once.
        if self.sink.borrow().is_some() {
            return Err(Error::from(CONNECT_E_ADVISELIMIT));
        }

        // Query the ITfLangBarItemSink interface and store it into sink.
        let Ok(sink) = punk.unwrap().cast() else {
            self.sink.replace(None);
            return Err(Error::from(E_NOINTERFACE));
        };

        self.sink.replace(Some(sink));

        // return our cookie.
        return Ok(TEXTSERVICE_LANGBARITEMSINK_COOKIE);
    }

    fn UnadviseSink(&self, _dwcookie: u32) -> Result<()> {
        todo!()
    }
}