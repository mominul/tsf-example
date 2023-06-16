use windows::{
    core::GUID,
    Win32::{
        Foundation::HMODULE,
        System::SystemServices::{LANG_JAPANESE, SUBLANG_DEFAULT},
    },
};

pub static mut DLL_INSTANCE: HMODULE = HMODULE(0);

pub const CLSID_TEXT_SERVICE: GUID = GUID::from_u128(0xe7ea138e_69f8_11d7_a6ea_00065b84435c);
pub const GUID_PROFILE: GUID = GUID::from_u128(0xe7ea138f_69f8_11d7_a6ea_00065b84435c);
pub const GUID_LANGBAR_ITEM_BUTTON: GUID = GUID::from_u128(0x41f46e67_86d5_49fb_a1d9_3dc0941a66a3);
pub const TEXTSERVICE_DESC: &str = "Sample Text Service";
pub const TEXTSERVICE_LANGID: u16 = (SUBLANG_DEFAULT << 10 | LANG_JAPANESE) as u16;
pub const TEXTSERVICE_ICON_INDEX: u32 = 0;
pub const LANGBAR_ITEM_DESC: &str = "Sample Text Service Button";
