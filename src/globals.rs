use windows::{
    core::GUID,
    Win32::{
        Foundation::HMODULE,
        System::SystemServices::{LANG_ENGLISH, SUBLANG_ENGLISH_US},
    },
};

pub static mut DLL_INSTANCE: HMODULE = HMODULE(0);

pub const CLSID_TEXT_SERVICE: GUID = GUID::from_u128(0xe7ea138e_69f8_11d7_a6ea_00065b84435c);
pub const GUID_PROFILE: GUID = GUID::from_u128(0xe7ea138f_69f8_11d7_a6ea_00065b84435c);
pub const TEXTSERVICE_DESC: &str = "Sample Text Service";
pub const TEXTSERVICE_LANGID: u16 = (SUBLANG_ENGLISH_US << 10 | LANG_ENGLISH) as u16;
pub const TEXTSERVICE_ICON_INDEX: u32 = 0;
