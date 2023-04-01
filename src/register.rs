use windows::{Win32::{System::Com::{CLSCTX_INPROC_SERVER, CoCreateInstance}, UI::TextServices::{CLSID_TF_InputProcessorProfiles, ITfInputProcessorProfiles}, Foundation::HMODULE}, core::{GUID, ComInterface}};
use winreg::{RegKey, enums::HKEY_CLASSES_ROOT};

use crate::{globals::{CLSID_TEXT_SERVICE, TEXTSERVICE_DESC, TEXTSERVICE_LANGID, GUID_PROFILE, TEXTSERVICE_ICON_INDEX}, dll::get_module_path};

pub fn create_instance<T: ComInterface>(clsid: &GUID) -> windows::core::Result<T> {
    unsafe { CoCreateInstance(clsid, None, CLSCTX_INPROC_SERVER) }
}

pub fn register_server(handle: HMODULE) -> std::io::Result<()> {
    let filename = get_module_path(handle).map_err(|_| std::io::Error::from(std::io::ErrorKind::InvalidData))?;

    let reg_path = format!("CLSID\\{{{CLSID_TEXT_SERVICE:?}}}");

    let (key, _) = RegKey::predef(HKEY_CLASSES_ROOT).create_subkey(reg_path)?;
    key.set_value("", &TEXTSERVICE_DESC)?;

    let (inproc_key, _) = key.create_subkey("InProcServer32")?;
    inproc_key.set_value("", &filename)?;
    inproc_key.set_value("ThreadingModel", &"Apartment")?;

    Ok(())
}

pub fn unregister_server() -> std::io::Result<()> {
    let reg_path = format!("CLSID\\{{{CLSID_TEXT_SERVICE:?}}}");
    RegKey::predef(HKEY_CLASSES_ROOT).delete_subkey_all(reg_path)
}

pub fn register_profile(handle: HMODULE) -> windows::core::Result<()> {
    let profiles: ITfInputProcessorProfiles = create_instance(&CLSID_TF_InputProcessorProfiles)?;

    unsafe { profiles.Register(&CLSID_TEXT_SERVICE)? };

    let icon_path: Vec<u16> = get_module_path(handle)?.encode_utf16().collect();
    let description: Vec<u16> = TEXTSERVICE_DESC.encode_utf16().collect();

    unsafe {
        profiles.AddLanguageProfile(&CLSID_TEXT_SERVICE, TEXTSERVICE_LANGID, &GUID_PROFILE, &description, &icon_path, TEXTSERVICE_ICON_INDEX)?;
    }

    Ok(())
}

pub fn unregister_profile() -> windows::core::Result<()> {
    let profiles: ITfInputProcessorProfiles = create_instance(&CLSID_TF_InputProcessorProfiles)?;

    unsafe {
        profiles.Unregister(&CLSID_TEXT_SERVICE)?;
    }

    Ok(())
}
