use std::ffi::c_void;

use windows::{Win32::Foundation::{MAX_PATH, E_FAIL, S_OK, CLASS_E_CLASSNOTAVAILABLE, E_UNEXPECTED}, Win32::System::{SystemServices::DLL_PROCESS_ATTACH, LibraryLoader::GetModuleFileNameW, Com::IClassFactory}, Win32::{UI::TextServices::ITfTextInputProcessor, Foundation::HMODULE}, core::{HRESULT, GUID, ComInterface, IUnknown}};

use crate::{globals::{DLL_INSTANCE, CLSID_TEXT_SERVICE}, register::{register_profile, register_server, unregister_server, unregister_profile}, factory::ClassFactory};

pub fn get_module_path(instance: HMODULE) -> Result<String, HRESULT> {
    let mut path = [0u16; MAX_PATH as usize];
    let path_len = unsafe { GetModuleFileNameW(instance, &mut path) } as usize;
    String::from_utf16(&path[0..path_len]).map_err(|_| E_FAIL)
}

#[no_mangle]
#[allow(non_snake_case)]
#[doc(hidden)]
pub unsafe extern "system" fn DllRegisterServer() -> HRESULT {
    // let module_path = {
    //     let result = get_module_path(DLL_INSTANCE);
    //     if let Err(err) = result {
    //         return err;
    //     }
    //     result.unwrap()
    // };

    if register_server(DLL_INSTANCE).is_ok() && register_profile(DLL_INSTANCE).is_ok() {
        S_OK
    } else {
        _ = DllUnregisterServer(); // cleanup
        E_FAIL
    }
}

#[no_mangle]
#[allow(non_snake_case)]
#[doc(hidden)]
pub unsafe extern "system" fn DllUnregisterServer() -> HRESULT {
    if unregister_server().is_ok() && unregister_profile().is_ok() {
        S_OK
    } else {
        E_FAIL
    }
}

#[no_mangle]
#[allow(non_snake_case)]
#[doc(hidden)]
pub extern "stdcall" fn DllMain(
    dll_instance: HMODULE,
    reason: u32,
    _reserved: *mut c_void,
) -> bool {
    if reason == DLL_PROCESS_ATTACH {
        unsafe {
            DLL_INSTANCE = dll_instance;
        }
    }
    true
}

#[no_mangle]
#[allow(non_snake_case)]
#[doc(hidden)]
pub unsafe extern "system" fn DllGetClassObject(
    rclsid: *const GUID,
    riid: *const GUID,
    pout: *mut *const core::ffi::c_void,
) -> HRESULT {
    // Sets up logging to the Cargo.toml directory for debug purposes.
    #[cfg(debug_assertions)]
    {
        // Set up logging to the project directory.
        simple_logging::log_to_file(
            &format!("{}\\debug.log", env!("CARGO_MANIFEST_DIR")),
            log::LevelFilter::Trace,
        )
        .unwrap();
    }
    log::trace!("DllGetClassObject");
    log::trace!("riid {:?}", *riid);
    log::trace!("rclsid {:?}", *rclsid);
    log::trace!("CLassFactory {:?}", IClassFactory::IID);
    log::trace!("IUnknown {:?}", IUnknown::IID);
    log::trace!("TextService {:?}", CLSID_TEXT_SERVICE);
    log::trace!("ITfTextInputProcessor {:?}", ITfTextInputProcessor::IID);
    
    if *riid != IClassFactory::IID || *riid != IUnknown::IID {
        return E_UNEXPECTED;
    }

    let factory = ClassFactory {};
    let unknown: IUnknown = factory.into();

    match *rclsid {
        CLSID_TEXT_SERVICE => unknown.query(&*riid, pout),
        _ => CLASS_E_CLASSNOTAVAILABLE,
    }
}