use std::{ffi::c_void, path::PathBuf, time::SystemTime};

use tracing::{Level, trace};
use tracing_appender::rolling::{RollingFileAppender, Rotation};
use tracing_subscriber::FmtSubscriber;
use windows::{
    core::{ComInterface, IUnknown, GUID, HRESULT},
    Win32::{
        Foundation::{
            CLASS_E_CLASSNOTAVAILABLE, E_FAIL, E_UNEXPECTED, HMODULE, MAX_PATH, S_FALSE, S_OK,
        },
        System::{
            Com::IClassFactory, LibraryLoader::GetModuleFileNameW,
            SystemServices::DLL_PROCESS_ATTACH,
        },
        UI::TextServices::ITfTextInputProcessor,
    },
};

use crate::{
    factory::ClassFactory,
    globals::{CLSID_TEXT_SERVICE, DLL_INSTANCE},
    register::{register_profile, register_server, unregister_profile, unregister_server},
};

pub fn get_module_path(instance: HMODULE) -> Result<String, HRESULT> {
    let mut path = [0u16; MAX_PATH as usize];
    let path_len = unsafe { GetModuleFileNameW(instance, &mut path) } as usize;
    String::from_utf16(&path[0..path_len]).map_err(|_| E_FAIL)
}

#[no_mangle]
#[allow(non_snake_case)]
#[doc(hidden)]
pub unsafe extern "system" fn DllRegisterServer() -> HRESULT {
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
        // Sets up logging to the dll's directory for debug purposes.
        #[cfg(debug_assertions)]
        {
            // Add some value to the name of the log file to prevent overwriting
            let time = SystemTime::now();
            let time = time
                .duration_since(SystemTime::UNIX_EPOCH)
                .unwrap()
                .as_secs();
            let mut path: PathBuf = get_module_path(dll_instance).unwrap().into();
            path.pop();

            let name = format!("debug-{time}.log");

            // Set up logging to the directory where the dll is present
            let file_appender = RollingFileAppender::new(Rotation::NEVER, path, name);
            let subscriber = FmtSubscriber::builder()
                .with_max_level(Level::TRACE)
                .with_ansi(false)
                .with_writer(file_appender)
                .finish();
            tracing::subscriber::set_global_default(subscriber).unwrap();
        }

        unsafe {
            DLL_INSTANCE = dll_instance;
        }
    }
    true
}

#[no_mangle]
#[allow(non_snake_case)]
#[doc(hidden)]
pub unsafe extern "stdcall" fn DllGetClassObject(
    rclsid: *const GUID,
    riid: *const GUID,
    pout: *mut *mut core::ffi::c_void,
) -> HRESULT {
    trace!("DllGetClassObject");
    trace!("riid {:?}", *riid);
    trace!("rclsid {:?}", *rclsid);
    trace!("IClassFactory {:?}", IClassFactory::IID);
    trace!("IUnknown {:?}", IUnknown::IID);
    trace!("TextService {:?}", CLSID_TEXT_SERVICE);
    trace!("ITfTextInputProcessor {:?}", ITfTextInputProcessor::IID);

    // Interface out pointer need to be set as null if error occurs.
    std::ptr::write(pout, std::ptr::null_mut());

    if *riid != IClassFactory::IID {
        trace!("E_UNEXPECTED");
        return E_UNEXPECTED;
    }

    if *rclsid != CLSID_TEXT_SERVICE {
        trace!("CLASS_E_CLASSNOTAVAILABLE");
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    let factory = ClassFactory {};
    let factory: IClassFactory = factory.into();

    std::ptr::write(pout, std::mem::transmute(factory));

    trace!("Done DllGetClassObject");

    S_OK
}

#[no_mangle]
pub extern "stdcall" fn DllCanUnloadNow() -> HRESULT {
    trace!("DllCanUnloadNow");

    S_FALSE
}
