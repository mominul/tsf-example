use windows::{
    core::{implement, IUnknown, Interface, Result, GUID},
    Win32::{
        Foundation::{BOOL, CLASS_E_NOAGGREGATION, E_NOINTERFACE, S_OK},
        System::Com::{IClassFactory, IClassFactory_Impl},
        UI::TextServices::ITfTextInputProcessor,
    },
};

use crate::service::TextService;

#[implement(IClassFactory)]
pub struct ClassFactory {
    //
}

impl IClassFactory_Impl for ClassFactory_Impl {
    fn CreateInstance(
        &self,
        punkouter: Option<&IUnknown>,
        riid: *const GUID,
        ppvobject: *mut *mut core::ffi::c_void,
    ) -> Result<()> {
        log::trace!("ClassFactory::CreateInstance");

        if punkouter.is_some() {
            return CLASS_E_NOAGGREGATION.ok();
        }

        unsafe {
            if *riid == ITfTextInputProcessor::IID {
                let unknown: IUnknown = TextService::new().into();
                unknown.query(&*riid, ppvobject as _).ok()
            } else {
                log::trace!("Unknown IID: {:?}", *riid);
                E_NOINTERFACE.ok()
            }
        }
    }

    fn LockServer(&self, _flock: BOOL) -> Result<()> {
        log::trace!("ClassFactory::LockServer");
        S_OK.ok()
    }
}
