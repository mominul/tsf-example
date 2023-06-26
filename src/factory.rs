use tracing::trace;
use windows::{
    core::{implement, ComInterface, IUnknown, Result, GUID},
    Win32::{
        Foundation::{BOOL, CLASS_E_NOAGGREGATION, E_NOINTERFACE, S_OK},
        System::Com::{IClassFactory, IClassFactory_Impl},
        UI::TextServices::ITfTextInputProcessor,
    },
};

use crate::service::TextService;

#[implement(IClassFactory)]
#[derive(Debug)]
pub struct ClassFactory {
    //
}

impl IClassFactory_Impl for ClassFactory {
    #[tracing::instrument]
    fn CreateInstance(
        &self,
        punkouter: Option<&IUnknown>,
        riid: *const GUID,
        ppvobject: *mut *mut core::ffi::c_void,
    ) -> Result<()> {
        trace!("Entered");

        if punkouter.is_some() {
            return CLASS_E_NOAGGREGATION.ok();
        }

        unsafe {
            if *riid == ITfTextInputProcessor::IID {
                let unknown: IUnknown = TextService::new().into();
                unknown.query(&*riid, ppvobject as _).ok()
            } else {
                trace!("Unknown IID: {:?}", *riid);
                E_NOINTERFACE.ok()
            }
        }
    }

    #[tracing::instrument(skip_all)]
    fn LockServer(&self, _flock: BOOL) -> Result<()> {
        trace!("Entered");
        S_OK.ok()
    }
}
