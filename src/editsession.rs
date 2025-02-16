use std::mem::ManuallyDrop;

use windows::{
    core::{implement, Interface},
    Win32::{
        Foundation::{FALSE, S_OK},
        UI::TextServices::{
            ITfCompositionSink, ITfContext, ITfContextComposition, ITfEditSession,
            ITfEditSession_Impl, ITfInsertAtSelection, TF_AE_NONE, TF_IAS_QUERYONLY, TF_SELECTION,
            TF_SELECTIONSTYLE,
        },
    },
};

use crate::service::TextService;

#[implement(ITfEditSession)]
pub struct StartCompositionEditSession<'a> {
    context: &'a ITfContext,
    service: &'a TextService,
}

impl<'a> StartCompositionEditSession<'a> {
    pub fn new(context: &'a ITfContext, service: &'a TextService) -> Self {
        StartCompositionEditSession { context, service }
    }
}

impl<'a> ITfEditSession_Impl for StartCompositionEditSession_Impl<'a> {
    fn DoEditSession(&self, ec: u32) -> windows_core::Result<()> {
        log::trace!("StartCompositionEditSession::DoEditSession");
        // we need a special interface to insert text at the selection
        let Ok(insert_at) = self.context.cast::<ITfInsertAtSelection>() else {
            return S_OK.ok();
        };

        // insert the text
        let Ok(range) = (unsafe { insert_at.InsertTextAtSelection(ec, TF_IAS_QUERYONLY, &[]) })
        else {
            return S_OK.ok();
        };

        // get an interface on the context we can use to deal with compositions
        let Ok(context_composition) = self.context.cast::<ITfContextComposition>() else {
            return S_OK.ok();
        };

        let sink: ITfCompositionSink = unsafe { self.service.cast_to().unwrap() };

        // start the new composition
        if let Ok(composition) = unsafe { context_composition.StartComposition(ec, &range, &sink) }
        {
            // Store the pointer of this new composition object in the instance
            // of the TextService struct. So this instance of the TextService
            // struct can know now it is in the composition stage.
            self.service.set_composition(composition);

            //  set selection to the adjusted range
            let range = ManuallyDrop::new(Some(range));
            let style = TF_SELECTIONSTYLE {
                ase: TF_AE_NONE,
                fInterimChar: FALSE,
            };
            let selection = TF_SELECTION { range, style };
            unsafe {
                _ = self.context.SetSelection(ec, &[selection]);
            }
        }

        S_OK.ok()
    }
}

#[implement(ITfEditSession)]
pub struct EndCompositionEditSession<'a> {
    context: &'a ITfContext,
    service: &'a TextService,
}

impl<'a> EndCompositionEditSession<'a> {
    pub fn new(service: &'a TextService, context: &'a ITfContext) -> Self {
        EndCompositionEditSession { context, service }
    }
}

impl<'a> ITfEditSession_Impl for EndCompositionEditSession_Impl<'a> {
    fn DoEditSession(&self, ec: u32) -> windows_core::Result<()> {
        log::trace!("EndCompositionEditSession::DoEditSession");
        self.service.terminate_composition(ec, &self.context);
        S_OK.ok()
    }
}
