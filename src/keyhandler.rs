use std::mem::ManuallyDrop;

use windows::core::implement;
use windows::Win32::UI::Input::KeyboardAndMouse::VK_SPACE;
use windows::Win32::UI::TextServices::ITfEditSession;
use windows::Win32::{
    Foundation::{LPARAM, S_OK, WPARAM},
    UI::{
        Input::KeyboardAndMouse::{VK_LEFT, VK_RETURN, VK_RIGHT},
        TextServices::{
            ITfContext, ITfEditSession_Impl, ITfRange, TF_ANCHOR_END, TF_ANCHOR_START,
            TF_DEFAULT_SELECTION, TF_ES_READWRITE, TF_ES_SYNC, TF_SELECTION,
        },
    },
};
use windows_core::Result;

use crate::service::TextService;

#[implement(ITfEditSession)]
pub struct KeyHandlerEditSession<'a> {
    service: &'a TextService,
    context: &'a ITfContext,
    param: WPARAM,
}

impl<'a> KeyHandlerEditSession<'a> {
    pub fn new(service: &'a TextService, context: &'a ITfContext, param: WPARAM) -> Self {
        KeyHandlerEditSession {
            service,
            context,
            param,
        }
    }
}

impl<'a> ITfEditSession_Impl for KeyHandlerEditSession_Impl<'a> {
    fn DoEditSession(&self, ec: u32) -> windows_core::Result<()> {
        log::trace!(
            "KeyHandlerEditSession::DoEditSession -> param: {:?}",
            self.param
        );
        if self.param.0 == VK_LEFT.0.into() || self.param.0 == VK_RIGHT.0.into() {
            self.service.handle_arrow_key(ec, &self.context, self.param)
        } else if self.param.0 == VK_RETURN.0.into() {
            self.service.handle_return_key(ec, &self.context)
        } else if self.param.0 == VK_SPACE.0.into() {
            self.service.handle_space_key(ec, &self.context)
        } else if self.param.0 >= b'A'.into() && self.param.0 <= b'Z'.into() {
            self.service
                .handle_character_key(ec, &self.context, self.param)
        } else {
            S_OK.ok()
        }
    }
}

impl TextService {
    /// If the keystroke happens within a composition, eat the key and return S_OK.
    pub fn handle_character_key(&self, ec: u32, context: &ITfContext, param: WPARAM) -> Result<()> {
        log::trace!("TextService::handle_character_key");
        // Start the new compositon if there is no composition.
        if self.is_composing() {
            self.start_composition(context);
        }

        // Assign VK_ value to the char. So the inserted the character is always
        // uppercase.
        let ch: u16 = param.0.try_into().unwrap();

        // first, test where a keystroke would go in the document if we did an insert
        let mut selection = [TF_SELECTION::default()];
        let mut fetched = 0;
        unsafe {
            context.GetSelection(ec, TF_DEFAULT_SELECTION, &mut selection, &mut fetched)?;
        }
        let [selection] = selection;
        let selection_range = ManuallyDrop::into_inner(selection.range.clone()).unwrap();

        // is the insertion point covered by a composition?
        if let Ok(range) = unsafe { self.composition.borrow().as_ref().unwrap().GetRange() } {
            let covered = is_range_covered(ec, &selection_range, &range);

            if !covered {
                return S_OK.ok();
            }
        }

        unsafe {
            // insert the text
            // we use SetText here instead of InsertTextAtSelection because we've already started a composition
            // we don't want to the app to adjust the insertion point inside our composition
            selection_range.SetText(ec, 0, &[ch])?;

            // update the selection, we'll make it an insertion point just past
            // the inserted text.
            _ = selection_range.Collapse(ec, TF_ANCHOR_END);
            _ = context.SetSelection(ec, &[selection]);
        }

        // set the display attribute to the composition range.
        _ = self.set_composition_display_attributes(ec, context, *self.display_attribute_input.borrow() as i32);

        S_OK.ok()
    }

    pub fn handle_return_key(&self, ec: u32, context: &ITfContext) -> Result<()> {
        log::trace!("TextService::handle_return_key");
        // just terminate the composition
        self.terminate_composition(ec, context);
        S_OK.ok()
    }

    pub fn handle_space_key(&self, ec: u32, context: &ITfContext) -> Result<()> {
        log::trace!("TextService::handle_space_key");
        // set the display attribute to the composition range.
        //
        // The real text service may have linguistic logic here and set 
        // the specific range to apply the display attribute rather than 
        // applying the display attribute to the entire composition range.
        _ = self.set_composition_display_attributes(ec, context, *self.display_attribute_converted.borrow() as i32);
        S_OK.ok()
    }

    pub fn handle_arrow_key(&self, ec: u32, context: &ITfContext, param: WPARAM) -> Result<()> {
        log::trace!("TextService::handle_arrow_key");
        // get the selection
        let mut selection = [TF_SELECTION::default()];
        let mut fetched = 0;
        let mut cch = 0;
        if unsafe {
            context
                .GetSelection(ec, TF_DEFAULT_SELECTION, &mut selection, &mut fetched)
                .is_err()
        } && fetched != 1
        {
            // no selection?
            return S_OK.ok();
        }
        let [selection] = selection;
        let selection_range = ManuallyDrop::into_inner(selection.range.clone()).unwrap();

        // get the composition range
        let range = unsafe { self.composition.borrow().as_ref().unwrap().GetRange()? };

        // adjust the selection, we won't do anything fancy
        if param.0 == VK_LEFT.0.into() {
            unsafe {
                if !selection_range
                    .IsEqualStart(ec, &range, TF_ANCHOR_START)
                    .unwrap()
                    .as_bool()
                {
                    _ = selection_range.ShiftStart(ec, -1, &mut cch, std::ptr::null());
                }
                _ = selection_range.Collapse(ec, TF_ANCHOR_START);
            }
        } else {
            // VK_RIGHT
            unsafe {
                if !selection_range
                    .IsEqualEnd(ec, &range, TF_ANCHOR_END)
                    .unwrap()
                    .as_bool()
                {
                    _ = selection_range.ShiftEnd(ec, 1, &mut cch, std::ptr::null());
                }
                _ = selection_range.Collapse(ec, TF_ANCHOR_END);
            }
        }

        unsafe {
            _ = context.SetSelection(ec, &[selection]);
        }

        return S_OK.ok(); // eat the keystroke
    }

    /// This text service is interested in handling keystrokes to demonstrate the
    /// use the compositions. Some apps will cancel compositions if they receive
    /// keystrokes while a compositions is ongoing.
    pub fn invoke_key_handler(
        &self,
        context: &ITfContext,
        wparam: WPARAM,
        _lparam: LPARAM,
    ) -> Result<()> {
        log::trace!("TextService::invoke_key_handler");
        let session = KeyHandlerEditSession::new(&self, context, wparam);
        let session: ITfEditSession = session.into();

        // we need a lock to do our work
        // nb: this method is one of the few places where it is legal to use
        // the TF_ES_SYNC flag
        unsafe {
            _ = context.RequestEditSession(
                self.client_id.borrow().clone(),
                &session,
                TF_ES_SYNC | TF_ES_READWRITE,
            )?;
        }

        return S_OK.ok();
    }
}

/// Returns TRUE if pRangeTest is entirely contained within pRangeCover.
pub fn is_range_covered(ec: u32, range_test: &ITfRange, range_cover: &ITfRange) -> bool {
    unsafe {
        if let Ok(result) = range_cover.CompareStart(ec, range_test, TF_ANCHOR_START) {
            if result > 0 {
                return false;
            }
        } else {
            return false;
        }

        if let Ok(result) = range_cover.CompareEnd(ec, range_test, TF_ANCHOR_END) {
            if result < 0 {
                return false;
            }
        } else {
            return false;
        }
    }

    return true;
}
