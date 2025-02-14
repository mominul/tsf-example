use std::cell::RefCell;

use windows::{
    core::implement,
    Win32::{
        Foundation::S_OK,
        UI::TextServices::{
            IEnumTfDisplayAttributeInfo, IEnumTfDisplayAttributeInfo_Impl, ITfDisplayAttributeInfo,
        },
    },
};

#[implement(IEnumTfDisplayAttributeInfo)]
struct EnumDisplayAttributeInfo {
    index: RefCell<u32>,
}

impl EnumDisplayAttributeInfo {
    pub fn new() -> Self {
        Self {
            index: RefCell::new(0),
        }
    }
}

impl IEnumTfDisplayAttributeInfo_Impl for EnumDisplayAttributeInfo_Impl {
    /// Returns a copy of the object.
    fn Clone(&self) -> windows_core::Result<IEnumTfDisplayAttributeInfo> {
        let clone = EnumDisplayAttributeInfo::new();

        // the clone should match this object's state
        *clone.index.borrow_mut() = *self.index.borrow();

        Ok(clone.into())
    }

    /// Returns an array of display attribute info objects supported by this service.
    fn Next(
        &self,
        ulcount: u32,
        rginfo: *mut Option<ITfDisplayAttributeInfo>,
        pcfetched: *mut u32,
    ) -> windows_core::Result<()> {
        // TODO: ITfDisplayAttributeInfo
        todo!()
    }

    // Resets the enumeration.
    fn Reset(&self) -> windows_core::Result<()> {
        *self.index.borrow_mut() = 0;
        S_OK.ok()
    }

    // Skips past objects in the enumeration.
    fn Skip(&self, ulcount: u32) -> windows_core::Result<()> {
        // we have only a single item to enum
        // so we can just skip it and avoid any overflow errors
        if ulcount > 0 && *self.index.borrow() == 0 {
            *self.index.borrow_mut() += 1;
        }

        S_OK.ok()
    }
}
