use windows::{
    core::{implement, Result},
    Win32::{UI::TextServices::{ITfTextInputProcessor, ITfTextInputProcessor_Impl, ITfThreadMgr}, Foundation::S_OK},
};

#[implement(ITfTextInputProcessor)]
pub struct TextService {}

impl ITfTextInputProcessor_Impl for TextService {
    fn Activate(&self, _ptim: &Option<ITfThreadMgr>, _tid: u32) -> Result<()> {
        log::trace!("TextService::Activate");
        S_OK.ok()
    }

    fn Deactivate(&self) -> Result<()> {
        log::trace!("TextService::Deactivate");
        S_OK.ok()
    }
}
