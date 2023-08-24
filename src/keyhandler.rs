use windows::Win32::UI::TextServices::{ITfRange, TF_ANCHOR_END, TF_ANCHOR_START};

pub fn is_range_covered(ec: u32, range_test: ITfRange, range_cover: ITfRange) -> bool {
    unsafe {
        if let Ok(result) = range_cover.CompareStart(ec, &range_test, TF_ANCHOR_START) {
            if result > 0 {
                return false;
            }
        } else {
            return false;
        }

        if let Ok(result) = range_cover.CompareEnd(ec, &range_test, TF_ANCHOR_END) {
            if result < 0 {
                return false;
            }
        } else {
            return false;
        }
    }

    return true;
}
