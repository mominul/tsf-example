use std::cell::RefCell;

use windows::{
    core::{implement, GUID},
    Win32::{
        Foundation::{COLORREF, FALSE, S_FALSE, S_OK},
        UI::TextServices::{
            IEnumTfDisplayAttributeInfo, IEnumTfDisplayAttributeInfo_Impl, ITfDisplayAttributeInfo,
            ITfDisplayAttributeInfo_Impl, TF_ATTR_INPUT, TF_ATTR_TARGET_CONVERTED, TF_CT_COLORREF,
            TF_CT_NONE, TF_DA_COLOR, TF_DA_COLOR_0, TF_DISPLAYATTRIBUTE, TF_LS_NONE, TF_LS_SOLID,
        },
    },
};
use winreg::{
    enums::{RegType::REG_BINARY, HKEY_CURRENT_USER, KEY_WRITE},
    RegKey, RegValue,
};

use crate::globals::{GUID_DISPLAY_ATTRIBUTE_CONVERTED, GUID_DISPLAY_ATTRIBUTE_INPUT};

// the registry key of this text service to save the custmized display attribute
const ATTRIBUTE_INFO_KEY: &str = "Software\\Sample Text Service";

const fn rgb(r: u8, g: u8, b: u8) -> u32 {
    (r as u32) | ((g as u32) << 8) | ((b as u32) << 16)
}

const DISPLAY_ATTRIBUTE_INFO_INPUT: TF_DISPLAYATTRIBUTE = TF_DISPLAYATTRIBUTE {
    // text color
    crText: TF_DA_COLOR {
        r#type: TF_CT_COLORREF,
        Anonymous: TF_DA_COLOR_0 {
            cr: COLORREF(rgb(255, 0, 0)),
        },
    },
    // background color (TF_CT_NONE => app default)
    crBk: TF_DA_COLOR {
        r#type: TF_CT_NONE,
        Anonymous: TF_DA_COLOR_0 { nIndex: 0 },
    },
    lsStyle: TF_LS_SOLID, // underline style
    fBoldLine: FALSE,     // underline boldness
    // underline color
    crLine: TF_DA_COLOR {
        r#type: TF_CT_COLORREF,
        Anonymous: TF_DA_COLOR_0 {
            cr: COLORREF(rgb(255, 0, 0)),
        },
    },
    bAttr: TF_ATTR_INPUT, // attribute info
};

const DISPLAY_ATTRIBUTE_INFO_CONVERTED: TF_DISPLAYATTRIBUTE = TF_DISPLAYATTRIBUTE {
    // text color
    crText: TF_DA_COLOR {
        r#type: TF_CT_COLORREF,
        Anonymous: TF_DA_COLOR_0 {
            cr: COLORREF(rgb(255, 255, 255)),
        },
    },
    // background color (TF_CT_NONE => app default)
    crBk: TF_DA_COLOR {
        r#type: TF_CT_COLORREF,
        Anonymous: TF_DA_COLOR_0 {
            cr: COLORREF(rgb(0, 255, 255)),
        },
    },
    lsStyle: TF_LS_NONE, // underline style
    fBoldLine: FALSE,    // underline boldness
    // underline color
    crLine: TF_DA_COLOR {
        r#type: TF_CT_NONE,
        Anonymous: TF_DA_COLOR_0 { nIndex: 0 },
    },
    bAttr: TF_ATTR_TARGET_CONVERTED, // attribute info
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
        let mut fetched = 0;

        if ulcount == 0 {
            return Ok(());
        }

        while fetched < ulcount {
            if *self.index.borrow() > 1 {
                break;
            }

            if *self.index.borrow() == 0 {
                let attribute = Some(DisplayAttributeInfo::new_input().into());
                unsafe {
                    rginfo.write(attribute);
                }
            } else if *self.index.borrow() == 1 {
                let attribute = Some(DisplayAttributeInfo::new_converted().into());
                unsafe {
                    rginfo.write(attribute);
                }
            }

            fetched += 1;
            *self.index.borrow_mut() += 1;
        }

        if !pcfetched.is_null() {
            unsafe {
                // technically this is only legal if ulcount == 1, but we won't check
                pcfetched.write(fetched);
            }
        }

        if fetched == ulcount {
            Ok(())
        } else {
            Err(S_FALSE.into()) 
        }
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

#[implement(ITfDisplayAttributeInfo)]
pub struct DisplayAttributeInfo {
    guid: GUID,
    name: String,
    description: String,
    attribute: TF_DISPLAYATTRIBUTE,
}

impl DisplayAttributeInfo {
    pub fn new_input() -> Self {
        Self {
            guid: GUID_DISPLAY_ATTRIBUTE_INPUT,
            name: "DisplayAttributeInput".to_owned(),
            description: "TextService Display Attribute Input".to_owned(),
            attribute: DISPLAY_ATTRIBUTE_INFO_INPUT,
        }
    }

    pub fn new_converted() -> Self {
        Self {
            guid: GUID_DISPLAY_ATTRIBUTE_CONVERTED,
            name: "DisplayAttributeConverted".to_owned(),
            description: "TextService Display Attribute Converted".to_owned(),
            attribute: DISPLAY_ATTRIBUTE_INFO_CONVERTED,
        }
    }
}

impl ITfDisplayAttributeInfo_Impl for DisplayAttributeInfo_Impl {
    fn GetGUID(&self) -> windows_core::Result<GUID> {
        Ok(self.guid)
    }

    fn GetDescription(&self) -> windows_core::Result<windows_core::BSTR> {
        Ok(self.description.clone().into())
    }

    fn GetAttributeInfo(&self, pda: *mut TF_DISPLAYATTRIBUTE) -> windows_core::Result<()> {
        let hkcu = RegKey::predef(HKEY_CURRENT_USER);
        let value = hkcu
            .open_subkey(ATTRIBUTE_INFO_KEY)
            .and_then(|key| key.get_raw_value(&self.name));

        match value {
            Ok(value) if value.bytes.len() == size_of::<TF_DISPLAYATTRIBUTE>() => unsafe {
                let attr: TF_DISPLAYATTRIBUTE =
                    *(value.bytes.as_ptr() as *const TF_DISPLAYATTRIBUTE);
                pda.write(attr);
            },

            _ => unsafe { pda.write(self.attribute) },
        }

        Ok(())
    }

    fn SetAttributeInfo(&self, pda: *const TF_DISPLAYATTRIBUTE) -> windows_core::Result<()> {
        let hkcu = RegKey::predef(HKEY_CURRENT_USER);
        let (key, _) = hkcu
            .create_subkey_with_flags(ATTRIBUTE_INFO_KEY, KEY_WRITE)
            .map_err(|_| windows::Win32::Foundation::E_FAIL)?;

        // Serialize TF_DISPLAYATTRIBUTE into bytes
        let bytes = unsafe {
            std::slice::from_raw_parts(pda as *const u8, size_of::<TF_DISPLAYATTRIBUTE>()).to_vec()
        };

        // Set the value as binary
        key.set_raw_value(
            &self.name,
            &RegValue {
                bytes,
                vtype: REG_BINARY,
            },
        )
        .map_err(|_| windows::Win32::Foundation::E_FAIL)?;

        Ok(())
    }

    fn Reset(&self) -> windows_core::Result<()> {
        self.SetAttributeInfo(&self.attribute)
    }
}
