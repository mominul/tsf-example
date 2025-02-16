use std::cell::RefCell;

use windows::{
    core::{implement, GUID},
    Win32::{
        Foundation::{COLORREF, E_INVALIDARG, FALSE, S_FALSE, S_OK},
        UI::TextServices::{
            CLSID_TF_CategoryMgr, IEnumTfDisplayAttributeInfo, IEnumTfDisplayAttributeInfo_Impl, ITfCategoryMgr, ITfContext, ITfDisplayAttributeInfo, ITfDisplayAttributeInfo_Impl, ITfDisplayAttributeProvider_Impl, GUID_PROP_ATTRIBUTE, TF_ATTR_INPUT, TF_ATTR_TARGET_CONVERTED, TF_CT_COLORREF, TF_CT_NONE, TF_DA_COLOR, TF_DA_COLOR_0, TF_DISPLAYATTRIBUTE, TF_LS_NONE, TF_LS_SOLID
        },
    },
};
use windows_core::VARIANT;
use winreg::{
    enums::{RegType::REG_BINARY, HKEY_CURRENT_USER, KEY_WRITE},
    RegKey, RegValue,
};

use crate::{globals::{GUID_DISPLAY_ATTRIBUTE_CONVERTED, GUID_DISPLAY_ATTRIBUTE_INPUT}, register::create_instance, service::{TextService, TextService_Impl}};

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
        log::trace!("EnumDisplayAttributeInfo::Clone");
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
        log::trace!("EnumDisplayAttributeInfo::Next");
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
        log::trace!("EnumDisplayAttributeInfo::Reset");
        *self.index.borrow_mut() = 0;
        S_OK.ok()
    }

    // Skips past objects in the enumeration.
    fn Skip(&self, ulcount: u32) -> windows_core::Result<()> {
        log::trace!("EnumDisplayAttributeInfo::Skip");
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
        log::trace!("DisplayAttributeInfo::GetGUID");
        Ok(self.guid)
    }

    fn GetDescription(&self) -> windows_core::Result<windows_core::BSTR> {
        log::trace!("DisplayAttributeInfo::GetDescription");
        Ok(self.description.clone().into())
    }

    fn GetAttributeInfo(&self, pda: *mut TF_DISPLAYATTRIBUTE) -> windows_core::Result<()> {
        log::trace!("DisplayAttributeInfo::GetAttributeInfo");
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
        log::trace!("DisplayAttributeInfo::SetAttributeInfo");
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
        log::trace!("DisplayAttributeInfo::Reset");
        self.SetAttributeInfo(&self.attribute)
    }
}

impl ITfDisplayAttributeProvider_Impl for TextService_Impl {
    fn EnumDisplayAttributeInfo(&self) -> windows_core::Result<IEnumTfDisplayAttributeInfo> {
        log::trace!("TextService::EnumDisplayAttributeInfo");
        let iter = EnumDisplayAttributeInfo::new();
        Ok(iter.into())
    }

    fn GetDisplayAttributeInfo(&self,guid: *const windows_core::GUID) -> windows_core::Result<ITfDisplayAttributeInfo> {
        log::trace!("TextService::GetDisplayAttributeInfo");
        let Some(guid) = (unsafe { guid.as_ref() }) else {
            return Err(E_INVALIDARG.into());
        };
        
        if *guid == GUID_DISPLAY_ATTRIBUTE_INPUT {
            let info = DisplayAttributeInfo::new_input();
            Ok(info.into())
        } else if *guid == GUID_DISPLAY_ATTRIBUTE_CONVERTED {
            let info = DisplayAttributeInfo::new_converted();
            Ok(info.into())
        } else {
            Err(E_INVALIDARG.into())
        }
    }
}

// apply the display attribute to the composition range.
impl TextService {
    /// Because it's expensive to map our display attribute GUID to a TSF
    /// TfGuidAtom, we do it once when Activate is called.
    pub fn init_display_attribute_guid_atom(&self) -> windows_core::Result<()> {
        log::trace!("TextService::init_display_attribute_guid_atom");
        let mgr: ITfCategoryMgr = create_instance(&CLSID_TF_CategoryMgr)?;

        unsafe {
            // register the display attribute for input text.
            *self.display_attribute_input.borrow_mut() = mgr.RegisterGUID(&GUID_DISPLAY_ATTRIBUTE_INPUT)?;

            // register the display attribute for the converted text.
            *self.display_attribute_converted.borrow_mut() = mgr.RegisterGUID(&GUID_DISPLAY_ATTRIBUTE_CONVERTED)?;
        }

        Ok(())
    }

    pub fn set_composition_display_attributes(&self, ec: u32, context: &ITfContext, attribute: i32) -> windows_core::Result<bool> {
        log::trace!("TextService::set_composition_display_attributes");
        let Ok(range) = (unsafe { self.composition.borrow().as_ref().unwrap().GetRange() }) else {
            return Ok(false);
        };

        let property = unsafe { context.GetProperty(&GUID_PROP_ATTRIBUTE)? };

        let var: VARIANT = attribute.into();
        
        unsafe {
            property.SetValue(ec, &range, &var)?;
        }

        Ok(true)
    }

    pub fn clear_composition_display_attributes(&self, ec: u32, context: &ITfContext) {
        log::trace!("TextService::clear_composition_display_attributes");
        let Ok(range) = (unsafe { self.composition.borrow().as_ref().unwrap().GetRange() }) else {
            return;
        };

        if let Ok(property) = unsafe { context.GetProperty(&GUID_PROP_ATTRIBUTE) } {
            // clear the value over the range
            unsafe {
                _ = property.Clear(ec, &range);
            }
        }
    }
}
