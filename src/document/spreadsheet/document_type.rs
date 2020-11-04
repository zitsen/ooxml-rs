

#[derive(Debug, Clone, Copy)]
pub enum SpreadsheetDocumentType
{
    /// Excel Workbook (*.xlsx).
    Workbook,
    /// Excel Template (*.xltx).
    Template,
    /// Excel Macro-Enabled Workbook (*.xlsm).
    MacroEnabledWorkbook,
    /// Excel Macro-Enabled Template (*.xltm).
    MacroEnabledTemplate,
    /// Excel Add-In (*.xlam).
    AddIn,
}

impl Default for SpreadsheetDocumentType {
    fn default() -> Self {
        SpreadsheetDocumentType::Workbook
    }
}

const WORKBOOK_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const TEMPLATE_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
const MACRO_ENABLED_WORKBOOK_CONTENT_TYPE: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
const MACRO_ENABLED_TEMPLATE_CONTENT_TYPE: &str = "application/vnd.ms-excel.template.macroEnabled.main+xml";
const ADDIN_CONTENT_TYPE: &str = "application/vnd.ms-excel.addin.macroEnabled.main+xml";
impl SpreadsheetDocumentType {
    pub fn content_type(&self) -> &'static str {
        match self {
            SpreadsheetDocumentType::Workbook => WORKBOOK_CONTENT_TYPE,
            SpreadsheetDocumentType::Template => TEMPLATE_CONTENT_TYPE,
            SpreadsheetDocumentType::MacroEnabledWorkbook => MACRO_ENABLED_WORKBOOK_CONTENT_TYPE,
            SpreadsheetDocumentType::MacroEnabledTemplate => MACRO_ENABLED_TEMPLATE_CONTENT_TYPE,
            SpreadsheetDocumentType::AddIn => ADDIN_CONTENT_TYPE,
        }
    }
    pub fn from_content_type(content_type: &str) -> Self {
        match content_type {
            WORKBOOK_CONTENT_TYPE => SpreadsheetDocumentType::Workbook,
            TEMPLATE_CONTENT_TYPE => SpreadsheetDocumentType::Template,
            MACRO_ENABLED_WORKBOOK_CONTENT_TYPE => SpreadsheetDocumentType::MacroEnabledWorkbook,
            MACRO_ENABLED_TEMPLATE_CONTENT_TYPE => SpreadsheetDocumentType::MacroEnabledTemplate,
            ADDIN_CONTENT_TYPE => SpreadsheetDocumentType::AddIn,
            _ => panic!("unsupported spread sheet document content type"),
        }

    }
}
