//! Documents implementations.
//!
//! A document is a OPC package with specific components.
mod presentation;
mod spreadsheet;
mod wordprocessing;

pub use spreadsheet::SpreadsheetDocument;
