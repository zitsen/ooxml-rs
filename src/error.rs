use thiserror::Error;

#[derive(Error, Debug)]
pub enum OoxmlError {
    #[error("zip error")]
    ZipError(#[from] zip::result::ZipError),
    #[error("io error")]
    IoError(#[from] std::io::Error),
    #[error("url parse error")]
    UriError(#[from] url::ParseError),
    #[error("xml error")]
    XmlError(#[from] quick_xml::Error),
    #[error("No content type in package")]
    PackageContentTypeError,
    #[error("unknown data store error")]
    Unknown,
}