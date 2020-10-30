use serde::{Deserialize, Serialize};
#[derive(Debug, Clone, Default, Deserialize, Serialize)]
pub struct OpenXmlLeafTextElement(String);
#[derive(Debug, Clone, Default, Deserialize, Serialize)]
pub struct OpenXmlLeafElement(String);

#[derive(Debug, Clone, Default, Deserialize, Serialize)]
pub struct OpenXmlMiscNode(String);


