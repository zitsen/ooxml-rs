pub struct ReferenceRelationship {
    //container: Rc<Cell<crate::packaging::part::Container>>,
    pub id: String,
    pub is_external: bool,
    pub relationship_type: String,
    pub uri: url::Url,
}

impl ReferenceRelationship {
    pub fn is_external(&self) -> bool {
        self.is_external
    }
}
