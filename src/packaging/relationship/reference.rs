pub struct ReferenceRelationship {
    //container: Rc<Cell<crate::packaging::part::Container>>,
    id: String,
    is_external: bool,
    relationship_type: String,
    uri: url::Url,
}

impl ReferenceRelationship {
    fn is_external(&self) -> bool {
        self.is_external
    }
}
