use serde::Deserialize;

#[derive(Deserialize, Debug)]
pub struct ColDefinition {
    pub width: f64,
    pub col_type: String,
    pub kbn_values: Option<Vec<u8>>,
}
