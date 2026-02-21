use serde::Deserialize;

#[derive(Debug, Clone)]
pub enum ColType {
    Str,
    Int,
    Date,
    KbnList,
}

#[derive(Deserialize, Debug)]
pub struct ColDefinition {
    pub width: f64,
    pub col_type: String,
    pub kbn_values: Option<Vec<u8>>,
}

impl ColType {
    pub fn from_str(s: &str) -> Self {
        match s.to_lowercase().as_str() {
            "int" => ColType::Int,
            "date" => ColType::Date,
            "kbn_list" => ColType::KbnList,
            _ => ColType::Str,
        }
    }
}
