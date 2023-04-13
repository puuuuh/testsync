use thiserror::Error;

use crate::cell_id::CellId;

pub struct Excel {
    data: office::Excel,
    columns: u32,
}

#[derive(Debug, Error)]
pub enum ImportError {
    #[error("Invalid input data at {pos}, expected: {expected}, found: {found}")]
    InvalidData {
        pos: CellId,
        name: &'static str,
        expected: &'static str,
        found: String,
    },
    #[error("Found unexpected trailing data at {pos}")]
    TrailingData { pos: CellId },
}

#[derive(Debug)]
pub struct Data {
    pub nick: String,
    pub data: Vec<Option<f64>>,
}

impl Excel {
    pub fn new<T: AsRef<std::path::Path>>(filename: T, columns: u32) -> Self {
        let ex = office::Excel::open(filename).unwrap();
        Self { data: ex, columns }
    }

    pub fn data(&mut self) -> Result<Vec<Data>, ImportError> {
        let range = self.data.worksheet_range("Sheet1").unwrap();
        let mut row_idx = 1;
        let mut eof = false;
        let mut res = Vec::with_capacity(50);

        for row in range.rows().skip(1) {
            row_idx += 1;
            let nickname = match &row[0] {
                office::DataType::Empty => {
                    eof = true;
                    continue;
                }
                office::DataType::String(nick) if !eof => nick.clone(),
                data => {
                    return Err(if eof {
                        ImportError::TrailingData {
                            pos: CellId(0, row_idx),
                        }
                    } else {
                        ImportError::InvalidData {
                            pos: CellId(0, row_idx),
                            name: "nickname",
                            expected: "string",
                            found: format!("{data:?}"),
                        }
                    })
                }
            };

            let mut data = Vec::with_capacity(self.columns as _);
            for (i, x) in row[1..].iter().enumerate() {
                if i > self.columns as _ {
                    break;
                }

                match x {
                    office::DataType::Int(x) => data.push(Some(*x as f64)),
                    office::DataType::Float(x) => data.push(Some(*x)),
                    office::DataType::Empty => data.push(None),
                    x => {
                        return Err(ImportError::InvalidData {
                            pos: CellId((i + 1) as _, row_idx),
                            name: "value",
                            expected: "number",
                            found: format!("{x:?}"),
                        })
                    }
                }
            }

            res.push(Data {
                nick: nickname,
                data,
            });
        }
        Ok(res)
    }
}
