use std::path::PathBuf;

use cell_id::CellId;
use serde::{Deserialize, Serialize};

mod cell_id;
pub mod export;
mod import;

const CONFIG_PATH: &str = "config.yaml";

#[derive(Debug, Default, Clone, Deserialize, Serialize)]
struct Config {
    columns: u32,
    sheet_id: String,
    input: PathBuf,
    output_pos: CellId,
}

async fn main_impl() -> Result<(), Box<dyn std::error::Error>> {
    let file = std::fs::OpenOptions::new().read(true).open(CONFIG_PATH);
    let file = match file {
        Ok(f) => f,
        Err(e) => {
            let f = std::fs::OpenOptions::new()
                .write(true)
                .create_new(true)
                .open(CONFIG_PATH)?;

            serde_yaml::to_writer(f, &Config::default())?;

            println!("[ERR] Can't open config file");

            return Err(Box::new(e));
        }
    };

    let config: Config = serde_yaml::from_reader(file)?;

    let mut import = import::Excel::new(config.input, config.columns);

    export::GSheets {
        sheet_id: config.sheet_id,
        start: config.output_pos,
        columns: config.columns,
    }
    .export(import.data()?)
    .await?;

    Ok(())
}

#[tokio::main]
async fn main() {
    match main_impl().await {
        Ok(_) => {}
        Err(e) => {
            eprintln!("{}", e)
        }
    }
}
