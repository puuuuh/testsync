use std::iter;

use google_sheets4::{api::ValueRange, hyper, hyper_rustls, oauth2, Sheets};
use serde_json::Number;

use crate::{cell_id::CellId, import::Data};

pub struct GSheets {
    pub sheet_id: String,
    pub start: CellId,
    pub columns: u32,
}

impl GSheets {
    pub async fn export(&self, data: Vec<Data>) -> Result<(), google_sheets4::Error> {
        // Get an ApplicationSecret instance by some means. It contains the `client_id` and
        // `client_secret`, among other things.
        let secret = oauth2::read_application_secret("creds.json")
            .await
            .expect("client secret could not be read");

        // Instantiate the authenticator. It will choose a suitable authentication flow for you,
        // unless you replace  `None` with the desired Flow.
        // Provide your own `AuthenticatorDelegate` to adjust the way it operates and get feedback about
        // what's going on. You probably want to bring in your own `TokenStorage` to persist tokens and
        // retrieve them from storage.
        let auth = oauth2::InstalledFlowAuthenticator::builder(
            secret,
            oauth2::InstalledFlowReturnMethod::HTTPRedirect,
        )
        .persist_tokens_to_disk("tokencache.json")
        .build()
        .await
        .unwrap();

        let hub = Sheets::new(
            hyper::Client::builder().build(
                hyper_rustls::HttpsConnectorBuilder::new()
                    .with_native_roots()
                    .https_or_http()
                    .enable_http1()
                    .enable_http2()
                    .build(),
            ),
            auth,
        );

        let data = data
            .into_iter()
            .map(|s| {
                iter::once(serde_json::Value::String(s.nick))
                    .chain(s.data.into_iter().map(|s| match s {
                        Some(s) => serde_json::Value::Number(Number::from_f64(s).unwrap()),
                        None => serde_json::Value::String(String::default()),
                    }))
                    .collect()
            })
            .collect::<Vec<Vec<_>>>();

        let start = self.start;
        let end = start.add_x(self.columns + 1).add_y(data.len() as _);

        hub.spreadsheets()
            .values_update(
                ValueRange {
                    major_dimension: None,
                    range: None,
                    values: Some(data),
                },
                &self.sheet_id,
                &format!("{start}:{end}"),
            )
            .value_input_option("RAW")
            .include_values_in_response(false)
            .doit()
            .await?;
        Ok(())
    }
}
