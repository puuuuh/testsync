use core::fmt;
use std::fmt::Display;

use serde::{
    de::{self, Visitor},
    Deserialize, Serialize,
};

#[derive(Debug, Default, Clone, Copy)]
pub struct CellId(pub u32, pub u32);

impl Serialize for CellId {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        serializer.serialize_str(&self.to_string())
    }
}

impl<'de> Deserialize<'de> for CellId {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct CellIdVisitor;

        impl<'de> Visitor<'de> for CellIdVisitor {
            type Value = CellId;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an excel cell id")
            }

            fn visit_str<E>(self, v: &str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                let mut pos = CellId::default();
                let mut first = true;

                for l in v.chars() {
                    let t = l as u32;

                    if first && t >= b'A' as u32 && t <= b'Z' as u32 {
                        pos.0 *= 26;
                        pos.0 += (t - b'A' as u32) + 1;
                        continue;
                    }

                    first = false;

                    if t >= b'0' as u32 && t <= b'9' as u32 {
                        pos.1 *= 10;
                        pos.1 += t - b'0' as u32;
                        continue;
                    }

                    return Err(E::custom("Invalid cell id"));
                }

                pos.0 -= 1;
                pos.1 -= 1;

                Ok(pos)
            }
        }

        deserializer.deserialize_str(CellIdVisitor)
    }
}

impl Display for CellId {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let len = f64::log(self.0 as f64, 26.).ceil() as usize;
        let mut res = Vec::with_capacity(len);
        let mut l = self.0 + 1;
        let r = self.1 + 1;
        while l > 0 {
            let m = (l - 1) % 26;
            res.push(65 + m as u8);
            l = (l - m) / 26;
        }
        res.reverse();
        let res = unsafe { String::from_utf8_unchecked(res) };
        f.write_str(&res)?;

        write!(f, "{r}")
    }
}

impl CellId {
    pub fn add_x(&self, x: u32) -> Self {
        Self(self.0 + x, self.1)
    }

    pub fn add_y(&self, y: u32) -> Self {
        Self(self.0, self.1 + y)
    }
}
