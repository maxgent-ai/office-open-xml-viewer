//! Shared OOXML unit conversion constants. EMU (English Metric Units) is the
//! coordinate unit for DrawingML transforms (`<a:off>`/`<a:ext>`, ECMA-376
//! §20.1.7.6 `ST_PositiveCoordinate`); several legacy or pixel-space inputs
//! (VML anchors, raw raster dimensions at the CSS-pixel default DPI) need to
//! be converted into it.

/// EMU per pixel at the 96 DPI default used by VML (`x:Anchor` offsets,
/// [MS-OI29500] 2.1.639) and by CSS pixels in general: 914400 EMU/inch ÷ 96
/// px/inch = 9525 EMU/px.
pub const EMU_PER_PX_96DPI: i64 = 9525;

/// Parse DrawingML `ST_TextPoint` (ECMA-376 §20.1.10.74) into points.
///
/// A unitless value is an XSD integer in hundredths of a point. The union also
/// accepts `ST_UniversalMeasure` (§22.9.2.15): a signed decimal followed by
/// `mm`, `cm`, `in`, `pt`, `pc`, or `pi` (`pc` and `pi` are both picas).
pub fn text_point_to_pt(value: &str) -> Option<f64> {
    let value = value.trim();
    let digits = value
        .strip_prefix('+')
        .or_else(|| value.strip_prefix('-'))
        .unwrap_or(value);
    if !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return finite_number(value).map(|number| number / 100.0);
    }

    let (number, scale) = if let Some(number) = value.strip_suffix("pt") {
        (number, 1.0)
    } else if let Some(number) = value.strip_suffix("in") {
        (number, 72.0)
    } else if let Some(number) = value
        .strip_suffix("pc")
        .or_else(|| value.strip_suffix("pi"))
    {
        (number, 12.0)
    } else if let Some(number) = value.strip_suffix("mm") {
        (number, 72.0 / 25.4)
    } else {
        (value.strip_suffix("cm")?, 72.0 / 2.54)
    };
    if !is_universal_measure_number(number) {
        return None;
    }
    finite_number(number).map(|number| number * scale)
}

fn is_universal_measure_number(value: &str) -> bool {
    let digits = value.strip_prefix('-').unwrap_or(value);
    if digits.is_empty() || value.starts_with('+') {
        return false;
    }
    let mut parts = digits.split('.');
    let whole = parts.next().unwrap_or_default();
    let fraction = parts.next();
    !whole.is_empty()
        && whole.bytes().all(|byte| byte.is_ascii_digit())
        && fraction
            .is_none_or(|part| !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit()))
        && parts.next().is_none()
}

fn finite_number(value: &str) -> Option<f64> {
    value
        .parse::<f64>()
        .ok()
        .filter(|number| number.is_finite())
}

#[cfg(test)]
mod tests {
    use super::text_point_to_pt;

    #[test]
    fn text_point_parses_hundredths_and_universal_measures() {
        for (value, expected) in [
            ("100", 1.0),
            ("-50", -0.5),
            ("0", 0.0),
            ("1pt", 1.0),
            ("1in", 72.0),
            ("1pc", 12.0),
            ("1pi", 12.0),
            ("25.4mm", 72.0),
            ("2.54cm", 72.0),
            ("-1.5pt", -1.5),
        ] {
            let actual = text_point_to_pt(value).expect("valid ST_TextPoint");
            assert!((actual - expected).abs() < 1e-9, "value={value}");
        }
    }

    #[test]
    fn text_point_rejects_non_schema_lexemes() {
        for value in ["", "1.5", "+1pt", "1px", "pt", "1e2", "NaN", "inf"] {
            assert_eq!(text_point_to_pt(value), None, "value={value}");
        }
    }
}
