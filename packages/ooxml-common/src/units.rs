//! Shared OOXML unit conversion constants. EMU (English Metric Units) is the
//! coordinate unit for DrawingML transforms (`<a:off>`/`<a:ext>`, ECMA-376
//! §20.1.7.6 `ST_PositiveCoordinate`); several legacy or pixel-space inputs
//! (VML anchors, raw raster dimensions at the CSS-pixel default DPI) need to
//! be converted into it.

/// EMU per pixel at the 96 DPI default used by VML (`x:Anchor` offsets,
/// [MS-OI29500] 2.1.639) and by CSS pixels in general: 914400 EMU/inch ÷ 96
/// px/inch = 9525 EMU/px.
pub const EMU_PER_PX_96DPI: i64 = 9525;

/// Parse DrawingML `ST_Percentage` into a fraction. Transitional OOXML uses
/// an `xsd:int` in thousandths of a percent (`100000` = 100%), while Strict
/// OOXML uses a signed decimal followed by `%` (`100%` = 100%).
pub fn drawingml_percentage_to_thousandths(value: &str) -> Option<f64> {
    if let Some(percent) = value.strip_suffix('%') {
        let digits = percent.strip_prefix('-').unwrap_or(percent);
        if digits.is_empty() || percent.starts_with('+') {
            return None;
        }
        let mut parts = digits.split('.');
        let whole = parts.next().unwrap_or_default();
        let fraction = parts.next();
        if whole.is_empty()
            || !whole.bytes().all(|byte| byte.is_ascii_digit())
            || fraction.is_some_and(|part| {
                part.is_empty() || !part.bytes().all(|byte| byte.is_ascii_digit())
            })
            || parts.next().is_some()
        {
            return None;
        }
        return finite_number(percent).map(|number| number * 1_000.0);
    }
    value.parse::<i32>().ok().map(f64::from)
}

pub fn drawingml_percentage_to_fraction(value: &str) -> Option<f64> {
    drawingml_percentage_to_thousandths(value).map(|value| value / 100_000.0)
}

/// Parse DrawingML `ST_Coordinate32` into EMU.
///
/// ECMA-376 defines this as a union of `xsd:int` (already expressed in EMU)
/// and `ST_UniversalMeasure`. The latter accepts a signed decimal followed by
/// `mm`, `cm`, `in`, `pt`, `pc`, or `pi`. Universal measures can resolve to a
/// fractional EMU; the integer wire model stores the nearest representable EMU.
pub fn coordinate32_to_emu(value: &str) -> Option<i64> {
    let value = value.trim();
    if !value.is_empty()
        && value
            .strip_prefix('+')
            .or_else(|| value.strip_prefix('-'))
            .unwrap_or(value)
            .bytes()
            .all(|byte| byte.is_ascii_digit())
    {
        return value.parse::<i32>().ok().map(i64::from);
    }

    universal_measure_to_emu(value)
}

/// Parse DrawingML `ST_Coordinate` into EMU. Its unitless branch is an
/// `xsd:long`, unlike `ST_Coordinate32`; the UniversalMeasure branch is shared.
pub fn coordinate_to_emu(value: &str) -> Option<i64> {
    let value = value.trim();
    if !value.is_empty()
        && value
            .strip_prefix('+')
            .or_else(|| value.strip_prefix('-'))
            .unwrap_or(value)
            .bytes()
            .all(|byte| byte.is_ascii_digit())
    {
        return value.parse::<i64>().ok();
    }
    universal_measure_to_emu(value)
}

fn universal_measure_to_emu(value: &str) -> Option<i64> {
    let (number, scale) = if let Some(number) = value.strip_suffix("pt") {
        (number, 12_700.0)
    } else if let Some(number) = value.strip_suffix("in") {
        (number, 914_400.0)
    } else if let Some(number) = value
        .strip_suffix("pc")
        .or_else(|| value.strip_suffix("pi"))
    {
        (number, 152_400.0)
    } else if let Some(number) = value.strip_suffix("mm") {
        (number, 36_000.0)
    } else {
        (value.strip_suffix("cm")?, 360_000.0)
    };
    if !is_universal_measure_number(number) {
        return None;
    }
    let emu = finite_number(number)? * scale;
    if emu < i64::MIN as f64 || emu > i64::MAX as f64 {
        return None;
    }
    Some(emu.round() as i64)
}

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
    use super::{
        coordinate32_to_emu, coordinate_to_emu, drawingml_percentage_to_fraction, text_point_to_pt,
    };

    #[test]
    fn drawingml_percentage_parses_strict_and_transitional_lexemes() {
        for (value, expected) in [
            ("25000", 0.25),
            ("25%", 0.25),
            ("-50000", -0.5),
            ("-50%", -0.5),
            ("125%", 1.25),
        ] {
            assert_eq!(drawingml_percentage_to_fraction(value), Some(expected));
        }
        for value in ["", "+25%", "25.%", ".25%", "1e2%", "1.5"] {
            assert_eq!(
                drawingml_percentage_to_fraction(value),
                None,
                "value={value}"
            );
        }
    }

    #[test]
    fn coordinate32_parses_integer_emu_and_universal_measures() {
        for (value, expected) in [
            ("12700", 12_700),
            ("-12700", -12_700),
            ("+12700", 12_700),
            ("1pt", 12_700),
            ("1in", 914_400),
            ("1pc", 152_400),
            ("1pi", 152_400),
            ("25.4mm", 914_400),
            ("2.54cm", 914_400),
            ("-0.5pt", -6_350),
        ] {
            assert_eq!(coordinate32_to_emu(value), Some(expected), "value={value}");
        }
    }

    #[test]
    fn coordinate32_rejects_non_schema_lexemes_and_unitless_non_integers() {
        for value in ["", "1.5", "+1pt", "1px", "pt", "1e2", "NaN", "inf"] {
            assert_eq!(coordinate32_to_emu(value), None, "value={value}");
        }
    }

    #[test]
    fn coordinate_parses_long_emu_and_universal_measures() {
        assert_eq!(coordinate_to_emu("9223372036854775807"), Some(i64::MAX));
        assert_eq!(coordinate_to_emu("-9223372036854775808"), Some(i64::MIN));
        assert_eq!(coordinate_to_emu("1pt"), Some(12_700));
        assert_eq!(coordinate_to_emu("-0.5pt"), Some(-6_350));
        assert_eq!(coordinate_to_emu("9223372036854775808"), None);
    }

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
