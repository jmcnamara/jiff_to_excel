//! This module provides functions to convert between Jiff's civil Date/Time and
//! Excel's serial date/time format.
//!
//! It is a prototype of the functions required to add `jiff` support to
//! `rust_xlsxwriter`.
//!
//! Excel doesn't support timezones so only the Jiff "Civil" types are handled
//! by this library.
//!
//! ## Datetimes in Excel
//!
//! Datetimes in Excel are serial dates with days counted from an epoch (usually
//! 1900-01-01) and where the time is a percentage/decimal of the milliseconds
//! in the day. Both the date and time are stored in the same f64 value. For
//! example, 2023/01/01 12:00:00 is stored as 44927.5.
//!
//! Datetimes in Excel must also be formatted with a number format like
//! `"yyyy/mm/dd hh:mm"` or otherwise they will appear as numbers (which
//! technically they are).
//!
//! Excel doesn't use timezones or try to convert or encode timezone information
//! in any way so they aren't supported by `rust_xlsxwriter`.
//!
//! Excel can also save dates in a text ISO 8601 format when the file is saved
//! using the "Strict Open XML Spreadsheet" option in the "Save" dialog. However
//! this is rarely used in practice and isn't supported by `rust_xlsxwriter`.
//!
#![warn(missing_docs)]
mod tests;

/// Convert a Jiff civil `DateTime` to an Excel serial datetime.
///
/// In Excel a serial date is the number of days since the epoch and a time is
/// the fraction of a day since midnight. The epoch if generally 1900-01-01.
///
/// # Examples
///
/// ```
/// use jiff::civil::DateTime;
/// use jiff_to_excel::jiff_datetime_to_excel;
///
/// let dt: DateTime = "2026-01-01 12:00".parse().unwrap();
/// assert_eq!(jiff_to_excel::jiff_datetime_to_excel(&dt), 46023.5);
/// ```
///
pub fn jiff_datetime_to_excel(datetime: &jiff::civil::DateTime) -> f64 {
    let date = jiff_date_to_excel(&datetime.date());
    let time = jiff_time_to_excel(&datetime.time());

    date + time
}

/// Convert a Jiff civil `Date` to an Excel serial datetime.
///
/// In Excel a serial date is the number of days since the epoch.
///
/// # Examples
///
/// ```
/// use jiff::civil::Date;
/// use jiff_to_excel::jiff_datetime_to_excel;
///
/// let d: Date = "2026-01-01".parse().unwrap();
/// assert_eq!(jiff_to_excel::jiff_date_to_excel(&d), 46023.0);
/// ```
///
pub fn jiff_date_to_excel(date: &jiff::civil::Date) -> f64 {
    let epoch = jiff::civil::date(1899, 12, 31);
    let duration = *date - epoch;

    let mut excel_date = f64::from(duration.get_days());

    // Excel treats 1900 as a leap year so we need to add an additional day for
    // dates after the leapday.
    if excel_date > 59.0 {
        excel_date += 1.0;
    }

    excel_date
}

/// Convert a Jiff civil `Time` to an Excel serial datetime.
///
/// In Excel a  time is the fraction of a day since midnight. The smallest unit
/// of time in Excel is the millisecond.
///
/// # Examples
///
/// ```
/// use jiff::civil::Time;
/// use jiff_to_excel::jiff_datetime_to_excel;
///
/// let t: Time = "12:00".parse().unwrap();
/// assert_eq!(jiff_to_excel::jiff_time_to_excel(&t), 0.5);
/// ```
///
pub fn jiff_time_to_excel(time: &jiff::civil::Time) -> f64 {
    let midnight = jiff::civil::time(0, 0, 0, 0);
    let duration = *time - midnight;

    duration.total(jiff::Unit::Millisecond).unwrap() / (24.0 * 60.0 * 60.0 * 1000.0)
}
