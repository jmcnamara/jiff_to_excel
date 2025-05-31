# jiff_to_excel

This module provides functions to convert between Jiff's civil Date/Time and
Excel's serial date/time format.

It is a prototype of the functions required to add `jiff` support to
`rust_xlsxwriter`.

Excel doesn't support timezones so only the Jiff "Civil" types are handled
by this library.

## Datetimes in Excel

Datetimes in Excel are serial dates with days counted from an epoch (usually
1900-01-01) and where the time is a percentage/decimal of the milliseconds
in the day. Both the date and time are stored in the same f64 value. For
example, 2023/01/01 12:00:00 is stored as 44927.5.

Excel doesn't use timezones or try to convert or encode timezone information
in any way so they aren't supported by `rust_xlsxwriter`.


## Examples

```rust
use jiff::civil::DateTime;
use jiff_to_excel::jiff_datetime_to_excel;

let dt: DateTime = "2026-01-01 12:00".parse().unwrap();
assert_eq!(jiff_to_excel::jiff_datetime_to_excel(&dt), 46023.5);
```

```rust
use jiff::civil::Date;
use jiff_to_excel::jiff_datetime_to_excel;

let d: Date = "2026-01-01".parse().unwrap();
assert_eq!(jiff_to_excel::jiff_date_to_excel(&d), 46023.0);
```

```rust
use jiff::civil::Time;
use jiff_to_excel::jiff_datetime_to_excel;

let t: Time = "12:00".parse().unwrap();
assert_eq!(jiff_to_excel::jiff_time_to_excel(&t), 0.5);
```
