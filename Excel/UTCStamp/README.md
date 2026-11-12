# UTCStamp

Cross-platform VBA module for non-volatile UTC timestamps with seed caching.

---

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform: Windows | Mac](https://img.shields.io/badge/Platform-Windows%20%7C%20Mac-brightgreen.svg)](#compatibility)
[![Excel Windows: 2007+](https://img.shields.io/badge/Excel_Windows-2007%2B-217346.svg)](#compatibility)
[![Excel Mac: 2016+](https://img.shields.io/badge/Excel_Mac-2016%2B-217346.svg)](#compatibility)

## Overview

`UTCStamp` is an Excel UDF that returns the current UTC date/time. It recalculates only when a designated trigger range changes, and supports an optional seed value that freezes the output to a specific date. An optional UTC offset parameter converts seed values entered in any local time zone to UTC automatically.

A per-cell cache bridges the transition when a seed is cleared, preventing the cell from jumping to the current UTC time until the next trigger-range change.

The module runs on both **Windows** and **Mac**. Platform selection is handled at compile time through `#If Mac Then` conditional compilation.

| Platform | UTC Source | Mechanism |
|---|---|---|
| Windows | `GetSystemTime` (kernel32) | `SYSTEMTIME` struct → `DateSerial` + `TimeSerial` |
| Mac | `CFAbsoluteTimeGetCurrent` (CoreFoundation) | Seconds since 2001-01-01 UTC → VBA serial date |

## Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [Function Reference](#function-reference)
- [State Machine](#state-machine)
- [Behavior Details](#behavior-details)
- [Workbook Open Workflow](#workbook-open-workflow)
- [Compatibility](#compatibility)
- [License](#license)

## Installation

1. Open your workbook in Excel.
2. Press <kbd>Alt</kbd>+<kbd>F11</kbd> (Windows) or <kbd>⌥ Option</kbd>+<kbd>F11</kbd> (Mac) to open the VB Editor.
3. In the Project Explorer, right-click your project → **Import File…**
4. Select `UTCStamp_v2.4.bas`.
5. Close the VB Editor and save the workbook as `.xlsm` (macro-enabled).

Alternatively, create a new standard module and paste the contents of `UTCStamp_v2.4.bas`.

## Quick Start

**Live UTC timestamp triggered by a single cell:**

```excel
=UTCStamp(B5)
```

Returns fresh UTC each time B5 changes.

**Trigger on a range:**

```excel
=UTCStamp(A1:A10)
```

Returns fresh UTC when any cell in A1:A10 changes.

**With a seed date (UTC):**

```excel
=UTCStamp(A1:A10, C1)
```

Returns C1's date while C1 is populated. When C1 is cleared, holds the last seed value until A1:A10 changes, then switches to live UTC.

**With a seed date and UTC offset:**

```excel
=UTCStamp(A1:A10, C1, -5)
```

Seed in C1 is treated as UTC−5 (EST). The function adds 5 hours to convert to UTC.

**With a cell-referenced offset:**

```excel
=UTCStamp(A1:A10, C1, D1)
```

Reads the offset from D1, allowing each row or region to specify its own UTC offset.

**Direct UTC retrieval:**

```excel
=GetUTC()
```

Returns current UTC. Non-volatile; recalculates only on a full workbook recalculation.

## Function Reference

### UTCStamp

```
UTCStamp(rngTrigger, [vSeed], [dblOffsetHours]) → Date
```

| Parameter | Type | Req | Default | Purpose |
|---|---|---|---|---|
| `rngTrigger` | `Range` | Yes | — | Dependency range; recalculates when any cell changes |
| `vSeed` | `Variant` | No | — | Seed date; when valid, returned as UTC instead of live time |
| `dblOffsetHours` | `Double` | No | `0` | UTC offset of seed in hours (e.g., −5 EST, +5.5 IST) |

- Non-volatile (`Application.Volatile False`). Recalculates only when `rngTrigger`, `vSeed`, or `dblOffsetHours` changes.
- When called from VBA rather than a cell formula, returns `GetUTC()` directly with no caching.

### GetUTC

```
GetUTC() → Date
```

Returns the current UTC date/time at second precision. No parameters. Public — callable from worksheet formulas and VBA. Non-volatile.

## State Machine

Each calling cell tracks two pieces of state in module-level collections: the last returned date and whether a seed was active on the previous call.

| State | `vSeed` | Trigger Δ | Result | Next State |
|---|---|---|---|---|
| Seeded | Has date | — | `vSeed` − offset/24 | Seeded |
| Seeded → Live | Just cleared | — | Cached seed (UTC) | Live |
| Live | Empty | Changed | Fresh UTC | Live |
| Fresh VBA state | Has date | — | `vSeed` − offset/24 | Seeded |
| Fresh VBA state | Empty | — | Fresh UTC | Live |

When `dblOffsetHours` is 0 (default), the seed is returned unchanged. The cached seed in the Seeded → Live transition is the UTC-converted value stored during seed mode.

## Behavior Details

### Seed Mode

When `vSeed` contains a valid date, the function computes `CDate(vSeed) - dblOffsetHours / 24`, caches the result, and marks the calling cell as seeded. The UTC source is never called. This persists as long as `vSeed` remains populated.

### Transitional State

When the user clears `vSeed`, the function detects that the cell was previously seeded and returns the cached seed value one final time. Because the cached value was already converted to UTC during seed mode, no further offset adjustment occurs. The seeded flag is then cleared so the next recalculation enters live mode.

This prevents the displayed value from jumping to the current UTC time the moment the seed is removed.

### Live Mode

With no seed and no prior-seed flag, the function calls `GetUTC()` and returns the current UTC date/time. The result is cached for potential future seed transitions. `dblOffsetHours` is ignored.

### Cache Implementation

The per-cell cache uses two module-level `Collection` objects (`mcolCachedDates`, `mcolHadSeed`). `Collection` replaces `Scripting.Dictionary` for cross-platform compatibility (Dictionary is unavailable on Mac).

Three private helpers abstract collection key operations:

| Helper | Purpose |
|---|---|
| `CacheSet` | Add or overwrite a key-value pair |
| `CacheExists` | Check if a key is present |
| `CacheGet` | Retrieve value for an existing key |

### Cache Reset

The module-level collections reset to `Nothing` when VBA state resets (project reset, unhandled errors, or closing/reopening the VB Editor). After a reset, cells with a seed return the seed (offset-adjusted); cells without a seed enter live mode directly.

### RegisterClipboardFormat (Mac Stub)

`RegisterClipboardFormat` is declared `Public` for use by other modules. On Mac, a conditional `#If Mac Then` block defines a stub function returning 0 (the API's failure value). On Windows, the original `user32` declaration is unchanged.

## Workbook Open Workflow

1. **Before saving**, populate the `vSeed` cells with the current `UTCStamp` values (manually or via a `Workbook_BeforeSave` event).
2. **On next open**, the function sees the seed dates and returns them without calling the UTC source.
3. **When ready for live updates**, clear the `vSeed` cells. The cached seed value holds until the trigger range changes.

> **Important:** When `dblOffsetHours` is non-zero, ensure that seed values populated by `Workbook_BeforeSave` are stored as UTC (the raw output of `UTCStamp`). The offset is applied to the seed on each recalculation; storing an already-converted UTC value with a non-zero offset would double-apply the conversion.

## Compatibility

| Requirement | Windows | Mac |
|---|---|---|
| Excel | 2007+ (32-bit or 64-bit) | 2016+ |
| OS API | kernel32 `GetSystemTime`, user32 `RegisterClipboardFormat` | CoreFoundation `CFAbsoluteTimeGetCurrent` |
| VBA | `Option Explicit` in all modules | `Option Explicit` in all modules |

On Windows, the `#If VBA7 Then` branch compiles with `PtrSafe` (Excel 2010+); the `#Else` branch compiles without it (Excel 2007). On Mac, VBA7 (`PtrSafe`) is required; Excel 2016 for Mac is the earliest version with VBA7 support.

## License

This project is licensed under the [MIT License](LICENSE).
