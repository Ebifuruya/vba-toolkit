# UTCStamp VBA Module

Cross-platform UTC timestamp with seed caching

---

## Overview

`UTCStamp` is a non-volatile Excel UDF that returns the current UTC date/time. It recalculates only when a designated trigger range changes, and supports an optional seed value that freezes the output to a specific date. An optional UTC offset parameter allows seed values entered in any local time zone to be converted to UTC automatically.

A per-cell cache bridges the transition when a seed is cleared, preventing the cell from jumping to the current UTC time until the next trigger-range change.

The module runs on both Windows and Mac. On Windows, UTC is obtained via the `GetSystemTime` kernel32 API. On Mac, UTC is obtained via `CFAbsoluteTimeGetCurrent` from the CoreFoundation framework. Platform selection is handled at compile time through `#If Mac Then` conditional compilation blocks. Cache state is stored in VBA `Collection` objects, which are available on both platforms.

## Function Signatures

### UTCStamp

```
UTCStamp(rngTrigger, [vSeed], [dblOffsetHours]) → Date
```

| Parameter | Type | Req | Default | Purpose |
|---|---|---|---|---|
| `rngTrigger` | `Range` | Yes | — | Dependency range. Recalculates when any cell in this range changes. Values not read. |
| `vSeed` | `Variant` | No | — | Seed date. When populated with a valid date, converted to UTC and returned instead of calling the UTC source. |
| `dblOffsetHours` | `Double` | No | `0` | UTC offset of the seed value in hours (e.g., −5 EST, +0 UTC, +5.5 IST). Applied only when `vSeed` is valid. Ignored in live mode. |

### GetUTC

```
GetUTC() → Date
```

Returns current UTC date/time at second precision. No parameters. Public — callable from worksheet formulas and VBA. Non-volatile; recalculates only on a full workbook recalculation.

| Platform | Source | Mechanism |
|---|---|---|
| Windows | `GetSystemTime` (kernel32) | Populates `SYSTEMTIME` struct; year/month/day/hour/minute/second converted via `DateSerial` + `TimeSerial` |
| Mac | `CFAbsoluteTimeGetCurrent` (CoreFoundation) | Returns seconds since 2001-01-01 00:00:00 UTC as `Double`; `Int()` truncates to whole seconds; converted to VBA serial: `CDate(36892# + Int(dblCFTime) / 86400#)` where 36892 = VBA serial for 2001-01-01 |

The `SYSTEMTIME` type and `GetSystemTime` declaration compile only on Windows. The `CFAbsoluteTimeGetCurrent` declaration compiles only on Mac.

## State Machine

Each calling cell tracks two pieces of state in module-level collections: the last returned date and whether a seed was active on the previous call.

| State | `vSeed` | Trigger Δ | Result | Next state |
|---|---|---|---|---|
| Seeded | Has date | — | `vSeed` − ofs/24 | Seeded |
| Seeded → Live | Just cleared | — | Cached seed (UTC) | Live |
| Live | Empty | Changed | Fresh UTC | Live |
| Open (fresh VBA state) | Has date | — | `vSeed` − ofs/24 | Seeded |
| Open (fresh VBA state) | Empty | — | Fresh UTC | Live |

ofs = `dblOffsetHours`. When 0 (default), the seed is returned unchanged. The cached seed in the Seeded → Live row is the UTC-converted value stored during seed mode.

## Behavior Details

### Non-Volatile Recalculation

`Application.Volatile False` is set explicitly. Excel skips this function during workbook open and unrelated recalculations. The only recalc triggers are changes to `rngTrigger` cells, the `vSeed` cell, or the `dblOffsetHours` cell.

### Seed Mode

When `vSeed` contains a valid date, the function converts it to UTC by computing `CDate(vSeed) - dblOffsetHours / 24`, caches the result, and marks the calling cell as seeded. When `dblOffsetHours` is 0 (the default), the seed is returned unchanged. The UTC source is never called. This persists across recalculations as long as `vSeed` remains populated.

### Transitional State

When the user clears `vSeed`, the function detects that the cell was previously seeded and returns the cached seed value one final time. Because the cached value was already converted to UTC during seed mode, no further offset adjustment occurs. It then clears the seeded flag so the next recalculation enters live mode.

This prevents the displayed value from jumping to the current UTC time the moment the seed is removed.

### Live Mode

With no seed and no prior-seed flag, the function calls `GetUTC` and returns the current UTC date/time. The result is cached for potential future seed transitions. `dblOffsetHours` is ignored in live mode.

### VBA Calls

When called from VBA code rather than a cell formula, the function calls `GetUTC` directly with no caching.

### Cache Implementation

The per-cell cache uses two module-level `Collection` objects (`mcolCachedDates`, `mcolHadSeed`). `Collection` replaces `Scripting.Dictionary`, which is unavailable on Mac.

Three private helpers abstract Collection key operations:

| Helper | Purpose |
|---|---|
| `CacheSet(col, strKey, vValue)` | Add or overwrite. Removes key first (error suppressed if absent), then adds. |
| `CacheExists(col, strKey) → Boolean` | `True` if key present. Uses `On Error Resume Next` to attempt retrieval. |
| `CacheGet(col, strKey) → Variant` | Returns value for existing key. Caller must verify existence first via `CacheExists`. |

### Cache Reset

The module-level collections (`mcolCachedDates`, `mcolHadSeed`) reset to `Nothing` when VBA state resets. This occurs on project reset, unhandled errors, or closing and reopening the VB Editor. After a reset, cells with a seed return the seed (offset-adjusted); cells without a seed go directly to live mode.

### RegisterClipboardFormat (Mac Stub)

`RegisterClipboardFormat` is declared `Public` for use by other modules. On Mac, the `user32` API is unavailable. A conditional `#If Mac Then` block defines a stub function with the same signature that returns 0 (the API's failure value). On Windows, the original `Declare` is unchanged.

## Usage Examples

### Basic: trigger on a single cell

```
=UTCStamp(B5)
```

Returns fresh UTC each time B5 changes.

### Trigger on a range

```
=UTCStamp(A1:A10)
```

Returns fresh UTC when any cell in A1:A10 changes.

### With seed date (UTC)

```
=UTCStamp(A1:A10, C1)
```

Returns C1's date while C1 is populated. When C1 is cleared, holds the last seed value until A1:A10 changes, then switches to live UTC. Offset defaults to 0 — seed treated as UTC.

### With seed date and fixed offset

```
=UTCStamp(A1:A10, C1, -5)
```

Seed in C1 is treated as UTC−5. The function adds 5 hours to convert to UTC before returning and caching.

### With seed date and cell-referenced offset

```
=UTCStamp(A1:A10, C1, D1)
```

Reads the offset from D1. Allows each row, region, or user to specify its own UTC offset.

### Direct UTC retrieval

```
=GetUTC()
```

Returns current UTC. Non-volatile; recalculates only when Excel performs a full recalculation.

## Workbook Open Workflow

1. Before saving, populate the `vSeed` cells with the current UTCStamp values (manually or via a `Workbook_BeforeSave` event).
2. On next open, the function sees the seed dates and returns them without calling the UTC source.
3. When ready for live updates, clear the `vSeed` cells. The cached seed value holds until the trigger range changes.

When `dblOffsetHours` is non-zero, ensure that seed values populated by `Workbook_BeforeSave` are stored as UTC (the raw output of `UTCStamp`). The offset is applied to the seed on each recalculation; storing an already-converted UTC value with a non-zero offset would double-apply the conversion.

## Module Code

The code below covers UTCStamp, GetUTC, and supporting infrastructure. Place in a standard module.

```vb
Option Explicit

' =================== API Declarations Section ===================
' Mac:     CFAbsoluteTimeGetCurrent from CoreFoundation (UTC)
' Windows: GetSystemTime from kernel32 (UTC)
' Windows: RegisterClipboardFormat from user32 (Mac stub below)
#If Mac Then
    Private Declare PtrSafe Function CFAbsoluteTimeGetCurrent _
        Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" _
        () As Double
#ElseIf VBA7 Then
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" _
        (lpSystemTime As SYSTEMTIME)
    Public Declare PtrSafe Function RegisterClipboardFormat _
        Lib "user32" Alias "RegisterClipboardFormatA" _
        (ByVal lpString As String) As Long
#Else
    Private Declare Sub GetSystemTime Lib "kernel32" _
        (lpSystemTime As SYSTEMTIME)
    Public Declare Function RegisterClipboardFormat _
        Lib "user32" Alias "RegisterClipboardFormatA" _
        (ByVal lpString As String) As Long
#End If
' ================================================================

' SYSTEMTIME structure — Windows GetSystemTime API only
#If Mac Then
    ' Not needed: Mac uses CFAbsoluteTimeGetCurrent (returns Double)
#Else
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
#End If

' Per-cell cache for UTCStamp seed tracking.
' VBA Collection replaces Scripting.Dictionary for
' cross-platform compatibility (Dictionary unavailable on Mac).
Private mcolCachedDates As Collection
Private mcolHadSeed As Collection

' =================== Cache Helpers ============================

Private Sub EnsureCache()
    If mcolCachedDates Is Nothing Then
        Set mcolCachedDates = New Collection
    End If
    If mcolHadSeed Is Nothing Then
        Set mcolHadSeed = New Collection
    End If
End Sub

Private Sub CacheSet(col As Collection, strKey As String, _
                     vValue As Variant)
    ' Add or overwrite a key-value pair.
    ' Collection has no native overwrite; remove first.
    On Error Resume Next
    col.Remove strKey
    On Error GoTo 0
    col.Add vValue, strKey
End Sub

Private Function CacheExists(col As Collection, _
                              strKey As String) As Boolean
    ' Returns True when strKey is present in the collection.
    Dim vTemp As Variant
    On Error Resume Next
    vTemp = col.Item(strKey)
    CacheExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function CacheGet(col As Collection, _
                           strKey As String) As Variant
    ' Caller must verify key exists via CacheExists first.
    CacheGet = col.Item(strKey)
End Function

' =================== UTC Time ================================

Public Function GetUTC() As Date
    ' Returns current UTC date/time at second precision.
    ' Usable from worksheet formulas and VBA.
    ' Mac:     CoreFoundation CFAbsoluteTimeGetCurrent
    ' Windows: kernel32 GetSystemTime
    #If Mac Then
        Dim dblCFTime As Double
        dblCFTime = CFAbsoluteTimeGetCurrent()
        ' CFAbsoluteTime epoch: 2001-01-01 00:00:00 UTC
        ' VBA serial for 2001-01-01 = 36892
        ' Int() truncates sub-second remainder
        GetUTC = CDate(36892# + Int(dblCFTime) / 86400#)
    #Else
        Dim st As SYSTEMTIME
        GetSystemTime st
        GetUTC = DateSerial(st.wYear, st.wMonth, st.wDay) + _
            TimeSerial(st.wHour, st.wMinute, st.wSecond)
    #End If
End Function

' =================== Public Functions =========================

Public Function UTCStamp(rngTrigger As Range, _
                         Optional vSeed As Variant, _
                         Optional dblOffsetHours As Double = 0) As Date
    ' dblOffsetHours: UTC offset of the seed value in hours.
    '   Examples: -5 (EST), +0 (UTC), +5.5 (IST), +9 (JST).
    '   Applied only when vSeed is a valid date. Ignored in live mode.
    '   Existing formulas without this argument are unaffected.
    Application.Volatile False

    Dim strKey As String
    Dim blnFromCell As Boolean
    Dim blnSeedValid As Boolean
    Dim blnWasSeeded As Boolean

    EnsureCache

    blnFromCell = (TypeName(Application.Caller) = "Range")

    If blnFromCell Then
        strKey = Application.Caller.Address(External:=True)

        blnSeedValid = False
        If Not IsMissing(vSeed) Then
            If Not IsEmpty(vSeed) Then
                If IsDate(vSeed) Then
                    blnSeedValid = True
                End If
            End If
        End If

        If blnSeedValid Then
            ' Seed mode: convert seed from local to UTC, cache, mark seeded
            UTCStamp = CDate(vSeed) - dblOffsetHours / 24#
            CacheSet mcolCachedDates, strKey, UTCStamp
            CacheSet mcolHadSeed, strKey, True
        Else
            ' vSeed empty or missing — check prior seed state
            blnWasSeeded = False
            If CacheExists(mcolHadSeed, strKey) Then
                blnWasSeeded = CBool(CacheGet(mcolHadSeed, strKey))
            End If

            If blnWasSeeded Then
                ' Transitional: seed just cleared, return cached value
                UTCStamp = CDate(CacheGet(mcolCachedDates, strKey))
                CacheSet mcolHadSeed, strKey, False
            Else
                ' Live mode: fetch fresh UTC
                UTCStamp = GetUTC()
                CacheSet mcolCachedDates, strKey, UTCStamp
            End If
        End If
    Else
        ' Called from VBA — always fresh UTC
        UTCStamp = GetUTC()
    End If
End Function

' =================== Mac Stubs ================================

#If Mac Then
Public Function RegisterClipboardFormat( _
        ByVal lpString As String) As Long
    ' Stub: user32 API not available on macOS.  Returns 0.
    RegisterClipboardFormat = 0
End Function
#End If
```

## Requirements

| Req | Windows | Mac |
|---|---|---|
| Excel | 2007 or later (32-bit or 64-bit) | 2016 or later |
| OS API | kernel32 `GetSystemTime`, user32 `RegisterClipboardFormat` | CoreFoundation `CFAbsoluteTimeGetCurrent` |
| VBA | `Option Explicit` in all modules | `Option Explicit` in all modules |

On Windows, the `#If VBA7 Then` branch compiles with `PtrSafe` (Excel 2010+); the `#Else` branch compiles without it (Excel 2007 and earlier). On Mac, VBA7 (`PtrSafe`) is required; Excel 2016 for Mac is the earliest version with VBA7 support.
