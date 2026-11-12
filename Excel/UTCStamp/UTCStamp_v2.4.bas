Attribute VB_Name = "Module1"
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

' SYSTEMTIME structure - Windows GetSystemTime API only
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
    '   Examples: -8 (PST), -7 (PDT), +5.5 (IST), 0 (UTC, default).
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
