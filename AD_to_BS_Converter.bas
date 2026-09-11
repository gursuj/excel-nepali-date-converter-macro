Attribute VB_Name = "AD_to_BS_Converter"
Option Explicit

'=====================================================================
' AD (Gregorian) -> BS (Bikram Sambat) date converter
'
' Data source: Project_Parva (github.com/dantwoashim/Project_Parva)
'   backend/app/calendar/constants.py  -> BS_MONTH_LENGTHS
'   backend/app/calendar/provenance.py -> confidence ranges
'
' Coverage: BS 2000-2099 (approx AD 1943-2043)
' Confidence:
'   BS 2078-2083 = "official"  (structured official source backing)
'   All other years in range  = "static_table_unverified"
'                                (in the table, but not source-verified
'                                 by the project itself)
' Output is flagged accordingly - see AppendConfidenceFlag below.
'=====================================================================

Private Const BS_REF_YEAR As Long = 2082
Private Const BS_REF_MONTH As Long = 1
Private Const BS_TABLE_MIN_YEAR As Long = 2000
Private Const BS_TABLE_MAX_YEAR As Long = 2099
Private Const BS_OFFICIAL_MIN_YEAR As Long = 2078
Private Const BS_OFFICIAL_MAX_YEAR As Long = 2083

Private BSTable(2000 To 2099, 1 To 12) As Integer
Private TableLoaded As Boolean
Private BSRefAD As Date

Private Sub LoadBSTable()
    If TableLoaded Then Exit Sub

    BSRefAD = DateSerial(2025, 4, 14) ' BS 2082-01-01 = AD 2025-04-14 (Hamro Patro reference)

    SetRow 2000, "30,32,31,32,31,30,30,30,29,30,29,31"
    SetRow 2001, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2002, "31,31,32,32,31,30,30,29,30,29,30,30"
    SetRow 2003, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2004, "30,32,31,32,31,30,30,30,29,30,29,31"
    SetRow 2005, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2006, "31,31,32,32,31,30,30,29,30,29,30,30"
    SetRow 2007, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2008, "31,31,31,32,31,31,29,30,30,29,29,31"
    SetRow 2009, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2010, "31,31,32,32,31,30,30,29,30,29,30,30"
    SetRow 2011, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2012, "31,31,31,32,31,31,29,30,30,29,30,30"
    SetRow 2013, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2014, "31,31,32,32,31,30,30,29,30,29,30,30"
    SetRow 2015, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2016, "31,31,31,32,31,31,29,30,30,29,30,30"
    SetRow 2017, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2018, "31,32,31,32,31,30,30,29,30,29,30,30"
    SetRow 2019, "31,32,31,32,31,30,30,30,29,30,29,31"
    SetRow 2020, "31,31,31,32,31,31,30,29,30,29,30,30"
    SetRow 2021, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2022, "31,32,31,32,31,30,30,30,29,29,30,30"
    SetRow 2023, "31,32,31,32,31,30,30,30,29,30,29,31"
    SetRow 2024, "31,31,31,32,31,31,30,29,30,29,30,30"
    SetRow 2025, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2026, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2027, "30,32,31,32,31,30,30,30,29,30,29,31"
    SetRow 2028, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2029, "31,31,32,31,32,30,30,29,30,29,30,30"
    SetRow 2030, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2031, "30,32,31,32,31,30,30,30,29,30,29,31"
    SetRow 2032, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2033, "31,31,32,32,31,30,30,29,30,29,30,30"
    SetRow 2034, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2035, "30,32,31,32,31,31,29,30,30,29,29,31"
    SetRow 2036, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2037, "31,31,32,32,31,30,30,29,30,29,30,30"
    SetRow 2038, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2039, "31,31,31,32,31,31,29,30,30,29,30,30"
    SetRow 2040, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2041, "31,31,32,32,31,30,30,29,30,29,30,30"
    SetRow 2042, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2043, "31,31,31,32,31,31,29,30,30,29,30,30"
    SetRow 2044, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2045, "31,32,31,32,31,30,30,29,30,29,30,30"
    SetRow 2046, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2047, "31,31,31,32,31,31,30,29,30,29,30,30"
    SetRow 2048, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2049, "31,32,31,32,31,30,30,30,29,29,30,30"
    SetRow 2050, "31,32,31,32,31,30,30,30,29,30,29,31"
    SetRow 2051, "31,31,31,32,31,31,30,29,30,29,30,30"
    SetRow 2052, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2053, "31,32,31,32,31,30,30,30,29,29,30,30"
    SetRow 2054, "31,32,31,32,31,30,30,30,29,30,29,31"
    SetRow 2055, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2056, "31,31,32,31,32,30,30,29,30,29,30,30"
    SetRow 2057, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2058, "30,32,31,32,31,30,30,30,29,30,29,31"
    SetRow 2059, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2060, "31,31,32,32,31,30,30,29,30,29,30,30"
    SetRow 2061, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2062, "31,31,31,32,31,31,29,30,29,30,29,31"
    SetRow 2063, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2064, "31,31,32,32,31,30,30,29,30,29,30,30"
    SetRow 2065, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2066, "31,31,31,32,31,31,29,30,30,29,29,31"
    SetRow 2067, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2068, "31,31,32,32,31,30,30,29,30,29,30,30"
    SetRow 2069, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2070, "31,31,31,32,31,31,29,30,30,29,30,30"
    SetRow 2071, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2072, "31,32,31,32,31,30,30,29,30,29,30,30"
    SetRow 2073, "31,32,31,32,31,30,30,30,29,29,30,31"
    SetRow 2074, "31,31,31,32,31,31,30,29,30,29,30,30"
    SetRow 2075, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2076, "31,32,31,32,31,30,30,30,29,29,30,30"
    SetRow 2077, "31,32,31,32,31,30,30,30,29,30,29,31"
    SetRow 2078, "31,31,31,32,31,31,30,29,30,29,30,30"
    SetRow 2079, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2080, "31,32,31,32,31,30,30,30,29,29,30,30"
    SetRow 2081, "31,32,31,32,31,30,30,30,29,30,29,31"
    SetRow 2082, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2083, "31,31,32,31,31,31,30,29,30,29,30,30"
    SetRow 2084, "31,31,32,31,31,30,30,30,29,30,30,30"
    SetRow 2085, "31,32,31,32,30,31,30,30,29,30,30,30"
    SetRow 2086, "30,32,31,32,31,30,30,30,29,30,30,30"
    SetRow 2087, "31,31,32,31,31,31,30,30,30,30,30,30"
    SetRow 2088, "30,31,32,32,30,31,30,30,29,30,30,30"
    SetRow 2089, "30,32,31,32,31,30,30,30,29,30,30,30"
    SetRow 2090, "30,32,31,32,31,30,30,30,29,30,30,30"
    SetRow 2091, "31,31,32,31,31,31,30,30,29,30,30,30"
    SetRow 2092, "30,31,32,32,31,30,30,30,29,30,30,30"
    SetRow 2093, "30,32,31,32,31,30,30,30,29,30,30,30"
    SetRow 2094, "31,31,32,31,31,30,30,30,29,30,30,30"
    SetRow 2095, "31,31,32,31,31,31,30,29,30,30,30,30"
    SetRow 2096, "30,31,32,32,31,30,30,29,30,29,30,30"
    SetRow 2097, "31,32,31,31,31,30,30,30,29,30,30,30"
    SetRow 2098, "31,31,32,31,31,31,29,30,29,30,29,31"
    SetRow 2099, "31,31,32,31,31,31,30,29,29,30,30,31"

    TableLoaded = True
End Sub

' Fills one year's row in BSTable from a comma-separated list of 12 month lengths.
' Using a helper per row (instead of one giant array literal) avoids VBA's line-continuation limit.
Private Sub SetRow(bsYear As Long, csv As String)
    Dim parts() As String
    Dim j As Long

    parts = Split(csv, ",")
    For j = 1 To 12
        BSTable(bsYear, j) = CInt(parts(j - 1))
    Next j
End Sub


' True if this BS year has official structured source backing (Project_Parva provenance.py)
Private Function IsOfficialYear(bsYear As Long) As Boolean
    IsOfficialYear = (bsYear >= BS_OFFICIAL_MIN_YEAR And bsYear <= BS_OFFICIAL_MAX_YEAR)
End Function

' Core conversion: Gregorian date -> BS (year, month, day)
Private Function GregorianToBS(adDate As Date, ByRef bsYear As Long, ByRef bsMonth As Long, ByRef bsDay As Long) As Boolean
    Dim remaining As Long
    Dim y As Long, m As Long, mLen As Long

    LoadBSTable

    remaining = CLng(adDate) - CLng(BSRefAD) ' can be negative
    y = BS_REF_YEAR
    m = BS_REF_MONTH

    If remaining >= 0 Then
        Do
            If y < BS_TABLE_MIN_YEAR Or y > BS_TABLE_MAX_YEAR Then
                GregorianToBS = False
                Exit Function
            End If
            mLen = BSTable(y, m)
            If remaining < mLen Then
                bsYear = y: bsMonth = m: bsDay = remaining + 1
                GregorianToBS = True
                Exit Function
            End If
            remaining = remaining - mLen
            m = m + 1
            If m > 12 Then m = 1: y = y + 1
        Loop
    Else
        Do
            m = m - 1
            If m < 1 Then m = 12: y = y - 1
            If y < BS_TABLE_MIN_YEAR Or y > BS_TABLE_MAX_YEAR Then
                GregorianToBS = False
                Exit Function
            End If
            mLen = BSTable(y, m)
            remaining = remaining + mLen
            If remaining >= 0 Then
                bsYear = y: bsMonth = m: bsDay = remaining + 1
                GregorianToBS = True
                Exit Function
            End If
        Loop
    End If
End Function

' Public: returns formatted "yyyy-mm-dd" BS string, flagged if outside the official-verified range
Public Function BSDateText(adDate As Date) As String
    Dim y As Long, m As Long, d As Long

    If Not GregorianToBS(adDate, y, m, d) Then
        BSDateText = "#OUT OF RANGE (BS " & BS_TABLE_MIN_YEAR & "-" & BS_TABLE_MAX_YEAR & " only)"
        Exit Function
    End If

    BSDateText = Format(y, "0000") & "-" & Format(m, "00") & "-" & Format(d, "00")

    If Not IsOfficialYear(y) Then
        BSDateText = BSDateText & " [unverified]"
    End If
End Function

'=====================================================================
' Entry point: select one or more cells containing Gregorian dates,
' run this macro. Each selected cell is overwritten IN PLACE with its
' BS equivalent (the original AD date is gone once this runs - running
' a macro clears Excel's undo history, so Ctrl+Z will not restore it).
' Result is forced to Text so Excel doesn't try to reinterpret it.
'=====================================================================
Public Sub ConvertSelectionToBS()
    Dim c As Range
    Dim convertedCount As Long, skippedCount As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select one or more cells containing dates first.", vbExclamation
        Exit Sub
    End If

    For Each c In Selection.Cells
        If IsDate(c.Value) And c.Value <> "" Then
            Dim bsText As String
            bsText = BSDateText(CDate(c.Value))
            c.NumberFormat = "@" ' force text before writing
            c.Value = bsText
            convertedCount = convertedCount + 1
        Else
            skippedCount = skippedCount + 1
        End If
    Next c

    ' Widen columns to fit the widest result (e.g. the "[unverified]" suffix),
    ' otherwise a column sized for a plain date can end up clipping that text.
    Selection.EntireColumn.AutoFit

    MsgBox "Converted: " & convertedCount & vbCrLf & "Skipped (not a date): " & skippedCount, vbInformation
End Sub
