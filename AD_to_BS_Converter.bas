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
    Dim raw As Variant
    Dim i As Long, j As Long

    If TableLoaded Then Exit Sub

    BSRefAD = DateSerial(2025, 4, 14) ' BS 2082-01-01 = AD 2025-04-14 (Hamro Patro reference)

    raw = Array( _
        Array(30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31), _ ' 2000
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2001
        Array(31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2002
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2003
        Array(30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31), _ ' 2004
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2005
        Array(31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2006
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2007
        Array(31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31), _ ' 2008
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2009
        Array(31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2010
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2011
        Array(31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30), _ ' 2012
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2013
        Array(31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2014
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2015
        Array(31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30), _ ' 2016
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2017
        Array(31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2018
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31), _ ' 2019
        Array(31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2020
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2021
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30), _ ' 2022
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31), _ ' 2023
        Array(31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2024
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2025
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2026
        Array(30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31), _ ' 2027
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2028
        Array(31, 31, 32, 31, 32, 30, 30, 29, 30, 29, 30, 30), _ ' 2029
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2030
        Array(30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31), _ ' 2031
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2032
        Array(31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2033
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2034
        Array(30, 32, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31), _ ' 2035
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2036
        Array(31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2037
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2038
        Array(31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30), _ ' 2039
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2040
        Array(31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2041
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2042
        Array(31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30), _ ' 2043
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2044
        Array(31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2045
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2046
        Array(31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2047
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2048
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30), _ ' 2049
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31), _ ' 2050
        Array(31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2051
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2052
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30), _ ' 2053
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31), _ ' 2054
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2055
        Array(31, 31, 32, 31, 32, 30, 30, 29, 30, 29, 30, 30), _ ' 2056
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2057
        Array(30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31), _ ' 2058
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2059
        Array(31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2060
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2061
        Array(31, 31, 31, 32, 31, 31, 29, 30, 29, 30, 29, 31), _ ' 2062
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2063
        Array(31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2064
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2065
        Array(31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31), _ ' 2066
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2067
        Array(31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2068
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2069
        Array(31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30), _ ' 2070
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2071
        Array(31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2072
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _ ' 2073
        Array(31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2074
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2075
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30), _ ' 2076
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31), _ ' 2077
        Array(31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2078
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2079
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30), _ ' 2080
        Array(31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31), _ ' 2081
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2082
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _ ' 2083
        Array(31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30), _ ' 2084
        Array(31, 32, 31, 32, 30, 31, 30, 30, 29, 30, 30, 30), _ ' 2085
        Array(30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30), _ ' 2086
        Array(31, 31, 32, 31, 31, 31, 30, 30, 30, 30, 30, 30), _ ' 2087
        Array(30, 31, 32, 32, 30, 31, 30, 30, 29, 30, 30, 30), _ ' 2088
        Array(30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30), _ ' 2089
        Array(30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30), _ ' 2090
        Array(31, 31, 32, 31, 31, 31, 30, 30, 29, 30, 30, 30), _ ' 2091
        Array(30, 31, 32, 32, 31, 30, 30, 30, 29, 30, 30, 30), _ ' 2092
        Array(30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30), _ ' 2093
        Array(31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30), _ ' 2094
        Array(31, 31, 32, 31, 31, 31, 30, 29, 30, 30, 30, 30), _ ' 2095
        Array(30, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30), _ ' 2096
        Array(31, 32, 31, 31, 31, 30, 30, 30, 29, 30, 30, 30), _ ' 2097
        Array(31, 31, 32, 31, 31, 31, 29, 30, 29, 30, 29, 31), _ ' 2098
        Array(31, 31, 32, 31, 31, 31, 30, 29, 29, 30, 30, 31) _  ' 2099
    )

    For i = 0 To UBound(raw)
        For j = 1 To 12
            BSTable(BS_TABLE_MIN_YEAR + i, j) = raw(i)(j - 1)
        Next j
    Next i

    TableLoaded = True
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
' ANY macro clears Excel's entire undo history, so Ctrl+Z will NOT
' restore it - confirmed via real testing). Keep a copy of the AD dates
' elsewhere first if you need to preserve them.
' Result is forced to Text so Excel doesn't try to reinterpret it.
'=====================================================================
Public Sub ConvertSelectionToBS()
    Dim c As Range
    Dim convertedCount As Long, skippedCount As Long, alreadyCount As Long
    Dim bsText As String
    Dim skippedRows As String
    Dim msg As String
    Dim maxRow As Long
    Dim statusCell As Range
    Dim rawText As String
    Dim sampleSkip As String
    Dim parsed As Boolean

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select one or more cells containing dates first.", vbExclamation
        Exit Sub
    End If

    maxRow = -1

    For Each c In Selection.Cells
        If c.Row > maxRow Then maxRow = c.Row
        rawText = Trim(CStr(c.Value))

        If rawText Like "####-##-##*" Then
            ' Already BS output from a previous run (yyyy-mm-dd, optionally
            ' "[unverified]") - re-running the macro is safe, don't
            ' re-convert or flag it as a problem.
            alreadyCount = alreadyCount + 1
        ElseIf rawText = "" Then
            skippedCount = skippedCount + 1
            If skippedRows <> "" Then skippedRows = skippedRows & ", "
            skippedRows = skippedRows & c.Row
        Else
            parsed = False
            ' Try the underlying value first, then fall back to the
            ' displayed text (.Text) in case .Value isn't what's expected -
            ' this was the root cause of the "everything skipped" bug.
            If IsDate(c.Value) Then
                bsText = BSDateText(CDate(c.Value))
                parsed = True
            ElseIf IsDate(c.Text) Then
                bsText = BSDateText(CDate(c.Text))
                parsed = True
            ElseIf IsDate(rawText) Then
                bsText = BSDateText(CDate(rawText))
                parsed = True
            End If

            If parsed Then
                c.NumberFormat = "@" ' force text before writing
                c.Value = bsText
                convertedCount = convertedCount + 1
            Else
                skippedCount = skippedCount + 1
                If skippedRows <> "" Then skippedRows = skippedRows & ", "
                skippedRows = skippedRows & c.Row
                If sampleSkip = "" Then
                    sampleSkip = "e.g. row " & c.Row & " raw=[" & rawText & "] type=" & TypeName(c.Value)
                End If
            End If
        End If
    Next c

    ' Widen columns to fit the widest result (e.g. the "[unverified]" suffix),
    ' otherwise a column sized for a plain date can end up clipping that text.
    Selection.EntireColumn.AutoFit

    msg = "Converted: " & convertedCount & vbCrLf & _
          "Already BS (skipped, safe): " & alreadyCount & vbCrLf & _
          "Skipped (not a date): " & skippedCount
    If skippedCount > 0 Then
        msg = msg & vbCrLf & "Skipped rows: " & skippedRows
        msg = msg & vbCrLf & "Sample: " & sampleSkip
    End If

    ' Same summary also written 2 rows below the selection (matches the
    ' OnlyOffice port's behavior, since confirm()-style dialogs aren't
    ' reliable there - kept consistent here even though MsgBox works fine
    ' in real Excel). Not a fixed cell like A1 - avoids overwriting
    ' unrelated data elsewhere on the sheet.
    If maxRow >= 0 Then
        Set statusCell = Selection.Worksheet.Cells(maxRow + 2, 1)
        statusCell.Value = Replace(msg, vbCrLf, " | ")
        statusCell.Select
    End If

    MsgBox msg, vbInformation
End Sub
