' =============================================================
' ImportFormPil - Macro to import FormPil .xlsm files
'
' Source General sheet structure (per block of 44 rows):
'   block_start + 4 : header row 1 (group labels, e.g. "Kiszolgálás ideje")
'   block_start + 5 : header row 2 (sub-labels, e.g. "#", "%")
'   block_start + 6 : header row 3 (column names)
'   block_start + 7 : first of 32 data rows (col B = time interval)
'   block_start + 38: last data row
'   block_start + 39..+42: Összeg/Átlag/Maximum/Minimum (skipped)
'
' Output sheets have 3 header rows, then 4 stat rows (Összeg/Átlag/
' Maximum/Minimum) highlighted in colour, then data rows.
' Rows where all data columns are empty are hidden (but groupable).
' Number formats are copied from the source.
' =============================================================

Option Explicit

Private Const SRC_SHEET         As String  = "General"
Private Const SHEET_PREFIX      As String  = "FP_"
Private Const BLOCK_COUNT       As Integer = 6
Private Const BLOCK_HEIGHT      As Integer = 44
Private Const HDR1_OFFSET       As Integer = 4   ' group label row
Private Const HDR2_OFFSET       As Integer = 5   ' sub-label row
Private Const HDR3_OFFSET       As Integer = 6   ' column name row
Private Const DATA_START_OFFSET As Integer = 7
Private Const DATA_ROW_COUNT    As Integer = 32
Private Const TOTAL_COLS        As Integer = 252

' Highlight colours for stat rows (ARGB hex, no alpha prefix needed for RGB)
Private Const CLR_TOTAL   As Long = 16764057   ' light blue   #FFD9D9 -> use Excel RGB()
Private Const CLR_AVG     As Long = 16763904   ' light green
Private Const CLR_MAX     As Long = 16777164   ' light yellow
Private Const CLR_MIN     As Long = 16764159   ' light purple

' Stat row labels
Private Const LBL_TOTAL As String = "Összeg"
Private Const LBL_AVG   As String = "Átlag"
Private Const LBL_MAX   As String = "Maximum"
Private Const LBL_MIN   As String = "Minimum"

' Output row layout
Private Const OUT_HDR_ROWS  As Integer = 3   ' rows 1-3 are headers
Private Const OUT_STAT_ROWS As Integer = 4   ' rows 4-7 are stat rows
' Data starts at row OUT_HDR_ROWS + OUT_STAT_ROWS + 1 = 8

' -----------------------------------------------------------------------
' Entry point
' -----------------------------------------------------------------------
Sub ImportFormPilFiles()
    Dim sFolder As String
    Dim sFile   As String
    Dim wbSrc   As Workbook
    Dim wbDst   As Workbook
    Dim wsNew   As Worksheet

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder containing FormPil files"
        If .Show <> -1 Then Exit Sub
        sFolder = .SelectedItems(1)
    End With
    If Right(sFolder, 1) <> "\" Then sFolder = sFolder & "\"

    Set wbDst = ThisWorkbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Remove previously imported sheets and old Summary
    Dim ws As Worksheet
    Dim toDelete() As String
    Dim delCount As Integer
    delCount = 0
    ReDim toDelete(0)
    For Each ws In wbDst.Worksheets
        If ws.Name = "Summary" Or Left(ws.Name, Len(SHEET_PREFIX)) = SHEET_PREFIX Then
            delCount = delCount + 1
            ReDim Preserve toDelete(delCount - 1)
            toDelete(delCount - 1) = ws.Name
        End If
    Next ws
    Dim k As Integer
    For k = 0 To delCount - 1
        wbDst.Sheets(toDelete(k)).Delete
    Next k

    ' Process each .xlsm in the folder
    Dim sheetNames() As String
    Dim sheetCount   As Integer
    sheetCount = 0
    ReDim sheetNames(0)

    sFile = Dir(sFolder & "*.xlsm")
    Do While sFile <> ""
        If LCase(sFile) <> LCase(wbDst.Name) Then
            Set wbSrc = Workbooks.Open(sFolder & sFile, ReadOnly:=True, UpdateLinks:=False)
            Set wsNew = ImportOneFile(wbSrc, wbDst, sFile)
            wbSrc.Close SaveChanges:=False
            If Not wsNew Is Nothing Then
                sheetCount = sheetCount + 1
                ReDim Preserve sheetNames(sheetCount - 1)
                sheetNames(sheetCount - 1) = wsNew.Name
            End If
        End If
        sFile = Dir()
    Loop

    If sheetCount > 0 Then
        BuildSummary wbDst, sheetNames, sheetCount
    Else
        MsgBox "No FormPil files found in the selected folder.", vbInformation
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Done! Imported " & sheetCount & " file(s).", vbInformation
End Sub

' -----------------------------------------------------------------------
' Import a single source file into a new sheet in wbDst.
' Layout: 3 header rows, 4 stat rows, then data rows.
' -----------------------------------------------------------------------
Private Function ImportOneFile(wbSrc As Workbook, wbDst As Workbook, _
                               sFileName As String) As Worksheet
    On Error GoTo ErrHandler

    If Not SheetExists(wbSrc, SRC_SHEET) Then
        Set ImportOneFile = Nothing
        Exit Function
    End If

    Dim wsSrc As Worksheet
    Set wsSrc = wbSrc.Sheets(SRC_SHEET)

    Dim shName As String
    shName = Left(SHEET_PREFIX & StripExtension(sFileName), 31)

    Dim wsNew As Worksheet
    Set wsNew = wbDst.Sheets.Add(After:=wbDst.Sheets(wbDst.Sheets.Count))
    wsNew.Name = shName

    ' Source header rows are at block1_start + offsets 4, 5, 6
    Dim srcHdr1 As Integer: srcHdr1 = 1 + HDR1_OFFSET  ' row 5
    Dim srcHdr2 As Integer: srcHdr2 = 1 + HDR2_OFFSET  ' row 6
    Dim srcHdr3 As Integer: srcHdr3 = 1 + HDR3_OFFSET  ' row 7

    ' --- Write 3 header rows ----------------------------------------
    ' Col A label only on row 3; rows 1-2 col A left blank
    wsNew.Cells(1, 1).Value = ""
    wsNew.Cells(2, 1).Value = ""
    wsNew.Cells(3, 1).Value = "Időszak"

    Dim c As Long
    For c = 3 To TOTAL_COLS
        Dim dstCol As Integer
        dstCol = c - 1   ' source col 3 -> dst col 2, etc.

        Dim v1 As Variant: v1 = wsSrc.Cells(srcHdr1, c).Value
        Dim v2 As Variant: v2 = wsSrc.Cells(srcHdr2, c).Value
        Dim v3 As Variant: v3 = wsSrc.Cells(srcHdr3, c).Value

        If Not IsEmpty(v1) And v1 <> "" Then wsNew.Cells(1, dstCol).Value = v1
        If Not IsEmpty(v2) And v2 <> "" Then wsNew.Cells(2, dstCol).Value = v2
        If Not IsEmpty(v3) And v3 <> "" Then wsNew.Cells(3, dstCol).Value = v3
    Next c

    ' --- Collect all data rows from all 6 blocks --------------------
    ' Store as array: dataVals(rowIdx, colIdx) and dataFmts(colIdx)
    Dim maxDataRows As Integer
    maxDataRows = BLOCK_COUNT * DATA_ROW_COUNT
    Dim dataVals()  As Variant
    Dim dataFmts()  As String
    ReDim dataVals(maxDataRows - 1, TOTAL_COLS - 1)
    ReDim dataFmts(TOTAL_COLS - 1)

    ' Capture number formats from first data row of block 1
    Dim fmtRow As Integer
    fmtRow = 1 + DATA_START_OFFSET  ' row 8
    For c = 3 To TOTAL_COLS
        dataFmts(c - 1) = wsSrc.Cells(fmtRow, c).NumberFormat
    Next c

    Dim rowCount As Integer
    rowCount = 0
    Dim blk As Integer
    For blk = 0 To BLOCK_COUNT - 1
        Dim blkStart As Integer
        blkStart = 1 + blk * BLOCK_HEIGHT
        Dim dataStart As Integer
        dataStart = blkStart + DATA_START_OFFSET
        Dim r As Integer
        For r = dataStart To dataStart + DATA_ROW_COUNT - 1
            Dim timeVal As Variant
            timeVal = wsSrc.Cells(r, 2).Value
            If Not IsEmpty(timeVal) And timeVal <> "" Then
                dataVals(rowCount, 1) = timeVal   ' col index 1 = dst col A (0-based: col 1)
                For c = 3 To TOTAL_COLS
                    dataVals(rowCount, c - 1) = wsSrc.Cells(r, c).Value
                Next c
                rowCount = rowCount + 1
            End If
        Next r
    Next blk

    ' --- Write stat rows (rows 4-7) ---------------------------------
    Dim statLabels(3)  As String
    Dim statColors(3)  As Long
    statLabels(0) = LBL_TOTAL: statColors(0) = RGB(173, 216, 230)  ' light blue
    statLabels(1) = LBL_AVG:   statColors(1) = RGB(144, 238, 144)  ' light green
    statLabels(2) = LBL_MAX:   statColors(2) = RGB(255, 255, 153)  ' light yellow
    statLabels(3) = LBL_MIN:   statColors(3) = RGB(216, 191, 216)  ' light purple

    Dim dataFirstRow As Integer
    dataFirstRow = OUT_HDR_ROWS + OUT_STAT_ROWS + 1  ' = 8

    Dim statRow As Integer
    For statRow = 0 To 3
        Dim outRow As Integer
        outRow = OUT_HDR_ROWS + 1 + statRow  ' rows 4,5,6,7
        wsNew.Cells(outRow, 1).Value = statLabels(statRow)

        Dim lastDataRow As Long
        lastDataRow = dataFirstRow + rowCount - 1

        For c = 2 To TOTAL_COLS - 1
            Dim colStr As String
            colStr = ColumnLetter(c)
            Dim rangeAddr As String
            rangeAddr = colStr & dataFirstRow & ":" & colStr & lastDataRow
            Dim isPct As Boolean
            isPct = (InStr(dataFmts(c), "%") > 0)

            Dim formula As String
            formula = ""
            Select Case statRow
                Case 0: If Not isPct Then formula = "=IFERROR(SUM("    & rangeAddr & "),"""")"
                Case 1: If Not isPct Then formula = "=IFERROR(AVERAGE(" & rangeAddr & "),"""")"
                Case 2: formula = "=IFERROR(MAX(" & rangeAddr & "),"""")"
                Case 3: formula = "=IFERROR(MIN(" & rangeAddr & "),"""")"
            End Select

            Dim statCell As Range
            Set statCell = wsNew.Cells(outRow, c)
            If formula <> "" Then
                statCell.formula = formula
                statCell.NumberFormat = dataFmts(c)
            End If
        Next c

        ' Highlight entire stat row
        wsNew.Rows(outRow).Interior.Color = statColors(statRow)
    Next statRow

    ' --- Write data rows (starting at row 8) ------------------------
    Dim dstRow As Long
    dstRow = dataFirstRow
    Dim ri As Integer
    For ri = 0 To rowCount - 1
        Dim hasAnyData As Boolean
        hasAnyData = False
        wsNew.Cells(dstRow, 1).Value = dataVals(ri, 1)
        For c = 3 To TOTAL_COLS
            Dim dstCell As Range
            Set dstCell = wsNew.Cells(dstRow, c - 1)
            dstCell.Value = dataVals(ri, c - 1)
            dstCell.NumberFormat = dataFmts(c - 1)
            Dim isColPct As Boolean
            isColPct = (InStr(dataFmts(c - 1), "%") > 0)
            If Not isColPct Then
                If Not IsEmpty(dataVals(ri, c - 1)) And dataVals(ri, c - 1) <> "" _
                   And dataVals(ri, c - 1) <> 0 Then
                    hasAnyData = True
                End If
            End If
        Next c

        ' Hide row if all non-percentage data columns are empty/zero
        If Not hasAnyData Then
            wsNew.Rows(dstRow).Hidden = True
        End If

        dstRow = dstRow + 1
    Next ri

    Set ImportOneFile = wsNew
    Exit Function

ErrHandler:
    MsgBox "Error importing " & sFileName & ": " & Err.Description, vbExclamation
    Set ImportOneFile = Nothing
End Function

' -----------------------------------------------------------------------
' Builds the Summary sheet.
' Same layout: 3 header rows, 4 stat rows, then one row per time slot.
' -----------------------------------------------------------------------
Private Sub BuildSummary(wbDst As Workbook, sheetNames() As String, sheetCount As Integer)
    Dim wsSumm As Worksheet
    Set wsSumm = wbDst.Sheets.Add(Before:=wbDst.Sheets(1))
    wsSumm.Name = "Summary"

    Dim wsRef As Worksheet
    Set wsRef = wbDst.Sheets(sheetNames(0))

    ' Find true last column by checking all 3 header rows
    Dim lastCol As Long
    Dim lastColCheck As Long
    Dim hrCheck As Integer
    lastCol = 1
    For hrCheck = 1 To OUT_HDR_ROWS
        lastColCheck = wsRef.Cells(hrCheck, wsRef.Columns.Count).End(xlToLeft).Column
        If lastColCheck > lastCol Then lastCol = lastColCheck
    Next hrCheck

    ' Copy 3 header rows from reference sheet
    Dim c As Long, hr As Integer
    For hr = 1 To OUT_HDR_ROWS
        For c = 1 To lastCol
            wsSumm.Cells(hr, c).Value = wsRef.Cells(hr, c).Value
        Next c
    Next hr

    ' Collect unique time intervals
    Dim col As New Collection
    Dim s As Integer, ws As Worksheet, r As Long
    On Error Resume Next
    For s = 0 To sheetCount - 1
        Set ws = wbDst.Sheets(sheetNames(s))
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For r = OUT_HDR_ROWS + OUT_STAT_ROWS + 1 To lastRow
            Dim tv As Variant
            tv = ws.Cells(r, 1).Value
            If Not IsEmpty(tv) And tv <> "" Then
                col.Add CStr(tv), CStr(tv)
            End If
        Next r
    Next s
    On Error GoTo 0

    If col.Count = 0 Then Exit Sub

    Dim n As Long: n = col.Count
    Dim arr() As String: ReDim arr(n - 1)
    Dim i As Long
    For i = 0 To n - 1: arr(i) = col(i + 1): Next i
    SortByStartTime arr, n

    ' Write data rows first (so stat formulas can reference them)
    Dim dataFirstRow As Long
    dataFirstRow = OUT_HDR_ROWS + OUT_STAT_ROWS + 1  ' row 8

    Dim dstRow As Long
    dstRow = dataFirstRow
    For i = 0 To n - 1
        Dim slot As String
        slot = arr(i)
        wsSumm.Cells(dstRow, 1).Value = slot

        Dim hasAnyData As Boolean
        hasAnyData = False

        For c = 2 To lastCol
            Dim total As Double: total = 0
            Dim hasData As Boolean: hasData = False
            For s = 0 To sheetCount - 1
                Set ws = wbDst.Sheets(sheetNames(s))
                lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                For r = OUT_HDR_ROWS + OUT_STAT_ROWS + 1 To lastRow
                    If CStr(ws.Cells(r, 1).Value) = slot Then
                        Dim v As Variant
                        v = ws.Cells(r, c).Value
                        If IsNumeric(v) And Not IsEmpty(v) Then
                            total = total + CDbl(v)
                            hasData = True
                        End If
                        Exit For
                    End If
                Next r
            Next s
            Dim isSummPct As Boolean
            isSummPct = (InStr(wsRef.Cells(dataFirstRow, c).NumberFormat, "%") > 0)
            If hasData Then
                Dim summCell As Range
                Set summCell = wsSumm.Cells(dstRow, c)
                If isSummPct Then
                    ' For percentage columns store average across files
                    Dim pctCount As Integer: pctCount = 0
                    Dim pctTotal As Double: pctTotal = 0
                    Dim s2 As Integer
                    For s2 = 0 To sheetCount - 1
                        Dim ws2 As Worksheet
                        Set ws2 = wbDst.Sheets(sheetNames(s2))
                        Dim lr2 As Long
                        lr2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
                        Dim r2 As Long
                        For r2 = OUT_HDR_ROWS + OUT_STAT_ROWS + 1 To lr2
                            If CStr(ws2.Cells(r2, 1).Value) = slot Then
                                Dim pv As Variant
                                pv = ws2.Cells(r2, c).Value
                                If IsNumeric(pv) And Not IsEmpty(pv) Then
                                    pctTotal = pctTotal + CDbl(pv)
                                    pctCount = pctCount + 1
                                End If
                                Exit For
                            End If
                        Next r2
                    Next s2
                    If pctCount > 0 Then summCell.Value = pctTotal / pctCount
                Else
                    summCell.Value = total
                    hasAnyData = True
                End If
                summCell.NumberFormat = wsRef.Cells(dataFirstRow, c).NumberFormat
            End If
        Next c

        If Not hasAnyData Then wsSumm.Rows(dstRow).Hidden = True
        dstRow = dstRow + 1
    Next i

    ' Write stat rows (rows 4-7)
    Dim statLabels(3)  As String
    Dim statColors(3)  As Long
    statLabels(0) = LBL_TOTAL: statColors(0) = RGB(173, 216, 230)
    statLabels(1) = LBL_AVG:   statColors(1) = RGB(144, 238, 144)
    statLabels(2) = LBL_MAX:   statColors(2) = RGB(255, 255, 153)
    statLabels(3) = LBL_MIN:   statColors(3) = RGB(216, 191, 216)

    Dim lastDataRow As Long
    lastDataRow = dataFirstRow + n - 1

    Dim statRow As Integer
    For statRow = 0 To 3
        Dim outRow As Long
        outRow = OUT_HDR_ROWS + 1 + statRow
        wsSumm.Cells(outRow, 1).Value = statLabels(statRow)

        For c = 2 To lastCol
            Dim colStr As String
            colStr = ColumnLetter(c)
            Dim rangeAddr As String
            rangeAddr = colStr & dataFirstRow & ":" & colStr & lastDataRow
            Dim isSumPct As Boolean
            isSumPct = (InStr(wsRef.Cells(dataFirstRow, c).NumberFormat, "%") > 0)
            Dim formula As String
            formula = ""
            Select Case statRow
                Case 0: If Not isSumPct Then formula = "=IFERROR(SUM("    & rangeAddr & "),"""")"
                Case 1: If Not isSumPct Then formula = "=IFERROR(AVERAGE(" & rangeAddr & "),"""")"
                Case 2: formula = "=IFERROR(MAX(" & rangeAddr & "),"""")"
                Case 3: formula = "=IFERROR(MIN(" & rangeAddr & "),"""")"
            End Select
            Dim statCell As Range
            Set statCell = wsSumm.Cells(outRow, c)
            If formula <> "" Then
                statCell.formula = formula
                statCell.NumberFormat = wsRef.Cells(dataFirstRow, c).NumberFormat
            End If
        Next c

        wsSumm.Rows(outRow).Interior.Color = statColors(statRow)
    Next statRow
End Sub

' -----------------------------------------------------------------------
' Convert column number to letter(s), e.g. 1->"A", 28->"AB"
' -----------------------------------------------------------------------
Private Function ColumnLetter(colNum As Long) As String
    Dim result As String
    Dim n As Long
    n = colNum
    Do While n > 0
        Dim remainder As Long
        remainder = (n - 1) Mod 26
        result = Chr(65 + remainder) & result
        n = (n - 1 - remainder) \ 26
    Loop
    ColumnLetter = result
End Function

' -----------------------------------------------------------------------
' Sort time-interval strings chronologically by start time
' -----------------------------------------------------------------------
Private Sub SortByStartTime(arr() As String, n As Long)
    Dim i As Long, j As Long, tmp As String
    For i = 1 To n - 1
        tmp = arr(i)
        j = i - 1
        Do While j >= 0 And StartMinutes(arr(j)) > StartMinutes(tmp)
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
End Sub

Private Function StartMinutes(s As String) As Long
    Dim dashPos As Integer: dashPos = InStr(s, "-")
    Dim startPart As String
    If dashPos > 0 Then
        startPart = Trim(Left(s, dashPos - 1))
    Else
        startPart = Trim(s)
    End If
    Dim colonPos As Integer: colonPos = InStr(startPart, ":")
    If colonPos > 0 Then
        StartMinutes = CLng(Left(startPart, colonPos - 1)) * 60 + _
                       CLng(Mid(startPart, colonPos + 1))
    Else
        StartMinutes = 0
    End If
End Function

' -----------------------------------------------------------------------
' Helpers
' -----------------------------------------------------------------------
Private Function SheetExists(wb As Workbook, sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function StripExtension(s As String) As String
    Dim pos As Integer: pos = InStrRev(s, ".")
    If pos > 0 Then StripExtension = Left(s, pos - 1) Else StripExtension = s
End Function
