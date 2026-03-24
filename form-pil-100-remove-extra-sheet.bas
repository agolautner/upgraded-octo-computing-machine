' =============================================================
' ImportFormPil - Macro to import FormPil .xlsm files
'
' Source column mapping (Excel letters = column numbers):
'   C=3, K=11, L=12, T=20, AJ=36, AU=47
'   Z=26, AA=27, AB=28, AC=29, AD=30, AE=31
'
' Output columns (destination):
'   1  = Időszak          (time interval, source col B)
'   2  = C  - Fogadott hívások
'   3  = K  - Várakozás nélkül kiszolgált
'   4  = L  - Várakozás után kiszolgált
'   5  = T  - Munkatárs által kiszolgált
'   6  = AJ - Üdvözlő hangbemondás
'   7  = AU - Híváselvesztések teljes száma
'   8  = Szolgáltatási színvonal 30 mp (%) = AA+AC+AE
'   9  = Szolgáltatási színvonal 30 mp (db) = Z+AB+AD
'   10 = Megválaszolási arány (%) = T/C
'   11 = Vesztett hívás (db) = AU-AJ
'   12 = Vesztett hívás (%) = AU/C
' =============================================================

Option Explicit

Private Const SRC_SHEET         As String  = "General"
Private Const SHEET_PREFIX      As String  = "FP_"
Private Const BLOCK_COUNT       As Integer = 6
Private Const BLOCK_HEIGHT      As Integer = 44
Private Const DATA_START_OFFSET As Integer = 7
Private Const DATA_ROW_COUNT    As Integer = 32

' Source column numbers
Private Const SC_C  As Integer = 3
Private Const SC_K  As Integer = 11
Private Const SC_L  As Integer = 12
Private Const SC_T  As Integer = 20
Private Const SC_Z  As Integer = 26
Private Const SC_AA As Integer = 27
Private Const SC_AB As Integer = 28
Private Const SC_AC As Integer = 29
Private Const SC_AD As Integer = 30
Private Const SC_AE As Integer = 31
Private Const SC_AJ As Integer = 36
Private Const SC_AU As Integer = 47

' Destination column numbers
Private Const DC_IDOSZAK    As Integer = 1
Private Const DC_C          As Integer = 2
Private Const DC_K          As Integer = 3
Private Const DC_L          As Integer = 4
Private Const DC_T          As Integer = 5
Private Const DC_AJ         As Integer = 6
Private Const DC_AU         As Integer = 7
Private Const DC_SZV_PCT    As Integer = 8   ' Szolgáltatási színvonal 30 mp (%)
Private Const DC_SZV_DB     As Integer = 9   ' Szolgáltatási színvonal 30 mp (db)
Private Const DC_MEGV_PCT   As Integer = 10  ' Megválaszolási arány (%)
Private Const DC_VESZT_DB   As Integer = 11  ' Vesztett hívás (db)
Private Const DC_VESZT_PCT  As Integer = 12  ' Vesztett hívás (%)
Private Const TOTAL_DST_COLS As Integer = 12

' Output layout
Private Const OUT_HDR_ROWS  As Integer = 1
Private Const OUT_STAT_ROWS As Integer = 4
' Data starts at row OUT_HDR_ROWS + OUT_STAT_ROWS + 1 = 6

' Stat row labels and colours
Private Const LBL_TOTAL As String = "Összeg"
Private Const LBL_AVG   As String = "Átlag"
Private Const LBL_MAX   As String = "Maximum"
Private Const LBL_MIN   As String = "Minimum"

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

    ' Remove empty default sheets (e.g. Sheet1, Sheet2) left over from workbook creation
    Dim wsClean As Worksheet
    Dim cleanNames() As String
    Dim cleanCount As Integer
    cleanCount = 0
    ReDim cleanNames(0)
    For Each wsClean In wbDst.Worksheets
        If wsClean.Name <> "Summary" And Left(wsClean.Name, Len(SHEET_PREFIX)) <> SHEET_PREFIX Then
            If wsClean.UsedRange.Rows.Count = 1 And wsClean.UsedRange.Columns.Count = 1 _
               And IsEmpty(wsClean.UsedRange.Cells(1, 1).Value) Then
                cleanCount = cleanCount + 1
                ReDim Preserve cleanNames(cleanCount - 1)
                cleanNames(cleanCount - 1) = wsClean.Name
            End If
        End If
    Next wsClean
    For k = 0 To cleanCount - 1
        wbDst.Sheets(cleanNames(k)).Delete
    Next k

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Done! Imported " & sheetCount & " file(s).", vbInformation
End Sub

' -----------------------------------------------------------------------
' Import one source file into a new sheet.
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

    ' --- Write header row -------------------------------------------
    wsNew.Cells(1, DC_IDOSZAK).Value   = "Időszak"
    wsNew.Cells(1, DC_C).Value         = "Fogadott hívások"
    wsNew.Cells(1, DC_K).Value         = "Várakozás nélkül kiszolgált"
    wsNew.Cells(1, DC_L).Value         = "Várakozás után kiszolgált"
    wsNew.Cells(1, DC_T).Value         = "Munkatárs által kiszolgált"
    wsNew.Cells(1, DC_AJ).Value        = "Üdvözlő hangbemondás"
    wsNew.Cells(1, DC_AU).Value        = "Híváselvesztések teljes száma"
    wsNew.Cells(1, DC_SZV_PCT).Value   = "Szolgáltatási színvonal 30 mp (%)"
    wsNew.Cells(1, DC_SZV_DB).Value    = "Szolgáltatási színvonal 30 mp (db)"
    wsNew.Cells(1, DC_MEGV_PCT).Value  = "Megválaszolási arány (%)"
    wsNew.Cells(1, DC_VESZT_DB).Value  = "Vesztett hívás (db)"
    wsNew.Cells(1, DC_VESZT_PCT).Value = "Vesztett hívás (%)"

    ' --- Collect all data rows from all 6 blocks --------------------
    Dim maxRows As Integer
    maxRows = BLOCK_COUNT * DATA_ROW_COUNT
    ' Store raw source values needed for calculations
    ' Indices: 0=timeVal, 1=C, 2=K, 3=L, 4=T, 5=Z, 6=AA, 7=AB, 8=AC, 9=AD, 10=AE, 11=AJ, 12=AU
    Dim rawVals(191, 12) As Variant
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
            Dim tv As Variant
            tv = wsSrc.Cells(r, 2).Value
            If Not IsEmpty(tv) And tv <> "" Then
                rawVals(rowCount, 0)  = tv
                rawVals(rowCount, 1)  = wsSrc.Cells(r, SC_C).Value
                rawVals(rowCount, 2)  = wsSrc.Cells(r, SC_K).Value
                rawVals(rowCount, 3)  = wsSrc.Cells(r, SC_L).Value
                rawVals(rowCount, 4)  = wsSrc.Cells(r, SC_T).Value
                rawVals(rowCount, 5)  = wsSrc.Cells(r, SC_Z).Value
                rawVals(rowCount, 6)  = wsSrc.Cells(r, SC_AA).Value
                rawVals(rowCount, 7)  = wsSrc.Cells(r, SC_AB).Value
                rawVals(rowCount, 8)  = wsSrc.Cells(r, SC_AC).Value
                rawVals(rowCount, 9)  = wsSrc.Cells(r, SC_AD).Value
                rawVals(rowCount, 10) = wsSrc.Cells(r, SC_AE).Value
                rawVals(rowCount, 11) = wsSrc.Cells(r, SC_AJ).Value
                rawVals(rowCount, 12) = wsSrc.Cells(r, SC_AU).Value
                rowCount = rowCount + 1
            End If
        Next r
    Next blk

    ' --- Write stat rows (rows 2-5) ---------------------------------
    Dim dataFirstRow As Integer
    dataFirstRow = OUT_HDR_ROWS + OUT_STAT_ROWS + 1  ' = 6

    Dim statLabels(3) As String
    Dim statColors(3) As Long
    statLabels(0) = LBL_TOTAL: statColors(0) = RGB(173, 216, 230)
    statLabels(1) = LBL_AVG:   statColors(1) = RGB(144, 238, 144)
    statLabels(2) = LBL_MAX:   statColors(2) = RGB(255, 255, 153)
    statLabels(3) = LBL_MIN:   statColors(3) = RGB(216, 191, 216)

    Dim lastDataRow As Long
    lastDataRow = dataFirstRow + rowCount - 1

    Dim statRow As Integer
    For statRow = 0 To 3
        Dim outRow As Integer
        outRow = OUT_HDR_ROWS + 1 + statRow  ' rows 2,3,4,5
        wsNew.Cells(outRow, DC_IDOSZAK).Value = statLabels(statRow)

        ' Simple numeric columns: SUM/AVG/MAX/MIN
        Dim dc As Long
        For dc = DC_C To DC_AU
            Dim cs As String: cs = ColumnLetter(dc)
            Dim ra As String: ra = cs & dataFirstRow & ":" & cs & lastDataRow
            Dim f As String
            Select Case statRow
                Case 0: f = "=IFERROR(SUM("     & ra & "),"""")"
                Case 1: f = "=IFERROR(AVERAGE("  & ra & "),"""")"
                Case 2: f = "=IFERROR(MAX("     & ra & "),"""")"
                Case 3: f = "=IFERROR(MIN("     & ra & "),"""")"
            End Select
            wsNew.Cells(outRow, dc).formula = f
        Next dc

        ' Calculated columns - only MAX and MIN make sense for these
        Dim cSzv As String:   cSzv  = ColumnLetter(DC_SZV_PCT)  & dataFirstRow & ":" & ColumnLetter(DC_SZV_PCT)  & lastDataRow
        Dim cSzvDb As String: cSzvDb = ColumnLetter(DC_SZV_DB)   & dataFirstRow & ":" & ColumnLetter(DC_SZV_DB)   & lastDataRow
        Dim cMegv As String:  cMegv = ColumnLetter(DC_MEGV_PCT) & dataFirstRow & ":" & ColumnLetter(DC_MEGV_PCT) & lastDataRow
        Dim cVDb As String:   cVDb  = ColumnLetter(DC_VESZT_DB)  & dataFirstRow & ":" & ColumnLetter(DC_VESZT_DB)  & lastDataRow
        Dim cVPct As String:  cVPct = ColumnLetter(DC_VESZT_PCT) & dataFirstRow & ":" & ColumnLetter(DC_VESZT_PCT) & lastDataRow

        Select Case statRow
            Case 0  ' Összeg - only for db columns
                wsNew.Cells(outRow, DC_SZV_DB).formula   = "=IFERROR(SUM("  & cSzvDb & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_DB).formula  = "=IFERROR(SUM("  & cVDb   & "),"""")"
            Case 1  ' Átlag - all calculated cols including pct
                wsNew.Cells(outRow, DC_SZV_PCT).formula   = "=IFERROR(AVERAGE(" & cSzv   & "),"""")"
                wsNew.Cells(outRow, DC_SZV_DB).formula    = "=IFERROR(AVERAGE(" & cSzvDb & "),"""")"
                wsNew.Cells(outRow, DC_MEGV_PCT).formula  = "=IFERROR(AVERAGE(" & cMegv  & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_DB).formula  = "=IFERROR(AVERAGE(" & cVDb   & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_PCT).formula = "=IFERROR(AVERAGE(" & cVPct  & "),"""")"
            Case 2  ' Maximum - all calculated cols
                wsNew.Cells(outRow, DC_SZV_PCT).formula   = "=IFERROR(MAX(" & cSzv   & "),"""")"
                wsNew.Cells(outRow, DC_SZV_DB).formula    = "=IFERROR(MAX(" & cSzvDb & "),"""")"
                wsNew.Cells(outRow, DC_MEGV_PCT).formula  = "=IFERROR(MAX(" & cMegv  & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_DB).formula  = "=IFERROR(MAX(" & cVDb   & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_PCT).formula = "=IFERROR(MAX(" & cVPct  & "),"""")"
            Case 3  ' Minimum - all calculated cols
                wsNew.Cells(outRow, DC_SZV_PCT).formula   = "=IFERROR(MIN(" & cSzv   & "),"""")"
                wsNew.Cells(outRow, DC_SZV_DB).formula    = "=IFERROR(MIN(" & cSzvDb & "),"""")"
                wsNew.Cells(outRow, DC_MEGV_PCT).formula  = "=IFERROR(MIN(" & cMegv  & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_DB).formula  = "=IFERROR(MIN(" & cVDb   & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_PCT).formula = "=IFERROR(MIN(" & cVPct  & "),"""")"
        End Select

        ' Apply number formats to stat row
        wsNew.Cells(outRow, DC_SZV_PCT).NumberFormat  = "0.00%"
        wsNew.Cells(outRow, DC_MEGV_PCT).NumberFormat = "0.00%"
        wsNew.Cells(outRow, DC_VESZT_PCT).NumberFormat = "0.00%"

        wsNew.Rows(outRow).Interior.Color = statColors(statRow)
    Next statRow

    ' --- Write data rows (starting at row 6) ------------------------
    Dim dstRow As Long
    dstRow = dataFirstRow
    Dim ri As Integer
    For ri = 0 To rowCount - 1
        Dim vC  As Double: vC  = ToD(rawVals(ri, 1))
        Dim vK  As Double: vK  = ToD(rawVals(ri, 2))
        Dim vL  As Double: vL  = ToD(rawVals(ri, 3))
        Dim vT  As Double: vT  = ToD(rawVals(ri, 4))
        Dim vZ  As Double: vZ  = ToD(rawVals(ri, 5))
        Dim vAA As Double: vAA = ToD(rawVals(ri, 6))
        Dim vAB As Double: vAB = ToD(rawVals(ri, 7))
        Dim vAC As Double: vAC = ToD(rawVals(ri, 8))
        Dim vAD As Double: vAD = ToD(rawVals(ri, 9))
        Dim vAE As Double: vAE = ToD(rawVals(ri, 10))
        Dim vAJ As Double: vAJ = ToD(rawVals(ri, 11))
        Dim vAU As Double: vAU = ToD(rawVals(ri, 12))

        wsNew.Cells(dstRow, DC_IDOSZAK).Value = rawVals(ri, 0)
        wsNew.Cells(dstRow, DC_C).Value       = vC
        wsNew.Cells(dstRow, DC_K).Value       = vK
        wsNew.Cells(dstRow, DC_L).Value       = vL
        wsNew.Cells(dstRow, DC_T).Value       = vT
        wsNew.Cells(dstRow, DC_AJ).Value      = vAJ
        wsNew.Cells(dstRow, DC_AU).Value      = vAU
        wsNew.Cells(dstRow, DC_SZV_DB).Value  = vZ + vAB + vAD
        wsNew.Cells(dstRow, DC_VESZT_DB).Value = vAU - vAJ

        ' Percentage calculated columns
        If vC <> 0 Then
            wsNew.Cells(dstRow, DC_SZV_PCT).Value   = (vAA + vAC + vAE)
            wsNew.Cells(dstRow, DC_MEGV_PCT).Value  = vT / vC
            wsNew.Cells(dstRow, DC_VESZT_PCT).Value = vAU / vC
        End If

        ' Apply number formats
        wsNew.Cells(dstRow, DC_SZV_PCT).NumberFormat   = "0.00%"
        wsNew.Cells(dstRow, DC_MEGV_PCT).NumberFormat  = "0.00%"
        wsNew.Cells(dstRow, DC_VESZT_PCT).NumberFormat = "0.00%"

        ' Hide row if no meaningful data (check non-pct cols only)
        Dim hasData As Boolean
        hasData = (vC <> 0 Or vK <> 0 Or vL <> 0 Or vT <> 0 Or vAJ <> 0 Or vAU <> 0)
        If Not hasData Then wsNew.Rows(dstRow).Hidden = True

        dstRow = dstRow + 1
    Next ri

    Set ImportOneFile = wsNew
    Exit Function

ErrHandler:
    MsgBox "Error importing " & sFileName & ": " & Err.Description, vbExclamation
    Set ImportOneFile = Nothing
End Function

' -----------------------------------------------------------------------
' Build Summary sheet
' -----------------------------------------------------------------------
Private Sub BuildSummary(wbDst As Workbook, sheetNames() As String, sheetCount As Integer)
    Dim wsSumm As Worksheet
    Set wsSumm = wbDst.Sheets.Add(Before:=wbDst.Sheets(1))
    wsSumm.Name = "Summary"

    ' Header row
    wsSumm.Cells(1, DC_IDOSZAK).Value   = "Időszak"
    wsSumm.Cells(1, DC_C).Value         = "Fogadott hívások"
    wsSumm.Cells(1, DC_K).Value         = "Várakozás nélkül kiszolgált"
    wsSumm.Cells(1, DC_L).Value         = "Várakozás után kiszolgált"
    wsSumm.Cells(1, DC_T).Value         = "Munkatárs által kiszolgált"
    wsSumm.Cells(1, DC_AJ).Value        = "Üdvözlő hangbemondás"
    wsSumm.Cells(1, DC_AU).Value        = "Híváselvesztések teljes száma"
    wsSumm.Cells(1, DC_SZV_PCT).Value   = "Szolgáltatási színvonal 30 mp (%)"
    wsSumm.Cells(1, DC_SZV_DB).Value    = "Szolgáltatási színvonal 30 mp (db)"
    wsSumm.Cells(1, DC_MEGV_PCT).Value  = "Megválaszolási arány (%)"
    wsSumm.Cells(1, DC_VESZT_DB).Value  = "Vesztett hívás (db)"
    wsSumm.Cells(1, DC_VESZT_PCT).Value = "Vesztett hívás (%)"

    ' Collect unique time slots
    Dim col As New Collection
    Dim s As Integer, ws As Worksheet, r As Long
    On Error Resume Next
    For s = 0 To sheetCount - 1
        Set ws = wbDst.Sheets(sheetNames(s))
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For r = OUT_HDR_ROWS + OUT_STAT_ROWS + 1 To lastRow
            Dim slotVal As Variant
            slotVal = ws.Cells(r, 1).Value
            If Not IsEmpty(slotVal) And slotVal <> "" Then
                col.Add CStr(slotVal), CStr(slotVal)
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

    Dim dataFirstRow As Long
    dataFirstRow = OUT_HDR_ROWS + OUT_STAT_ROWS + 1  ' = 6

    ' Write data rows
    Dim dstRow As Long
    dstRow = dataFirstRow
    For i = 0 To n - 1
        Dim slot As String: slot = arr(i)
        wsSumm.Cells(dstRow, DC_IDOSZAK).Value = slot

        ' Accumulate raw values across all sheets for this slot
        Dim sumC  As Double: sumC  = 0
        Dim sumK  As Double: sumK  = 0
        Dim sumL  As Double: sumL  = 0
        Dim sumT  As Double: sumT  = 0
        Dim sumAJ As Double: sumAJ = 0
        Dim sumAU As Double: sumAU = 0
        Dim sumSzvDb As Double: sumSzvDb = 0
        Dim hasAnyData As Boolean: hasAnyData = False

        For s = 0 To sheetCount - 1
            Set ws = wbDst.Sheets(sheetNames(s))
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            For r = OUT_HDR_ROWS + OUT_STAT_ROWS + 1 To lastRow
                If CStr(ws.Cells(r, 1).Value) = slot Then
                    sumC    = sumC  + ToD(ws.Cells(r, DC_C).Value)
                    sumK    = sumK  + ToD(ws.Cells(r, DC_K).Value)
                    sumL    = sumL  + ToD(ws.Cells(r, DC_L).Value)
                    sumT    = sumT  + ToD(ws.Cells(r, DC_T).Value)
                    sumAJ   = sumAJ + ToD(ws.Cells(r, DC_AJ).Value)
                    sumAU   = sumAU + ToD(ws.Cells(r, DC_AU).Value)
                    sumSzvDb = sumSzvDb + ToD(ws.Cells(r, DC_SZV_DB).Value)
                    If ToD(ws.Cells(r, DC_C).Value) <> 0 Then hasAnyData = True
                    Exit For
                End If
            Next r
        Next s

        wsSumm.Cells(dstRow, DC_C).Value      = sumC
        wsSumm.Cells(dstRow, DC_K).Value      = sumK
        wsSumm.Cells(dstRow, DC_L).Value      = sumL
        wsSumm.Cells(dstRow, DC_T).Value      = sumT
        wsSumm.Cells(dstRow, DC_AJ).Value     = sumAJ
        wsSumm.Cells(dstRow, DC_AU).Value     = sumAU
        wsSumm.Cells(dstRow, DC_SZV_DB).Value = sumSzvDb
        wsSumm.Cells(dstRow, DC_VESZT_DB).Value = sumAU - sumAJ

        ' Recalculate percentages from summed values
        If sumC <> 0 Then
            wsSumm.Cells(dstRow, DC_MEGV_PCT).Value  = sumT / sumC
            wsSumm.Cells(dstRow, DC_VESZT_PCT).Value = sumAU / sumC
        End If
        ' Szolgáltatási színvonal %: sum of AA+AC+AE is not directly available,
        ' so derive as SZV_DB / C (calls served within 30s / total calls)
        If sumC <> 0 Then
            wsSumm.Cells(dstRow, DC_SZV_PCT).Value = sumSzvDb / sumC
        End If

        wsSumm.Cells(dstRow, DC_SZV_PCT).NumberFormat   = "0.00%"
        wsSumm.Cells(dstRow, DC_MEGV_PCT).NumberFormat  = "0.00%"
        wsSumm.Cells(dstRow, DC_VESZT_PCT).NumberFormat = "0.00%"

        If Not hasAnyData Then wsSumm.Rows(dstRow).Hidden = True
        dstRow = dstRow + 1
    Next i

    ' Stat rows (rows 2-5)
    Dim statLabels(3) As String
    Dim statColors(3) As Long
    statLabels(0) = LBL_TOTAL: statColors(0) = RGB(173, 216, 230)
    statLabels(1) = LBL_AVG:   statColors(1) = RGB(144, 238, 144)
    statLabels(2) = LBL_MAX:   statColors(2) = RGB(255, 255, 153)
    statLabels(3) = LBL_MIN:   statColors(3) = RGB(216, 191, 216)

    Dim lastDataRow As Long: lastDataRow = dataFirstRow + n - 1

    Dim statRow As Integer
    For statRow = 0 To 3
        Dim outRow As Long: outRow = OUT_HDR_ROWS + 1 + statRow

        wsSumm.Cells(outRow, DC_IDOSZAK).Value = statLabels(statRow)

        Dim dc As Long
        For dc = DC_C To DC_AU
            Dim cs As String: cs = ColumnLetter(dc)
            Dim ra As String: ra = cs & dataFirstRow & ":" & cs & lastDataRow
            Dim f As String
            Select Case statRow
                Case 0: f = "=IFERROR(SUM("     & ra & "),"""")"
                Case 1: f = "=IFERROR(AVERAGE("  & ra & "),"""")"
                Case 2: f = "=IFERROR(MAX("     & ra & "),"""")"
                Case 3: f = "=IFERROR(MIN("     & ra & "),"""")"
            End Select
            wsSumm.Cells(outRow, dc).formula = f
        Next dc

        ' Calculated columns in stat rows
        Dim cSzv   As String: cSzv   = ColumnLetter(DC_SZV_PCT)   & dataFirstRow & ":" & ColumnLetter(DC_SZV_PCT)   & lastDataRow
        Dim cSzvDb As String: cSzvDb = ColumnLetter(DC_SZV_DB)    & dataFirstRow & ":" & ColumnLetter(DC_SZV_DB)    & lastDataRow
        Dim cMegv  As String: cMegv  = ColumnLetter(DC_MEGV_PCT)  & dataFirstRow & ":" & ColumnLetter(DC_MEGV_PCT)  & lastDataRow
        Dim cVDb   As String: cVDb   = ColumnLetter(DC_VESZT_DB)  & dataFirstRow & ":" & ColumnLetter(DC_VESZT_DB)  & lastDataRow
        Dim cVPct  As String: cVPct  = ColumnLetter(DC_VESZT_PCT) & dataFirstRow & ":" & ColumnLetter(DC_VESZT_PCT) & lastDataRow

        Select Case statRow
            Case 0
                wsSumm.Cells(outRow, DC_SZV_DB).formula  = "=IFERROR(SUM("  & cSzvDb & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_DB).formula = "=IFERROR(SUM("  & cVDb   & "),"""")"
            Case 1
                wsSumm.Cells(outRow, DC_SZV_PCT).formula   = "=IFERROR(AVERAGE(" & cSzv   & "),"""")"
                wsSumm.Cells(outRow, DC_SZV_DB).formula    = "=IFERROR(AVERAGE(" & cSzvDb & "),"""")"
                wsSumm.Cells(outRow, DC_MEGV_PCT).formula  = "=IFERROR(AVERAGE(" & cMegv  & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_DB).formula  = "=IFERROR(AVERAGE(" & cVDb   & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_PCT).formula = "=IFERROR(AVERAGE(" & cVPct  & "),"""")"
            Case 2
                wsSumm.Cells(outRow, DC_SZV_PCT).formula   = "=IFERROR(MAX(" & cSzv   & "),"""")"
                wsSumm.Cells(outRow, DC_SZV_DB).formula    = "=IFERROR(MAX(" & cSzvDb & "),"""")"
                wsSumm.Cells(outRow, DC_MEGV_PCT).formula  = "=IFERROR(MAX(" & cMegv  & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_DB).formula  = "=IFERROR(MAX(" & cVDb   & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_PCT).formula = "=IFERROR(MAX(" & cVPct  & "),"""")"
            Case 3
                wsSumm.Cells(outRow, DC_SZV_PCT).formula   = "=IFERROR(MIN(" & cSzv   & "),"""")"
                wsSumm.Cells(outRow, DC_SZV_DB).formula    = "=IFERROR(MIN(" & cSzvDb & "),"""")"
                wsSumm.Cells(outRow, DC_MEGV_PCT).formula  = "=IFERROR(MIN(" & cMegv  & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_DB).formula  = "=IFERROR(MIN(" & cVDb   & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_PCT).formula = "=IFERROR(MIN(" & cVPct  & "),"""")"
        End Select

        wsSumm.Cells(outRow, DC_SZV_PCT).NumberFormat   = "0.00%"
        wsSumm.Cells(outRow, DC_MEGV_PCT).NumberFormat  = "0.00%"
        wsSumm.Cells(outRow, DC_VESZT_PCT).NumberFormat = "0.00%"

        wsSumm.Rows(outRow).Interior.Color = statColors(statRow)
    Next statRow
End Sub

' -----------------------------------------------------------------------
' Helpers
' -----------------------------------------------------------------------
Private Function ToD(v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then ToD = CDbl(v) Else ToD = 0
End Function

Private Function ColumnLetter(colNum As Long) As String
    Dim result As String
    Dim n As Long: n = colNum
    Do While n > 0
        Dim remainder As Long
        remainder = (n - 1) Mod 26
        result = Chr(65 + remainder) & result
        n = (n - 1 - remainder) \ 26
    Loop
    ColumnLetter = result
End Function

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
    If dashPos > 0 Then startPart = Trim(Left(s, dashPos - 1)) Else startPart = Trim(s)
    Dim colonPos As Integer: colonPos = InStr(startPart, ":")
    If colonPos > 0 Then
        StartMinutes = CLng(Left(startPart, colonPos - 1)) * 60 + CLng(Mid(startPart, colonPos + 1))
    Else
        StartMinutes = 0
    End If
End Function

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
