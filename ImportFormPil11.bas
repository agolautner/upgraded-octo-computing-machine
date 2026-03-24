Attribute VB_Name = "ImportFormPil"
Option Explicit

Private Const SRC_SHEET         As String  = "General"
Private Const SHEET_PREFIX      As String  = "FP_"
Private Const BLOCK_COUNT       As Integer = 6
Private Const BLOCK_HEIGHT      As Integer = 44
Private Const DATA_START_OFFSET As Integer = 7
Private Const DATA_ROW_COUNT    As Integer = 32

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

Private Const DC_IDOSZAK   As Integer = 1
Private Const DC_C         As Integer = 2
Private Const DC_K         As Integer = 3
Private Const DC_L         As Integer = 4
Private Const DC_T         As Integer = 5
Private Const DC_AJ        As Integer = 6
Private Const DC_AU        As Integer = 7
Private Const DC_SZV_PCT   As Integer = 8
Private Const DC_SZV_DB    As Integer = 9
Private Const DC_MEGV_PCT  As Integer = 10
Private Const DC_VESZT_DB  As Integer = 11
Private Const DC_VESZT_PCT As Integer = 12

Private Const OUT_HDR_ROWS  As Integer = 1
Private Const OUT_STAT_ROWS As Integer = 4

' -----------------------------------------------------------------------
' Header strings via Chr() to avoid encoding issues
' -----------------------------------------------------------------------
Private Function H_Idoszak() As String
    H_Idoszak = "Id" & Chr(337) & "szak"
End Function
Private Function H_Fogadott() As String
    H_Fogadott = "Fogadott h" & Chr(237) & "v" & Chr(225) & "sok"
End Function
Private Function H_VarakNelkul() As String
    H_VarakNelkul = "V" & Chr(225) & "rakoz" & Chr(225) & "s n" & Chr(233) & "lk" & Chr(252) & "l kiszolg" & Chr(225) & "lt"
End Function
Private Function H_VarakUtan() As String
    H_VarakUtan = "V" & Chr(225) & "rakoz" & Chr(225) & "s ut" & Chr(225) & "n kiszolg" & Chr(225) & "lt"
End Function
Private Function H_Munkatars() As String
    H_Munkatars = "Munkat" & Chr(225) & "rs " & Chr(225) & "ltal kiszolg" & Chr(225) & "lt"
End Function
Private Function H_Udvozlo() As String
    H_Udvozlo = Chr(220) & "dv" & Chr(246) & "zl" & Chr(337) & " hangbemond" & Chr(225) & "s"
End Function
Private Function H_Hivashiv() As String
    H_Hivashiv = "H" & Chr(237) & "v" & Chr(225) & "selveszt" & Chr(233) & "sek teljes sz" & Chr(225) & "ma"
End Function
Private Function H_SzvPct() As String
    H_SzvPct = "Szolg" & Chr(225) & "ltat" & Chr(225) & "si sz" & Chr(237) & "nvonal 30 mp (%)"
End Function
Private Function H_SzvDb() As String
    H_SzvDb = "Szolg" & Chr(225) & "ltat" & Chr(225) & "si sz" & Chr(237) & "nvonal 30 mp (db)"
End Function
Private Function H_MegvPct() As String
    H_MegvPct = "Megv" & Chr(225) & "laszol" & Chr(225) & "si ar" & Chr(225) & "ny (%)"
End Function
Private Function H_VesztDb() As String
    H_VesztDb = "Vesztett h" & Chr(237) & "v" & Chr(225) & "s (db)"
End Function
Private Function H_VesztPct() As String
    H_VesztPct = "Vesztett h" & Chr(237) & "v" & Chr(225) & "s (%)"
End Function
Private Function H_Osszeg() As String
    H_Osszeg = Chr(214) & "sszeg"
End Function
Private Function H_Atlag() As String
    H_Atlag = Chr(193) & "tlag"
End Function

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

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Done! Imported " & sheetCount & " file(s).", vbInformation
End Sub

' -----------------------------------------------------------------------
' Import one source file into a new sheet
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

    ' Write header row
    wsNew.Cells(1, DC_IDOSZAK).Value   = H_Idoszak()
    wsNew.Cells(1, DC_C).Value         = H_Fogadott()
    wsNew.Cells(1, DC_K).Value         = H_VarakNelkul()
    wsNew.Cells(1, DC_L).Value         = H_VarakUtan()
    wsNew.Cells(1, DC_T).Value         = H_Munkatars()
    wsNew.Cells(1, DC_AJ).Value        = H_Udvozlo()
    wsNew.Cells(1, DC_AU).Value        = H_Hivashiv()
    wsNew.Cells(1, DC_SZV_PCT).Value   = H_SzvPct()
    wsNew.Cells(1, DC_SZV_DB).Value    = H_SzvDb()
    wsNew.Cells(1, DC_MEGV_PCT).Value  = H_MegvPct()
    wsNew.Cells(1, DC_VESZT_DB).Value  = H_VesztDb()
    wsNew.Cells(1, DC_VESZT_PCT).Value = H_VesztPct()

    ' Collect data rows from all 6 blocks
    Dim dataFirstRow As Long
    dataFirstRow = OUT_HDR_ROWS + OUT_STAT_ROWS + 1

    Dim dstRow As Long
    dstRow = dataFirstRow

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
                Dim vC  As Double: vC  = ToD(wsSrc.Cells(r, SC_C).Value)
                Dim vK  As Double: vK  = ToD(wsSrc.Cells(r, SC_K).Value)
                Dim vL  As Double: vL  = ToD(wsSrc.Cells(r, SC_L).Value)
                Dim vT  As Double: vT  = ToD(wsSrc.Cells(r, SC_T).Value)
                Dim vZ  As Double: vZ  = ToD(wsSrc.Cells(r, SC_Z).Value)
                Dim vAA As Double: vAA = ToD(wsSrc.Cells(r, SC_AA).Value)
                Dim vAB As Double: vAB = ToD(wsSrc.Cells(r, SC_AB).Value)
                Dim vAC As Double: vAC = ToD(wsSrc.Cells(r, SC_AC).Value)
                Dim vAD As Double: vAD = ToD(wsSrc.Cells(r, SC_AD).Value)
                Dim vAE As Double: vAE = ToD(wsSrc.Cells(r, SC_AE).Value)
                Dim vAJ As Double: vAJ = ToD(wsSrc.Cells(r, SC_AJ).Value)
                Dim vAU As Double: vAU = ToD(wsSrc.Cells(r, SC_AU).Value)

                wsNew.Cells(dstRow, DC_IDOSZAK).Value  = tv
                wsNew.Cells(dstRow, DC_C).Value        = vC
                wsNew.Cells(dstRow, DC_K).Value        = vK
                wsNew.Cells(dstRow, DC_L).Value        = vL
                wsNew.Cells(dstRow, DC_T).Value        = vT
                wsNew.Cells(dstRow, DC_AJ).Value       = vAJ
                wsNew.Cells(dstRow, DC_AU).Value       = vAU
                wsNew.Cells(dstRow, DC_SZV_DB).Value   = vZ + vAB + vAD
                wsNew.Cells(dstRow, DC_VESZT_DB).Value = vAU - vAJ

                If vC <> 0 Then
                    wsNew.Cells(dstRow, DC_SZV_PCT).Value   = vAA + vAC + vAE
                    wsNew.Cells(dstRow, DC_MEGV_PCT).Value  = vT / vC
                    wsNew.Cells(dstRow, DC_VESZT_PCT).Value = vAU / vC
                End If

                wsNew.Cells(dstRow, DC_SZV_PCT).NumberFormat   = "0.00%"
                wsNew.Cells(dstRow, DC_MEGV_PCT).NumberFormat  = "0.00%"
                wsNew.Cells(dstRow, DC_VESZT_PCT).NumberFormat = "0.00%"

                Dim hasData As Boolean
                hasData = (vC <> 0 Or vK <> 0 Or vL <> 0 Or vT <> 0 Or vAJ <> 0 Or vAU <> 0)
                If Not hasData Then wsNew.Rows(dstRow).Hidden = True

                dstRow = dstRow + 1
            End If
        Next r
    Next blk

    Dim lastDataRow As Long
    lastDataRow = dstRow - 1

    ' Write stat rows (rows 2-5)
    Dim statColors(3) As Long
    statColors(0) = RGB(173, 216, 230)
    statColors(1) = RGB(144, 238, 144)
    statColors(2) = RGB(255, 255, 153)
    statColors(3) = RGB(216, 191, 216)

    Dim statLabels(3) As String
    statLabels(0) = H_Osszeg()
    statLabels(1) = H_Atlag()
    statLabels(2) = "Maximum"
    statLabels(3) = "Minimum"

    Dim statRow As Integer
    For statRow = 0 To 3
        Dim outRow As Integer
        outRow = OUT_HDR_ROWS + 1 + statRow

        wsNew.Cells(outRow, DC_IDOSZAK).Value = statLabels(statRow)

        Dim dc As Long
        For dc = DC_C To DC_AU
            Dim cs As String: cs = ColumnLetter(dc)
            Dim ra As String: ra = cs & dataFirstRow & ":" & cs & lastDataRow
            Dim f As String
            Select Case statRow
                Case 0: f = "=IFERROR(SUM("     & ra & "),"""")"
                Case 1: f = "=IFERROR(AVERAGE(" & ra & "),"""")"
                Case 2: f = "=IFERROR(MAX("     & ra & "),"""")"
                Case 3: f = "=IFERROR(MIN("     & ra & "),"""")"
            End Select
            wsNew.Cells(outRow, dc).Formula = f
        Next dc

        Dim rSzvPct   As String: rSzvPct   = ColumnLetter(DC_SZV_PCT)   & dataFirstRow & ":" & ColumnLetter(DC_SZV_PCT)   & lastDataRow
        Dim rSzvDb    As String: rSzvDb    = ColumnLetter(DC_SZV_DB)    & dataFirstRow & ":" & ColumnLetter(DC_SZV_DB)    & lastDataRow
        Dim rMegvPct  As String: rMegvPct  = ColumnLetter(DC_MEGV_PCT)  & dataFirstRow & ":" & ColumnLetter(DC_MEGV_PCT)  & lastDataRow
        Dim rVesztDb  As String: rVesztDb  = ColumnLetter(DC_VESZT_DB)  & dataFirstRow & ":" & ColumnLetter(DC_VESZT_DB)  & lastDataRow
        Dim rVesztPct As String: rVesztPct = ColumnLetter(DC_VESZT_PCT) & dataFirstRow & ":" & ColumnLetter(DC_VESZT_PCT) & lastDataRow

        Select Case statRow
            Case 0  ' Sum: db columns only, skip pct
                wsNew.Cells(outRow, DC_SZV_DB).Formula    = "=IFERROR(SUM(" & rSzvDb   & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_DB).Formula  = "=IFERROR(SUM(" & rVesztDb & "),"""")"
            Case 1  ' Average: all columns including pct
                wsNew.Cells(outRow, DC_SZV_PCT).Formula   = "=IFERROR(AVERAGE(" & rSzvPct   & "),"""")"
                wsNew.Cells(outRow, DC_SZV_DB).Formula    = "=IFERROR(AVERAGE(" & rSzvDb    & "),"""")"
                wsNew.Cells(outRow, DC_MEGV_PCT).Formula  = "=IFERROR(AVERAGE(" & rMegvPct  & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_DB).Formula  = "=IFERROR(AVERAGE(" & rVesztDb  & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_PCT).Formula = "=IFERROR(AVERAGE(" & rVesztPct & "),"""")"
            Case 2  ' Max
                wsNew.Cells(outRow, DC_SZV_PCT).Formula   = "=IFERROR(MAX(" & rSzvPct   & "),"""")"
                wsNew.Cells(outRow, DC_SZV_DB).Formula    = "=IFERROR(MAX(" & rSzvDb    & "),"""")"
                wsNew.Cells(outRow, DC_MEGV_PCT).Formula  = "=IFERROR(MAX(" & rMegvPct  & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_DB).Formula  = "=IFERROR(MAX(" & rVesztDb  & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_PCT).Formula = "=IFERROR(MAX(" & rVesztPct & "),"""")"
            Case 3  ' Min
                wsNew.Cells(outRow, DC_SZV_PCT).Formula   = "=IFERROR(MIN(" & rSzvPct   & "),"""")"
                wsNew.Cells(outRow, DC_SZV_DB).Formula    = "=IFERROR(MIN(" & rSzvDb    & "),"""")"
                wsNew.Cells(outRow, DC_MEGV_PCT).Formula  = "=IFERROR(MIN(" & rMegvPct  & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_DB).Formula  = "=IFERROR(MIN(" & rVesztDb  & "),"""")"
                wsNew.Cells(outRow, DC_VESZT_PCT).Formula = "=IFERROR(MIN(" & rVesztPct & "),"""")"
        End Select

        wsNew.Cells(outRow, DC_SZV_PCT).NumberFormat   = "0.00%"
        wsNew.Cells(outRow, DC_MEGV_PCT).NumberFormat  = "0.00%"
        wsNew.Cells(outRow, DC_VESZT_PCT).NumberFormat = "0.00%"

        wsNew.Rows(outRow).Interior.Color = statColors(statRow)
    Next statRow

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

    wsSumm.Cells(1, DC_IDOSZAK).Value   = H_Idoszak()
    wsSumm.Cells(1, DC_C).Value         = H_Fogadott()
    wsSumm.Cells(1, DC_K).Value         = H_VarakNelkul()
    wsSumm.Cells(1, DC_L).Value         = H_VarakUtan()
    wsSumm.Cells(1, DC_T).Value         = H_Munkatars()
    wsSumm.Cells(1, DC_AJ).Value        = H_Udvozlo()
    wsSumm.Cells(1, DC_AU).Value        = H_Hivashiv()
    wsSumm.Cells(1, DC_SZV_PCT).Value   = H_SzvPct()
    wsSumm.Cells(1, DC_SZV_DB).Value    = H_SzvDb()
    wsSumm.Cells(1, DC_MEGV_PCT).Value  = H_MegvPct()
    wsSumm.Cells(1, DC_VESZT_DB).Value  = H_VesztDb()
    wsSumm.Cells(1, DC_VESZT_PCT).Value = H_VesztPct()

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
    dataFirstRow = OUT_HDR_ROWS + OUT_STAT_ROWS + 1

    Dim dstRow As Long
    dstRow = dataFirstRow
    For i = 0 To n - 1
        Dim slot As String: slot = arr(i)
        wsSumm.Cells(dstRow, DC_IDOSZAK).Value = slot

        Dim sumC     As Double: sumC     = 0
        Dim sumK     As Double: sumK     = 0
        Dim sumL     As Double: sumL     = 0
        Dim sumT     As Double: sumT     = 0
        Dim sumAJ    As Double: sumAJ    = 0
        Dim sumAU    As Double: sumAU    = 0
        Dim sumSzvDb As Double: sumSzvDb = 0
        Dim hasAnyData As Boolean: hasAnyData = False

        For s = 0 To sheetCount - 1
            Set ws = wbDst.Sheets(sheetNames(s))
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            For r = OUT_HDR_ROWS + OUT_STAT_ROWS + 1 To lastRow
                If CStr(ws.Cells(r, 1).Value) = slot Then
                    sumC     = sumC     + ToD(ws.Cells(r, DC_C).Value)
                    sumK     = sumK     + ToD(ws.Cells(r, DC_K).Value)
                    sumL     = sumL     + ToD(ws.Cells(r, DC_L).Value)
                    sumT     = sumT     + ToD(ws.Cells(r, DC_T).Value)
                    sumAJ    = sumAJ    + ToD(ws.Cells(r, DC_AJ).Value)
                    sumAU    = sumAU    + ToD(ws.Cells(r, DC_AU).Value)
                    sumSzvDb = sumSzvDb + ToD(ws.Cells(r, DC_SZV_DB).Value)
                    If ToD(ws.Cells(r, DC_C).Value) <> 0 Then hasAnyData = True
                    Exit For
                End If
            Next r
        Next s

        wsSumm.Cells(dstRow, DC_C).Value        = sumC
        wsSumm.Cells(dstRow, DC_K).Value        = sumK
        wsSumm.Cells(dstRow, DC_L).Value        = sumL
        wsSumm.Cells(dstRow, DC_T).Value        = sumT
        wsSumm.Cells(dstRow, DC_AJ).Value       = sumAJ
        wsSumm.Cells(dstRow, DC_AU).Value       = sumAU
        wsSumm.Cells(dstRow, DC_SZV_DB).Value   = sumSzvDb
        wsSumm.Cells(dstRow, DC_VESZT_DB).Value = sumAU - sumAJ

        If sumC <> 0 Then
            wsSumm.Cells(dstRow, DC_SZV_PCT).Value   = sumSzvDb / sumC
            wsSumm.Cells(dstRow, DC_MEGV_PCT).Value  = sumT / sumC
            wsSumm.Cells(dstRow, DC_VESZT_PCT).Value = sumAU / sumC
        End If

        wsSumm.Cells(dstRow, DC_SZV_PCT).NumberFormat   = "0.00%"
        wsSumm.Cells(dstRow, DC_MEGV_PCT).NumberFormat  = "0.00%"
        wsSumm.Cells(dstRow, DC_VESZT_PCT).NumberFormat = "0.00%"

        If Not hasAnyData Then wsSumm.Rows(dstRow).Hidden = True
        dstRow = dstRow + 1
    Next i

    Dim lastDataRow As Long
    lastDataRow = dstRow - 1

    ' Stat rows
    Dim statColors(3) As Long
    statColors(0) = RGB(173, 216, 230)
    statColors(1) = RGB(144, 238, 144)
    statColors(2) = RGB(255, 255, 153)
    statColors(3) = RGB(216, 191, 216)

    Dim statLabels(3) As String
    statLabels(0) = H_Osszeg()
    statLabels(1) = H_Atlag()
    statLabels(2) = "Maximum"
    statLabels(3) = "Minimum"

    Dim statRow As Integer
    For statRow = 0 To 3
        Dim outRow As Integer
        outRow = OUT_HDR_ROWS + 1 + statRow
        wsSumm.Cells(outRow, DC_IDOSZAK).Value = statLabels(statRow)

        Dim dc As Long
        For dc = DC_C To DC_AU
            Dim cs As String: cs = ColumnLetter(dc)
            Dim ra As String: ra = cs & dataFirstRow & ":" & cs & lastDataRow
            Dim f As String
            Select Case statRow
                Case 0: f = "=IFERROR(SUM("     & ra & "),"""")"
                Case 1: f = "=IFERROR(AVERAGE(" & ra & "),"""")"
                Case 2: f = "=IFERROR(MAX("     & ra & "),"""")"
                Case 3: f = "=IFERROR(MIN("     & ra & "),"""")"
            End Select
            wsSumm.Cells(outRow, dc).Formula = f
        Next dc

        Dim rSzvPct   As String: rSzvPct   = ColumnLetter(DC_SZV_PCT)   & dataFirstRow & ":" & ColumnLetter(DC_SZV_PCT)   & lastDataRow
        Dim rSzvDb    As String: rSzvDb    = ColumnLetter(DC_SZV_DB)    & dataFirstRow & ":" & ColumnLetter(DC_SZV_DB)    & lastDataRow
        Dim rMegvPct  As String: rMegvPct  = ColumnLetter(DC_MEGV_PCT)  & dataFirstRow & ":" & ColumnLetter(DC_MEGV_PCT)  & lastDataRow
        Dim rVesztDb  As String: rVesztDb  = ColumnLetter(DC_VESZT_DB)  & dataFirstRow & ":" & ColumnLetter(DC_VESZT_DB)  & lastDataRow
        Dim rVesztPct As String: rVesztPct = ColumnLetter(DC_VESZT_PCT) & dataFirstRow & ":" & ColumnLetter(DC_VESZT_PCT) & lastDataRow

        Select Case statRow
            Case 0
                wsSumm.Cells(outRow, DC_SZV_DB).Formula    = "=IFERROR(SUM(" & rSzvDb   & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_DB).Formula  = "=IFERROR(SUM(" & rVesztDb & "),"""")"
            Case 1
                wsSumm.Cells(outRow, DC_SZV_PCT).Formula   = "=IFERROR(AVERAGE(" & rSzvPct   & "),"""")"
                wsSumm.Cells(outRow, DC_SZV_DB).Formula    = "=IFERROR(AVERAGE(" & rSzvDb    & "),"""")"
                wsSumm.Cells(outRow, DC_MEGV_PCT).Formula  = "=IFERROR(AVERAGE(" & rMegvPct  & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_DB).Formula  = "=IFERROR(AVERAGE(" & rVesztDb  & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_PCT).Formula = "=IFERROR(AVERAGE(" & rVesztPct & "),"""")"
            Case 2
                wsSumm.Cells(outRow, DC_SZV_PCT).Formula   = "=IFERROR(MAX(" & rSzvPct   & "),"""")"
                wsSumm.Cells(outRow, DC_SZV_DB).Formula    = "=IFERROR(MAX(" & rSzvDb    & "),"""")"
                wsSumm.Cells(outRow, DC_MEGV_PCT).Formula  = "=IFERROR(MAX(" & rMegvPct  & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_DB).Formula  = "=IFERROR(MAX(" & rVesztDb  & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_PCT).Formula = "=IFERROR(MAX(" & rVesztPct & "),"""")"
            Case 3
                wsSumm.Cells(outRow, DC_SZV_PCT).Formula   = "=IFERROR(MIN(" & rSzvPct   & "),"""")"
                wsSumm.Cells(outRow, DC_SZV_DB).Formula    = "=IFERROR(MIN(" & rSzvDb    & "),"""")"
                wsSumm.Cells(outRow, DC_MEGV_PCT).Formula  = "=IFERROR(MIN(" & rMegvPct  & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_DB).Formula  = "=IFERROR(MIN(" & rVesztDb  & "),"""")"
                wsSumm.Cells(outRow, DC_VESZT_PCT).Formula = "=IFERROR(MIN(" & rVesztPct & "),"""")"
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
