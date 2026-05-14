Option Explicit

' Ghep cac sheet trong workbook hien tai vao sheet TongHop.
' Cach dung:
' 1. Mo file Excel can ghep.
' 2. Nhan Alt + F11, Insert > Module.
' 3. Import file .bas nay hoac copy toan bo code vao Module.
' 4. Chay macro GhepSheet_TongHop.

Public Sub GhepSheet_TongHop()
    Dim wb As Workbook
    Dim wsTongHop As Worksheet
    Dim ws As Worksheet
    Dim selectedSheets As Collection
    Dim item As Variant
    Dim headerStartRow As Long
    Dim headerEndRow As Long
    Dim dataStartRow As Long
    Dim dataEndRow As Long
    Dim actualDataEndRow As Long
    Dim targetRow As Long
    Dim sequenceNumber As Long
    Dim copiedSheetCount As Long
    Dim renumberFirstColumn As Boolean

    Set wb = ActiveWorkbook
    Set selectedSheets = New Collection

    headerStartRow = AskLong("Nhap dong bat dau tieu de:", 1)
    If headerStartRow = 0 Then Exit Sub

    headerEndRow = AskLong("Nhap dong ket thuc tieu de:", 7)
    If headerEndRow = 0 Then Exit Sub

    dataStartRow = AskLong("Nhap dong bat dau du lieu:", 8)
    If dataStartRow = 0 Then Exit Sub

    dataEndRow = AskLong("Nhap dong ket thuc du lieu (nhap 0 de tu dong dong cuoi co du lieu):", 0)

    If headerStartRow < 1 Or headerEndRow < headerStartRow Then
        MsgBox "Dai dong tieu de khong hop le.", vbExclamation, "Canh bao"
        Exit Sub
    End If

    If dataStartRow < 1 Or (dataEndRow <> 0 And dataEndRow < dataStartRow) Then
        MsgBox "Dai dong du lieu khong hop le.", vbExclamation, "Canh bao"
        Exit Sub
    End If

    renumberFirstColumn = (MsgBox("Ban co muon danh lai STT cot A khong?", vbYesNo + vbQuestion, "Danh lai STT") = vbYes)

    For Each ws In wb.Worksheets
        If ws.Name <> "TongHop" Then
            If MsgBox("Ghep sheet '" & ws.Name & "' vao TongHop?", vbYesNo + vbQuestion, "Chon sheet") = vbYes Then
                selectedSheets.Add ws.Name
            End If
        End If
    Next ws

    If selectedSheets.Count = 0 Then
        MsgBox "Chua chon sheet nao de ghep.", vbExclamation, "Canh bao"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo CleanFail

    DeleteSheetIfExists wb, "TongHop"
    Set wsTongHop = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    wsTongHop.Name = "TongHop"

    targetRow = 1
    CopyRowsWithFormat wb.Worksheets(CStr(selectedSheets(1))), wsTongHop, headerStartRow, headerEndRow, targetRow
    targetRow = targetRow + (headerEndRow - headerStartRow + 1)

    sequenceNumber = 1
    copiedSheetCount = 0

    For Each item In selectedSheets
        Set ws = wb.Worksheets(CStr(item))

        If dataEndRow = 0 Then
            actualDataEndRow = LastDisplayRow(ws, dataStartRow)
        Else
            actualDataEndRow = dataEndRow
        End If

        If actualDataEndRow >= dataStartRow Then
            CopyRowsWithFormat ws, wsTongHop, dataStartRow, actualDataEndRow, targetRow

            If renumberFirstColumn Then
                Dim rowIndex As Long
                For rowIndex = targetRow To targetRow + (actualDataEndRow - dataStartRow)
                    wsTongHop.Cells(rowIndex, 1).Value = sequenceNumber
                    sequenceNumber = sequenceNumber + 1
                Next rowIndex
            End If

            targetRow = targetRow + (actualDataEndRow - dataStartRow + 1)
            copiedSheetCount = copiedSheetCount + 1
        End If
    Next item

    wsTongHop.Activate
    wsTongHop.Range("A1").Select

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    MsgBox "Da ghep xong " & copiedSheetCount & " sheet vao TongHop.", vbInformation, "Thanh cong"
    Exit Sub

CleanFail:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    MsgBox "Co loi xay ra: " & Err.Description, vbCritical, "Loi"
End Sub

Private Function AskLong(ByVal prompt As String, ByVal defaultValue As Long) As Long
    Dim answer As Variant

    answer = Application.InputBox(prompt, "Thong so ghep sheet", defaultValue, Type:=1)
    If answer = False Then
        AskLong = 0
    Else
        AskLong = CLng(answer)
    End If
End Function

Private Sub DeleteSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            ws.Delete
            Exit Sub
        End If
    Next ws
End Sub

Private Sub CopyRowsWithFormat( _
    ByVal sourceWs As Worksheet, _
    ByVal targetWs As Worksheet, _
    ByVal sourceStartRow As Long, _
    ByVal sourceEndRow As Long, _
    ByVal targetStartRow As Long _
)
    Dim lastCol As Long
    Dim sourceRange As Range
    Dim targetCell As Range
    Dim colIndex As Long
    Dim rowOffset As Long

    If sourceEndRow < sourceStartRow Then Exit Sub

    lastCol = LastDisplayColumn(sourceWs)
    If lastCol < 1 Then Exit Sub

    Set sourceRange = sourceWs.Range(sourceWs.Cells(sourceStartRow, 1), sourceWs.Cells(sourceEndRow, lastCol))
    Set targetCell = targetWs.Cells(targetStartRow, 1)

    sourceRange.Copy Destination:=targetCell

    For colIndex = 1 To lastCol
        targetWs.Columns(colIndex).ColumnWidth = sourceWs.Columns(colIndex).ColumnWidth
        targetWs.Columns(colIndex).Hidden = sourceWs.Columns(colIndex).Hidden
    Next colIndex

    For rowOffset = 0 To sourceEndRow - sourceStartRow
        targetWs.Rows(targetStartRow + rowOffset).RowHeight = sourceWs.Rows(sourceStartRow + rowOffset).RowHeight
        targetWs.Rows(targetStartRow + rowOffset).Hidden = sourceWs.Rows(sourceStartRow + rowOffset).Hidden
    Next rowOffset
End Sub

Private Function LastDisplayRow(ByVal ws As Worksheet, ByVal startRow As Long) As Long
    Dim usedRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowIndex As Long
    Dim colIndex As Long

    Set usedRange = ws.UsedRange
    lastRow = usedRange.Row + usedRange.Rows.Count - 1
    lastCol = usedRange.Column + usedRange.Columns.Count - 1

    For rowIndex = lastRow To startRow Step -1
        For colIndex = 1 To lastCol
            If Trim$(CStr(ws.Cells(rowIndex, colIndex).Text)) <> "" Then
                LastDisplayRow = rowIndex
                Exit Function
            End If
        Next colIndex
    Next rowIndex

    LastDisplayRow = startRow - 1
End Function

Private Function LastDisplayColumn(ByVal ws As Worksheet) As Long
    Dim usedRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowIndex As Long
    Dim colIndex As Long

    Set usedRange = ws.UsedRange
    lastRow = usedRange.Row + usedRange.Rows.Count - 1
    lastCol = usedRange.Column + usedRange.Columns.Count - 1

    For colIndex = lastCol To 1 Step -1
        For rowIndex = 1 To lastRow
            If Trim$(CStr(ws.Cells(rowIndex, colIndex).Text)) <> "" Then
                LastDisplayColumn = colIndex
                Exit Function
            End If
        Next rowIndex
    Next colIndex

    LastDisplayColumn = 0
End Function
