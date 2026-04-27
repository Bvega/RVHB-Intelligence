Attribute VB_Name = "HB_Module1_Ingest"
' ============================================================
' RVHB Intelligence Toolkit
' HB_Module1_Ingest.bas  —  Annual Report Ingestion
' Version: 1.0
' ============================================================
Option Explicit

Private Const SRC_ROOMS_FULL  As String = "Contracted Room Nights"
Private Const DST_BOOKING_ID  As String = "Booking ID"
Private Const DST_DATE        As String = "Booking Date"
Private Const DST_MONTH       As String = "Booking Month"
Private Const DST_QUARTER     As String = "Booking Quarter"
Private Const DST_YEAR        As String = "Booking Year"
Private Const DST_PROGRAM     As String = "Program Name"
Private Const DST_CLIENT      As String = "Client"
Private Const DST_CONTRACT    As String = "Contract Date"
Private Const DST_START       As String = "Program Start Date"
Private Const DST_END         As String = "Program End Date"
Private Const DST_PROPERTY    As String = "Property"
Private Const DST_PARENT_BRD  As String = "Parent Brand"
Private Const DST_BRAND       As String = "Brand"
Private Const DST_CITY        As String = "Property City"
Private Const DST_STATE       As String = "Property State"
Private Const DST_COUNTRY     As String = "Property Country"
Private Const DST_ROOMS       As String = "Room Nights"
Private Const DST_REVENUE     As String = "Contracted Revenue (USD)"

Private Function DestHeaders() As Variant
    DestHeaders = Array( _
        DST_BOOKING_ID, DST_DATE,     DST_MONTH,      DST_QUARTER, _
        DST_YEAR,       DST_PROGRAM,  DST_CLIENT,     DST_CONTRACT, _
        DST_START,      DST_END,      DST_PROPERTY,   DST_PARENT_BRD, _
        DST_BRAND,      DST_CITY,     DST_STATE,      DST_COUNTRY, _
        DST_ROOMS,      DST_REVENUE)
End Function

' ── Main entry point — assign to Home button ─────────────────
Public Sub IngestNewReport()
    Dim srcPath      As String
    Dim srcWb        As Workbook
    Dim srcWs        As Worksheet
    Dim destWs       As Worksheet
    Dim srcColMap    As Object
    Dim destColMap   As Object
    Dim existingIDs  As Object
    Dim addedCount   As Long
    Dim skippedCount As Long
    Dim destNewRow   As Long
    Dim i            As Long

    srcPath = HB_PickFile()
    If srcPath = "" Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts  = False
    Application.StatusBar = HB_PRODUCT & ": Opening report..."

    On Error GoTo ErrorHandler
    Set srcWb = Workbooks.Open(Filename:=srcPath, ReadOnly:=True, _
                UpdateLinks:=False, IgnoreReadOnlyRecommended:=True)

    Set srcWs = HB_FindMainSheet(srcWb)
    If srcWs Is Nothing Then
        srcWb.Close False
        MsgBox "Could not find a Main sheet." & vbNewLine & _
               "Sheets found: " & HB_SheetList(srcWb), vbExclamation, HB_PRODUCT
        GoTo Cleanup
    End If

    Set srcColMap  = HB_BuildColMap(srcWs)
    Call HB_NormalizeRoomsCol(srcColMap)
    Set destWs     = ThisWorkbook.Sheets(SHEET_RAW)
    Set destColMap = HB_BuildColMap(destWs)

    Application.StatusBar = HB_PRODUCT & ": Scanning existing records..."
    Set existingIDs = HB_LoadExistingIDs(destWs)

    Application.StatusBar = HB_PRODUCT & ": Importing new records..."
    destNewRow   = HB_GetLastRow() + 1
    addedCount   = 0
    skippedCount = 0

    Dim srcLastRow As Long
    Dim headers    As Variant
    srcLastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    headers    = DestHeaders()

    For i = 2 To srcLastRow
        Dim bookingID As String
        bookingID = Trim(CStr(srcWs.Cells(i, 1).Value))
        If bookingID = "" Then GoTo NextRow
        If existingIDs.Exists(bookingID) Then
            skippedCount = skippedCount + 1
            GoTo NextRow
        End If
        Call HB_CopyRow(srcWs, i, destWs, destNewRow, srcColMap, destColMap, headers)
        Call HB_StyleDataRow(destWs, destNewRow)
        addedCount   = addedCount + 1
        destNewRow   = destNewRow + 1
        existingIDs(bookingID) = True
NextRow:
    Next i

    srcWb.Close False
    Application.StatusBar = HB_PRODUCT & ": Recalculating dashboards..."
    Application.CalculateFullRebuild
    Call HB_UpdateHomeStatus(addedCount, skippedCount)
    GoTo Cleanup

ErrorHandler:
    MsgBox "Import error: " & Err.Description, vbCritical, HB_PRODUCT
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close False

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True
    Application.StatusBar      = False
    If addedCount > 0 Or skippedCount > 0 Then
        MsgBox "Import complete!" & vbNewLine & vbNewLine & _
               "  Added:   " & addedCount   & " new booking(s)" & vbNewLine & _
               "  Skipped: " & skippedCount & " duplicate(s)"   & vbNewLine & vbNewLine & _
               "All dashboards refreshed.", vbInformation, HB_PRODUCT
    End If
End Sub

' ── Private helpers ──────────────────────────────────────────
Private Function HB_PickFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select HelmsBriscoe Annual Master Report"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.xlsm;*.xlsb"
        .AllowMultiSelect  = False
        .InitialFileName   = Environ("USERPROFILE") & "\Desktop\"
    End With
    If fd.Show = -1 Then HB_PickFile = fd.SelectedItems(1) Else HB_PickFile = ""
End Function

Private Function HB_FindMainSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Select Case Trim(LCase(ws.Name))
            Case "main", "main data", "main ", "bookings"
                Set HB_FindMainSheet = ws
                Exit Function
        End Select
    Next ws
    Set HB_FindMainSheet = Nothing
End Function

Private Sub HB_NormalizeRoomsCol(colMap As Object)
    If colMap.Exists(SRC_ROOMS_FULL) And Not colMap.Exists(DST_ROOMS) Then
        colMap(DST_ROOMS) = colMap(SRC_ROOMS_FULL)
    End If
End Sub

Private Function HB_LoadExistingIDs(ws As Worksheet) As Object
    Dim dict    As Object
    Dim lastRow As Long
    Dim i       As Long
    Dim id      As String
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        id = Trim(CStr(ws.Cells(i, 1).Value))
        If id <> "" Then dict(id) = True
    Next i
    Set HB_LoadExistingIDs = dict
End Function

Private Sub HB_CopyRow(srcWs As Worksheet, srcRow As Long, _
    destWs As Worksheet, destRow As Long, _
    srcMap As Object, destMap As Object, headers As Variant)
    Dim k   As Integer
    Dim hdr As String
    For k = 0 To UBound(headers)
        hdr = headers(k)
        If srcMap.Exists(hdr) And destMap.Exists(hdr) Then
            destWs.Cells(destRow, destMap(hdr)).Value = _
                srcWs.Cells(srcRow, srcMap(hdr)).Value
            If InStr(LCase(hdr), "date") > 0 Then
                destWs.Cells(destRow, destMap(hdr)).NumberFormat = "MM/DD/YYYY"
            End If
        End If
    Next k
    If destMap.Exists(DST_REVENUE) Then
        destWs.Cells(destRow, destMap(DST_REVENUE)).NumberFormat = """$""#,##0.00"
    End If
End Sub

Private Sub HB_StyleDataRow(ws As Worksheet, rowNum As Long)
    With ws.Rows(rowNum)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Interior.Color = IIf(rowNum Mod 2 = 0, CLR_GRAY, CLR_WHITE)
        .RowHeight = 15
    End With
End Sub