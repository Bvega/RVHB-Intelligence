Attribute VB_Name = "HB_Intelligence"
' ============================================================
' HelmsBriscoe Intelligence Toolkit  |  VBA Core Module
' Version 1.0  |  Module 1: Annual Report Ingestion
' ============================================================
' MODULES IN THIS FILE:
'   IngestNewReport   — import a new year's master report
'   RefreshDashboard  — force-recalculate all formulas
'   FormatRFP         — Module 2 stub (RFP Formatter)
'   HB_Navigate       — sheet navigation helper
'   HB_Version        — returns toolkit version string
' ============================================================

Option Explicit

' ── Constants ──────────────────────────────────────────────
Private Const HB_VERSION    As String = "1.0"
Private Const RAW_SHEET     As String = "📥 Raw Data"
Private Const HOME_SHEET    As String = "🏠 Home"
Private Const HEADER_ROW    As Long = 1

' Destination column names (must match Raw Data headers exactly)
Private Const DC_BOOKING_ID  As String = "Booking ID"
Private Const DC_DATE        As String = "Booking Date"
Private Const DC_MONTH       As String = "Booking Month"
Private Const DC_QUARTER     As String = "Booking Quarter"
Private Const DC_YEAR        As String = "Booking Year"
Private Const DC_PROGRAM     As String = "Program Name"
Private Const DC_CLIENT      As String = "Client"
Private Const DC_CONTRACT    As String = "Contract Date"
Private Const DC_START       As String = "Program Start Date"
Private Const DC_END         As String = "Program End Date"
Private Const DC_PROPERTY    As String = "Property"
Private Const DC_PARENT_BRD  As String = "Parent Brand"
Private Const DC_BRAND       As String = "Brand"
Private Const DC_CITY        As String = "Property City"
Private Const DC_STATE       As String = "Property State"
Private Const DC_COUNTRY     As String = "Property Country"
Private Const DC_ROOMS       As String = "Room Nights"
Private Const DC_REVENUE     As String = "Contracted Revenue (USD)"

' Source column aliases (handles slight name variations across report years)
Private Const SC_ROOMS_ALT   As String = "Contracted Room Nights"


' ============================================================
' MODULE 1 — Annual Report Ingestion
' ============================================================
Public Sub IngestNewReport()

    Dim fd          As FileDialog
    Dim srcPath     As String
    Dim srcWb       As Workbook
    Dim srcWs       As Worksheet
    Dim destWs      As Worksheet
    Dim srcColMap   As Object    ' source header -> column index
    Dim destColMap  As Object    ' dest header -> column index
    Dim existingIDs As Object    ' Scripting.Dictionary for dedup
    Dim addedCount  As Long
    Dim skippedCount As Long
    Dim srcLastRow  As Long
    Dim destNewRow  As Long
    Dim i           As Long
    Dim j           As Long
    Dim bookingID   As String
    Dim hdr         As String
    Dim srcCol      As Long
    Dim destLastCol As Long

    ' ── File picker ──────────────────────────────────────
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select HelmsBriscoe Annual Master Report"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.xlsm;*.xlsb"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
    End With

    If fd.Show <> -1 Then Exit Sub
    srcPath = fd.SelectedItems(1)

    ' ── Open source quietly ──────────────────────────────
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "HB Toolkit: Opening report..."

    On Error GoTo ErrorHandler
    Set srcWb = Workbooks.Open(Filename:=srcPath, ReadOnly:=True, _
                               UpdateLinks:=False, IgnoreReadOnlyRecommended:=True)

    ' ── Find Main sheet (tolerates trailing spaces) ──────
    Set srcWs = Nothing
    Dim sn As String
    Dim ws As Worksheet
    For Each ws In srcWb.Worksheets
        sn = Trim(LCase(ws.Name))
        If sn = "main" Or sn = "main data" Or sn = "bookings" Then
            Set srcWs = ws
            Exit For
        End If
    Next ws

    If srcWs Is Nothing Then
        srcWb.Close False
        MsgBox "Could not find a 'Main' sheet in:" & vbNewLine & srcPath & _
               vbNewLine & vbNewLine & "Sheet names found: " & _
               Join(HB_SheetNames(srcWb), ", "), _
               vbExclamation, "HB Toolkit — Sheet Not Found"
        GoTo Cleanup
    End If

    ' ── Build source column map from headers ─────────────
    Set srcColMap = CreateObject("Scripting.Dictionary")
    srcColMap.CompareMode = vbTextCompare
    Dim srcLastCol As Long
    srcLastCol = srcWs.Cells(HEADER_ROW, srcWs.Columns.Count).End(xlToLeft).Column
    For j = 1 To srcLastCol
        hdr = Trim(CStr(srcWs.Cells(HEADER_ROW, j).Value))
        If hdr <> "" Then srcColMap(hdr) = j
    Next j

    ' ── Map "Contracted Room Nights" to our dest name ────
    If srcColMap.Exists(SC_ROOMS_ALT) And Not srcColMap.Exists(DC_ROOMS) Then
        srcColMap(DC_ROOMS) = srcColMap(SC_ROOMS_ALT)
    End If

    ' ── Build destination column map ─────────────────────
    Set destColMap = CreateObject("Scripting.Dictionary")
    destColMap.CompareMode = vbTextCompare
    Set destWs = ThisWorkbook.Sheets(RAW_SHEET)
    destLastCol = destWs.Cells(HEADER_ROW, destWs.Columns.Count).End(xlToLeft).Column
    For j = 1 To destLastCol
        hdr = Trim(CStr(destWs.Cells(HEADER_ROW, j).Value))
        If hdr <> "" Then destColMap(hdr) = j
    Next j

    ' ── Load existing Booking IDs for deduplication ──────
    Application.StatusBar = "HB Toolkit: Scanning existing records..."
    Set existingIDs = CreateObject("Scripting.Dictionary")
    existingIDs.CompareMode = vbTextCompare
    Dim destLastRow As Long
    destLastRow = destWs.Cells(destWs.Rows.Count, 1).End(xlUp).Row
    For i = HEADER_ROW + 1 To destLastRow
        bookingID = Trim(CStr(destWs.Cells(i, 1).Value))
        If bookingID <> "" Then existingIDs(bookingID) = True
    Next i

    ' ── Walk source rows and append new ones ─────────────
    Application.StatusBar = "HB Toolkit: Importing new records..."
    srcLastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    destNewRow = destLastRow + 1
    addedCount = 0
    skippedCount = 0

    Dim destColHeaders(1 To 18) As String
    destColHeaders(1)  = DC_BOOKING_ID
    destColHeaders(2)  = DC_DATE
    destColHeaders(3)  = DC_MONTH
    destColHeaders(4)  = DC_QUARTER
    destColHeaders(5)  = DC_YEAR
    destColHeaders(6)  = DC_PROGRAM
    destColHeaders(7)  = DC_CLIENT
    destColHeaders(8)  = DC_CONTRACT
    destColHeaders(9)  = DC_START
    destColHeaders(10) = DC_END
    destColHeaders(11) = DC_PROPERTY
    destColHeaders(12) = DC_PARENT_BRD
    destColHeaders(13) = DC_BRAND
    destColHeaders(14) = DC_CITY
    destColHeaders(15) = DC_STATE
    destColHeaders(16) = DC_COUNTRY
    destColHeaders(17) = DC_ROOMS
    destColHeaders(18) = DC_REVENUE

    For i = HEADER_ROW + 1 To srcLastRow

        ' Get booking ID from source (column 1 assumed to be Booking ID)
        bookingID = Trim(CStr(srcWs.Cells(i, 1).Value))
        If bookingID = "" Then GoTo NextSrcRow

        If existingIDs.Exists(bookingID) Then
            skippedCount = skippedCount + 1
            GoTo NextSrcRow
        End If

        ' Copy each mapped column
        Dim k As Long
        For k = 1 To 18
            Dim dh As String
            dh = destColHeaders(k)
            If destColMap.Exists(dh) And srcColMap.Exists(dh) Then
                Dim dCol As Long, sCol As Long
                dCol = destColMap(dh)
                sCol = srcColMap(dh)
                destWs.Cells(destNewRow, dCol).Value = srcWs.Cells(i, sCol).Value
                ' Preserve date formatting
                If InStr(LCase(dh), "date") > 0 Then
                    destWs.Cells(destNewRow, dCol).NumberFormat = "MM/DD/YYYY"
                End If
            End If
        Next k

        ' Row styling (alternating)
        With destWs.Rows(destNewRow)
            .Font.Name = "Arial"
            .Font.Size = 9
            If destNewRow Mod 2 = 0 Then
                .Interior.Color = RGB(242, 242, 242)
            Else
                .Interior.Color = RGB(255, 255, 255)
            End If
            .RowHeight = 15
        End With

        ' Revenue number format
        If destColMap.Exists(DC_REVENUE) Then
            destWs.Cells(destNewRow, destColMap(DC_REVENUE)).NumberFormat = """$""#,##0.00"
        End If

        addedCount = addedCount + 1
        destNewRow = destNewRow + 1
        existingIDs(bookingID) = True

NextSrcRow:
    Next i

    ' ── Close source workbook ────────────────────────────
    srcWb.Close False

    ' ── Recalculate all formulas ─────────────────────────
    Application.StatusBar = "HB Toolkit: Recalculating dashboards..."
    Application.CalculateFullRebuild

    ' ── Update Home sheet status ─────────────────────────
    Call HB_UpdateHomeStatus(addedCount, skippedCount)

    ' ── Done ─────────────────────────────────────────────
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim msg As String
    msg = "Import Complete!" & vbNewLine & vbNewLine
    msg = msg & Chr(10) & "  Added:    " & addedCount & " new booking(s)" & vbNewLine
    msg = msg & "  Skipped: " & skippedCount & " duplicate(s)" & vbNewLine & vbNewLine
    msg = msg & "All dashboards have been refreshed."

    MsgBox msg, vbInformation, "HB Intelligence Toolkit v" & HB_VERSION
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    MsgBox "An error occurred: " & Err.Description & vbNewLine & _
           "Error number: " & Err.Number, vbCritical, "HB Toolkit Error"
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close False
    Exit Sub

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False

End Sub


' ============================================================
' Dashboard Refresh
' ============================================================
Public Sub RefreshDashboard()
    Application.ScreenUpdating = False
    Application.StatusBar = "HB Toolkit: Refreshing all dashboards..."
    Application.CalculateFullRebuild
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Call HB_UpdateHomeStatus(-1, -1)
    MsgBox "All dashboards refreshed.", vbInformation, "HB Intelligence Toolkit"
End Sub


' ============================================================
' MODULE 2 — RFP Formatter  (stub — coming soon)
' ============================================================
Public Sub FormatRFP()
    Dim msg As String
    msg = "RFP Formatter — Module 2" & vbNewLine & vbNewLine
    msg = msg & "This module will:" & vbNewLine
    msg = msg & "  • Pull client data from Raw Data" & vbNewLine
    msg = msg & "  • Auto-generate sourcing-ready RFP layout" & vbNewLine
    msg = msg & "  • Export as formatted PDF or Word doc" & vbNewLine & vbNewLine
    msg = msg & "Coming in the next release."
    MsgBox msg, vbInformation, "HB Intelligence Toolkit — Module 2"
End Sub


' ============================================================
' Navigation Helper
' ============================================================
Public Sub HB_Navigate(sheetName As String)
    On Error Resume Next
    ThisWorkbook.Sheets(sheetName).Activate
    If Err.Number <> 0 Then
        MsgBox "Sheet '" & sheetName & "' not found.", vbExclamation, "HB Toolkit"
    End If
    On Error GoTo 0
End Sub

Public Sub GoToRawData()      Call HB_Navigate(RAW_SHEET)      End Sub
Public Sub GoToYoYDashboard() Call HB_Navigate("📊 YoY Dashboard") End Sub
Public Sub GoToMonthly()      Call HB_Navigate("📅 Monthly Trends") End Sub
Public Sub GoToClients()      Call HB_Navigate("👥 Top Clients")    End Sub
Public Sub GoToBrands()       Call HB_Navigate("🏨 Brand Mix")      End Sub


' ============================================================
' Home Sheet Status Update
' ============================================================
Private Sub HB_UpdateHomeStatus(addedCount As Long, skippedCount As Long)
    On Error Resume Next
    Dim hws As Worksheet
    Set hws = ThisWorkbook.Sheets(HOME_SHEET)
    If hws Is Nothing Then Exit Sub

    hws.Range("status_date").Value = "Last updated:  " & Format(Now(), "MMM DD, YYYY  h:MM AM/PM")

    If addedCount >= 0 Then
        hws.Range("status_import").Value = _
            "Added " & addedCount & " booking(s)  ·  Skipped " & skippedCount & " duplicate(s)"
    Else
        hws.Range("status_import").Value = "Dashboard refreshed manually."
    End If

    ' Update total bookings count
    Dim destWs As Worksheet
    Set destWs = ThisWorkbook.Sheets(RAW_SHEET)
    Dim totalRows As Long
    totalRows = destWs.Cells(destWs.Rows.Count, 1).End(xlUp).Row - 1
    On Error Resume Next
    hws.Range("status_total").Value = totalRows & " total bookings in tracker"
    On Error GoTo 0
End Sub


' ============================================================
' Utility: Return sheet names as array
' ============================================================
Private Function HB_SheetNames(wb As Workbook) As String()
    Dim names() As String
    Dim ws As Worksheet
    Dim i As Long
    ReDim names(0 To wb.Worksheets.Count - 1)
    i = 0
    For Each ws In wb.Worksheets
        names(i) = ws.Name
        i = i + 1
    Next ws
    HB_SheetNames = names
End Function


' ============================================================
' Version
' ============================================================
Public Function HB_Version() As String
    HB_Version = HB_VERSION
End Function
