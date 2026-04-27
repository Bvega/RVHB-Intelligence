Attribute VB_Name = "HB_Core"
' ============================================================
' RVHB Intelligence Toolkit
' HB_Core.bas  —  Shared constants, utilities & navigation
' Version: 1.0
' ============================================================
Option Explicit

Public Const HB_VERSION     As String = "1.0"
Public Const HB_PRODUCT     As String = "RVHB Intelligence Toolkit"
Public Const HB_YEAR        As Integer = 2024

Public Const SHEET_HOME     As String = "🏠 Home"
Public Const SHEET_RAW      As String = "📥 Raw Data"
Public Const SHEET_DASH     As String = "📊 YoY Dashboard"
Public Const SHEET_MONTHLY  As String = "📅 Monthly Trends"
Public Const SHEET_CLIENTS  As String = "👥 Top Clients"
Public Const SHEET_BRANDS   As String = "🏨 Brand Mix"

Public Const COL_BOOKING_ID  As Integer = 1
Public Const COL_BOOK_DATE   As Integer = 2
Public Const COL_BOOK_MONTH  As Integer = 3
Public Const COL_BOOK_QTR    As Integer = 4
Public Const COL_BOOK_YEAR   As Integer = 5
Public Const COL_PROGRAM     As Integer = 6
Public Const COL_CLIENT      As Integer = 7
Public Const COL_CONTRACT    As Integer = 8
Public Const COL_START_DATE  As Integer = 9
Public Const COL_END_DATE    As Integer = 10
Public Const COL_PROPERTY    As Integer = 11
Public Const COL_PARENT_BRD  As Integer = 12
Public Const COL_BRAND       As Integer = 13
Public Const COL_CITY        As Integer = 14
Public Const COL_STATE       As Integer = 15
Public Const COL_COUNTRY     As Integer = 16
Public Const COL_ROOMS       As Integer = 17
Public Const COL_REVENUE     As Integer = 18

Public Const NR_STATUS_DATE   As String = "status_date"
Public Const NR_STATUS_IMPORT As String = "status_import"
Public Const NR_STATUS_TOTAL  As String = "status_total"

Public Const CLR_NAVY   As Long = 3958820
Public Const CLR_BLUE   As Long = 6250403
Public Const CLR_GREEN  As Long = 1738561
Public Const CLR_GRAY   As Long = 15921906
Public Const CLR_WHITE  As Long = 16777215
Public Const CLR_INPUT  As Long = 16775620

' ── Navigation ──────────────────────────────────────────────
Public Sub HB_GoTo(sheetName As String)
    On Error Resume Next
    ThisWorkbook.Sheets(sheetName).Activate
    If Err.Number <> 0 Then
        MsgBox "Sheet '" & sheetName & "' not found.", vbExclamation, HB_PRODUCT
    End If
    On Error GoTo 0
End Sub

Public Sub GoToHome()      Call HB_GoTo(SHEET_HOME)    End Sub
Public Sub GoToRawData()   Call HB_GoTo(SHEET_RAW)     End Sub
Public Sub GoToDashboard() Call HB_GoTo(SHEET_DASH)    End Sub
Public Sub GoToMonthly()   Call HB_GoTo(SHEET_MONTHLY) End Sub
Public Sub GoToClients()   Call HB_GoTo(SHEET_CLIENTS) End Sub
Public Sub GoToBrands()    Call HB_GoTo(SHEET_BRANDS)  End Sub

' ── Dashboard refresh ────────────────────────────────────────
Public Sub RefreshDashboard()
    Application.ScreenUpdating = False