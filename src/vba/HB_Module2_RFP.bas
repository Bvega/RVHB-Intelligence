Attribute VB_Name = "HB_Module2_RFP"
' ============================================================
' RVHB Intelligence Toolkit
' HB_Module2_RFP.bas  —  RFP Formatter
' Version: 1.0  (stub — releasing in v1.1)
' ============================================================
Option Explicit

Public Sub FormatRFP()
    Dim msg As String
    msg = "RFP Formatter  -  Module 2" & vbNewLine & vbNewLine
    msg = msg & "This module will:" & vbNewLine & vbNewLine
    msg = msg & "  1.  Select a client from your Raw Data" & vbNewLine
    msg = msg & "  2.  Choose a date range" & vbNewLine
    msg = msg & "  3.  Auto-fill RFP fields (property, city," & vbNewLine
    msg = msg & "      room nights, dates, revenue target)" & vbNewLine
    msg = msg & "  4.  Generate a formatted, print-ready RFP" & vbNewLine
    msg = msg & "  5.  Export as PDF or separate workbook" & vbNewLine & vbNewLine
    msg = msg & "Coming in Version 1.1."
    MsgBox msg, vbInformation, HB_PRODUCT & "  -  Module 2"
End Sub

' ── Stubs (fleshed out in v1.1) ──────────────────────────────
Private Function GetClientList() As Variant
    ' TODO: read COL_CLIENT, return unique sorted array
    GetClientList = Array()
End Function

Private Function GetClientBookings(clientName As String, _
    dateFrom As Date, dateTo As Date) As Object
    ' TODO: return filtered rows from SHEET_RAW
    Set GetClientBookings = Nothing
End Function

Private Sub BuildRFPSheet(bookings As Object, clientName As String)
    ' TODO: create RFP sheet, populate header + rows
End Sub

Private Sub ExportToPDF(clientName As String)
    ' TODO: ExportAsFixedFormat with date-stamped filename
End Sub