Attribute VB_Name = "modCalendar"
Option Explicit

' Pop-up month calendar for the date cells on Settings. Double-click a date cell
' to open it; click a day to write it back. Kept on a worksheet rather than an
' ActiveX date picker so it works on every Excel install.
Private Const CAL_SHEET As String = "Calendar"

Private gTargetSheet As String
Private gTargetAddress As String
Private gMonthStart As Date
Private gRendering As Boolean

Public Sub ShowCalendarFor(ByVal target As Range)
    Dim seed As Date, ws As Worksheet

    Set ws = CalendarSheet()
    If ws Is Nothing Then Exit Sub

    gTargetSheet = target.Parent.Name
    gTargetAddress = target.Address(False, False)
    If IsDate(target.Value) Then
        seed = DateValue(target.Value)
    Else
        seed = Date
    End If
    gMonthStart = DateSerial(Year(seed), Month(seed), 1)

    RenderCalendar
    ws.Visible = xlSheetVisible
    ws.Activate
    gRendering = True
    ws.Range("B1").Select
    gRendering = False
End Sub

' Double-clicking a row on the Expiry List sheet makes that expiry active.
Public Sub UseExpiryFromList(ByVal target As Range)
    Dim ws As Worksheet, picked As Variant

    If target Is Nothing Then Exit Sub
    If target.Row < 2 Then Exit Sub

    picked = target.Parent.Cells(target.Row, 1).Value
    If Not IsDate(picked) Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ws.Range("B7").Value = DateValue(picked)
    ws.Range("B7").NumberFormat = "dd-mmm-yyyy"
    ws.Activate
    ws.Range("B7").Select
End Sub

Public Sub CalendarClicked(ByVal target As Range)
    Dim picked As Variant

    If gRendering Then Exit Sub
    If target Is Nothing Then Exit Sub
    If target.Cells.CountLarge > 1 Then Exit Sub

    If target.Row = 1 Then
        If target.Column = 1 Then
            gMonthStart = DateAdd("m", -1, gMonthStart)
            RenderCalendar
        ElseIf target.Column = 7 Then
            gMonthStart = DateAdd("m", 1, gMonthStart)
            RenderCalendar
        End If
        Exit Sub
    End If

    If target.Row >= 4 And target.Row <= 9 And target.Column <= 7 Then
        picked = target.Value
        If IsDate(picked) Then ApplyPickedDate DateValue(picked)
    End If
End Sub

Public Sub CloseCalendar()
    Dim ws As Worksheet
    Set ws = CalendarSheet()
    If ws Is Nothing Then Exit Sub
    gRendering = True
    On Error Resume Next
    ThisWorkbook.Worksheets("Settings").Activate
    ws.Visible = xlSheetHidden
    On Error GoTo 0
    gRendering = False
End Sub

Private Sub ApplyPickedDate(ByVal picked As Date)
    Dim ws As Worksheet

    If gTargetSheet <> "" And gTargetAddress <> "" Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(gTargetSheet)
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Range(gTargetAddress).Value = picked
            ws.Range(gTargetAddress).NumberFormat = "dd-mmm-yyyy"
        End If
    End If

    CloseCalendar
    If Not ws Is Nothing Then ws.Range(gTargetAddress).Select
End Sub

Private Sub RenderCalendar()
    Dim ws As Worksheet, cursor As Date, r As Long, c As Long, leadingDays As Long

    Set ws = CalendarSheet()
    If ws Is Nothing Then Exit Sub
    If gMonthStart = 0 Then gMonthStart = DateSerial(Year(Date), Month(Date), 1)

    gRendering = True
    ws.Range("A4:G9").ClearContents
    ws.Range("B1").NumberFormat = "@"
    ws.Range("B1").Value = Format$(gMonthStart, "mmmm yyyy")

    leadingDays = Weekday(gMonthStart, vbMonday) - 1
    cursor = DateAdd("d", -leadingDays, gMonthStart)
    For r = 4 To 9
        For c = 1 To 7
            If Month(cursor) = Month(gMonthStart) Then
                ws.Cells(r, c).Value = cursor
            End If
            cursor = DateAdd("d", 1, cursor)
        Next c
    Next r

    ws.Range("A4:G9").NumberFormat = "d"
    gRendering = False
End Sub

Private Function CalendarSheet() As Worksheet
    On Error Resume Next
    Set CalendarSheet = ThisWorkbook.Worksheets(CAL_SHEET)
End Function
