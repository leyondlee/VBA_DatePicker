Attribute VB_Name = "CalendarMod"
Function getLastDate(monthInt As Integer, yearInt As Integer) As Date
    Dim firstDate As Date
    Dim lastDate As Date
    
    firstDate = DateSerial(yearInt, monthInt, 1)
    lastDate = WorksheetFunction.EoMonth(firstDate, 0)
    
    getLastDate = lastDate
End Function

Function selectDate() As Variant
    Dim calendarForm As FormCalendar
    Dim result As Variant
    
    Set calendarForm = New FormCalendar
    calendarForm.Show
    
    result = calendarForm.getSelectedDate()
    Unload calendarForm
    
    selectDate = result
End Function

Sub test()
    Dim selDate As Variant
    
    selDate = selectDate()
    If IsNull(selDate) Then
        Exit Sub
    End If
    
    MsgBox "SELECTED DATE: " & Day(selDate) & " " & UCase(MonthName(Month(selDate))) & " " & Year(selDate)
End Sub
