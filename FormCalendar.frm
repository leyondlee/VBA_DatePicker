VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCalendar 
   Caption         =   "Calendar"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4635
   OleObjectBlob   =   "FormCalendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LABEL_DEFAULT_FORECOLOR As Long = &H80000012
Private Const LABEL_EDGE_FORECOLOR As Long = &H80000010
Private Const LABEL_DEFAULT_BACKCOLOR As Long = &H8000000F
Private Const LABEL_HOVER_BACKCOLOR As Long = &H80000016
Private Const LABEL_SELECTED_BACKCOLOR As Long = &H8000000A
Private Const LABEL_DEFAULT_BORDERCOLOR As Long = &H80000006
Private Const LABEL_TODAY_BORDERCOLOR As Long = &HFF&

Private disableEvents As Boolean
Private labelEventCol As Collection
Private hoverLabel As CalendarDayLabel
Private selectedLabel As CalendarDayLabel

Private Sub ButtonLeft_Click()
    Dim monthInt As Integer
    Dim yearInt As Integer
    
    monthInt = CmbMonth.listIndex + 1
    yearInt = SpinYear.value
    If monthInt = 1 Then
        If yearInt <= SpinYear.Min Then
            Exit Sub
        End If
        
        monthInt = 12
        yearInt = yearInt - 1
    Else
        monthInt = monthInt - 1
    End If
    
    CmbMonth.listIndex = monthInt - 1
    SpinYear.value = yearInt
End Sub

Private Sub ButtonRight_Click()
    Dim monthInt As Integer
    Dim yearInt As Integer
    
    monthInt = CmbMonth.listIndex + 1
    yearInt = SpinYear.value
    If monthInt = 12 Then
        If yearInt >= SpinYear.Max Then
            Exit Sub
        End If
        
        monthInt = 1
        yearInt = yearInt + 1
    Else
        monthInt = monthInt + 1
    End If
    
    CmbMonth.listIndex = monthInt - 1
    SpinYear.value = yearInt
End Sub

Private Sub ButtonCancel_Click()
    Set selectedLabel = Nothing
    Me.Hide
End Sub

Private Sub ButtonOk_Click()
    If selectedLabel Is Nothing Then
        MsgBox "PLEASE SELECT A DATE"
        Exit Sub
    End If
    
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 1 Then
        Set selectedLabel = Nothing
        Me.Hide
        Cancel = 1
    End If
End Sub

Private Sub CmbMonth_Change()
    If disableEvents Then
        Exit Sub
    End If
    
    Call resetSelected
    Call refresh
End Sub

Private Sub SpinYear_Change()
    TxtYear.value = SpinYear.value
    
    If disableEvents Then
        Exit Sub
    End If
    
    Call resetSelected
    Call refresh
End Sub

Private Sub UserForm_Initialize()
    Dim labelObj As MSForms.Label
    Dim labelEventObj As CalendarDayLabel
    
    disableEvents = False
    
    Set labelEventCol = New Collection
    
    Set hoverLabel = Nothing
    Set selectedLabel = Nothing
    
    For i = 1 To 42
        Set labelObj = Me.Controls("LblDay" & i)
        
        Set labelEventObj = New CalendarDayLabel
        labelEventObj.setFormObj Me
        labelEventObj.setLabelObj labelObj
        
        labelEventCol.Add labelEventObj, Str(i)
    Next i
    
    For i = 1 To 12
        CmbMonth.AddItem UCase(MonthName(i))
    Next i
    
    CmbMonth.listIndex = Month(Now()) - 1
    SpinYear.value = Year(Now())
End Sub

Private Sub TxtYear_KeyDown(ByVal KeyCode As ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        
        Call updateYear
    End If
End Sub

Private Sub TxtYear_Exit(ByVal Cancel As ReturnBoolean)
    Call updateYear
End Sub

Private Sub updateYear()
    Dim minVal As Integer
    Dim maxVal As Integer
    Dim value As String
    Dim yearInt As Integer
    
    minVal = SpinYear.Min
    maxVal = SpinYear.Max
    
    value = TxtYear.value
    
    If Len(value) = 4 And IsNumeric(value) Then
        yearInt = Int(value)
        
        If yearInt >= minVal And yearInt <= maxVal Then
            SpinYear.value = yearInt
            Exit Sub
        End If
    End If
    
    TxtYear.value = SpinYear.value
End Sub

Public Sub refresh()
    Dim dateToday As Date
    Dim selectedDate As Variant
    Dim yearInt As Integer
    Dim monthInt As Integer
    Dim firstWeekday As Integer
    Dim lastDay As Integer
    Dim prevMonthInt As Integer
    Dim prevYearInt As Integer
    Dim prevLastDay As Integer
    Dim nextMonthInt As Integer
    Dim nextYearInt As Integer
    Dim labelEventObj As CalendarDayLabel
    Dim labelObj As MSForms.Label
    Dim dayInt As Integer
    Dim colorCode As Variant
    Dim cMonthInt As Integer
    Dim cYearInt As Integer
    Dim dateStr As String
    Dim curDate As Date
    
    dateToday = Date
    selectedDate = getSelectedDate()
    
    yearInt = SpinYear.value
    monthInt = CmbMonth.listIndex + 1
    
    firstWeekday = Weekday(DateSerial(yearInt, monthInt, 1))
    lastDay = Day(CalendarMod.getLastDate(monthInt, yearInt))
    
    If monthInt = 1 Then
        prevYearInt = yearInt - 1
        
        If prevYearInt >= SpinYear.Min Then
            prevMonthInt = 12
            prevLastDay = Day(CalendarMod.getLastDate(prevMonthInt, prevYearInt))
        Else
            prevYearInt = -1
        End If
    Else
        prevMonthInt = monthInt - 1
        prevYearInt = yearInt
        prevLastDay = Day(CalendarMod.getLastDate(prevMonthInt, prevYearInt))
    End If
    
    If monthInt = 12 Then
        nextYearInt = yearInt + 1
        
        If nextYearInt <= SpinYear.Max Then
            nextMonthInt = 1
        Else
            nextYearInt = -1
        End If
    Else
        nextMonthInt = monthInt + 1
        nextYearInt = yearInt
    End If
    
    For i = 1 To 42
        Set labelEventObj = labelEventCol.Item(Str(i))
        Set labelObj = labelEventObj.getLabelObj()
        
        dayInt = i - firstWeekday + 1
        
        colorCode = LABEL_EDGE_FORECOLOR
        
        If dayInt <= 0 Then
            dayInt = prevLastDay + dayInt
            cMonthInt = prevMonthInt
            cYearInt = prevYearInt
        ElseIf dayInt > lastDay Then
            dayInt = i - lastDay - firstWeekday + 1
            cMonthInt = nextMonthInt
            cYearInt = nextYearInt
        Else
            cMonthInt = monthInt
            cYearInt = yearInt
            
            colorCode = LABEL_DEFAULT_FORECOLOR
        End If
        
        If cYearInt = -1 Then
            labelEventObj.setLabelDate Null
            labelObj.Caption = ""
            labelObj.BackColor = LABEL_DEFAULT_BACKCOLOR
        Else
            curDate = DateSerial(cYearInt, cMonthInt, dayInt)
            
            If Not IsNull(selectedDate) Then
                If selectedDate = curDate Then
                    Set selectedLabel = labelEventObj
                    labelObj.BackColor = LABEL_SELECTED_BACKCOLOR
                Else
                    labelObj.BackColor = LABEL_DEFAULT_BACKCOLOR
                End If
            End If
            
            labelEventObj.setLabelDate curDate
            labelObj.Caption = dayInt
            labelObj.ForeColor = colorCode
            
            If curDate = dateToday Then
                labelObj.BorderColor = LABEL_TODAY_BORDERCOLOR
            Else
                labelObj.BorderColor = LABEL_DEFAULT_BORDERCOLOR
            End If
        End If
    Next i
End Sub
    
Private Sub DayFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call resetHover
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call resetHover
End Sub

Public Sub resetHover()
    Dim selectedDate As Variant
    
    If hoverLabel Is Nothing Then
        Exit Sub
    End If
    
    selectedDate = getSelectedDate()
    If IsNull(selectedDate) Or hoverLabel.getLabelDate() <> selectedDate Then
        hoverLabel.getLabelObj.BackColor = LABEL_DEFAULT_BACKCOLOR
        Set hoverLabel = Nothing
    End If
End Sub

Public Sub resetSelected()
    If selectedLabel Is Nothing Then
        Exit Sub
    End If
    
    selectedLabel.getLabelObj.BackColor = LABEL_DEFAULT_BACKCOLOR
    Set selectedLabel = Nothing
End Sub

Public Function getSelectedDate() As Variant
    getSelectedDate = Null
    If selectedLabel Is Nothing Then
        Exit Function
    End If
    
    If IsNull(selectedLabel.getLabelDate()) Then
        Exit Function
    End If
    
    getSelectedDate = selectedLabel.getLabelDate()
End Function

Public Sub setDisableEvents(bool As Boolean)
    disableEvents = bool
End Sub

Public Sub setHoverLabel(obj As CalendarDayLabel)
    Dim labelObj As MSForms.Label
    
    Set hoverLabel = obj
    
    Set labelObj = obj.getLabelObj()
    selectedDate = getSelectedDate()
    If IsNull(selectedDate) Or obj.getLabelDate() <> selectedDate Then
        labelObj.BackColor = LABEL_HOVER_BACKCOLOR
    End If
End Sub

Public Sub setSelectedLabel(labelObj As CalendarDayLabel)
    Set selectedLabel = labelObj
End Sub
