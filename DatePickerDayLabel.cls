VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatePickerDayLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private formObj As FormDatePicker
Private WithEvents labelObj As MSForms.Label
Attribute labelObj.VB_VarHelpID = -1
Private labelDate As Variant

Private Sub Class_Initialize()
    Set formObj = Nothing
    Set labelObj = Nothing
    labelDate = Null
End Sub

Private Sub labelObj_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim selectedDate As Variant
    
    Call formObj.resetHover
    
    If IsNull(labelDate) Then
        Exit Sub
    End If
    
    formObj.setHoverLabel Me
End Sub

Private Sub labelObj_click()
    Dim obj As MSForms.Label
    Dim ldate As Date
    Dim monthInt As Integer
    Dim yearInt As Integer
    
    If IsNull(labelDate) Then
        Exit Sub
    End If
    
    monthInt = Month(labelDate)
    yearInt = Year(labelDate)
    
    formObj.setSelectedLabel Me
    
    formObj.setDisableRefresh True
    formObj.CmbMonth.listIndex = monthInt - 1
    formObj.SpinYear.value = yearInt
    formObj.setDisableRefresh False
    
    Call formObj.refresh
End Sub

Public Sub setLabelObj(obj As MSForms.Label)
    Set labelObj = obj
End Sub

Public Function getLabelObj() As MSForms.Label
    Set getLabelObj = labelObj
End Function

Public Sub setFormObj(obj As UserForm)
    Set formObj = obj
End Sub

Public Sub setLabelDate(dateVar As Variant)
    labelDate = dateVar
End Sub

Public Function getLabelDate() As Variant
    getLabelDate = labelDate
End Function
