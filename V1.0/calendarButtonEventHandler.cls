VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "calendarButtonEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this is the reference to the calendar buttons
Public WithEvents calendarButtonEvent As MSForms.CommandButton
Attribute calendarButtonEvent.VB_VarHelpID = -1
    
'this will be set in the form to the desired date to show
Public dateToReturn As String

'this will be called only by calendar buttons
Private Sub calendarButtonEvent_click()
    'MsgBox dateToReturn
    Sheets(ThisWorkbook.callingSheet).Cells(ThisWorkbook.callingCellX, ThisWorkbook.callingCellY) = dateToReturn
    Unload calendar
End Sub
