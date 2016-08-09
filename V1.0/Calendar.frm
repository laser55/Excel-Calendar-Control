VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendar 
   Caption         =   "Calendar"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4305
   OleObjectBlob   =   "Calendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim calendarButtonEvents(0 To 5, 0 To 6) As New calendarButtonEventHandler
Dim referenceDate As Date

Private Sub nextMonth_Click()
    setupForm (DateAdd("m", 1, referenceDate))
End Sub

Private Sub prevMonth_Click()
    setupForm (DateAdd("m", -1, referenceDate))
End Sub

Private Sub UserForm_Initialize()
    setupForm (Date)
End Sub

Function setupForm(refDate As Date)
    Dim i, j As Integer
    Dim calendar As Variant
    Dim thisMonth As Integer
    Dim monthName As Variant
    monthName = ",January,February,March,April,May,June,July,August,September,October,November,December"
    monthName = Split(monthName, ",")
    
    Me.Caption = "Calendar"
    sundayLabel.Caption = "Su"
    mondayLabel.Caption = "Mo"
    tusedayLabel.Caption = "Tu"
    wednesdayLabel.Caption = "We"
    thursedayLabel.Caption = "Th"
    fridayLabel.Caption = "Fr"
    saturdayLabel.Caption = "Sa"
    monthLabel.Caption = monthName(Month(refDate)) & " " & Year(refDate)
      
    calendar = getDates(refDate)
    thisMonth = Month(refDate)
    
    For i = 0 To 5
        For j = 0 To 6
            Me.Controls("d" & i & j).Caption = Day(calendar(i, j))
            
            If Month(calendar(i, j)) <> thisMonth Then
                Me.Controls("d" & i & j).Enabled = False
            Else
                'setting event handler to button
                Me.Controls("d" & i & j).Enabled = True
                Set calendarButtonEvents(i, j).calendarButtonEvent = Me.Controls("d" & i & j)
                calendarButtonEvents(i, j).dateToReturn = Day(calendar(i, j)) & "/" & Month(calendar(i, j)) & "/" & Year(calendar(i, j))
            End If
        Next j
    Next i
    
    referenceDate = refDate
End Function

'this function gets the day numbers for each month display
Function getDates(refDate As Date)
    Dim initialDate As Date
    Dim refDay As Integer
    Dim calendar(0 To 5, 0 To 6) As Date
    
    Dim daysToRemove, daysToAdd As Integer
    Dim i, j As Integer

    'get day number of the month
    refDay = Day(refDate)
    
    'going back to the beginning of the month
    daysToRemove = refDay - 1
    initialDate = DateAdd("d", -daysToRemove, refDate)
    
    'going back to the beginning of the month
    daysToRemove = Weekday(initialDate) - 1
    initialDate = DateAdd("d", -daysToRemove, initialDate)
    
    'filling the dates array
    daysToAdd = 0
    For i = 0 To 5
        For j = 0 To 6
            calendar(i, j) = DateAdd("d", daysToAdd, initialDate)
            daysToAdd = daysToAdd + 1
        Next j
    Next i
    
    getDates = calendar
End Function
