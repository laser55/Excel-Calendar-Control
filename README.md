# Excel-Calendar-Control
A calendar form for excel which can be easily controlled through VBA

## Setup
To load the calandar form code (the .frm file), on youe excel sheet, open the VBA view (ALT+F11) and on the project window, right click on your workbook VBA project name > import file > then choose the "Calendar.frm" file.
You also need to import the class module "calendarButtonEventHandler.cls" to enable clicking on your calandar form buttons*
Again right click on the workbook VBA project name > import file > then choose the "calendarButtonEventHandler.cls" file.

Now you should have 2 new folders in your VBA project: 1- Forms (containing the Calendar.frm) and 2- Class Modules (containing the calendarButtonEventHandler.cls)

## setting up the workbook object
You are not quite done yet. If you open the calendarButtonEventHandler class, you will notice that I am using 3 variables from the workbook object:
1- ThisWorkbook.callingSheet
2- ThisWorkbook.callingCellX
3- ThisWorkbook.callingCellY

I am using these variables to define where should the date chosen in the calandar go in the workbook (which sheet and which cell). Therefore, you need to defice these public variables in the workbook object. Copy and past the following code in the workbook object:
```VB
'this is for the calendar
Public callingCellX As Integer
Public callingCellY As Integer
Public callingSheet As String
```
That is all for setting up

## Using the calendar form
To use the form simply use this command anywhere in the workbook VBA project
```VB
calendar.Show
```
## example sheet
You also just use the already setup example excel sheet. The example is simple. The calender control will pop-up when you select the C2 cell in sheet one. click on the day you want and the control will automatically close and write the date in C2.


*if anyone has a better and more convenient way to implement this, please contribute to this small project
