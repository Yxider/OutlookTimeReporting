Sub TotalCategories()

Dim app As New Outlook.Application
Dim namespace As Outlook.namespace
Dim calendar As Outlook.Folder
Dim appt As Outlook.AppointmentItem
Dim apptList As Outlook.Items
Dim apptListFiltered As Outlook.Items
Dim explorer As Outlook.explorer
Dim view As Outlook.view
Dim calView As Outlook.CalendarView
Dim startDate As String
Dim endDate As String
Dim category As String
Dim duration As Integer
Dim outMsg As String

' Access appointment list
Set namespace = app.GetNamespace("MAPI")
Set calendar = namespace.GetDefaultFolder(olFolderCalendar)
Set apptList = calendar.Items

' Include recurring appointments and sort the list
apptList.IncludeRecurrences = True
apptList.Sort "[Start]"

firstDay = DateSerial(Year(DateAdd("m", 0, Now)), Month(DateAdd("m", 0, Now)), 1)
lastDay = DateAdd("d", -1, DateSerial(Year(Now), Month(DateAdd("m", 1, Now)), 1))
If Month(lastDay) = 12 Then
    lastDay = DateAdd("yyyy", 1, lastDay)
End If


Dim Message, Title, Default, MyValue
Message = "Input"    ' Set prompt.
Title = "Start date"    ' Set title.
Default = firstDay    ' Set default.
' Display message, title, and default value.
firstDay = InputBox(Message, Title, Default)

Title = "Start date"    ' Set title.
Default = lastDay    ' Set default.
lastDay = InputBox(Message, Title, Default)

' firstDay = DateAdd("d", -1, firstDay)
lastDay = DateAdd("d", 1, lastDay)


' Get selected date
Set explorer = app.ActiveExplorer()
Set view = explorer.CurrentView()
Set calView = view

' Filter the appointment list
strFilter = "[Start] >= '" & firstDay & "'" & " AND [End] < '" & lastDay & "'"

Set apptListFiltered = apptList.Restrict(strFilter)

' Loop through the appointments and total for each category
Set catHours = CreateObject("Scripting.Dictionary")
For Each appt In apptListFiltered
    category = appt.Categories
    duration = appt.duration
    If catHours.Exists(category) Then
        catHours(category) = catHours(category) + duration
    Else
        catHours.Add category, duration
    End If
Next


Dim sum As Double
Dim toWork As Integer
Dim outoOfOfficeTime As Integer
Dim day As Date
Dim dayName As String
day = firstDay
toWork = 0
outoOfOfficeTime = 0

keyArray = catHours.Keys
For Each key In keyArray
    If key <> "" And key <> "Holiday" And key <> "OoO" Then
        outMsgClipboard = outMsgClipboard & key & Chr(9) & (catHours(key) / 60) & vbCrLf
        sum = sum + (catHours(key) / 60)
    End If

    dayName = WeekdayName(Weekday(day, vbMonday))
    If dayName <> "lördag" And dayName <> "söndag" Then
        toWork = toWork + 1
    End If

    If key = "OoO" Then
        outoOfOfficeTime = outoOfOfficeTime + (catHours(key) / 60)
    End If

    day = day + 1
Next
'Copy to clipboard
CopyText (outMsgClipboard)

toWork = Return_workdays_between_dates(firstDay, lastDay)

Dim TotalWorkHours As Integer
Dim msg As String
TotalWorkHours = toWork * 8 - outoOfOfficeTime
msg = "Worked: " & sum & " hours." & vbCrLf & "Planned: " & TotalWorkHours & " hours."
MsgBox msg


' Clean up objects
Set app = Nothing
Set namespace = Nothing
Set calendar = Nothing
Set appt = Nothing
Set apptList = Nothing
Set apptListFiltered = Nothing
Set explorer = Nothing
Set view = Nothing
Set calView = Nothing

End Sub

Public Function Return_workdays_between_dates(startDate, endDate)
    Dim day As Date
    Dim daysCount As Integer
    day = startDate
    daysCount = 0

    Do While day < endDate
        If WeekdayName(Weekday(day, vbMonday)) <> "lördag" And WeekdayName(Weekday(day, vbMonday)) <> "söndag" Then
            daysCount = daysCount + 1
        End If
        day = DateAdd("d", 1, day)
    Loop

    Return_workdays_between_dates = daysCount

End Function


Sub CopyText(Text As String)
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub
