Attribute VB_Name = "OutlookAccess"
Option Explicit

Sub FindApptsToday()
    Dim appt As Variant 'Appointment
    For Each appt In FindAppts(Date, Date + 1)
        Debug.Print appt.ToString()
    Next
End Sub

Function FindAppts(ByVal FromDate As Date, ByVal ToDate As Date) As Appointment()

    Dim CalFolder As Outlook.Folder
    Dim OutItems As Outlook.Items
    Dim OutItemsInDateRange As Outlook.Items
    Dim OutAppt As Outlook.AppointmentItem
    Dim StrRestriction As String
    Dim OutApp As New Outlook.Application

    'Construct filter for the next 30-day date range
    StrRestriction = "[Start] >= '" & _
        Format$(FromDate, "mm/dd/yyyy hh:mm AMPM") _
        & "' AND [End] <= '" & _
        Format$(ToDate, "mm/dd/yyyy hh:mm AMPM") & "'"
    'Check the restriction string
    Debug.Print StrRestriction

    Set CalFolder = OutApp.Session.GetDefaultFolder(olFolderCalendar)
    Set OutItems = CalFolder.Items
    OutItems.IncludeRecurrences = True
    OutItems.Sort "[Start]"
    'Restrict the Items collection for the 30-day date range
    Set OutItemsInDateRange = OutItems.Restrict(StrRestriction)
    'Construct filter for Subject containing 'team'
    'Const PropTag  As String = "http://schemas.microsoft.com/mapi/proptag/"
    'strRestriction = "@SQL=" & Chr(34) & PropTag _
    '    & "0x0037001E" & Chr(34) & " like '%team%'"
    ''Restrict the last set of filtered items for the subject
    'Set oFinalItems = outItemsInDateRange.Restrict(strRestriction)
    ''Sort and Debug.Print final results
    'oFinalItems.Sort "[Start]"
    Dim apptFactory  As Appointment: Set apptFactory = New Appointment
    Dim results() As Appointment, apptCount As Integer
    For Each OutAppt In OutItemsInDateRange
        ReDim Preserve results(apptCount)
        apptCount = apptCount + 1
        Set results(UBound(results)) = apptFactory.NewAppointment(OutAppt.Start, OutAppt.End, OutAppt.Subject)
    Next
    FindAppts = results
End Function

