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
    Dim filter As String
    Dim OutApp As New Outlook.Application

    filter = "[Start] >= '" & _
        Format$(FromDate, "mm/dd/yyyy hh:mm AMPM") _
        & "' AND [End] <= '" & _
        Format$(ToDate, "mm/dd/yyyy hh:mm AMPM") & "'"
    Debug.Print filter

    Set CalFolder = OutApp.Session.GetDefaultFolder(olFolderCalendar)
    Set OutItems = CalFolder.Items
    OutItems.IncludeRecurrences = True
    OutItems.Sort "[Start]"
    Set OutItemsInDateRange = OutItems.Restrict(filter)
    Dim apptFactory  As Appointment: Set apptFactory = New Appointment
    Dim results() As Appointment, apptCount As Integer
    For Each OutAppt In OutItemsInDateRange
        ReDim Preserve results(apptCount)
        apptCount = apptCount + 1
        Set results(UBound(results)) = apptFactory.NewAppointment(OutAppt.Start, OutAppt.End, OutAppt.Subject)
    Next
    FindAppts = results
End Function

