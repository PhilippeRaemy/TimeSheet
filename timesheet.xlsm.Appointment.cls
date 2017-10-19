VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Appointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public StartTime As Date
Public EndTime As Date
Public Subject As String

Public Property Get StartTick(TickDuration As Single) As Integer
    StartTick = Tick(StartTime, TickDuration)
End Property

Public Property Get EndTick(TickDuration As Single) As Integer
    EndTick = Tick(EndTime, TickDuration) - 1
End Property

Private Function Tick(datetime As Date, TickDuration As Single) As Integer
    Tick = 1 + Round((datetime - Int(datetime)) * 24 / TickDuration)
End Function

Public Function ToString() As String
    ToString = Format(StartTime, "yyyy-MM-dd") & " [" & Format(StartTime, "HH:mm") & "-" & Format(EndTime, "HH:mm") & "]: " & Subject
End Function

Public Function NewAppointment(StartTime As Date, EndTime As Date, Subject As String) As Appointment
    Set NewAppointment = New Appointment
    NewAppointment.StartTime = StartTime
    NewAppointment.EndTime = EndTime
    NewAppointment.Subject = Subject
End Function


