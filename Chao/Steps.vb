Public Class Steps
    Public step_str As String
    Public step_label As Label

    Public NextStep As Steps
    Public LastStep As Boolean
    Public Time As Integer

    Sub New(ByRef ss As String, ByRef sl As Label, ByRef ns As Steps, ByRef ls As Boolean, ByVal sec As Integer)
        step_str = ss
        step_label = sl
        NextStep = ns
        LastStep = ls
        Time = sec
    End Sub

    Public Function HasNext()
        Return Not LastStep
    End Function
End Class
