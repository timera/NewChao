Public Class Steps
    Dim step_str As String
    Dim step_label As Label
    Public NextStep As Steps
    Public LastStep As Boolean

    Sub New(ByRef ss As String, ByRef sl As Label, ByRef ns As Steps, ByRef ls As Boolean)
        step_str = ss
        step_label = sl
        NextStep = ns
        LastStep = ls
    End Sub

    Public Function HasNext()
        Return Not LastStep
    End Function
End Class
