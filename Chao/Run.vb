Public Class Run_Unit

    Public WithEvents Link As LinkLabel
    Dim Panel As Panel
    Dim Data As List(Of Double)
    Public NextUnit As Run_Unit
    Public PrevUnit As Run_Unit
    Public Time As Integer
    Public Steps As Steps
    Public HeadStep As Steps
    Public CurStep As Integer
    Public StartStep As Integer
    Public EndStep As Integer
    Public Name As String

    Sub New(ByRef l As LinkLabel, ByRef p As Panel, ByRef nu As Run_Unit, ByRef pu As Run_Unit, ByVal sec As Integer, ByRef na As String, ByVal cur As Integer, ByVal start As Integer, ByVal en As Integer)
        Link = l
        Panel = p
        NextUnit = nu
        PrevUnit = pu
        Time = sec
        Data = New List(Of Double)
        CurStep = cur
        StartStep = start
        EndStep = en
        Name = na
    End Sub

    Public Function HasNext()
        Return Not CurStep = EndStep
    End Function

    Sub Append_Data(ByVal d() As Double)
        If Not IsNothing(d) And d.Length > 0 Then
            For i = 0 To d.Length - 1
                Data.Add(d(i))
            Next
        End If
    End Sub

    Sub Set_BackColor(ByVal c As Color)
        Panel.BackColor = c
    End Sub

    Private Sub LinkLabel_Clicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles Link.LinkClicked

    End Sub
    'Function L_Average()

    'End Function

    'Function Set_Time()

    'End Function
End Class
