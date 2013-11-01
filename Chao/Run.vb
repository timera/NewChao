﻿Public Class Run_Unit
    Dim Link As LinkLabel
    Dim Panel As Panel
    Dim Data As List(Of Double)
    Public NextUnit As Run_Unit
    Public PrevUnit As Run_Unit
    Dim Time As Integer
    Public Steps(8) As String
    Public CurStep As Integer
    Public StartStep As Integer
    Public EndStep As Integer

    Sub New(ByRef l As LinkLabel, ByRef p As Panel, ByRef nu As Run_Unit, ByRef pu As Run_Unit, ByVal sec As Integer)
        Link = l
        Panel = p
        NextUnit = nu
        PrevUnit = pu
        Time = sec
        Data = New List(Of Double)
        CurStep = 0
        StartStep = 0
        EndStep = 0
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

    Function L_Average()

    End Function

    Function Set_Time()

    End Function
End Class
