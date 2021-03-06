﻿'Run_Unit is a class that represents each big step on the main testing page. Each light(green or red) is represented by one 
'Run_Unit object
Public Class Run_Unit

    Public WithEvents Link As LinkLabel
    Dim Panel As Panel
    'Dim Data As List(Of Double)
    Public NextUnit As Run_Unit
    Public PrevUnit As Run_Unit
    Public Time As Integer
    Public Steps As Steps
    Public HeadStep As Steps
    Public CurStep As Integer
    Public StartStep As Integer
    Public EndStep As Integer
    Public Name As String
    Public Executed As Boolean = False

    Public GRU As Grid_Run_Unit

    Sub New(ByRef l As LinkLabel, ByRef p As Panel, ByRef nu As Run_Unit, ByRef pu As Run_Unit, ByVal sec As Integer, ByRef na As String, ByVal cur As Integer, ByVal start As Integer, ByVal en As Integer)
        Link = l
        Panel = p
        NextUnit = nu
        PrevUnit = pu
        Time = sec
        'Data = New List(Of Double)
        CurStep = cur
        StartStep = start
        EndStep = en
        Name = na
    End Sub

    Public Function HasNextStep()
        Return Not CurStep = EndStep
    End Function

    Public Sub Deallocate()
        Link = Nothing
        Panel = Nothing
        'Dim Data As List(Of Double)
        NextUnit = Nothing
        PrevUnit = Nothing
        Time = Nothing
        Steps = Nothing
        HeadStep = Nothing
        CurStep = Nothing
        StartStep = Nothing
        EndStep = Nothing
        Name = Nothing
        Executed = Nothing
        If GRU IsNot Nothing Then
            GRU.ParentRU = Nothing
        End If
        GRU = Nothing
    End Sub

    Sub Set_BackColor(ByVal c As Color)
        Panel.BackColor = c
    End Sub

    '被按下去後的反應
    Private Sub LinkLabel_Clicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles Link.LinkClicked
        Program.startButton.Enabled = True
        Program.Accept_Button.Enabled = False
        Program.All_Panel_Disable()

        If Me.Link.Name = "LinkLabel_PreCal_1st" Then
            Countdown_False_before_Jump()
            'jump
            Program.CurRun = Program.PreCal_1st
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_PreCal_2nd" Then
            Countdown_False_before_Jump()
            'jump
            Program.CurRun = Program.PreCal_2nd
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_PreCal_3rd" Then
            Countdown_False_before_Jump()
            'jump
            Program.CurRun = Program.PreCal_3rd
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_PreCal_4th" Then
            Countdown_False_before_Jump()
            'jump
            Program.CurRun = Program.PreCal_4th
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_PreCal_5th" Then
            Countdown_False_before_Jump()
            'jump
            Program.CurRun = Program.PreCal_5th
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_PreCal_6th" Then
            Countdown_False_before_Jump()
            'jump
            Program.CurRun = Program.PreCal_6th
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_BG" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.BG
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_RSS" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.RSS
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_PostCal_1st" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.PostCal_1st
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_PostCal_2nd" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.PostCal_2nd
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_PostCal_3rd" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.PostCal_3rd
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_PostCal_4th" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.PostCal_4th
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_PostCal_5th" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.PostCal_5th
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_PostCal_6th" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.PostCal_6th
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_ExA1_Fst_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.ExA1_Fst
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_ExA1_Sec_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.ExA1_Sec
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_ExA1_Thd_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.ExA1_Thd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_ExA1_Add_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.ExA1_Add
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_ExA2_Fst_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.ExA2_Fst
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_ExA2_Sec_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.ExA2_Sec
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_ExA2_Thd_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.ExA2_Thd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_ExA2_Add_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.ExA2_Add
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA1_Fst_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA1_Fst
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA1_Sec_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA1_Sec
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA1_Thd_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA1_Thd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA1_Add_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA1_Add
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA2_Fst_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA2_Fst
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA2_Sec_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA2_Sec
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA2_Thd_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA2_Thd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA2_Add_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA2_Add
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA3_Fst_fwd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA3_Fst_fwd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA3_Fst_bkd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA3_Fst_bkd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA3_Sec_fwd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA3_Sec_fwd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA3_Sec_bkd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA3_Sec_bkd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA3_Thd_fwd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA3_Thd_fwd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA3_Thd_bkd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA3_Thd_bkd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA3_Add_fwd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA3_Add_fwd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_LoA3_Add_bkd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.LoA3_Add_bkd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_TrA1_Fst_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.TrA1_Fst
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_TrA1_Sec_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.TrA1_Sec
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_TrA1_Thd_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.TrA1_Thd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_TrA1_Add_1st" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.TrA1_Add
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_TrA3_Fst_fwd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.TrA3_Fst_fwd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_TrA3_Fst_bkd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.TrA3_Fst_bkd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_TrA3_Sec_fwd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.TrA3_Sec_fwd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_TrA3_Sec_bkd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.TrA3_Sec_bkd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_TrA3_Thd_fwd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.TrA3_Thd_fwd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_TrA3_Thd_bkd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.TrA3_Thd_bkd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_TrA3_Add_fwd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.TrA3_Add_fwd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_TrA3_Add_bkd" Then
            Countdown_True_before_Jump()
            Program.CurRun = Program.TrA3_Add_bkd
            Countdown_True_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_A4_Fst" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.A4_Fst
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_A4_Sec" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.A4_Sec
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_A4_Thd" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.A4_Thd
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_A4_Add" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.A4_Add
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_A4_Sec_Mid_Background" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.A4_Sec_Mid
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_A4_Thd_Mid_Background" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.A4_Thd_Mid
            Countdown_False_after_Jump()
        ElseIf Me.Link.Name = "LinkLabel_A4_Add_Mid_Background" Then
            Countdown_False_before_Jump()
            Program.CurRun = Program.A4_Add_Mid
            Countdown_False_after_Jump()
        End If
        Program.SendNonChangeSignalToMobile()
    End Sub

    '在跳步驟前，如果要跳往的步驟是以倒數的方式計時所要叫的function
    Sub Countdown_True_before_Jump()
        If Program.CurRun IsNot Nothing Then
            If Program.CurRun.Name = "ExA2_2nd_3rd" Or Program.CurRun.Name = "LoA2_2nd_3rd" Or Program.CurRun.Name = "ExA2_2nd_3rd_Add" Or Program.CurRun.Name = "LoA2_2nd_3rd_Add" Then
                If Program.CurRun.PrevUnit.Name = "ExA2_1st" Or Program.CurRun.PrevUnit.Name = "LoA2_1st" Or Program.CurRun.PrevUnit.Name = "ExA2_1st_Add" Or Program.CurRun.PrevUnit.Name = "LoA2_1st_Add" Then
                    Program.CurRun.Set_BackColor(Color.IndianRed)
                    Program.CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
                    Program.CurRun.CurStep = 1
                    Program.CurRun.PrevUnit.CurStep = 1
                    Program.Temp_CurRun = Program.CurRun.PrevUnit
                ElseIf Program.CurRun.PrevUnit.Name = "ExA2_2nd_3rd" Or Program.CurRun.PrevUnit.Name = "LoA2_2nd_3rd" Or Program.CurRun.PrevUnit.Name = "ExA2_2nd_3rd_Add" Or Program.CurRun.PrevUnit.Name = "LoA2_2nd_3rd_Add" Then
                    Program.CurRun.Set_BackColor(Color.IndianRed)
                    Program.CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
                    Program.CurRun.PrevUnit.PrevUnit.Set_BackColor(Color.IndianRed)
                    Program.CurRun.CurStep = 1
                    Program.CurRun.PrevUnit.CurStep = 1
                    Program.CurRun.PrevUnit.PrevUnit.CurStep = 1
                    Program.Temp_CurRun = Program.CurRun.PrevUnit.PrevUnit
                End If
            ElseIf Program.CurRun.Name = "ExA2_1st" Or Program.CurRun.Name = "ExA2_1st_Add" Or Program.CurRun.Name = "LoA2_1st" Or Program.CurRun.Name = "LoA2_1st_Add" Then
                Program.CurRun.Set_BackColor(Color.IndianRed)
                Program.Temp_CurRun = Program.CurRun
            Else
                Program.CurRun.Set_BackColor(Color.IndianRed)
                Program.Temp_CurRun = Program.CurRun
            End If
        Else
            Program.Temp_CurRun = Nothing
        End If
        Program.Temp_Countdown = Program.Countdown

    End Sub

    '在跳步驟後，如果要跳往的步驟是以倒數的方式計時所要叫的function
    Sub Countdown_True_after_Jump()
        Program.Countdown = True
        Program.Load_Array_Step_S()
        Program.LoadInputTime(3, Nothing)
        Program.CurRun.CurStep = 1
        Program.CurRun.Steps = Program.CurRun.HeadStep
        For index = Program.CurRun.StartStep To Program.CurRun.EndStep
            If Program.CurRun.Steps.Time = -1 Then
                Program.CurRun.Steps = Program.CurRun.Steps.NextStep
                Program.CurRun.CurStep += 1
            Else
                Exit For
            End If
        Next

        Program.CurStep = Program.CurRun.HeadStep
        Program.timeLeft = Program.CurRun.Steps.Time
        Program.timeLabel.Text = Program.timeLeft & " s"
        Program.CurRun.Set_BackColor(Color.Yellow)
        If Program.CurRun.Name = "ExA2_1st" Or Program.CurRun.Name = "ExA2_1st_Add" Or Program.CurRun.Name = "LoA2_1st" Or Program.CurRun.Name = "LoA2_1st_Add" Then
            Program.CurRun.NextUnit.Set_BackColor(Color.IndianRed)
            Program.CurRun.NextUnit.CurStep = 1
            Program.CurRun.NextUnit.NextUnit.Set_BackColor(Color.IndianRed)
            Program.CurRun.NextUnit.NextUnit.CurStep = 1
        End If

        'load steps
        Program.Clear_Steps()
        Program.Load_Steps()
        For index = Program.CurRun.StartStep To Program.CurRun.EndStep
            Program.array_step_display(index - 1).Text = Program.CurStep.Time
            If Program.array_step_display(index - 1).Text = "-1" Then
                Program.array_step_display(index - 1).Text = "0"
            End If
            If Program.CurStep.HasNext() Then
                Program.CurStep = Program.CurStep.NextStep
            End If
        Next
        Program.CurStep = Program.CurRun.HeadStep

        'dispose old graph and create new graph
        Program.Load_New_Graph_CD_True()
    End Sub

    '在跳步驟前，如果要跳往的步驟是不是以倒數的方式計時所要叫的function
    Sub Countdown_False_before_Jump()
        'before jump       
        If Program.CurRun IsNot Nothing Then

            If Program.CurRun.Name = "ExA2_2nd_3rd" Or Program.CurRun.Name = "LoA2_2nd_3rd" Or Program.CurRun.Name = "ExA2_2nd_3rd_Add" Or Program.CurRun.Name = "LoA2_2nd_3rd_Add" Then
                If Program.CurRun.PrevUnit.Name = "ExA2_1st" Or Program.CurRun.PrevUnit.Name = "LoA2_1st" Or Program.CurRun.PrevUnit.Name = "ExA2_1st_Add" Or Program.CurRun.PrevUnit.Name = "LoA2_1st_Add" Then
                    Program.CurRun.Set_BackColor(Color.IndianRed)
                    Program.CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
                    Program.CurRun.PrevUnit.Link.Enabled = False
                    Program.CurRun.CurStep = 1
                    Program.CurRun.PrevUnit.CurStep = 1
                    Program.Temp_CurRun = Program.CurRun.PrevUnit
                ElseIf Program.CurRun.PrevUnit.Name = "ExA2_2nd_3rd" Or Program.CurRun.PrevUnit.Name = "LoA2_2nd_3rd" Or Program.CurRun.PrevUnit.Name = "ExA2_2nd_3rd_Add" Or Program.CurRun.PrevUnit.Name = "LoA2_2nd_3rd_Add" Then
                    Program.CurRun.Set_BackColor(Color.IndianRed)
                    Program.CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
                    Program.CurRun.PrevUnit.PrevUnit.Set_BackColor(Color.IndianRed)
                    Program.CurRun.PrevUnit.PrevUnit.Link.Enabled = False
                    Program.CurRun.CurStep = 1
                    Program.CurRun.PrevUnit.CurStep = 1
                    Program.CurRun.PrevUnit.PrevUnit.CurStep = 1
                    Program.Temp_CurRun = Program.CurRun.PrevUnit.PrevUnit
                End If
            ElseIf Program.CurRun.Name = "ExA2_1st" Or Program.CurRun.Name = "ExA2_1st_Add" Or Program.CurRun.Name = "LoA2_1st" Or Program.CurRun.Name = "LoA2_1st_Add" Then
                Program.CurRun.Set_BackColor(Color.IndianRed)
                Program.CurRun.Link.Enabled = False
                Program.Temp_CurRun = Program.CurRun
            Else
                Program.CurRun.Set_BackColor(Color.IndianRed)
                Program.Temp_CurRun = Program.CurRun
            End If
        Else
            Program.Temp_CurRun = Nothing
        End If
        Program.Temp_Countdown = Program.Countdown

    End Sub

    '在跳步驟後，如果要跳往的步驟不是以倒數的方式計時所要叫的function
    Sub Countdown_False_after_Jump()
        Program.Countdown = False
        Program.Clear_Steps()
        For index = 0 To 8
            Program.array_step_display(index).Text = ""
        Next
        'dispose old graph and create new graph
        Program.Load_New_Graph_CD_False()
        Program.timeLeft = Program.CurRun.Time
        Program.timeLabel.Text = Program.timeLeft & " s"
        Program.CurRun.Set_BackColor(Color.Yellow)
    End Sub
End Class


