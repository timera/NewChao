Imports System.Drawing.Printing
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Partial Public Class Program

    '當要run A2 or A3 時，我們需要先測量即將跑的時間，這時就會call SetTimerTabOn(true), 要回到普通畫面則call SetTimerTabOff(false)
    Sub SetTimerTabOn(ByVal timerOn As Boolean)
        'Tabcontrol_Changed(timerOn=true)=>jump to tabpagetimer(enable) and tabpage3(disable),use for input seconds
        'Tabcontrol_Changed(timerOn=false)=>jump to tabpageProcedure(enable) and tabpage4(disable),use for button confirm
        Dim tabpage As TabPage
        '先將此STEP的步驟LOAD出來
        CurStep = Load_Steps_helper(CurRun)
        If timerOn Then 'TIMER模式
            Me.TabControl2.SelectedIndex = 1
            For Each tabpage In TabControl2.TabPages
                If tabpage.Name = "TabPageProcedure" Then
                    tabpage.Enabled = False
                ElseIf tabpage.Name = "TabPageTimer" Then
                    tabpage.Enabled = True

                End If
            Next
            Test_StartButton.Enabled = True
            Test_NextButton.Enabled = False
            Test_ConfirmButton.Enabled = False
            For i = 0 To 8
                array_step_s(i).Text = 0
                array_step_s(i).Enabled = False
            Next
            For i = 0 To CurRun.EndStep - 1
                If Not CurStep.Time = -1 Then
                    array_step_s(i).Enabled = True
                End If
                CurStep = CurStep.NextStep
            Next
        ElseIf timerOn = False Then '回到原來畫面模式
            Me.TabControl2.SelectedIndex = 0
            For Each tabpage In TabControl2.TabPages
                If tabpage.Name = "TabPageTimer" Then
                    tabpage.Enabled = False
                ElseIf tabpage.Name = "TabPageProcedure" Then
                    tabpage.Enabled = True
                    stopButton.Enabled = False
                    startButton.Enabled = True
                End If
            Next
            For i = 0 To 8
                array_light(i).BackColor = Color.DarkGray
            Next
        End If
    End Sub

    '雖然VB有garbage collector, 但因為RUN_UNIT裡有指到前後RUN_UNIT的POINTER，讓VB的garbage collector 無法正常得將釋放出的RUN_UNIT殺掉
    '所以需要此步驟的幫忙才能釋出舊的RUN_UNIT的MEMORY
    Private Sub DisposeAllRuns()
        If HeadRun IsNot Nothing Then
            Dim tempcurRun As Run_Unit = HeadRun
            Dim tempRU As Run_Unit
            While tempcurRun.NextUnit IsNot Nothing
                tempRU = tempcurRun.NextUnit
                tempcurRun.Deallocate()
                tempcurRun = Nothing
                tempcurRun = tempRU
            End While
            tempcurRun.Deallocate()
            tempcurRun = Nothing
        End If
    End Sub

    '要換機具時要叫此FUNCTION，他會先問是否要儲存，然後再將需要歸零的歸零，更新的更新
    Private Function Change_Machine()
        If SaveToolStripMenuItem.Enabled Then
            Dim answer As DialogResult = MessageBox.Show("尚未儲存資料，是否儲存?", "Save?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If answer = DialogResult.Yes Then
                If Save() Then
                    SaveToolStripMenuItem.Enabled = False
                Else
                    Return False
                End If
            ElseIf answer = DialogResult.Cancel Then
                Return False
            End If
        End If
        'can choose machine again
        BackToZero()



        'dispose graph
        If MainLineGraph IsNot Nothing Then
            MainLineGraph.Dispose()
        End If
        If MainBarGraph IsNot Nothing Then
            MainBarGraph.Dispose()
        End If

        'dispose chart
        DisposeChart()
        DisposeAllRuns()
        'Set invisible property
        Panel_PreCal.Visible = False
        Panel_Bkg.Visible = False
        Panel_RSS.Visible = False
        Panel_PostCal.Visible = False
        Panel_PreCal_Sub.Visible = False
        Panel_PostCal_Sub.Visible = False
        PanelA4.Visible = False
        PanelExcavatorA1.Visible = False
        PanelExcavatorA2.Visible = False
        PanelLoaderA3.Visible = False
        PanelLoaderA1.Visible = False
        PanelLoaderA2.Visible = False
        PanelTractorA1.Visible = False
        PanelTractorA3.Visible = False

        'set groupbox enable property
        GroupBox_A1_A2_A3.Enabled = False
        GroupBox_A4.Enabled = False

        'set TabPage enable property
        Me.TabPageProcedure.Enabled = True
        Me.TabControl2.SelectedIndex = 0

        'set button enable property
        startButton.Enabled = True
        stopButton.Enabled = False
        Accept_Button.Enabled = False
        Test_NextButton.Enabled = False
        Test_StartButton.Enabled = False
        Test_ConfirmButton.Enabled = False

        'clear textbox content
        TextBox_L.Text = Nothing
        TextBox_L1.Text = Nothing
        TextBox_L2.Text = Nothing
        TextBox_L3.Text = Nothing
        TextBox_r1.Text = Nothing
        TextBox_r2.Text = Nothing

        Countdown = False

        For i = 0 To array_step_s.Length - 1
            array_step_s(i).Text = 0
            array_step_s(i).Enabled = False
        Next

        For i = 0 To array_step_display.Length - 1
            array_step_display(i).Text = ""
        Next

        For i = 0 To array_time.Length - 1
            array_time(i) = 0
        Next

        A4_step_text = Nothing

        Label_A4_Hint1.Visible = False
        Label_A4_Hint2.Visible = False
        Return True
    End Function


    'Given a machine name, it will decide what type of machine it is based on the selection, if not valid machine name, then set machine = nothing, return ""
    Private Function Decide_Machine() As String
        If choice = "開挖機(Excavator)" Then
            Picture_machine.Image = My.Resources.小型開挖機_compact_excavator_
            Machine = Machines.Excavator
            Return "A1,A2"
        ElseIf choice = "推土機(Crawler and wheel tractor)" Then  'A1+A3
            Picture_machine.Image = My.Resources.履帶式推土機_crawler_dozer_
            Machine = Machines.Tractor
            Return "A1,A3"
        ElseIf choice = "鐵輪壓路機(Road roller)" Then
            Picture_machine.Image = My.Resources.壓路機_rollers_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "膠輪壓路機(Wheel roller)" Then
            Picture_machine.Image = My.Resources.壓路機_rollers_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "振動式壓路機(Vibrating roller)" Then
            Picture_machine.Image = My.Resources.壓路機_rollers_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "裝料機(Crawler and wheel loader)" Then
            Picture_machine.Image = My.Resources.履帶式裝料機_crawler_loader_
            Machine = Machines.Loader
            Return "A1,A2,A3"
        ElseIf choice = "裝料開挖機" Then
            Picture_machine.Image = My.Resources.履帶式開挖裝料機_crawler_backhoe_loader_
            Machine = Machines.Loader_Excavator
            Return "A1,A2,A1,A2,A3"
        ElseIf choice = "履帶起重機(Crawler crane)" Then
            Picture_machine.Image = My.Resources.履帶起重機
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "卡車起重機(Truck crane)" Then
            Picture_machine.Image = My.Resources.卡車起重機
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "輪形起重機(Wheel crane)" Then
            Picture_machine.Image = My.Resources.輪式起重機
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "振動式樁錘(Vibrating hammer)" Then
            Picture_machine.Image = My.Resources.vibrating_hammer
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "油壓式打樁機(Hydraulic pile driver)" Then
            Picture_machine.Image = My.Resources.油壓式打樁機_Hydraulic_pile_driver_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "拔樁機" Then
            Picture_machine.Image = My.Resources.拔樁機
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "油壓式拔樁機" Then
            Picture_machine.Image = My.Resources.油壓式打樁機_Hydraulic_pile_driver_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "土壤取樣器(地鑽) (Earth auger)" Then
            Picture_machine.Image = My.Resources.土壤取樣器_地鑽___Earth_auger_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "全套管鑽掘機" Then
            Picture_machine.Image = My.Resources.全套管鑽掘機
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "鑽土機(Earth drill)" Then
            Picture_machine.Image = My.Resources.鑽土機_Earth_drill_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "鑽岩機(Rock breaker)" Then
            Picture_machine.Image = My.Resources.鑽岩機_Rock_breaker_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "混凝土泵車(Concrete pump)" Then
            Picture_machine.Image = My.Resources.混凝土泵車_Concrete_pump_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "混凝土破碎機(Concrete breaker)" Then
            Picture_machine.Image = My.Resources.混凝土破碎機_Concrete_breaker_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "瀝青混凝土舖築機(Asphalt finisher)" Then
            Picture_machine.Image = My.Resources.瀝青混凝土舖築機_Asphalt_finisher_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "混凝土割切機(Concrete cutter)" Then
            Picture_machine.Image = My.Resources.混凝土割切機_Concrete_cutter_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "發電機(Generator)" Then
            Picture_machine.Image = My.Resources.發電機_Generator_
            Machine = Machines.Others
            Return "A4"
        ElseIf choice = "空氣壓縮機(Compressor)" Then
            Picture_machine.Image = My.Resources.空氣壓縮機_Compressor_
            Machine = Machines.Others
            Return "A4"
        Else
            Machine = Nothing

        End If
        Return ""
    End Function

    'given the radius, it creates the chart for recording data
    Private Sub CreateChart(ByVal r As Double)
        DataGrid = New Grid(TabPageCharts, A_Unit_Size, New Point(10, 10), r, Machine, HeadRun)
    End Sub

    'DataGrid 是在第三個TAB中的主要表格，在要換新表格前我們要call此function去移除就表格
    Private Sub DisposeChart()
        If DataGrid IsNot Nothing Then
            Try
                DataGrid.Form.Dispose()
                DataGrid = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub



    '要測Tractor, Loader, Excavator, or Loader Excavator 時，將r輸入後要叫此function
    Private Sub A123_Prepare()
        'Set up Plotting
        'Points = New ArrayList
        pos = {New CoorPoint(), New CoorPoint(), New CoorPoint(), New CoorPoint(), New CoorPoint(), New CoorPoint()}

        Dim r1 As Integer

        If TextBox_L.Text < 1.5 Then
            TextBox_r1.Text = "4"
            r1 = 4
        End If
        If TextBox_L.Text >= 1.5 And TextBox_L.Text < 4 Then
            TextBox_r1.Text = "10"
            r1 = 10
        End If
        If TextBox_L.Text >= 4 Then
            TextBox_r1.Text = "16"
            r1 = 16
        End If
        pos(0).Coors = New ThreeDPoint(r1 * 0.7, r1 * 0.7, 1.5)
        pos(1).Coors = New ThreeDPoint(-1 * r1 * 0.7, r1 * 0.7, 1.5)
        pos(2).Coors = New ThreeDPoint(-1 * r1 * 0.7, -1 * r1 * 0.7, 1.5)
        pos(3).Coors = New ThreeDPoint(r1 * 0.7, -1 * r1 * 0.7, 1.5)
        pos(4).Coors = New ThreeDPoint(r1 * -0.27, r1 * 0.65, r1 * 0.71)
        pos(5).Coors = New ThreeDPoint(r1 * 0.27, r1 * -0.65, r1 * 0.71)
        R = r1
        GP.Clear(Color.DarkGray)
        plot(False, GP, xCor, yCor, GroupBox_Plot)

        If choice = "開挖機(Excavator)" Then
            Load_Excavator()
            MachChosen = True
        ElseIf choice = "推土機(Crawler and wheel tractor)" Then  'A1+A3
            Load_Tractor()
            MachChosen = True
        ElseIf choice = "裝料機(Crawler and wheel loader)" Then
            Load_Loader()
            MachChosen = True
        ElseIf choice = "裝料開挖機" Then
            Load_Loader_Excavator()
            MachChosen = True
        End If

        CreateChart(r1)
        '選擇機具後可更換
        If MachChosen = True Then
            Button_change_machine.Enabled = True
            ComboBox_machine_list.Enabled = False
            Button_L_check.Enabled = False
        End If
        TextBox_L.Enabled = False
        SendChangeSignalForMobile()
    End Sub

    '測A4機具時要叫此Function
    Private Sub A4_Prepare()
        TextBox_r2.ReadOnly = True
        Label_A4_Hint1.Visible = True
        Label_A4_Hint2.Visible = True
        'Set up Plotting
        pos = {New CoorPoint(), New CoorPoint(), New CoorPoint(), New CoorPoint()}
        Dim L1 As Double
        Dim L2 As Double
        Dim L3 As Double
        Dim r2 As Double
        L1 = TextBox_L1.Text
        L2 = TextBox_L2.Text
        L3 = TextBox_L3.Text
        r2 = Math.Ceiling(2 * Math.Sqrt(((L1 / 2) ^ 2) + ((L2 / 2) ^ 2) + (L3 ^ 2)))
        TextBox_r2.Text = r2



        pos(0).Coors = New ThreeDPoint(r2 * -0.45, r2 * 0.77, r2 * 0.45)
        pos(1).Coors = New ThreeDPoint(r2 * -0.45, r2 * -0.77, r2 * 0.45)
        pos(2).Coors = New ThreeDPoint(r2 * 0.89, 0, r2 * 0.45)
        pos(3).Coors = New ThreeDPoint(0, 0, r2)

        R = r2
        GP.Clear(Color.DarkGray)
        plot(False, GP, xCor, yCor, GroupBox_Plot)
        Load_Others(choice)
        MachChosen = True
        CreateChart(r2)

        '選擇機具後可更換
        Button_change_machine.Enabled = True
        ComboBox_machine_list.Enabled = False
        Button_L1_L2_L3_check.Enabled = False
        TextBox_L1.Enabled = False
        TextBox_L2.Enabled = False
        TextBox_L3.Enabled = False
        SendChangeSignalForMobile()
    End Sub

    '這個是設定主畫面中每個步驟上的字、大小、顏色等 例如:P2,P4,P6
    Sub Set_Panel(ByRef p As Panel, ByRef l As Label)
        If p.Name = "Panel_Bkg" Or p.Name = "Panel_RSS" Or p.Name.Contains("Panel_P") Then
            p.Size = New Size(80, 26)
            If p.Name = "Panel_PreCal_1st" Then
                p.BackColor = Color.Yellow
            Else
                p.BackColor = Color.IndianRed
            End If
            p.Controls.Add(l)
            If p.Name.Contains("Panel_P") Or p.Name = "Panel_RSS" Then
                l.Location = New Point(28, 7)
            Else
                l.Location = New Point(0, 5)
            End If
            l.ForeColor = Color.IndianRed
        ElseIf p.Name.Contains("A1") Or p.Name.Contains("A4_Fst") Then
            p.Size = New Size(43, 106)
            p.BackColor = Color.IndianRed
            p.Controls.Add(l)
            l.Location = New Point(1, 46)
            l.ForeColor = Color.IndianRed
        ElseIf p.Name.Contains("A3") Or p.Name.Contains("A4_Sec") Or p.Name.Contains("A4_Thd") Or p.Name.Contains("A4_Add") Then
            p.Size = New Size(43, 52)
            p.BackColor = Color.IndianRed
            p.Controls.Add(l)
            l.Location = New Point(1, 19)
            l.ForeColor = Color.IndianRed
        Else
            p.Size = New Size(43, 34)
            p.BackColor = Color.IndianRed
            p.Controls.Add(l)
            l.Location = New Point(1, 10)
            l.ForeColor = Color.IndianRed
        End If
    End Sub

    '設定A123 or A4的Precal and Postcal的綠燈，因為A4較少
    Sub LinkLabel_PreCal_PostCal_visible()
        Panel_PreCal_Sub.Visible = True
        Panel_PostCal_Sub.Visible = True

        If Machine = Machines.Others Then
            Panel_PreCal_5th.Visible = False
            Panel_PreCal_6th.Visible = False
            Panel_PostCal_5th.Visible = False
            Panel_PostCal_6th.Visible = False
        Else
            Panel_PreCal_5th.Visible = True
            Panel_PreCal_6th.Visible = True
            Panel_PostCal_5th.Visible = True
            Panel_PostCal_6th.Visible = True
        End If
    End Sub


    Sub LinkLabel_PreCal_PostCal_BG_RSS_enable()
        LinkLabel_PreCal_1st.Enabled = False
        LinkLabel_PreCal_2nd.Enabled = False
        LinkLabel_PreCal_3rd.Enabled = False
        LinkLabel_PreCal_4th.Enabled = False
        LinkLabel_PostCal_1st.Enabled = False
        LinkLabel_PostCal_2nd.Enabled = False
        LinkLabel_PostCal_3rd.Enabled = False
        LinkLabel_PostCal_4th.Enabled = False
        LinkLabel_BG.Enabled = False
        LinkLabel_RSS.Enabled = False

        If Not Machine = Machines.Others Then
            LinkLabel_PreCal_5th.Enabled = False
            LinkLabel_PreCal_6th.Enabled = False
            LinkLabel_PostCal_5th.Enabled = False
            LinkLabel_PostCal_6th.Enabled = False
        Else
            LinkLabel_PreCal_5th.Enabled = True
            LinkLabel_PreCal_6th.Enabled = True
            LinkLabel_PostCal_5th.Enabled = True
            LinkLabel_PostCal_6th.Enabled = True
        End If
    End Sub

    '將前校正的RUN_UNIT先讀入
    Function Load_PreCal(ByRef tempRun As Run_Unit) As Run_Unit
        'Precal

        PreCal_1st = tempRun
        HeadRun = tempRun
        CurRun = HeadRun
        timeLeft = CurRun.Time
        timeLabel.Text = timeLeft & " s"

        Set_Panel(Panel_PreCal_2nd, LinkLabel_PreCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_2nd, Panel_PreCal_2nd, Nothing, tempRun, 0, "PreCal_P4", 0, 0, 0)
        tempRun = tempRun.NextUnit
        PreCal_2nd = tempRun

        Set_Panel(Panel_PreCal_3rd, LinkLabel_PreCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_3rd, Panel_PreCal_3rd, Nothing, tempRun, 0, "PreCal_P6", 0, 0, 0)
        tempRun = tempRun.NextUnit
        PreCal_3rd = tempRun

        Set_Panel(Panel_PreCal_4th, LinkLabel_PreCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_4th, Panel_PreCal_4th, Nothing, tempRun, 0, "PreCal_P8", 0, 0, 0)
        tempRun = tempRun.NextUnit
        PreCal_4th = tempRun

        Set_Panel(Panel_PreCal_5th, LinkLabel_PreCal_5th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_5th, Panel_PreCal_5th, Nothing, tempRun, 0, "PreCal_P10", 0, 0, 0)
        tempRun = tempRun.NextUnit
        PreCal_5th = tempRun

        Set_Panel(Panel_PreCal_6th, LinkLabel_PreCal_6th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_6th, Panel_PreCal_6th, Nothing, tempRun, 0, "PreCal_P12", 0, 0, 0)
        tempRun = tempRun.NextUnit
        PreCal_6th = tempRun

        'Background
        Set_Panel(Panel_Bkg, LinkLabel_BG)
        tempRun.NextUnit = New Run_Unit(LinkLabel_BG, Panel_Bkg, Nothing, tempRun, 0, "Background", 0, 0, 0)
        tempRun = tempRun.NextUnit
        BG = tempRun
        Return tempRun
    End Function

    '將後校正的RUN_UNIT讀入
    Sub Load_PostCal(ByRef temprun As Run_Unit)
        'PostCal
        'Set_Panel(Panel_PostCal, LinkLabel_postCal)
        Set_Panel(Panel_PostCal_1st, LinkLabel_PostCal_1st)
        temprun.NextUnit = New Run_Unit(LinkLabel_PostCal_1st, Panel_PostCal_1st, Nothing, temprun, 0, "PostCal_P2", 0, 0, 0)
        temprun = temprun.NextUnit
        PostCal_1st = temprun

        Set_Panel(Panel_PostCal_2nd, LinkLabel_PostCal_2nd)
        temprun.NextUnit = New Run_Unit(LinkLabel_PostCal_2nd, Panel_PostCal_2nd, Nothing, temprun, 0, "PostCal_P4", 0, 0, 0)
        temprun = temprun.NextUnit
        PostCal_2nd = temprun

        Set_Panel(Panel_PostCal_3rd, LinkLabel_PostCal_3rd)
        temprun.NextUnit = New Run_Unit(LinkLabel_PostCal_3rd, Panel_PostCal_3rd, Nothing, temprun, 0, "PostCal_P6", 0, 0, 0)
        temprun = temprun.NextUnit
        PostCal_3rd = temprun

        Set_Panel(Panel_PostCal_4th, LinkLabel_PostCal_4th)
        temprun.NextUnit = New Run_Unit(LinkLabel_PostCal_4th, Panel_PostCal_4th, Nothing, temprun, 0, "PostCal_P8", 0, 0, 0)
        temprun = temprun.NextUnit
        PostCal_4th = temprun

        Set_Panel(Panel_PostCal_5th, LinkLabel_PostCal_5th)
        temprun.NextUnit = New Run_Unit(LinkLabel_PostCal_5th, Panel_PostCal_5th, Nothing, temprun, 0, "PostCal_P10", 0, 0, 0)
        temprun = temprun.NextUnit
        PostCal_5th = temprun

        Set_Panel(Panel_PostCal_6th, LinkLabel_PostCal_6th)
        temprun.NextUnit = New Run_Unit(LinkLabel_PostCal_6th, Panel_PostCal_6th, Nothing, temprun, 0, "PostCal_P12", 0, 0, 0)
        temprun = temprun.NextUnit
        PostCal_6th = temprun

    End Sub

    '讀入Excavator的必備物件
    Sub Load_Excavator()

        PanelExcavatorA1.Visible = True
        PanelExcavatorA2.Visible = True
        Panel_PreCal.Visible = True
        Panel_Bkg.Visible = True
        Panel_RSS.Visible = True
        Panel_PostCal.Visible = True

        PanelA4.Visible = False
        PanelLoaderA3.Visible = False
        PanelLoaderA1.Visible = False
        PanelLoaderA2.Visible = False
        PanelTractorA1.Visible = False
        PanelTractorA3.Visible = False
        PanelExcavatorA1.Size = ASize
        PanelExcavatorA1.Location = New Point(135, 55)
        PanelExcavatorA2.Size = ASize
        PanelExcavatorA2.Location = New Point(185, 55)

        LinkLabel_PreCal_PostCal_BG_RSS_enable()
        LinkLabel_PreCal_PostCal_visible()

        'A1's linklabel
        LinkLabel_ExA1_Fst_1st.Enabled = False
        LinkLabel_ExA1_Sec_1st.Enabled = False
        LinkLabel_ExA1_Thd_1st.Enabled = False
        LinkLabel_ExA1_Add_1st.Enabled = False
        'A2's linklabel
        LinkLabel_ExA2_Fst_1st.Enabled = False
        LinkLabel_ExA2_Sec_1st.Enabled = False
        LinkLabel_ExA2_Thd_1st.Enabled = False
        LinkLabel_ExA2_Add_1st.Enabled = False



        PanelExcavatorA1.Controls.Add(Panel_ExA1_Fst_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Sec_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Thd_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Add_1st)

        PanelExcavatorA2.Controls.Add(Panel_ExA2_Fst_1st)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Fst_2nd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Fst_3rd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Sec_1st)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Sec_2nd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Sec_3rd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Thd_1st)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Thd_2nd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Thd_3rd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Add_1st)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Add_2nd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Add_3rd)


        'Create an object for each step
        Dim tempRun As Run_Unit
        Set_Panel(Panel_PreCal_1st, LinkLabel_PreCal_1st)
        tempRun = New Run_Unit(LinkLabel_PreCal_1st, Panel_PreCal_1st, Nothing, Nothing, 0, "PreCal_P2", 0, 0, 0)
        tempRun = Load_PreCal(tempRun)

        'Main Steps
        tempRun = Load_Excavator_Helper(tempRun)

        'RSS
        Set_Panel(Panel_RSS, LinkLabel_RSS)
        tempRun.NextUnit = New Run_Unit(LinkLabel_RSS, Panel_RSS, Nothing, tempRun, 0, "RSS", 0, 0, 0)
        tempRun = tempRun.NextUnit
        RSS = tempRun

        Load_PostCal(tempRun)

        MainLineGraph = New LineGraph(TabPage2, CGraph.Modes.A1A2A3, HeadRun.Time)
        MainBarGraph = New BarGraph(New Point(810, 93), New Size(380, 450), TabPage2, CGraph.Modes.A1A2A3)
        For i = 0 To 6
            NoisesArray(i).Visible = True
        Next
    End Sub

    '讀入Excavator的RUN_UNIT
    Function Load_Excavator_Helper(ByRef run As Run_Unit)
        'Create an object for each step
        Dim tempRun As Run_Unit = run
        ''A1
        'A1 First 
        Set_Panel(Panel_ExA1_Fst_1st, LinkLabel_ExA1_Fst_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Fst_1st, Panel_ExA1_Fst_1st, Nothing, tempRun, 5, "ExA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        ExA1_Fst = tempRun
        'A1 Second 
        Set_Panel(Panel_ExA1_Sec_1st, LinkLabel_ExA1_Sec_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Sec_1st, Panel_ExA1_Sec_1st, Nothing, tempRun, 5, "ExA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        ExA1_Sec = tempRun
        'A1 Third 
        Set_Panel(Panel_ExA1_Thd_1st, LinkLabel_ExA1_Thd_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Thd_1st, Panel_ExA1_Thd_1st, Nothing, tempRun, 5, "ExA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        ExA1_Thd = tempRun
        'A1 Add 
        Set_Panel(Panel_ExA1_Add_1st, LinkLabel_ExA1_Add_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Add_1st, Panel_ExA1_Add_1st, Nothing, tempRun, 5, "ExA1_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        ExA1_Add = tempRun

        ''A2
        'A2 First 1
        Set_Panel(Panel_ExA2_Fst_1st, LinkLabel_ExA2_Fst_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Fst_1st, Panel_ExA2_Fst_1st, Nothing, tempRun, 1, "ExA2_1st", 1, 1, 9)
        tempRun = tempRun.NextUnit
        ExA2_Fst = tempRun

        'A2 First 2
        Set_Panel(Panel_ExA2_Fst_2nd, LinkLabel_ExA2_Fst_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Fst_2nd, Panel_ExA2_Fst_2nd, Nothing, tempRun, 1, "ExA2_2nd_3rd", 1, 1, 8)
        tempRun = tempRun.NextUnit

        'A2 First 3
        Set_Panel(Panel_ExA2_Fst_3rd, LinkLabel_ExA2_Fst_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Fst_3rd, Panel_ExA2_Fst_3rd, Nothing, tempRun, 1, "ExA2_2nd_3rd", 1, 1, 8)
        tempRun = tempRun.NextUnit

        'A2 Second 1
        Set_Panel(Panel_ExA2_Sec_1st, LinkLabel_ExA2_Sec_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Sec_1st, Panel_ExA2_Sec_1st, Nothing, tempRun, 1, "ExA2_1st", 1, 1, 9)
        tempRun = tempRun.NextUnit
        ExA2_Sec = tempRun

        'A2 Second 2
        Set_Panel(Panel_ExA2_Sec_2nd, LinkLabel_ExA2_Sec_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Sec_2nd, Panel_ExA2_Sec_2nd, Nothing, tempRun, 1, "ExA2_2nd_3rd", 1, 1, 8)
        tempRun = tempRun.NextUnit

        'A2 Second 3
        Set_Panel(Panel_ExA2_Sec_3rd, LinkLabel_ExA2_Sec_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Sec_3rd, Panel_ExA2_Sec_3rd, Nothing, tempRun, 1, "ExA2_2nd_3rd", 1, 1, 8)
        tempRun = tempRun.NextUnit

        'A2 Third 1
        Set_Panel(Panel_ExA2_Thd_1st, LinkLabel_ExA2_Thd_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Thd_1st, Panel_ExA2_Thd_1st, Nothing, tempRun, 1, "ExA2_1st", 1, 1, 9)
        tempRun = tempRun.NextUnit
        ExA2_Thd = tempRun
        'A2 Third 2
        Set_Panel(Panel_ExA2_Thd_2nd, LinkLabel_ExA2_Thd_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Thd_2nd, Panel_ExA2_Thd_2nd, Nothing, tempRun, 1, "ExA2_2nd_3rd", 1, 1, 8)
        tempRun = tempRun.NextUnit
        'A2 Third 3
        Set_Panel(Panel_ExA2_Thd_3rd, LinkLabel_ExA2_Thd_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Thd_3rd, Panel_ExA2_Thd_3rd, Nothing, tempRun, 1, "ExA2_2nd_3rd", 1, 1, 8)
        tempRun = tempRun.NextUnit

        'A2 Add 1
        Set_Panel(Panel_ExA2_Add_1st, LinkLabel_ExA2_Add_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Add_1st, Panel_ExA2_Add_1st, Nothing, tempRun, 1, "ExA2_1st_Add", 1, 1, 9)
        tempRun = tempRun.NextUnit
        ExA2_Add = tempRun
        'A2 Add 2
        Set_Panel(Panel_ExA2_Add_2nd, LinkLabel_ExA2_Add_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Add_2nd, Panel_ExA2_Add_2nd, Nothing, tempRun, 1, "ExA2_2nd_3rd_Add", 1, 1, 8)
        tempRun = tempRun.NextUnit
        'A2 Add 3
        Set_Panel(Panel_ExA2_Add_3rd, LinkLabel_ExA2_Add_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Add_3rd, Panel_ExA2_Add_3rd, Nothing, tempRun, 1, "ExA2_2nd_3rd_Add", 1, 1, 8)
        tempRun = tempRun.NextUnit

        Return tempRun

    End Function

    '讀入Loader的必備物件
    Sub Load_Loader()

        PanelLoaderA1.Visible = True
        PanelLoaderA2.Visible = True
        PanelLoaderA3.Visible = True
        Panel_PreCal.Visible = True
        Panel_Bkg.Visible = True
        Panel_RSS.Visible = True
        Panel_PostCal.Visible = True
        PanelA4.Visible = False
        PanelExcavatorA1.Visible = False
        PanelExcavatorA2.Visible = False
        PanelTractorA1.Visible = False
        PanelTractorA3.Visible = False
        PanelLoaderA1.Size = ASize
        PanelLoaderA1.Location = New Point(135, 55)
        PanelLoaderA2.Size = ASize
        PanelLoaderA2.Location = New Point(185, 55)
        PanelLoaderA3.Size = ASize
        PanelLoaderA3.Location = New Point(235, 55)

        LinkLabel_PreCal_PostCal_BG_RSS_enable()
        LinkLabel_PreCal_PostCal_visible()

        'A1's linklabel
        LinkLabel_LoA1_Fst_1st.Enabled = False
        LinkLabel_LoA1_Sec_1st.Enabled = False
        LinkLabel_LoA1_Thd_1st.Enabled = False
        LinkLabel_LoA1_Add_1st.Enabled = False
        'A2's linklabel
        LinkLabel_LoA2_Fst_1st.Enabled = False
        LinkLabel_LoA2_Sec_1st.Enabled = False
        LinkLabel_LoA2_Thd_1st.Enabled = False
        LinkLabel_LoA2_Add_1st.Enabled = False
        'A3's linklabel
        LinkLabel_LoA3_Fst_fwd.Enabled = False
        LinkLabel_LoA3_Fst_bkd.Enabled = False
        LinkLabel_LoA3_Sec_fwd.Enabled = False
        LinkLabel_LoA3_Sec_bkd.Enabled = False
        LinkLabel_LoA3_Thd_fwd.Enabled = False
        LinkLabel_LoA3_Thd_bkd.Enabled = False
        LinkLabel_LoA3_Add_fwd.Enabled = False
        LinkLabel_LoA3_Add_bkd.Enabled = False

        PanelLoaderA1.Controls.Add(Panel_LoA1_Fst_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Sec_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Thd_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Add_1st)

        PanelLoaderA2.Controls.Add(Panel_LoA2_Fst_1st)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Fst_2nd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Fst_3rd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Sec_1st)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Sec_2nd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Sec_3rd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Thd_1st)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Thd_2nd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Thd_3rd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Add_1st)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Add_2nd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Add_3rd)

        PanelLoaderA3.Controls.Add(Panel_LoA3_Fst_bkd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Fst_fwd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Sec_bkd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Sec_fwd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Thd_bkd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Thd_bkd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Add_bkd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Add_bkd)

        'Create an object for each step
        Dim tempRun As Run_Unit
        'Precal
        'Set_Panel(Panel_PreCal, LinkLabel_preCal)
        Set_Panel(Panel_PreCal_1st, LinkLabel_PreCal_1st)
        tempRun = New Run_Unit(LinkLabel_PreCal_1st, Panel_PreCal_1st, Nothing, Nothing, 0, "PreCal_P2", 0, 0, 0)
        tempRun = Load_PreCal(tempRun)

        tempRun = Load_Loader_Helper(tempRun)

        'RSS
        Set_Panel(Panel_RSS, LinkLabel_RSS)
        tempRun.NextUnit = New Run_Unit(LinkLabel_RSS, Panel_RSS, Nothing, tempRun, 0, "RSS", 0, 0, 0)
        tempRun = tempRun.NextUnit
        RSS = tempRun

        'PostCal
        Load_PostCal(tempRun)


        MainLineGraph = New LineGraph(TabPage2, CGraph.Modes.A1A2A3, HeadRun.Time)
        MainBarGraph = New BarGraph(New Point(765, 93), New Size(380, 450), TabPage2, CGraph.Modes.A1A2A3)
        For i = 0 To 6
            NoisesArray(i).Visible = True
        Next
    End Sub

    '讀入Loader的RUN_UNIT
    Function Load_Loader_Helper(ByRef run As Run_Unit)
        'Create an object for each step
        Dim tempRun As Run_Unit = run
        ''A1
        'A1 First 
        Set_Panel(Panel_LoA1_Fst_1st, LinkLabel_LoA1_Fst_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Fst_1st, Panel_LoA1_Fst_1st, Nothing, tempRun, 3, "LoA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        LoA1_Fst = tempRun

        'A1 Second 
        Set_Panel(Panel_LoA1_Sec_1st, LinkLabel_LoA1_Sec_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Sec_1st, Panel_LoA1_Sec_1st, Nothing, tempRun, 3, "LoA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        LoA1_Sec = tempRun

        'A1 Third 
        Set_Panel(Panel_LoA1_Thd_1st, LinkLabel_LoA1_Thd_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Thd_1st, Panel_LoA1_Thd_1st, Nothing, tempRun, 3, "LoA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        LoA1_Thd = tempRun

        'A1 Add 
        Set_Panel(Panel_LoA1_Add_1st, LinkLabel_LoA1_Add_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Add_1st, Panel_LoA1_Add_1st, Nothing, tempRun, 3, "LoA1_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        LoA1_Add = tempRun

        ''A2
        'A2 First 1
        Set_Panel(Panel_LoA2_Fst_1st, LinkLabel_LoA2_Fst_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Fst_1st, Panel_LoA2_Fst_1st, Nothing, tempRun, 1, "LoA2_1st", 1, 1, 3)
        tempRun = tempRun.NextUnit
        LoA2_Fst = tempRun
        'A2 First 2
        Set_Panel(Panel_LoA2_Fst_2nd, LinkLabel_LoA2_Fst_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Fst_2nd, Panel_LoA2_Fst_2nd, Nothing, tempRun, 1, "LoA2_2nd_3rd", 1, 1, 2)
        tempRun = tempRun.NextUnit
        'A2 First 3
        Set_Panel(Panel_LoA2_Fst_3rd, LinkLabel_LoA2_Fst_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Fst_3rd, Panel_LoA2_Fst_3rd, Nothing, tempRun, 1, "LoA2_2nd_3rd", 1, 1, 2)
        tempRun = tempRun.NextUnit

        'A2 Second 1
        Set_Panel(Panel_LoA2_Sec_1st, LinkLabel_LoA2_Sec_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Sec_1st, Panel_LoA2_Sec_1st, Nothing, tempRun, 1, "LoA2_1st", 1, 1, 3)
        tempRun = tempRun.NextUnit
        LoA2_Sec = tempRun
        'A2 Second 2
        Set_Panel(Panel_LoA2_Sec_2nd, LinkLabel_LoA2_Sec_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Sec_2nd, Panel_LoA2_Sec_2nd, Nothing, tempRun, 1, "LoA2_2nd_3rd", 1, 1, 2)
        tempRun = tempRun.NextUnit
        'A2 Second 3
        Set_Panel(Panel_LoA2_Sec_3rd, LinkLabel_LoA2_Sec_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Sec_3rd, Panel_LoA2_Sec_3rd, Nothing, tempRun, 1, "LoA2_2nd_3rd", 1, 1, 2)
        tempRun = tempRun.NextUnit

        'A2 Third 1
        Set_Panel(Panel_LoA2_Thd_1st, LinkLabel_LoA2_Thd_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Thd_1st, Panel_LoA2_Thd_1st, Nothing, tempRun, 1, "LoA2_1st", 1, 1, 3)
        tempRun = tempRun.NextUnit
        LoA2_Thd = tempRun
        'A2 Third 2
        Set_Panel(Panel_LoA2_Thd_2nd, LinkLabel_LoA2_Thd_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Thd_2nd, Panel_LoA2_Thd_2nd, Nothing, tempRun, 1, "LoA2_2nd_3rd", 1, 1, 2)
        tempRun = tempRun.NextUnit
        'A2 Third 3
        Set_Panel(Panel_LoA2_Thd_3rd, LinkLabel_LoA2_Thd_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Thd_3rd, Panel_LoA2_Thd_3rd, Nothing, tempRun, 1, "LoA2_2nd_3rd", 1, 1, 2)
        tempRun = tempRun.NextUnit

        'A2 Add 1
        Set_Panel(Panel_LoA2_Add_1st, LinkLabel_LoA2_Add_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Add_1st, Panel_LoA2_Add_1st, Nothing, tempRun, 1, "LoA2_1st_Add", 1, 1, 3)
        tempRun = tempRun.NextUnit
        LoA2_Add = tempRun
        'A2 Add 2
        Set_Panel(Panel_LoA2_Add_2nd, LinkLabel_LoA2_Add_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Add_2nd, Panel_LoA2_Add_2nd, Nothing, tempRun, 1, "LoA2_2nd_3rd_Add", 1, 1, 2)
        tempRun = tempRun.NextUnit
        'A2 Add 3
        Set_Panel(Panel_LoA2_Add_3rd, LinkLabel_LoA2_Add_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Add_3rd, Panel_LoA2_Add_3rd, Nothing, tempRun, 1, "LoA2_2nd_3rd_Add", 1, 1, 2)
        tempRun = tempRun.NextUnit

        ''A3
        'A3 First forward
        Set_Panel(Panel_LoA3_Fst_fwd, LinkLabel_LoA3_Fst_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Fst_fwd, Panel_LoA3_Fst_fwd, Nothing, tempRun, 1, "LoA3_fwd", 1, 1, 3)
        tempRun = tempRun.NextUnit
        LoA3_Fst_fwd = tempRun

        'A3 First backward
        Set_Panel(Panel_LoA3_Fst_bkd, LinkLabel_LoA3_Fst_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Fst_bkd, Panel_LoA3_Fst_bkd, Nothing, tempRun, 1, "LoA3_bkd", 1, 1, 1)
        tempRun = tempRun.NextUnit
        LoA3_Fst_bkd = tempRun

        'A3 Second fwd
        Set_Panel(Panel_LoA3_Sec_fwd, LinkLabel_LoA3_Sec_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Sec_fwd, Panel_LoA3_Sec_fwd, Nothing, tempRun, 1, "LoA3_fwd", 1, 1, 3)
        tempRun = tempRun.NextUnit
        LoA3_Sec_fwd = tempRun

        'A3 Second backward
        Set_Panel(Panel_LoA3_Sec_bkd, LinkLabel_LoA3_Sec_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Sec_bkd, Panel_LoA3_Sec_bkd, Nothing, tempRun, 1, "LoA3_bkd", 1, 1, 1)
        tempRun = tempRun.NextUnit
        LoA3_Sec_bkd = tempRun


        'A3 Third fwd
        Set_Panel(Panel_LoA3_Thd_fwd, LinkLabel_LoA3_Thd_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Thd_fwd, Panel_LoA3_Thd_fwd, Nothing, tempRun, 1, "LoA3_fwd", 1, 1, 3)
        tempRun = tempRun.NextUnit
        LoA3_Thd_fwd = tempRun

        'A3 Third bkd
        Set_Panel(Panel_LoA3_Thd_bkd, LinkLabel_LoA3_Thd_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Thd_bkd, Panel_LoA3_Thd_bkd, Nothing, tempRun, 1, "LoA3_bkd", 1, 1, 1)
        tempRun = tempRun.NextUnit
        LoA3_Thd_bkd = tempRun

        'A3 Add fwd
        Set_Panel(Panel_LoA3_Add_fwd, LinkLabel_LoA3_Add_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Add_fwd, Panel_LoA3_Add_fwd, Nothing, tempRun, 1, "LoA3_fwd_Add", 1, 1, 3)
        tempRun = tempRun.NextUnit
        LoA3_Add_fwd = tempRun
        'A3 Add bkd
        Set_Panel(Panel_LoA3_Add_bkd, LinkLabel_LoA3_Add_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Add_bkd, Panel_LoA3_Add_bkd, Nothing, tempRun, 1, "LoA3_bkd_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        LoA3_Add_bkd = tempRun

        Return tempRun

    End Function

    '讀入Loader Excavator的必要物件
    Sub Load_Loader_Excavator()
        PanelExcavatorA1.Visible = True
        PanelExcavatorA2.Visible = True
        PanelLoaderA1.Visible = True
        PanelLoaderA2.Visible = True
        PanelLoaderA3.Visible = True
        Panel_PreCal.Visible = True
        Panel_Bkg.Visible = True
        Panel_RSS.Visible = True
        Panel_PostCal.Visible = True
        PanelA4.Visible = False
        PanelTractorA1.Visible = False
        PanelTractorA3.Visible = False

        LinkLabel_PreCal_PostCal_BG_RSS_enable()
        LinkLabel_PreCal_PostCal_visible()

        'A1's linklabel
        LinkLabel_LoA1_Fst_1st.Enabled = False
        LinkLabel_LoA1_Sec_1st.Enabled = False
        LinkLabel_LoA1_Thd_1st.Enabled = False
        LinkLabel_LoA1_Add_1st.Enabled = False
        LinkLabel_ExA1_Fst_1st.Enabled = False
        LinkLabel_ExA1_Sec_1st.Enabled = False
        LinkLabel_ExA1_Thd_1st.Enabled = False
        LinkLabel_ExA1_Add_1st.Enabled = False
        'A2's linklabel
        LinkLabel_ExA2_Fst_1st.Enabled = False
        LinkLabel_ExA2_Sec_1st.Enabled = False
        LinkLabel_ExA2_Thd_1st.Enabled = False
        LinkLabel_ExA2_Add_1st.Enabled = False
        LinkLabel_LoA2_Fst_1st.Enabled = False
        LinkLabel_LoA2_Sec_1st.Enabled = False
        LinkLabel_LoA2_Thd_1st.Enabled = False
        LinkLabel_LoA2_Add_1st.Enabled = False
        'A3's linklabel
        LinkLabel_LoA3_Fst_fwd.Enabled = False
        LinkLabel_LoA3_Fst_bkd.Enabled = False
        LinkLabel_LoA3_Sec_fwd.Enabled = False
        LinkLabel_LoA3_Sec_bkd.Enabled = False
        LinkLabel_LoA3_Thd_fwd.Enabled = False
        LinkLabel_LoA3_Thd_bkd.Enabled = False
        LinkLabel_LoA3_Add_fwd.Enabled = False
        LinkLabel_LoA3_Add_bkd.Enabled = False

        PanelExcavatorA1.Size = ASize
        PanelExcavatorA1.Location = New Point(135, 55)
        PanelExcavatorA2.Size = ASize
        PanelExcavatorA2.Location = New Point(185, 55)
        PanelLoaderA1.Size = ASize
        PanelLoaderA1.Location = New Point(235, 55)
        PanelLoaderA2.Size = ASize
        PanelLoaderA2.Location = New Point(285, 55)
        PanelLoaderA3.Size = ASize
        PanelLoaderA3.Location = New Point(335, 55)

        PanelExcavatorA1.Controls.Add(Panel_ExA1_Fst_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Sec_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Thd_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Add_1st)

        PanelExcavatorA2.Controls.Add(Panel_ExA2_Fst_1st)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Fst_2nd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Fst_3rd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Sec_1st)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Sec_2nd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Sec_3rd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Thd_1st)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Thd_2nd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Thd_3rd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Add_1st)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Add_2nd)
        PanelExcavatorA2.Controls.Add(Panel_ExA2_Add_3rd)

        PanelLoaderA1.Controls.Add(Panel_LoA1_Fst_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Sec_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Thd_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Add_1st)

        PanelLoaderA2.Controls.Add(Panel_LoA2_Fst_1st)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Fst_2nd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Fst_3rd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Sec_1st)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Sec_2nd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Sec_3rd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Thd_1st)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Thd_2nd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Thd_3rd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Add_1st)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Add_2nd)
        PanelLoaderA2.Controls.Add(Panel_LoA2_Add_3rd)

        PanelLoaderA3.Controls.Add(Panel_LoA3_Fst_bkd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Fst_fwd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Sec_bkd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Sec_fwd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Thd_bkd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Thd_bkd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Add_bkd)
        PanelLoaderA3.Controls.Add(Panel_LoA3_Add_bkd)

        'Create an object for each step
        Dim tempRun As Run_Unit
        'Precal
        'Set_Panel(Panel_PreCal, LinkLabel_preCal)
        Set_Panel(Panel_PreCal_1st, LinkLabel_PreCal_1st)
        tempRun = New Run_Unit(LinkLabel_PreCal_1st, Panel_PreCal_1st, Nothing, Nothing, 0, "PreCal_P2", 0, 0, 0)
        tempRun = Load_PreCal(tempRun)

        tempRun = Load_Excavator_Helper(tempRun)
        tempRun = Load_Loader_Helper(tempRun)

        'RSS
        Set_Panel(Panel_RSS, LinkLabel_RSS)
        tempRun.NextUnit = New Run_Unit(LinkLabel_RSS, Panel_RSS, Nothing, tempRun, 0, "RSS", 0, 0, 0)
        tempRun = tempRun.NextUnit
        RSS = tempRun

        Load_PostCal(tempRun)

        MainLineGraph = New LineGraph(TabPage2, CGraph.Modes.A1A2A3, HeadRun.Time)
        MainBarGraph = New BarGraph(New Point(765, 93), New Size(380, 450), TabPage2, CGraph.Modes.A1A2A3)
        For i = 0 To 6
            NoisesArray(i).Visible = True
        Next
    End Sub

    '讀入Tractor必要物件
    Sub Load_Tractor()

        PanelTractorA1.Visible = True
        PanelTractorA3.Visible = True
        Panel_PreCal.Visible = True
        Panel_Bkg.Visible = True
        Panel_RSS.Visible = True
        Panel_PostCal.Visible = True
        PanelA4.Visible = False
        PanelExcavatorA1.Visible = False
        PanelExcavatorA2.Visible = False
        PanelLoaderA3.Visible = False
        PanelLoaderA1.Visible = False
        PanelLoaderA2.Visible = False
        PanelTractorA1.Size = ASize
        PanelTractorA1.Location = New Point(135, 55)
        PanelTractorA3.Size = ASize
        PanelTractorA3.Location = New Point(185, 55)

        LinkLabel_PreCal_PostCal_BG_RSS_enable()
        LinkLabel_PreCal_PostCal_visible()

        'A1's linklabel
        LinkLabel_TrA1_Fst_1st.Enabled = False
        LinkLabel_TrA1_Sec_1st.Enabled = False
        LinkLabel_TrA1_Thd_1st.Enabled = False
        LinkLabel_TrA1_Add_1st.Enabled = False
        'A3's linklabel
        LinkLabel_TrA3_Fst_fwd.Enabled = False
        LinkLabel_TrA3_Fst_bkd.Enabled = False
        LinkLabel_TrA3_Sec_fwd.Enabled = False
        LinkLabel_TrA3_Sec_bkd.Enabled = False
        LinkLabel_TrA3_Thd_fwd.Enabled = False
        LinkLabel_TrA3_Thd_bkd.Enabled = False
        LinkLabel_TrA3_Add_fwd.Enabled = False
        LinkLabel_TrA3_Add_bkd.Enabled = False

        PanelTractorA1.Controls.Add(Panel_TrA1_Fst_1st)
        PanelTractorA1.Controls.Add(Panel_TrA1_Sec_1st)
        PanelTractorA1.Controls.Add(Panel_TrA1_Thd_1st)
        PanelTractorA1.Controls.Add(Panel_TrA1_Add_1st)

        PanelTractorA3.Controls.Add(Panel_TrA3_Fst_bkd)
        PanelTractorA3.Controls.Add(Panel_TrA3_Fst_fwd)
        PanelTractorA3.Controls.Add(Panel_TrA3_Sec_bkd)
        PanelTractorA3.Controls.Add(Panel_TrA3_Sec_fwd)
        PanelTractorA3.Controls.Add(Panel_TrA3_Thd_bkd)
        PanelTractorA3.Controls.Add(Panel_TrA3_Thd_bkd)
        PanelTractorA3.Controls.Add(Panel_TrA3_Add_bkd)
        PanelTractorA3.Controls.Add(Panel_TrA3_Add_bkd)



        'Create an object for each step
        Dim tempRun As Run_Unit

        'Precal
        'Set_Panel(Panel_PreCal, LinkLabel_preCal)
        Set_Panel(Panel_PreCal_1st, LinkLabel_PreCal_1st)
        tempRun = New Run_Unit(LinkLabel_PreCal_1st, Panel_PreCal_1st, Nothing, Nothing, 0, "PreCal_P2", 0, 0, 0)
        tempRun = Load_PreCal(tempRun)

        tempRun = Load_Tractor_Helper(tempRun)

        'RSS
        Set_Panel(Panel_RSS, LinkLabel_RSS)
        tempRun.NextUnit = New Run_Unit(LinkLabel_RSS, Panel_RSS, Nothing, tempRun, 0, "RSS", 0, 0, 0)
        tempRun = tempRun.NextUnit
        RSS = tempRun

        'PostCal
        Load_PostCal(tempRun)

        MainLineGraph = New LineGraph(TabPage2, CGraph.Modes.A1A2A3, HeadRun.Time)
        MainBarGraph = New BarGraph(New Point(765, 93), New Size(380, 450), TabPage2, CGraph.Modes.A1A2A3)
        For i = 0 To 6
            NoisesArray(i).Visible = True
        Next
    End Sub

    '讀入Tractor的RUN_UNIT
    Function Load_Tractor_Helper(ByRef run As Run_Unit)
        'Create an object for each step
        Dim tempRun As Run_Unit = run
        ''A1
        'A1 First 
        Set_Panel(Panel_TrA1_Fst_1st, LinkLabel_TrA1_Fst_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Fst_1st, Panel_TrA1_Fst_1st, Nothing, tempRun, 3, "TrA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        TrA1_Fst = tempRun

        'A1 Second 
        Set_Panel(Panel_TrA1_Sec_1st, LinkLabel_TrA1_Sec_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Sec_1st, Panel_TrA1_Sec_1st, Nothing, tempRun, 3, "TrA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        TrA1_Sec = tempRun

        'A1 Third 
        Set_Panel(Panel_TrA1_Thd_1st, LinkLabel_TrA1_Thd_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Thd_1st, Panel_TrA1_Thd_1st, Nothing, tempRun, 3, "TrA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        TrA1_Thd = tempRun

        'A1 Add 
        Set_Panel(Panel_TrA1_Add_1st, LinkLabel_TrA1_Add_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Add_1st, Panel_TrA1_Add_1st, Nothing, tempRun, 3, "TrA1_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        TrA1_Add = tempRun


        ''A3
        'A3 First forward
        Set_Panel(Panel_TrA3_Fst_fwd, LinkLabel_TrA3_Fst_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Fst_fwd, Panel_TrA3_Fst_fwd, Nothing, tempRun, 3, "TrA3_fwd", 1, 1, 3)
        tempRun = tempRun.NextUnit
        TrA3_Fst_fwd = tempRun
        'tempRun.Steps = Load_Steps_helper(tempRun)
        'A3 First backward
        Set_Panel(Panel_TrA3_Fst_bkd, LinkLabel_TrA3_Fst_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Fst_bkd, Panel_TrA3_Fst_bkd, Nothing, tempRun, 3, "TrA3_bkd", 1, 1, 1)
        tempRun = tempRun.NextUnit
        'tempRun.Steps = Load_Steps_helper(tempRun)
        TrA3_Fst_bkd = tempRun

        'A3 Second fwd
        Set_Panel(Panel_TrA3_Sec_fwd, LinkLabel_TrA3_Sec_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Sec_fwd, Panel_TrA3_Sec_fwd, Nothing, tempRun, 3, "TrA3_fwd", 1, 1, 3)
        tempRun = tempRun.NextUnit
        TrA3_Sec_fwd = tempRun

        'A3 Second backward
        Set_Panel(Panel_TrA3_Sec_bkd, LinkLabel_TrA3_Sec_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Sec_bkd, Panel_TrA3_Sec_bkd, Nothing, tempRun, 3, "TrA3_bkd", 1, 1, 1)
        tempRun = tempRun.NextUnit
        TrA3_Sec_bkd = tempRun

        'A3 Third fwd
        Set_Panel(Panel_TrA3_Thd_fwd, LinkLabel_TrA3_Thd_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Thd_fwd, Panel_TrA3_Thd_fwd, Nothing, tempRun, 3, "TrA3_fwd", 1, 1, 3)
        tempRun = tempRun.NextUnit
        TrA3_Thd_fwd = tempRun

        'A3 Third bkd
        Set_Panel(Panel_TrA3_Thd_bkd, LinkLabel_TrA3_Thd_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Thd_bkd, Panel_TrA3_Thd_bkd, Nothing, tempRun, 3, "TrA3_bkd", 1, 1, 1)
        tempRun = tempRun.NextUnit
        TrA3_Thd_bkd = tempRun

        'A3 Add fwd
        Set_Panel(Panel_TrA3_Add_fwd, LinkLabel_TrA3_Add_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Add_fwd, Panel_TrA3_Add_fwd, Nothing, tempRun, 3, "TrA3_fwd_Add", 1, 1, 3)
        tempRun = tempRun.NextUnit
        TrA3_Add_fwd = tempRun
        'A3 Add bkd
        Set_Panel(Panel_TrA3_Add_bkd, LinkLabel_TrA3_Add_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Add_bkd, Panel_TrA3_Add_bkd, Nothing, tempRun, 3, "TrA3_bkd_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        TrA3_Add_bkd = tempRun

        Return tempRun

    End Function

    '讀入A4的必須物件
    Sub Load_Others(ByVal name As String)

        PanelA4.Visible = True
        Panel_PreCal.Visible = True
        Panel_Bkg.Visible = True
        Panel_RSS.Visible = True
        Panel_PostCal.Visible = True
        PanelExcavatorA1.Visible = False
        PanelExcavatorA2.Visible = False
        PanelLoaderA3.Visible = False
        PanelLoaderA1.Visible = False
        PanelLoaderA2.Visible = False
        PanelTractorA1.Visible = False
        PanelTractorA3.Visible = False
        PanelA4.Size = ASize
        PanelA4.Location = New Point(135, 55)
        PanelA4.Controls.Add(Panel_A4_Fst)
        PanelA4.Controls.Add(Panel_A4_Sec)
        PanelA4.Controls.Add(Panel_A4_Thd)
        PanelA4.Controls.Add(Panel_A4_Add)

        LinkLabel_PreCal_PostCal_BG_RSS_enable()
        LinkLabel_PreCal_PostCal_visible()

        'A4's linklabel
        LinkLabel_A4_Fst.Enabled = False
        LinkLabel_A4_Sec.Enabled = False
        LinkLabel_A4_Thd.Enabled = False
        LinkLabel_A4_Add.Enabled = False
        LinkLabel_A4_Sec_Mid_Background.Enabled = False
        LinkLabel_A4_Thd_Mid_Background.Enabled = False
        LinkLabel_A4_Add_Mid_Background.Enabled = False

        'Create an object for each step
        Dim tempRun As Run_Unit
        'Precal
        Set_Panel(Panel_PreCal_1st, LinkLabel_PreCal_1st)
        tempRun = New Run_Unit(LinkLabel_PreCal_1st, Panel_PreCal_1st, Nothing, Nothing, 0, "PreCal_P2", 0, 0, 0)
        PreCal_1st = tempRun
        HeadRun = tempRun
        CurRun = HeadRun
        timeLeft = CurRun.Time
        timeLabel.Text = timeLeft & " s"

        Set_Panel(Panel_PreCal_2nd, LinkLabel_PreCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_2nd, Panel_PreCal_2nd, Nothing, tempRun, 0, "PreCal_P4", 0, 0, 0)
        tempRun = tempRun.NextUnit
        PreCal_2nd = tempRun

        Set_Panel(Panel_PreCal_3rd, LinkLabel_PreCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_3rd, Panel_PreCal_3rd, Nothing, tempRun, 0, "PreCal_P6", 0, 0, 0)
        tempRun = tempRun.NextUnit
        PreCal_3rd = tempRun

        Set_Panel(Panel_PreCal_4th, LinkLabel_PreCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_4th, Panel_PreCal_4th, Nothing, tempRun, 0, "PreCal_P8", 0, 0, 0)
        tempRun = tempRun.NextUnit
        PreCal_4th = tempRun

        'Background
        Set_Panel(Panel_Bkg, LinkLabel_BG)
        tempRun.NextUnit = New Run_Unit(LinkLabel_BG, Panel_Bkg, Nothing, tempRun, 0, "Background", 0, 0, 0)
        tempRun = tempRun.NextUnit
        BG = tempRun

        'A4 1
        Set_Panel(Panel_A4_Fst, LinkLabel_A4_Fst)
        tempRun.NextUnit = New Run_Unit(LinkLabel_A4_Fst, Panel_A4_Fst, Nothing, tempRun, 0, "A4", 1, 1, 1)
        tempRun = tempRun.NextUnit
        A4_Fst = tempRun

        'A4 2 Mid_Background
        Set_Panel(Panel_A4_Sec_Mid_Background, LinkLabel_A4_Sec_Mid_Background)
        tempRun.NextUnit = New Run_Unit(LinkLabel_A4_Sec_Mid_Background, Panel_A4_Sec_Mid_Background, Nothing, tempRun, 0, "A4_Mid_BG", 0, 0, 0)
        tempRun = tempRun.NextUnit
        A4_Sec_Mid = tempRun

        'A4 2
        Set_Panel(Panel_A4_Sec, LinkLabel_A4_Sec)
        tempRun.NextUnit = New Run_Unit(LinkLabel_A4_Sec, Panel_A4_Sec, Nothing, tempRun, 0, "A4", 1, 1, 1)
        tempRun = tempRun.NextUnit
        A4_Sec = tempRun

        'A4 3 Mid_Background
        Set_Panel(Panel_A4_Thd_Mid_Background, LinkLabel_A4_Thd_Mid_Background)
        tempRun.NextUnit = New Run_Unit(LinkLabel_A4_Thd_Mid_Background, Panel_A4_Thd_Mid_Background, Nothing, tempRun, 0, "A4_Mid_BG", 0, 0, 0)
        tempRun = tempRun.NextUnit
        A4_Thd_Mid = tempRun

        'A4 3
        Set_Panel(Panel_A4_Thd, LinkLabel_A4_Thd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_A4_Thd, Panel_A4_Thd, Nothing, tempRun, 0, "A4", 1, 1, 1)
        tempRun = tempRun.NextUnit
        A4_Thd = tempRun


        'A4 add Mid_Background
        Set_Panel(Panel_A4_Add_Mid_Background, LinkLabel_A4_Add_Mid_Background)
        tempRun.NextUnit = New Run_Unit(LinkLabel_A4_Add_Mid_Background, Panel_A4_Add_Mid_Background, Nothing, tempRun, 0, "A4_Mid_BG_Add", 0, 0, 0)
        tempRun = tempRun.NextUnit
        A4_Add_Mid = tempRun
        'A4 add
        Set_Panel(Panel_A4_Add, LinkLabel_A4_Add)
        tempRun.NextUnit = New Run_Unit(LinkLabel_A4_Add, Panel_A4_Add, Nothing, tempRun, 0, "A4_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        A4_Add = tempRun

        'RSS
        Set_Panel(Panel_RSS, LinkLabel_RSS)
        tempRun.NextUnit = New Run_Unit(LinkLabel_RSS, Panel_RSS, Nothing, tempRun, 0, "RSS", 0, 0, 0)
        tempRun = tempRun.NextUnit
        RSS = tempRun

        'PostCal
        Set_Panel(Panel_PostCal_1st, LinkLabel_PostCal_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_1st, Panel_PostCal_1st, Nothing, tempRun, 0, "PostCal_P2", 0, 0, 0)
        tempRun = tempRun.NextUnit
        PostCal_1st = tempRun

        Set_Panel(Panel_PostCal_2nd, LinkLabel_PostCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_2nd, Panel_PostCal_2nd, Nothing, tempRun, 0, "PostCal_P4", 0, 0, 0)
        tempRun = tempRun.NextUnit
        PostCal_2nd = tempRun

        Set_Panel(Panel_PostCal_3rd, LinkLabel_PostCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_3rd, Panel_PostCal_3rd, Nothing, tempRun, 0, "PostCal_P6", 0, 0, 0)
        tempRun = tempRun.NextUnit
        PostCal_3rd = tempRun

        Set_Panel(Panel_PostCal_4th, LinkLabel_PostCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_4th, Panel_PostCal_4th, Nothing, tempRun, 0, "PostCal_P8", 0, 0, 0)
        tempRun = tempRun.NextUnit
        PostCal_4th = tempRun

        MainLineGraph = New LineGraph(TabPage2, CGraph.Modes.A4, HeadRun.Time)
        MainBarGraph = New BarGraph(New Point(765, 93), New Size(380, 450), TabPage2, CGraph.Modes.A4)
        For i = 0 To 6
            NoisesArray(i).Visible = True
        Next

        'display them on screen-set position and size

        If name = "空氣壓縮機(Compressor)" Then
            A4_step_text = My.Resources.A4_Compressor
        ElseIf name = "混凝土割切機(Concrete cutter)" Then
            A4_step_text = My.Resources.A4_Concrete_Cutter
        ElseIf name = "瀝青混凝土舖築機(Asphalt finisher)" Then
            A4_step_text = My.Resources.A4_Asphalt_Finisher
        ElseIf name = "混凝土破碎機(Concrete breaker)" Then
            A4_step_text = My.Resources.A4_Concrete_Breaker
        ElseIf name = "混凝土泵車(Concrete pump)" Then
            A4_step_text = My.Resources.A4_Concrete_Pump
        ElseIf name = "鑽土機(Earth drill)" Then
            A4_step_text = My.Resources.A4_Auger_Drill_Driver
        ElseIf name = "全套管鑽掘機" Then
            A4_step_text = My.Resources.A4_Auger_Drill_Driver
        ElseIf name = "土壤取樣器(地鑽) (Earth auger)" Then
            A4_step_text = My.Resources.A4_Auger_Drill_Driver
        ElseIf name = "油壓式拔樁機" Then
            A4_step_text = My.Resources.A4_Auger_Drill_Driver
        ElseIf name = "油壓式打樁機(Hydraulic pile driver)" Then
            A4_step_text = My.Resources.A4_Auger_Drill_Driver
        ElseIf name = "振動式樁錘(Vibrating hammer)" Then
            A4_step_text = My.Resources.A4_Vibrating_Hammer
        ElseIf name = "輪形起重機(Wheel crane)" Then
            A4_step_text = My.Resources.A4_Crane
        ElseIf name = "卡車起重機(Truck crane)" Then
            A4_step_text = My.Resources.A4_Crane
        ElseIf name = "履帶起重機(Crawler crane)" Then
            A4_step_text = My.Resources.A4_Crane
        ElseIf name = "振動式壓路機(Vibrating roller)" Then
            A4_step_text = My.Resources.A4_Roller
        ElseIf name = "膠輪壓路機(Wheel roller)" Then
            A4_step_text = My.Resources.A4_Roller
        ElseIf name = "鐵輪壓路機(Road roller)" Then
            A4_step_text = My.Resources.A4_Roller
        ElseIf name = "拔樁機" Then
            A4_step_text = My.Resources.A4_Auger_Drill_Driver
        ElseIf name = "鑽岩機(Rock breaker)" Then
            A4_step_text = My.Resources.A4_Auger_Drill_Driver
        ElseIf name = "發電機(Generator)" Then
            A4_step_text = My.Resources.A4_Generator
        End If
        tempRun = HeadRun
        While Not IsNothing(tempRun)
            tempRun = tempRun.NextUnit
        End While
    End Sub


    'before using Load_Steps_helper, array_time need to be set to the correct values
    'creates steps for run and returns the pointer to the first step
    Function Load_Steps_helper(ByRef run As Run_Unit) As Steps
        Dim tempRun As Run_Unit = run
        Dim tempStep As Steps
        If tempRun.Name = "ExA1" Or tempRun.Name = "LoA1" Or tempRun.Name = "TrA1" Or tempRun.Name = "ExA1_Add" Or tempRun.Name = "LoA1_Add" Or tempRun.Name = "TrA1_Add" Then
            tempRun.Steps = New Steps(My.Resources.A1_step1, Step1, Nothing, True, A1Time)
            tempRun.HeadStep = tempRun.Steps
        ElseIf tempRun.Name = "ExA2_1st" Or tempRun.Name = "ExA2_1st_Add" Or tempRun.Name = "ExA2_2nd_3rd" Or tempRun.Name = "ExA2_2nd_3rd_Add" Then
            If tempRun.Name = "ExA2_1st" Or tempRun.Name = "ExA2_1st_Add" Then
                tempRun.Steps = New Steps(My.Resources.A2_Excavator_step1, Step1, Nothing, False, -1)
                tempRun.HeadStep = tempRun.Steps
                tempRun.Steps.NextStep = New Steps(My.Resources.A2_Excavator_step2, Step2, Nothing, False, -1)
                tempStep = tempRun.Steps.NextStep
                tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step3, Step3, Nothing, False, array_time(2))
                tempStep = tempStep.NextStep
            ElseIf tempRun.Name = "ExA2_2nd_3rd" Or tempRun.Name = "ExA2_2nd_3rd_Add" Then
                tempRun.Steps = New Steps(My.Resources.A2_Excavator_step2, Step1, Nothing, False, -1)
                tempRun.HeadStep = tempRun.Steps
                tempRun.Steps.NextStep = New Steps(My.Resources.A2_Excavator_step3, Step2, Nothing, False, array_time(2))
                tempStep = tempRun.Steps.NextStep
            End If
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step4, Step4, Nothing, False, array_time(3))
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step5, Step5, Nothing, False, array_time(4))
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step6, Step6, Nothing, False, array_time(5))
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step7, Step7, Nothing, False, array_time(6))
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step8, Step8, Nothing, False, array_time(7))
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step9, Step9, Nothing, True, array_time(8))
        ElseIf tempRun.Name = "LoA2_1st" Or tempRun.Name = "LoA2_1st_Add" Then
            tempRun.Steps = New Steps(My.Resources.A2_Loader_step1, Step1, Nothing, False, -1)
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A2_Loader_step2, Step2, Nothing, False, array_time(1))
            tempStep = tempRun.Steps.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Loader_step3, Step3, Nothing, True, array_time(2))
        ElseIf tempRun.Name = "LoA2_2nd_3rd" Or tempRun.Name = "LoA2_2nd_3rd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A2_Loader_step2, Step1, Nothing, False, array_time(1))
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A2_Loader_step3, Step2, Nothing, True, array_time(2))
        ElseIf tempRun.Name = "LoA3_fwd" Or tempRun.Name = "LoA3_fwd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A3_Loader_step1, Step1, Nothing, False, -1)
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A3_Loader_step2, Step2, Nothing, False, -1)
            tempStep = tempRun.Steps.NextStep
            tempStep.NextStep = New Steps(My.Resources.A3_Loader_step3, Step3, Nothing, True, array_time(2))
        ElseIf tempRun.Name = "LoA3_bkd" Or tempRun.Name = "LoA3_bkd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A3_Loader_step4, Step1, Nothing, True, array_time(3))
            tempRun.HeadStep = tempRun.Steps
        ElseIf tempRun.Name = "TrA3_fwd" Or tempRun.Name = "TrA3_fwd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A3_Tractor_step1, Step1, Nothing, False, -1)
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A3_Tractor_step2, Step2, Nothing, False, -1)
            tempStep = tempRun.Steps.NextStep
            tempStep.NextStep = New Steps(My.Resources.A3_Tractor_step3, Step3, Nothing, True, array_time(2))
        ElseIf tempRun.Name = "TrA3_bkd" Or tempRun.Name = "TrA3_bkd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A3_Tractor_step4, Step1, Nothing, True, array_time(3))
            tempRun.HeadStep = tempRun.Steps
        ElseIf tempRun.Name = "A4" Or tempRun.Name = "A4_Add" Then
            tempRun.Steps = New Steps(A4_step_text, Step1, Nothing, True, 0)
            tempRun.HeadStep = tempRun.Steps
        End If

        Return tempRun.Steps

    End Function


    '從METER抓即時資訊
    Private Function GetInstantData()
        Dim result(6 - 1) As Double
        Dim temp() As String =
        Comm.GetMeasurementsFromMeters(Communication.Measurements.Lp)
        'TEMP
        For i = 0 To temp.Length - 1
            Double.TryParse(temp(i), result(i))
        Next
        For i = temp.Length To result.Length - 1
            result(i) = "0"
        Next
        Return result
    End Function

    '更新螢幕上的即時圖和數據
    Private Sub SetScreenValuesAndGraphs(ByVal vals() As Double)
        Dim num As Integer
        Dim sum As Long = 0
        If Not Machine = Machines.Others Then
            num = 6
        Else
            num = 4
        End If
        'if Calibration
        Dim onlyCal As Boolean = False
        Dim meter As Integer = -1
        If CurRun.Name.Contains("Cal") Then
            onlyCal = True
            If CurRun.Name.Contains("Cal_P2") Then
                meter = Communication.Meters.p2
            ElseIf CurRun.Name.Contains("Cal_P4") Then
                meter = Communication.Meters.p4
            ElseIf CurRun.Name.Contains("Cal_P6") Then
                meter = Communication.Meters.p6
            ElseIf CurRun.Name.Contains("Cal_P8") Then
                meter = Communication.Meters.p8
            ElseIf CurRun.Name.Contains("Cal_P10") Then
                meter = Communication.Meters.p10
            ElseIf CurRun.Name.Contains("Cal_P12") Then
                meter = Communication.Meters.p12
            End If
        End If


        Dim valsAndAvg() As Double = New Double(6) {}
        Array.Copy(vals, valsAndAvg, 6)
        Dim avg As Double = GetAverageFromVals(vals)

        SetNoiseLabelValues(vals, onlyCal, num, meter, avg)
        valsAndAvg(6) = avg
        'Set graphs
        MainBarGraph.Update(valsAndAvg)
        MainLineGraph.Update(vals)
    End Sub

    Private Function GetAverageFromVals(ByRef vals() As Double) As Double
        Dim sum As Long = 0
        For i = 0 To vals.Length - 1
            sum += 10 ^ (0.1 * vals(i))
        Next
        Return 10 * (Math.Log10(sum / vals.Length))
    End Function

    Private Sub SetNoiseLabelValues(ByRef vals() As Double, ByVal onlyCal As Boolean, ByVal numOfMeters As Integer, ByVal meter As Integer, ByVal avg As Double)
        For i = 0 To numOfMeters - 1
            If Not onlyCal Then
                NoisesArray(i).Text = vals(i)
            ElseIf i = meter Then
                NoisesArray(i).Text = vals(i)
            End If
        Next
        NoisesArray(6).Text = CStr(Math.Round(avg, 1))
    End Sub

    '要換機具時要寄此訊號
    Private Sub SendChangeSignalForMobile()
        Try
            If PhoneSocket IsNot Nothing Then
                If PhoneSocket.Connected Then
                    Dim Buffer() As Byte = System.Text.Encoding.ASCII.GetBytes(ADictionary(ComboBox_machine_list.Text) & "," & timeLeft & ",T" & Environment.NewLine)
                    PhoneSocket.Send(Buffer)
                Else
                    LabelConnectMobile.Text = "Not Connected"
                End If
            Else
                LabelConnectMobile.Text = "Not Connected"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    '送出TCP signal到ANDROID，更新新步驟，和秒數
    Sub SendNonChangeSignalToMobile()
        If PhoneSocket IsNot Nothing Then
            If PhoneSocket.Connected Then
                Dim buffer() As Byte

                'A4
                If Machine = Machines.Others Then
                    buffer = System.Text.Encoding.ASCII.GetBytes(ADictionary(ComboBox_machine_list.Text) & "," & timeLeft & ",F" & Environment.NewLine)

                Else 'others
                    Dim aWhat As String = Regex.Match(CurRun.Name, "A.").ToString()
                    Dim cStep As Integer = CurRun.CurStep
                    If CurRun.Name.Contains("A2_2nd_3rd") Then
                        cStep = cStep + 1
                    ElseIf CurRun.Name.Contains("bkd") Then 'A3 backward
                        cStep = 4
                    End If
                    If Machine = Machines.Loader_Excavator And CurRun.Name.Contains("Lo") Then
                        buffer = System.Text.Encoding.ASCII.GetBytes(aWhat & "-" & ADictionary(ComboBox_machine_list.Text) & cStep & "-2," & timeLeft & ",F" & Environment.NewLine)
                    Else
                        buffer = System.Text.Encoding.ASCII.GetBytes(aWhat & "-" & ADictionary(ComboBox_machine_list.Text) & cStep & "," & timeLeft & ",F" & Environment.NewLine)
                    End If
                End If
                Try
                    PhoneSocket.Send(buffer)
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Else
                LabelConnectMobile.Text = "Not Connected"
            End If
        Else
            LabelConnectMobile.Text = "Not Connected"
        End If

    End Sub

    Sub Set_Panel_BackColor()
        CurRun.Set_BackColor(Color.Green)
        CurRun.NextUnit.Set_BackColor(Color.Yellow)
    End Sub
    Sub Set_Step_BackColor()
        array_step(CurRun.CurStep - 1).BackColor = Color.Green
        array_step(CurRun.CurStep).BackColor = Color.Yellow
        If TimerTesting = True Then
            array_light(CurRun.CurStep - 1).BackColor = Color.Green
            array_light(CurRun.CurStep).BackColor = Color.Yellow
        End If
    End Sub
    Sub All_Panel_Disable()
        Panel_PreCal_Sub.Enabled = False
        Panel_Bkg.Enabled = False
        Panel_RSS.Enabled = False
        Panel_PostCal_Sub.Enabled = False
        PanelExcavatorA1.Enabled = False
        PanelExcavatorA2.Enabled = False
        PanelLoaderA1.Enabled = False
        PanelLoaderA2.Enabled = False
        PanelLoaderA3.Enabled = False
        PanelTractorA1.Enabled = False
        PanelTractorA3.Enabled = False
        PanelA4.Enabled = False
    End Sub
    Sub All_Panel_Enable()
        Panel_PreCal_Sub.Enabled = True
        Panel_Bkg.Enabled = True
        Panel_RSS.Enabled = True
        Panel_PostCal_Sub.Enabled = True
        PanelExcavatorA1.Enabled = True
        PanelExcavatorA2.Enabled = True
        PanelLoaderA1.Enabled = True
        PanelLoaderA2.Enabled = True
        PanelLoaderA3.Enabled = True
        PanelTractorA1.Enabled = True
        PanelTractorA3.Enabled = True
        PanelA4.Enabled = True
    End Sub
    Sub Clear_Steps()
        For index = 0 To 9
            array_step(index).Text = ""
            array_step(index).BackColor = Color.DarkGray
        Next
    End Sub

    '讀入步驟所需的物件及細節
    Sub Load_Steps()
        Set_Second_for_Steps()
        For index = CurRun.StartStep To CurRun.EndStep
            sum_steps += CurStep.Time
            array_step(index - 1).Text = CurStep.step_str
            If CurStep.Time = -1 Then
                array_step(index - 1).BackColor = Color.Green
                If TimerTesting = True Then
                    array_light(index - 1).BackColor = Color.Green
                End If
            Else
                If CurRun.EndStep = 1 Or index = 1 Then
                    array_step(index - 1).BackColor = Color.Yellow
                    If TimerTesting = True Then
                        array_light(index - 1).BackColor = Color.Yellow
                    End If
                Else
                    If array_step(index - 2).BackColor = Color.Green Then
                        array_step(index - 1).BackColor = Color.Yellow
                        If TimerTesting = True Then
                            array_light(index - 1).BackColor = Color.Yellow
                        End If
                    Else
                        array_step(index - 1).BackColor = Color.IndianRed
                        If TimerTesting = True Then
                            array_light(index - 1).BackColor = Color.IndianRed
                        End If
                    End If
                End If
            End If
            If CurStep.HasNext() = True Then
                CurStep = CurStep.NextStep
            End If
        Next
        CurStep = CurRun.HeadStep
    End Sub

    Sub Load_New_Graph_CD_False()
        'load new graph (when variable countdown = false)
        MainLineGraph.Dispose()
        If Machine = Machines.Others Then
            MainLineGraph = New LineGraph(TabPage2, CGraph.Modes.A4, CurRun.Time)  'A4 mode
        Else
            MainLineGraph = New LineGraph(TabPage2, CGraph.Modes.A1A2A3, CurRun.Time) 'A1 A2 A3 mode
        End If
    End Sub
    Sub Load_New_Graph_CD_True()
        'load new graph(when variable countdown = true)
        MainLineGraph.Dispose()
        If Machine = Machines.Others Then
            MainLineGraph = New LineGraph(TabPage2, CGraph.Modes.A4, sum_steps)  'A4 mode
        Else
            MainLineGraph = New LineGraph(TabPage2, CGraph.Modes.A1A2A3, sum_steps) 'A1 A2 A3 mode
        End If
        sum_steps = 0
    End Sub

    '在按STOP或時間到，要Bar graph 顯示最後的數據
    Private Sub updateFinalBarGraph()
        Dim cal As Boolean = False
        Dim meter As Integer = -1
        Dim num As Integer
        If Not Machine = Machines.Others Then
            num = 6
        Else
            num = 4
        End If

        If CurRun.Name.Contains("Cal") Then
            cal = True
            If CurRun.Name.Contains("Cal_P2") Then
                meter = 0
            ElseIf CurRun.Name.Contains("Cal_P4") Then
                meter = 1
            ElseIf CurRun.Name.Contains("Cal_P6") Then
                meter = 2
            ElseIf CurRun.Name.Contains("Cal_P8") Then
                meter = 3
            ElseIf CurRun.Name.Contains("Cal_P10") Then
                meter = 4
            Else
                meter = 5
            End If
        End If
        Dim temps() As String = Comm.ConsumeMeasurementsFromBuffer(Communication.Measurements.Leq)
        Dim Leqs(6) As Double
        Dim sum As Double = 0
        For i = 0 To temps.Length - 1
            If Not cal Then
                Dim number = Convert.ToDouble(temps(i))
                Leqs(i) = number
                sum += 10 ^ (0.1 * number)
            ElseIf i = meter Then
                Dim number = Convert.ToDouble(temps(i))
                Leqs(i) = number
                sum = number
            Else
                Leqs(i) = 0
            End If
        Next
        For i = temps.Length To 6 - 1
            Leqs(i) = 0
        Next
        Dim avg As Double
        If Not cal Then
            avg = 10 * (Math.Log10(sum / num))
        Else
            avg = sum
        End If
        Leqs(6) = avg
        Dim vals(6 - 1) As Double
        Array.Copy(Leqs, vals, 6)
        SetNoiseLabelValues(Leqs, False, num, 0, avg)
        MainBarGraph.Update(Leqs)

    End Sub

    '
    Private Sub ShowResultsOnForm()

        'Precal,postcal,RSS,background
        Dim series As List(Of DataVisualization.Charting.Series) = MainLineGraph.GetSeries()

        Dim Leqpoints As DataVisualization.Charting.DataPointCollection = MainBarGraph.GetSeries(0).Points()

        'precal and postcal
        If CurRun.Name.Contains("Cal") Then
            If CurRun.Name.Contains("Cal_P2") Then
                CurRun.GRU.SetM(Meter_Measure_Unit.SeriesToMMU(series(0), Leqpoints(0).YValues(0)), 2)
            ElseIf CurRun.Name.Contains("Cal_P4") Then
                CurRun.GRU.SetM(Meter_Measure_Unit.SeriesToMMU(series(1), Leqpoints(1).YValues(0)), 4)
            ElseIf CurRun.Name.Contains("Cal_P6") Then
                CurRun.GRU.SetM(Meter_Measure_Unit.SeriesToMMU(series(2), Leqpoints(2).YValues(0)), 6)
            ElseIf CurRun.Name.Contains("Cal_P8") Then
                CurRun.GRU.SetM(Meter_Measure_Unit.SeriesToMMU(series(3), Leqpoints(3).YValues(0)), 8)
            ElseIf CurRun.Name.Contains("Cal_P10") Then
                CurRun.GRU.SetM(Meter_Measure_Unit.SeriesToMMU(series(4), Leqpoints(4).YValues(0)), 10)
            ElseIf CurRun.Name.Contains("Cal_P12") Then
                CurRun.GRU.SetM(Meter_Measure_Unit.SeriesToMMU(series(5), Leqpoints(5).YValues(0)), 12)

            End If
        End If


        'everything else
        Dim tempGRU As Grid_Run_Unit = CurRun.GRU
        While Not IsNothing(tempGRU.NextGRU) 'for more than one additional runs
            tempGRU = tempGRU.NextGRU
        End While
        Try
            If series.Count = 4 + 1 Then
                tempGRU.SetMs(Meter_Measure_Unit.SeriesToMMU(series(0), Leqpoints(0).YValues(0)),
                                Meter_Measure_Unit.SeriesToMMU(series(1), Leqpoints(1).YValues(0)),
                                Meter_Measure_Unit.SeriesToMMU(series(2), Leqpoints(2).YValues(0)),
                                Meter_Measure_Unit.SeriesToMMU(series(3), Leqpoints(3).YValues(0)))
            Else
                tempGRU.SetMs(Meter_Measure_Unit.SeriesToMMU(series(0), Leqpoints(0).YValues(0)),
                                Meter_Measure_Unit.SeriesToMMU(series(1), Leqpoints(1).YValues(0)),
                                Meter_Measure_Unit.SeriesToMMU(series(2), Leqpoints(2).YValues(0)),
                                Meter_Measure_Unit.SeriesToMMU(series(3), Leqpoints(3).YValues(0)),
                                Meter_Measure_Unit.SeriesToMMU(series(4), Leqpoints(4).YValues(0)),
                                Meter_Measure_Unit.SeriesToMMU(series(5), Leqpoints(5).YValues(0)))
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        DataGrid.ShowGRUonForm(tempGRU)
        If CurRun.Name.Contains("bkd") Then
            Dim latestFwdGRU As Grid_Run_Unit = CurRun.PrevUnit.GRU
            Dim latestBkdGRU As Grid_Run_Unit = CurRun.GRU
            While Not IsNothing(latestFwdGRU.NextGRU)
                latestFwdGRU = latestFwdGRU.NextGRU
                latestBkdGRU = latestBkdGRU.NextGRU
            End While
            DataGrid.ShowA3Overall(latestFwdGRU, latestBkdGRU)
        End If
    End Sub

    Sub ClearSeconds()
        For i = 0 To array_step_display.Length - 1
            array_step_display(i).Text = ""
        Next
    End Sub

    '處理加測的Grid_Run_Unit
    'two approaches, if there's already a GRU existent, add as next GRU, if not, attach to the RU
    Sub Set_Add_GRU(ByVal whichA As Integer, ByVal colName As String, ByVal subHeader As String)
        Dim tempGRU = New Grid_Run_Unit(colName) 'the final GRU added
        tempGRU.Subheader = subHeader
        Dim tempCol As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn()
        Dim lastGRU As Grid_Run_Unit 'used for getting the nearest column
        Dim firstAdd As Boolean = False 'only used for A3, check if it's the first add, if not it will help jump 3 columns at a time
        Dim tempRun As Run_Unit = CurRun 'to get the right parent
        If whichA = 2 Then
            tempRun = tempRun.NextUnit.NextUnit
        End If

        If IsNothing(tempRun.GRU) Then 'condition for first addtional run
            If whichA = 2 Then
                lastGRU = CurRun.PrevUnit.GRU
            ElseIf whichA = 3 And tempRun.Name.Contains("fwd") Then
                lastGRU = tempRun.PrevUnit.GRU.OverallGRU
            Else
                lastGRU = tempRun.PrevUnit.GRU
            End If
            tempRun.GRU = tempGRU
            firstAdd = True
        Else 'condition for more than one additional run
            lastGRU = tempRun.GRU

            While lastGRU.NextGRU IsNot Nothing
                lastGRU = lastGRU.NextGRU
            End While
            lastGRU.NextGRU = tempGRU
        End If

        'i is the number of columns increased each additional run has, A3 has fwd and bwd and overall so there's three
        'A4 has background and regular so there's 2
        Dim i = 1
        If whichA = 3 And Not firstAdd Then
            i = 3
        ElseIf whichA = 4 And Not firstAdd Then
            i = 2
        End If

        DataGrid.Form.Columns.Insert(lastGRU.Column.Index + i, tempCol)
        DataGrid.Form.Rows(0).Cells(tempCol.Index).Value = subHeader
        tempGRU.Column = tempCol
        tempGRU.Background = DataGrid.Background
        tempCol.HeaderText = colName
        tempGRU.ParentRU = tempRun
    End Sub

    'Load steps based on current Run, start from the effective start step
    Sub Set_Run_Unit()
        CurRun.Steps = Load_Steps_helper(CurRun)
        CurStep = CurRun.HeadStep
        CurRun.CurStep = 1
        For index = CurRun.StartStep To CurRun.EndStep
            If CurRun.Steps.Time = -1 Then 'not a real step to run
                CurRun.Steps = CurRun.Steps.NextStep
                CurRun.CurStep += 1
            Else
                Exit For
            End If
        Next
        timeLeft = CurRun.Steps.Time
        timeLabel.Text = timeLeft & " s"
    End Sub
    Sub Restart_from_1st_Previous_Run_Unit()
        CurRun.PrevUnit.Set_BackColor(Color.Yellow)
        CurRun.Set_BackColor(Color.IndianRed)
        MoveToRun(CurRun.PrevUnit)
        Set_Run_Unit()
        CurRun.NextUnit.CurStep = 1
    End Sub
    Sub Restart_from_2nd_Previous_Run_Unit()
        CurRun.PrevUnit.PrevUnit.Set_BackColor(Color.Yellow)
        CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
        CurRun.Set_BackColor(Color.IndianRed)
        MoveToRun(CurRun.PrevUnit.PrevUnit)
        Set_Run_Unit()
        CurRun.NextUnit.CurStep = 1
        CurRun.NextUnit.NextUnit.CurStep = 1
    End Sub


    Public Sub Load_Array_Step_S()
        Dim tempStep As Steps = CurRun.HeadStep
        Dim i = 0
        While tempStep IsNot Nothing
            array_step_s(i).Text = tempStep.Time
            tempStep = tempStep.NextStep
            i += 1
        End While
    End Sub

    ' This is called when 1) user input is entered for A2 and A3's time or 2) loading saved file that includes A2 or A3
    Public Sub LoadInputTime(ByVal mode As Integer, ByRef secList As List(Of Double))
        'mode 1 is manually set
        'mode 2 is load from file
        'mode 3 is jump step

        'First, set array_time()
        If mode = 1 Then 'used when manually set time is used
            For i = 0 To array_time.Length - 1
                array_time(i) = array_step_s(i).Text
            Next
        ElseIf mode = 2 Then 'used when we load from file
            If secList IsNot Nothing Then
                For i = 0 To secList.Count - 1
                    array_time(i) = secList(i)
                Next
            End If
        Else 'jump step meaning the steps have been loaded already, so we only need to load into array_step_s
            If Not CurRun.Name.Contains("bkd") Then 'not A3 backward
                Dim tempStep As Steps = CurRun.HeadStep
                Dim i As Integer = 0
                While tempStep IsNot Nothing
                    array_time(i) = tempStep.Time
                    tempStep = tempStep.NextStep
                    i += 1
                End While
            Else 'A3 backward
                'first load A3 forward
                Dim tempStep As Steps = CurRun.PrevUnit.HeadStep
                Dim i As Integer = 0
                While tempStep IsNot Nothing
                    array_time(i) = tempStep.Time
                    tempStep = tempStep.NextStep
                    i += 1
                End While
                'then load A3 backward
                tempStep = CurRun.HeadStep
                While tempStep IsNot Nothing
                    array_time(i) = tempStep.Time
                    tempStep = tempStep.NextStep
                    i += 1
                End While
            End If

        End If

        'Second, set up the steps
        If mode = 2 Then
            CurStep = Load_Steps_helper(CurRun)
        Else
            CurStep = CurRun.HeadStep
        End If

        'Last, Set up array_step_s
        If Not CurRun.Name.Contains("A3") Then ' not A3
            'this for loop leaves out the -1 for display
            If Not mode = 1 Then
                For i = 0 To CurRun.EndStep - 1
                    If CurStep.Time = -1 Then
                        array_step_s(i).Text = 0
                    Else
                        array_step_s(i).Text = array_time(i)
                    End If
                    CurStep = CurStep.NextStep
                Next
            End If
        Else 'A3
            If CurRun.Name.Contains("bkd") Then 'if already running backward
                If mode = 1 Then 'manually set
                    'bkd 的時間和 fwd的時間都放在同一個array裡
                    array_time(CurRun.PrevUnit.EndStep) = array_step_s(CurRun.PrevUnit.EndStep).Text
                ElseIf mode = 2 Then 'load from file
                    array_step_s(CurRun.PrevUnit.EndStep).Text = secList(secList.Count - 1)
                    array_time(CurRun.PrevUnit.EndStep) = secList(secList.Count - 1)
                Else 'jump step
                    CurRun.Steps = CurRun.HeadStep
                    array_step_s(CurRun.PrevUnit.EndStep).Text = CurRun.Steps.Time
                    array_time(CurRun.PrevUnit.EndStep) = CurRun.Steps.Time
                End If
            Else 'running forward
                If mode = 1 Then 'manually set
                    'bkd 的時間和 fwd的時間都放在同一個array裡
                    Dim tempstep As Steps = CurRun.HeadStep
                    Dim numSteps As Integer = 1
                    While tempstep IsNot Nothing
                        tempstep = tempstep.NextStep
                        numSteps += 1
                    End While
                    CurRun.EndStep = numSteps - 1
                ElseIf mode = 2 Then 'load from file
                    Dim tempStep As Steps = CurRun.HeadStep
                    For i = 0 To CurRun.EndStep - 1
                        If tempStep.Time = -1 Then
                            array_step_s(i).Text = 0
                            CurStep = tempStep
                        Else
                            array_step_s(i).Text = array_time(i)
                        End If
                        tempStep = tempStep.NextStep
                    Next
                    array_step_s(CurRun.EndStep).Text = array_time(CurRun.EndStep)
                Else 'jump step
                    Dim tempStep As Steps = CurRun.HeadStep
                    For i = 0 To CurRun.EndStep - 1
                        If tempStep.Time = -1 Then
                            array_step_s(i).Text = 0
                            CurStep = tempStep
                        Else
                            array_step_s(i).Text = array_time(i)
                        End If
                        tempStep = tempStep.NextStep
                    Next
                End If
            End If

        End If
    End Sub

    '要設步驟的秒數前，要確定array_step_s有設好
    Sub Set_Second_for_Steps()
        Dim isJumpStep As Boolean = False
        If Temp_CurRun Is Nothing Then
            isJumpStep = True
        ElseIf Not Temp_CurRun.Link.Name = "LinkLabel_Temp" Then
            isJumpStep = True
        End If

        If CurRun.Name.Contains("ExA2_1st") Then

            If isJumpStep Then 'if jump step
                For index = 8 To 1
                    array_step_display(index).Text = array_step_display(index - 1).Text
                Next
                array_step_display(0).Text = ""
            Else
                For index = 0 To 8
                    array_step_display(index).Text = ""
                Next
                For index = 0 To 8
                    array_step_display(index).Text = array_step_s(index).Text
                Next
            End If
        ElseIf CurRun.Name.Contains("ExA2_2nd") Then

            If isJumpStep Then
                If CurRun.PrevUnit.Name.Contains("ExA2_1st") Then
                    For index = 0 To 7
                        array_step_display(index).Text = array_step_display(index + 1).Text
                    Next
                    array_step_display(8).Text = ""
                End If
            Else
                For index = 0 To 8
                    array_step_display(index).Text = ""
                Next
                For index = 0 To 7
                    array_step_display(index).Text = array_step_s(index + 1).Text
                Next
            End If

        ElseIf CurRun.Name.Contains("LoA2_1st") Then
            If isJumpStep Then
                For index = 2 To 1
                    array_step_display(index).Text = array_step_display(index - 1).Text
                Next
                array_step_display(0).Text = ""
            Else
                For index = 0 To 8
                    array_step_display(index).Text = ""
                Next
                For index = 0 To 2
                    array_step_display(index).Text = array_step_s(index).Text
                Next
            End If

        ElseIf CurRun.Name.Contains("LoA2_2nd") Then
            If isJumpStep Then
                If CurRun.PrevUnit.Name.Contains("LoA2_1st") Then
                    For index = 0 To 1
                        array_step_display(index).Text = array_step_display(index + 1).Text
                    Next
                    array_step_display(2).Text = ""
                End If
            Else
                For index = 0 To 8
                    array_step_display(index).Text = ""
                Next
                For index = 0 To 1
                    array_step_display(index).Text = array_step_s(index + 1).Text
                Next
            End If
        ElseIf CurRun.Name.Contains("A1") Then
            For index = 0 To 8
                array_step_display(index).Text = ""
            Next
            array_step_display(0).Text = CurRun.Steps.Time
        ElseIf CurRun.Name.Contains("fwd") Then
            For index = 0 To 2
                array_step_display(index).Text = array_step_s(index).Text
            Next
        ElseIf CurRun.Name.Contains("bkd") Then
            Label_step1_second.Text = Input_S_Step4.Text
            Label_step2_second.Text = ""
            Label_step3_second.Text = ""
        Else
            For index = 0 To 8
                array_step_display(index).Text = ""
            Next
        End If
    End Sub

    Sub Reset_Test_Time()
        CurRun.Steps = Load_Steps_helper(CurRun)
        TimerTesting = True
        SetTimerTabOn(TimerTesting)
        Countdown = False
        CurStep = CurRun.HeadStep
        CurRun.CurStep = 1
        For index = CurRun.StartStep To CurRun.EndStep
            If CurRun.Steps.Time = -1 Then
                CurRun.Steps = CurRun.Steps.NextStep
                CurRun.CurStep += 1
                Index_for_Setup_Time += 1
            Else
                Exit For
            End If
        Next
        timeLeft = 0
        timeLabel.Text = timeLeft & " s"
    End Sub

    Sub Jump_Back_Countdown_True(ByVal Result As DialogResult)
        If CurRun.Name = "ExA2_1st" Or CurRun.Name = "ExA2_1st_Add" Or CurRun.Name = "LoA2_1st" Or CurRun.Name = "LoA2_1st_Add" Or CurRun.Name = "ExA2_2nd_3rd" And CurRun.NextUnit.Name = "ExA2_2nd_3rd" Or CurRun.Name = "ExA2_2nd_3rd_Add" And CurRun.NextUnit.Name = "ExA2_2nd_3rd_Add" Or CurRun.Name = "LoA2_2nd_3rd" And CurRun.NextUnit.Name = "LoA2_2nd_3rd" Or CurRun.Name = "LoA2_2nd_3rd_Add" And CurRun.NextUnit.Name = "LoA2_2nd_3rd_Add" Then
            'MessageBox.Show("here")
            Set_Panel_BackColor()
            MoveToRun(CurRun.NextUnit)
            Set_Run_Unit()

            'load A2's steps
            Clear_Steps()
            Load_Steps()

            'dispose old graph and create new graph
            Load_New_Graph_CD_True()
        Else
            Dim keyGRU As Grid_Run_Unit = CurRun.GRU
            While Not IsNothing(keyGRU.NextGRU) 'for more than one additional runs
                keyGRU = keyGRU.NextGRU
            End While
            If Result = DialogResult.Yes Then
                'save data
                keyGRU.Accept()
                DataGrid.ShowGRUonForm(keyGRU)
                If CurRun.Name.Contains("bkd") Then
                    keyGRU.OverallGRU.Accept()
                    DataGrid.ShowGRUonForm(keyGRU)
                End If


                All_Panel_Enable()
                CurRun.Set_BackColor(Color.Green)
                If Temp_CurRun IsNot Nothing Then
                    Temp_CurRun.Set_BackColor(Color.Yellow)
                End If
                CurRun = Temp_CurRun
                Temp_CurRun = Null_CurRun
                If CurRun IsNot Nothing Then
                    If Temp_Countdown = True Then
                        Countdown = True
                        Set_Run_Unit()
                        Clear_Steps()
                        Load_Steps()
                        'dispose old graph and create new graph
                        Load_New_Graph_CD_True()
                    Else
                        Countdown = False
                        timeLeft = CurRun.Time
                        timeLabel.Text = timeLeft & " s"
                        Clear_Steps()
                        'dispose old graph and create new graph
                        Load_New_Graph_CD_False()
                    End If
                End If
                timeLabel.Text = timeLeft & " s"
            ElseIf Result = DialogResult.No Then
                'Accept_No()
                'don't save data
                keyGRU.Deny()
                DataGrid.ShowGRUonForm(keyGRU)
                If CurRun.Name.Contains("bkd") Then
                    keyGRU.OverallGRU.Deny()
                    DataGrid.ShowGRUonForm(keyGRU)
                End If


                All_Panel_Enable()

                CurRun.Set_BackColor(Color.Green)
                If Temp_CurRun IsNot Nothing Then
                    Temp_CurRun.Set_BackColor(Color.Yellow)
                End If
                CurRun = Temp_CurRun
                Temp_CurRun = Null_CurRun
                If CurRun IsNot Nothing Then
                    If Temp_Countdown = True Then
                        Countdown = True
                        Set_Run_Unit()
                        Clear_Steps()
                        Load_Steps()
                        'dispose old graph and create new graph
                        Load_New_Graph_CD_True()
                    Else
                        Countdown = False
                        timeLeft = CurRun.Time
                        timeLabel.Text = timeLeft & " s"
                        Clear_Steps()
                        'dispose old graph and create new graph
                        Load_New_Graph_CD_False()
                    End If
                End If
                timeLabel.Text = timeLeft & " s"

            ElseIf Result = DialogResult.Cancel Then
                Accept_Cancel()
            End If
        End If

    End Sub


    Sub Jump_Back_Countdown_False(ByVal Result As DialogResult)
        Dim keyGRU As Grid_Run_Unit = CurRun.GRU
        While Not IsNothing(keyGRU.NextGRU) 'for more than one additional runs
            keyGRU = keyGRU.NextGRU
        End While
        If Result = DialogResult.Yes Then
            'save data
            keyGRU.Accept()
            DataGrid.ShowGRUonForm(keyGRU)
            If CurRun.Name.Contains("bkd") Then
                keyGRU.OverallGRU.Accept()
                DataGrid.ShowGRUonForm(keyGRU)
            End If

            If CurRun.Name.Contains("RSS") Then
                'Show calculated results for chart
                DataGrid.ShowCalculated()
            End If

            All_Panel_Enable()
            CurRun.Set_BackColor(Color.Green)
            If Temp_CurRun IsNot Nothing Then
                Temp_CurRun.Set_BackColor(Color.Yellow)
            End If
            CurRun = Temp_CurRun
            If CurRun IsNot Nothing Then
                If Temp_Countdown = True Then
                    Countdown = True
                    CurRun.Steps = Load_Steps_helper(CurRun)
                    CurRun.CurStep = 1
                    CurStep = CurRun.Steps
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"
                    Clear_Steps()
                    Load_Steps()
                    'dispose old graph and create new graph
                    Load_New_Graph_CD_True()
                Else
                    Countdown = False
                    timeLeft = CurRun.Time
                    timeLabel.Text = timeLeft & " s"
                    'dispose old graph and create new graph
                    Load_New_Graph_CD_False()
                End If
            End If
            Temp_CurRun = Null_CurRun
            timeLabel.Text = timeLeft & " s"
        ElseIf Result = DialogResult.No Then
            'Accept_No()
            'don't save data
            keyGRU.Deny()
            DataGrid.ShowGRUonForm(keyGRU)
            If CurRun.Name.Contains("bkd") Then
                keyGRU.OverallGRU.Deny()
                DataGrid.ShowGRUonForm(keyGRU)
            End If

            All_Panel_Enable()
            CurRun.Set_BackColor(Color.Green)
            If Temp_CurRun IsNot Nothing Then
                Temp_CurRun.Set_BackColor(Color.Yellow)
            End If
            CurRun = Temp_CurRun
            If CurRun IsNot Nothing Then
                If Temp_Countdown = True Then
                    Countdown = True
                    CurRun.Steps = Load_Steps_helper(CurRun)
                    CurRun.CurStep = 1
                    CurStep = CurRun.Steps
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"
                    Clear_Steps()
                    Load_Steps()
                    'dispose old graph and create new graph
                    Load_New_Graph_CD_True()
                Else
                    Countdown = False
                    timeLeft = CurRun.Time
                    timeLabel.Text = timeLeft & " s"
                    'dispose old graph and create new graph
                    Load_New_Graph_CD_False()
                End If
            End If
            Temp_CurRun = Null_CurRun
            timeLabel.Text = timeLeft & " s"
        ElseIf Result = DialogResult.Cancel Then
            Accept_Cancel()
        End If
    End Sub

    Sub Accept_No()
        If Countdown = False Then
            All_Panel_Enable()
            'dispose data here

            timeLeft = CurRun.Time
            timeLabel.Text = timeLeft & " s"
            'dispose old graph and create new graph
            Load_New_Graph_CD_False()
        Else
            All_Panel_Enable()
            'dispose data

            'reset run_unit
            Set_Run_Unit()
            'load A1's steps
            Clear_Steps()
            Load_Steps()
            'dispose old graph and create new graph
            Load_New_Graph_CD_True()
        End If
    End Sub


    Sub Accept_Cancel()
        startButton.Enabled = False
        Accept_Button.Enabled = True
        Button_Skip_Add.Enabled = False
    End Sub

    '當每次要到下一步時，都會call這個步驟去決定下一步要怎麼走，無論是跳步驟或是暫停等
    Sub MoveOnToNextRun(ByVal showGraph As Boolean, ByVal needDetermineAdd As Boolean, ByVal needAdd As Boolean)
        If IsNothing(CurRun.NextUnit) Then
            All_Panel_Enable()
            CurRun.Set_BackColor(Color.Green)
            CurRun.Link.Enabled = True
            MessageBox.Show("End")
            startButton.Enabled = False
            stopButton.Enabled = False
            Accept_Button.Enabled = False
            CurRun = Nothing
        ElseIf CurRun.Name.Contains("PreCal") Then
            Countdown = False
            'didn't move
            All_Panel_Enable()
            'Case1: now is PreCal
            ' change light
            Set_Panel_BackColor()
            CurRun.Link.Enabled = True
            'dispose old graph and create new graph
            If showGraph Then
                Load_New_Graph_CD_False()
            End If
            'save info here

            'jump to next Run_Unit and set second to zero(should be zero here)
            MoveToRun(CurRun.NextUnit)
            timeLeft = CurRun.Time
            timeLabel.Text = timeLeft & " s"


        ElseIf CurRun.Name = "Background" Then
            Countdown = False
            All_Panel_Enable()
            'Case2: now is Background
            ' change light
            Set_Panel_BackColor()
            CurRun.Link.Enabled = True
            'jump to next Run_Unit and change countdown False to True(for A1 A2 A3 A4 test)
            MoveToRun(CurRun.NextUnit)
            Set_Run_Unit()
            If CurRun.Name.Contains("A4") Then
                Countdown = False
            Else
                Countdown = True
            End If
            'load A1's steps
            Clear_Steps()
            Load_Steps()

            'dispose old graph and create new graph
            Load_New_Graph_CD_True()

        ElseIf CurRun.Name = "RSS" Then
            Countdown = False
            All_Panel_Enable()
            'Case3: now is RSS
            ' change light
            Set_Panel_BackColor()
            CurRun.Link.Enabled = True
            'dispose old graph and create new graph
            If showGraph Then
                Load_New_Graph_CD_False()
            End If
            'Show calculated results for chart
            DataGrid.ShowCalculated()

            'save info here
            'jump to next Run_Unit and set second to zero(should be zero here)
            MoveToRun(CurRun.NextUnit)
            timeLeft = CurRun.Time
            timeLabel.Text = timeLeft & " s"

        ElseIf CurRun.Name.Contains("PostCal") Then
            Countdown = False
            All_Panel_Enable()
            'Case4: now is PostCal
            ' change light
            Set_Panel_BackColor()
            CurRun.Link.Enabled = True
            'dispose old graph and create new graph
            If showGraph Then
                Load_New_Graph_CD_False()
            End If
            'save info here

            'jump to next Run_Unit and set second to zero(should be zero here)
            MoveToRun(CurRun.NextUnit)
            timeLeft = CurRun.Time
            timeLabel.Text = timeLeft & " s"


        ElseIf CurRun.Name = "A4" Then
            Countdown = False
            All_Panel_Enable()
            If CurRun.NextUnit.Name = "A4_Mid_BG" Then
                'Case: now is A4 and next is A4_Mid_BG
                ' change light
                Set_Panel_BackColor()
                CurRun.Link.Enabled = True
                'jump to next Run_Unit
                MoveToRun(CurRun.NextUnit)
                timeLeft = CurRun.Time
                timeLabel.Text = timeLeft & " s"
                Clear_Steps()
                'dispose old graph and create new graph
                If showGraph Then
                    Load_New_Graph_CD_False()
                End If
            ElseIf CurRun.NextUnit.Name = "A4_Mid_BG_Add" Then
                'have an additional test?
                'call a function 
                'if want_add()=true then ... elseif want_add()=false then ... endif

                CurRun.Set_BackColor(Color.Green)
                CurRun.Link.Enabled = True
                'CurRun.Link.Enabled = True
                'jump to next Run_Unit and change light
                'Add?
                Dim reallyNeedAdd As Boolean
                If needDetermineAdd Then
                    Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                    list.Add(CurRun.GRU) 'adding test 3 
                    list.Add(CurRun.PrevUnit.PrevUnit.GRU) 'adding testing 2
                    list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding test 1
                    reallyNeedAdd = DataGrid.NeedAdd(list)
                Else
                    reallyNeedAdd = needAdd
                End If

                If reallyNeedAdd Then
                    'True: add test
                    MoveToRun(CurRun.NextUnit)
                    Set_Add_GRU(0, "Background", "")    'True: add test
                    ' change light
                    CurRun.Set_BackColor(Color.Yellow)
                    timeLeft = CurRun.Time
                    timeLabel.Text = timeLeft & " s"
                    Clear_Steps()
                    'dispose old graph and create new graph
                    If showGraph Then
                        Load_New_Graph_CD_False()
                    End If
                    'skip button enable
                    Button_Skip_Add.Enabled = True
                Else
                    'False: not add test , jump to RSS
                    MoveToRun(CurRun.NextUnit.NextUnit.NextUnit)
                    CurRun.Set_BackColor(Color.Yellow)
                    timeLeft = CurRun.Time
                    timeLabel.Text = timeLeft & " s"
                    Clear_Steps()
                    For index = 0 To 8
                        array_step_display(index).Text = ""
                    Next
                    'dispose old graph and create new graph
                    If showGraph Then
                        Load_New_Graph_CD_False()
                    End If
                End If
            End If

        ElseIf CurRun.Name = "A4_Add" Then
            Countdown = False
            All_Panel_Enable()
            Add_Test_Record = True
            'have an additional test?
            'call a function 
            'if want_add()=true then ... elseif want_add()=false then ... endif
            CurRun.Link.Enabled = True
            'CurRun.Link.Enabled = True
            'Add?
            Dim reallyNeedAdd As Boolean
            If needDetermineAdd Then
                Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                list.Add(CurRun.PrevUnit.GRU) 'adding test 3 
                list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding testing 2
                list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding test 1
                Dim tempGRU = CurRun.GRU
                Dim i As Integer = 4
                While Not IsNothing(tempGRU)
                    list.Add(tempGRU)
                    tempGRU = tempGRU.NextGRU
                    i += 1
                End While
                reallyNeedAdd = DataGrid.NeedAdd(list)
            Else
                reallyNeedAdd = needAdd
            End If

            If reallyNeedAdd Then
                CurRun.PrevUnit.Set_BackColor(Color.Yellow)
                CurRun.Set_BackColor(Color.IndianRed)
                MoveToRun(CurRun.PrevUnit)
                'True: add test
                Set_Add_GRU(4, "Background", "")    'True: add test
                'True: add test , jump to A4_Mid_BG_Add

                timeLeft = CurRun.Time
                timeLabel.Text = timeLeft & " s"

                Clear_Steps()

                'dispose old graph and create new graph
                If showGraph Then
                    Load_New_Graph_CD_True()
                End If
                'skip button enable
                Button_Skip_Add.Enabled = True
            Else
                'False: not add test , jump to RSS
                Set_Panel_BackColor()
                MoveToRun(CurRun.NextUnit)
                CurRun.Steps = Load_Steps_helper(CurRun)
                timeLeft = CurRun.Time
                timeLabel.Text = timeLeft & " s"
                Clear_Steps()
                For index = 0 To 8
                    array_step_display(index).Text = ""
                Next
                'dispose old graph and create new graph
                If showGraph Then
                    Load_New_Graph_CD_True()
                End If
                'skip button disable
                Button_Skip_Add.Enabled = False
            End If

        ElseIf CurRun.Name = "A4_Mid_BG" Or CurRun.Name = "A4_Mid_BG_Add" Then
            Countdown = False
            All_Panel_Enable()
            'Case: now is A4_Mid_BG or A4_Mid_BG_Add and next is A4 or A4_Add
            ' change light
            Set_Panel_BackColor()
            CurRun.Link.Enabled = True

            MoveToRun(CurRun.NextUnit)
            'jump to next Run_Unit
            'case: next is A4_Add
            If CurRun.Name.Contains("Add") Then
                Dim tempGRU = CurRun.GRU
                Dim i As Integer = 4
                While Not IsNothing(tempGRU)
                    tempGRU = tempGRU.NextGRU
                    i += 1
                End While

                Set_Add_GRU(4, "Run" & i, "")
            End If
            Set_Run_Unit()
            'load A4's steps
            Clear_Steps()
            Load_Steps()

            'dispose old graph and create new graph
            If showGraph Then
                Load_New_Graph_CD_True()
            End If

        ElseIf CurRun.Name = "ExA1" Or CurRun.Name = "TrA1" Or CurRun.Name = "LoA1" Then
            Countdown = True
            All_Panel_Enable()
            If CurRun.NextUnit.Name = "ExA1" Or CurRun.NextUnit.Name = "TrA1" Or CurRun.NextUnit.Name = "LoA1" Then
                'Case1: now is ExA1 and next is also ExA1 or now is TrA1 and next is also TrA1 or now is LoA1 and next is also LoA1
                ' change light
                Set_Panel_BackColor()
                CurRun.Link.Enabled = True
                'jump to next Run_Unit
                MoveToRun(CurRun.NextUnit)
                Set_Run_Unit()

                'load A1's steps
                Clear_Steps()
                Load_Steps()

                'dispose old graph and create new graph
                If showGraph Then
                    Load_New_Graph_CD_True()
                End If
            ElseIf CurRun.NextUnit.Name.Contains("A1_Add") Then
                'Case2: now is ExA1 and next is ExA1_Add or now is TrA1 and next is TrA1_Add or now is LoA1 and next is LoA1_Add
                'have an additional test?
                'call a function 

                CurRun.Set_BackColor(Color.Green)
                CurRun.Link.Enabled = True
                'jump to next Run_Unit and change light
                Dim reallyNeedAdd As Boolean
                If needDetermineAdd Then
                    Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                    list.Add(CurRun.GRU) 'adding test 3
                    list.Add(CurRun.PrevUnit.GRU) 'adding testing 2
                    list.Add(CurRun.PrevUnit.PrevUnit.GRU) 'adding test 1
                    reallyNeedAdd = DataGrid.NeedAdd(list)
                Else
                    reallyNeedAdd = needAdd
                End If
                'Add?

                If reallyNeedAdd Then
                    'True: add test
                    MoveToRun(CurRun.NextUnit)
                    Set_Add_GRU(0, "Run4", "")
                    CurRun.Set_BackColor(Color.Yellow)
                    Set_Run_Unit()
                    'skip button enable
                    Button_Skip_Add.Enabled = True
                Else
                    'False: not add test , jump to ExA2_1st or jump to TrA3_fwd or jump to LoA2_1st
                    MoveToRun(CurRun.NextUnit.NextUnit)
                    CurRun.Set_BackColor(Color.Yellow)
                    Reset_Test_Time()
                End If
                'load steps
                Clear_Steps()
                Load_Steps()

                'dispose old graph and create new graph
                If showGraph Then
                    Load_New_Graph_CD_True()
                End If
            End If

        ElseIf CurRun.Name = "ExA1_Add" Or CurRun.Name = "TrA1_Add" Or CurRun.Name = "LoA1_Add" Then
            Countdown = True
            All_Panel_Enable()
            Add_Test_Record = True
            'Case2: now is ExA1_Add and next is ExA2_1st or now is TrA1_Add and next is TrA3_fwd or now is LoA1_Add and next is LoA2_1st
            'have an additional test?
            'call a function 
            'if want_add()=true then ... elseif want_add()=false then ... endif


            'CurRun.Link.Enabled = True
            'Add?
            Dim reallyNeedAdd As Boolean
            Dim i As Integer = 4
            If needDetermineAdd Then
                Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                list.Add(CurRun.PrevUnit.GRU) 'adding test 3
                list.Add(CurRun.PrevUnit.PrevUnit.GRU) 'adding testing 2
                list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding test 1
                Dim tempGRU = CurRun.GRU

                While Not IsNothing(tempGRU)
                    list.Add(tempGRU)
                    tempGRU = tempGRU.NextGRU
                    i += 1
                End While
                reallyNeedAdd = DataGrid.NeedAdd(list)
            Else
                reallyNeedAdd = needAdd
            End If

            If reallyNeedAdd Then
                'True: add test
                CurRun.Set_BackColor(Color.Yellow)
                Set_Add_GRU(0, "Run" & i, "")
                Set_Run_Unit()
                'skip button enable
                Button_Skip_Add.Enabled = True
            Else
                'False: not add test
                Set_Panel_BackColor()
                MoveToRun(CurRun.NextUnit)
                Reset_Test_Time()
                'skip button disable
                Button_Skip_Add.Enabled = False
            End If
            'load steps
            Clear_Steps()
            Load_Steps()

            'dispose old graph and create new graph
            If showGraph Then
                Load_New_Graph_CD_True()
            End If
        ElseIf CurRun.Name = "ExA2_1st" Or CurRun.Name = "ExA2_2nd_3rd" Or CurRun.Name = "LoA2_1st" Or CurRun.Name = "LoA2_2nd_3rd" Then
            Countdown = True
            All_Panel_Enable()
            CurRun.PrevUnit.PrevUnit.Link.Enabled = True
            If CurRun.NextUnit.Name = "ExA2_1st_Add" Or CurRun.NextUnit.Name = "LoA2_1st_Add" Then
                'Case: ExA2_2nd_3rd to ExA2_1st_Add or LoA2_2nd_3rd to LoA2_1st_Add
                'have an additional test?
                'call a function 
                'if want_add()=true then ... elseif want_add()=false then ... endif

                CurRun.Set_BackColor(Color.Green)

                'jump to next Run_Unit and change light
                'Add?
                Dim reallyNeedAdd As Boolean
                If needDetermineAdd Then
                    Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                    list.Add(CurRun.GRU) 'adding test 3
                    list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding testing 2
                    list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding test 1
                    reallyNeedAdd = DataGrid.NeedAdd(list)
                Else
                    reallyNeedAdd = needAdd
                End If

                If reallyNeedAdd Then
                    'True: add test
                    MoveToRun(CurRun.NextUnit)
                    Set_Add_GRU(2, "Run4", "")    'True: add test
                    Set_Run_Unit()
                    CurRun.Set_BackColor(Color.Yellow)
                    'load A2's steps
                    Clear_Steps()
                    Load_Steps()
                    'dispose old graph and create new graph
                    If showGraph Then
                        Load_New_Graph_CD_True()
                    End If
                    'skip button enable
                    Button_Skip_Add.Enabled = True
                Else
                    'False: not add test , jump to RSS(Ex) or LoA1(Ex) or LoA3_fwd(Lo)
                    MoveToRun(CurRun.NextUnit.NextUnit.NextUnit.NextUnit)
                    CurRun.Set_BackColor(Color.Yellow)
                    If CurRun.Name = "RSS" Then
                        CurRun.Steps = Load_Steps_helper(CurRun)
                        timeLeft = CurRun.Time
                        timeLabel.Text = timeLeft & " s"
                        Countdown = False
                        If showGraph Then

                            Load_New_Graph_CD_False()
                        End If
                        Clear_Steps()
                        For index = 0 To 8
                            array_step_display(index).Text = ""
                        Next
                    ElseIf CurRun.Name = "LoA1" Then
                        Set_Run_Unit()
                        'load steps
                        Clear_Steps()
                        Load_Steps()
                        'dispose old graph and create new graph
                        If showGraph Then
                            Load_New_Graph_CD_True()
                        End If
                    ElseIf CurRun.Name = "LoA3_fwd" Then
                        Reset_Test_Time()
                        Clear_Steps()
                        Load_Steps()
                        'dispose old graph and create new graph
                        If showGraph Then
                            Load_New_Graph_CD_True()
                        End If
                    End If
                End If
            ElseIf CurRun.NextUnit.Name = "ExA2_1st" Or CurRun.NextUnit.Name = "LoA2_1st" Then
                Set_Panel_BackColor()
                MoveToRun(CurRun.NextUnit)
                Set_Run_Unit()
                'load A2's steps
                Clear_Steps()
                Load_Steps()

                'dispose old graph and create new graph
                If showGraph Then
                    Load_New_Graph_CD_True()
                End If
            End If


        ElseIf CurRun.Name = "ExA2_1st_Add" Or CurRun.Name = "ExA2_2nd_3rd_Add" Or CurRun.Name = "LoA2_1st_Add" Or CurRun.Name = "LoA2_2nd_3rd_Add" Then
            Countdown = True
            All_Panel_Enable()
            Add_Test_Record = True
            'CurRun.PrevUnit.PrevUnit.Link.Enabled = True
            If CurRun.NextUnit.Name = "RSS" Or CurRun.NextUnit.Name = "LoA1" Or CurRun.NextUnit.Name = "LoA3_fwd" Then
                'have an additional test?
                'call a function 
                'if want_add()=true then ... elseif want_add()=false then ... endif

                'jump to next Run_Unit
                'Add?
                Dim reallyNeedAdd As Boolean
                Dim i As Integer = 4
                If needDetermineAdd Then
                    Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                    list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding test 3
                    list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding testing 2
                    list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding test 1
                    Dim tempGRU = CurRun.GRU

                    While Not IsNothing(tempGRU)
                        list.Add(tempGRU)
                        tempGRU = tempGRU.NextGRU
                        i += 1
                    End While
                    reallyNeedAdd = DataGrid.NeedAdd(list)
                Else
                    reallyNeedAdd = needAdd
                End If


                If reallyNeedAdd Then
                    'True: add test
                    Restart_from_2nd_Previous_Run_Unit()
                    Set_Add_GRU(2, "Run" & i, "")
                    'load A2's steps
                    Clear_Steps()
                    Load_Steps()

                    'dispose old graph and create new graph
                    If showGraph Then
                        Load_New_Graph_CD_True()
                    End If
                    'skip button enable
                    Button_Skip_Add.Enabled = True
                Else
                    'False: not add test ,  jump to RSS or LoA1
                    Set_Panel_BackColor()
                    MoveToRun(CurRun.NextUnit)
                    'skip button disable
                    Button_Skip_Add.Enabled = False
                    If CurRun.Name = "RSS" Then
                        CurRun.Steps = Load_Steps_helper(CurRun)
                        timeLeft = CurRun.Time
                        timeLabel.Text = timeLeft & " s"
                        Countdown = False
                        If showGraph Then
                            Load_New_Graph_CD_False()
                        End If
                        Clear_Steps()
                        For index = 0 To 8
                            array_step_display(index).Text = ""
                        Next
                    ElseIf CurRun.Name = "LoA1" Then
                        Set_Run_Unit()
                        'load LoA1's steps
                        Clear_Steps()
                        Load_Steps()
                        'dispose old graph and create new graph
                        If showGraph Then
                            Load_New_Graph_CD_True()
                        End If
                    ElseIf CurRun.Name = "LoA3_fwd" Then
                        Reset_Test_Time()
                        Clear_Steps()
                        Load_Steps()
                        'dispose old graph and create new graph
                        If showGraph Then
                            Load_New_Graph_CD_True()
                        End If
                    End If
                End If
            End If

        ElseIf CurRun.Name = "LoA3_fwd" Or CurRun.Name = "LoA3_bkd" Or CurRun.Name = "TrA3_fwd" Or CurRun.Name = "TrA3_bkd" Then
            Countdown = True
            All_Panel_Enable()
            If CurRun.NextUnit.Name = "LoA3_fwd" Or CurRun.NextUnit.Name = "LoA3_bkd" Or CurRun.NextUnit.Name = "TrA3_fwd" Or CurRun.NextUnit.Name = "TrA3_bkd" Then
                'Case: LoA3_fwd to LoA3_bkd or LoA3_bkd to LoA3_fwd or TrA3_bkd to TrA3_fwd or TrA3_bkd to TrA3_fwd
                ' change light
                Set_Panel_BackColor()
                CurRun.Link.Enabled = True

                'Case:LoA3_bkd to LoA3_fwd or TrA3_bkd to TrA3_fwd or TrA3_bkd to TrA3_fwd
                'If CurRun.NextUnit.Name.Contains("fwd") Then
                '    DataGrid.AddA3Overall(CurRun.PrevUnit.GRU, CurRun.GRU)
                'End If
                ''jump to next Run_Unit
                MoveToRun(CurRun.NextUnit)
                Set_Run_Unit()
                'load LoA3 or TrA3's steps
                Clear_Steps()
                Load_Steps()

                'dispose old graph and create new graph
                If showGraph Then
                    Load_New_Graph_CD_True()
                End If
            ElseIf CurRun.NextUnit.Name = "LoA3_fwd_Add" Or CurRun.NextUnit.Name = "TrA3_fwd_Add" Then
                'Case: LoA3_bkd to LoA3_fwd_Add or TrA3_bkd to TrA3_fwd_Add
                'have an additional test?
                'call a function 
                'if want_add()=true then ... elseif want_add()=false then ... endif

                CurRun.Set_BackColor(Color.Green)
                CurRun.Link.Enabled = True
                'jump to next Run_Unit and change light

                'DataGrid.AddA3Overall(CurRun.PrevUnit.GRU, CurRun.GRU) 'adding overall column
                'Add?
                Dim reallyNeedAdd As Boolean
                If needDetermineAdd Then
                    Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                    list.Add(CurRun.GRU.NextGRU) 'adding test 3 overall
                    list.Add(CurRun.PrevUnit.PrevUnit.GRU.NextGRU) 'adding testing 2 overall
                    list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU.NextGRU) 'adding test 1
                    reallyNeedAdd = DataGrid.NeedAdd(list)
                Else
                    reallyNeedAdd = needAdd
                End If


                If reallyNeedAdd Then
                    'True: add test
                    MoveToRun(CurRun.NextUnit)
                    Set_Add_GRU(3, "Run4", "前進")    'True: add test

                    Set_Run_Unit()
                    CurRun.Set_BackColor(Color.Yellow)
                    'load LoA3_fwd_Add or TrA3_fwd_Add's steps
                    Clear_Steps()
                    Load_Steps()

                    'dispose old graph and create new graph
                    Load_New_Graph_CD_True()
                    'skip button enable
                    Button_Skip_Add.Enabled = True
                Else
                    'False: not add test , jump to RSS
                    MoveToRun(CurRun.NextUnit.NextUnit.NextUnit)
                    CurRun.Set_BackColor(Color.Yellow)
                    timeLeft = CurRun.Time
                    timeLabel.Text = timeLeft & " s"
                    Countdown = False
                    Clear_Steps()
                    For index = 0 To 8
                        array_step_display(index).Text = ""
                    Next
                    'dispose old graph and create new graph
                    Load_New_Graph_CD_False()
                End If
            End If

        ElseIf CurRun.Name = "LoA3_fwd_Add" Or CurRun.Name = "LoA3_bkd_Add" Or CurRun.Name = "TrA3_fwd_Add" Or CurRun.Name = "TrA3_bkd_Add" Then
            Countdown = True
            All_Panel_Enable()
            If CurRun.NextUnit.Name = "LoA3_bkd_Add" Or CurRun.NextUnit.Name = "TrA3_bkd_Add" Then
                Button_Skip_Add.Enabled = True
                'Case: LoA3_fwd_Add to LoA3_bkd_Add or TrA3_fwd_Add to TrA3_bkd_Add
                ' change light
                Set_Panel_BackColor()
                'CurRun.Link.Enabled = True
                'jump to next Run_Unit
                MoveToRun(CurRun.NextUnit)

                If Not IsNothing(CurRun.GRU) Then
                    Dim i As Integer = 5
                    Dim tempGRU As Grid_Run_Unit = CurRun.GRU
                    While tempGRU.NextGRU IsNot Nothing
                        tempGRU = tempGRU.NextGRU
                        i += 1
                    End While
                    Set_Add_GRU(3, "Run" & i, "後退")
                    tempGRU = tempGRU.NextGRU
                    DataGrid.AddA3OverallColumn(tempGRU)
                Else
                    Set_Add_GRU(3, "Run4", "後退")
                    DataGrid.AddA3OverallColumn(CurRun.GRU)
                End If

                Set_Run_Unit()
                'load LoA3 or TrA3's steps
                Clear_Steps()
                Load_Steps()

                'dispose old graph and create new graph
                If showGraph Then
                    Load_New_Graph_CD_True()
                End If
            ElseIf CurRun.NextUnit.Name = "RSS" Then
                Add_Test_Record = True
                'Case: LoA3_bkd_Add to RSS or TrA3_bkd_Add to RSS
                'have an additional test?
                'call a function 
                'if want_add()=true then ... elseif want_add()=false then ... endif

                'CurRun.Link.Enabled = True
                'Add?
                Dim reallyNeedAdd As Boolean
                Dim i As Integer = 5
                Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                list.Add(CurRun.PrevUnit.PrevUnit.GRU.OverallGRU) 'adding test 3
                list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU.OverallGRU) 'adding testing 2
                list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU.OverallGRU) 'adding test 1
                Dim tempGRU = CurRun.GRU
                list.Add(tempGRU)

                While tempGRU.NextGRU IsNot Nothing
                    list.Add(tempGRU.NextGRU)
                    tempGRU = tempGRU.NextGRU
                    i += 1
                End While

                If needDetermineAdd Then
                    reallyNeedAdd = DataGrid.NeedAdd(list)
                Else
                    reallyNeedAdd = needAdd
                End If



                If reallyNeedAdd Then 'True: add test

                    Restart_from_1st_Previous_Run_Unit()

                    If tempGRU.Header.Equals("Run" & i) Then '前進 already exists
                        MoveToRun(CurRun.NextUnit)
                        Set_Add_GRU(3, "Run" & i, "後退")
                    Else
                        Set_Add_GRU(3, "Run" & i, "前進")
                    End If

                    'load LoA3_fwd or TrA3_fwd's steps
                    Clear_Steps()
                    Load_Steps()

                    'dispose old graph and create new graph
                    If showGraph Then
                        Load_New_Graph_CD_True()
                    End If
                    'skip button enable
                    Button_Skip_Add.Enabled = True
                Else
                    'False: not add test , jump to RSS
                    Set_Panel_BackColor()
                    MoveToRun(CurRun.NextUnit)
                    timeLeft = CurRun.Time
                    timeLabel.Text = timeLeft & " s"
                    Countdown = False
                    Clear_Steps()
                    For index = 0 To 8
                        array_step_display(index).Text = ""
                    Next
                    'dispose old graph and create new graph
                    If showGraph Then
                        Load_New_Graph_CD_False()
                    End If
                    'skip button disable
                    Button_Skip_Add.Enabled = False
                End If
            End If

        End If
    End Sub

    '已經做完所有的步驟
    Private Sub BackToEnd()
        startButton.Enabled = False
        stopButton.Enabled = False
        Clear_Steps()
        ClearSeconds()
        timeLeft = 0
        timeLabel.Text = timeLeft & "s"
    End Sub

    '此function是在表格中的資料有變化後會被叫，然後他會去以表格裡有沒有已被接受且未存檔的去決定需不需要儲存
    Private Sub EnableSave()
        'look for the right GRU
        If CurRun IsNot Nothing Then
            If CurRun.GRU IsNot Nothing Then
                If Not CurRun.GRU.NotYetAccepted Then 'accepted so we can allow saving
                    SaveToolStripMenuItem.Enabled = True
                    ButtonExport.Enabled = True
                    Return
                Else 'if more than one additional
                    'search for the right addition
                    Dim tempgru As Grid_Run_Unit = CurRun.GRU
                    'search regular
                    While tempgru IsNot Nothing
                        If Not tempgru.NotYetAccepted Then
                            SaveToolStripMenuItem.Enabled = True
                            ButtonExport.Enabled = True
                            Return
                        End If
                        tempgru = tempgru.NextGRU
                    End While
                End If
            End If
        End If
    End Sub



    'plot the coordinate system on startup
    Private Sub plotCor(ByRef gp As Graphics, ByVal xCor As Double(,), ByVal yCor As Double(,))
        Dim fn As Font = New Font("Microsoft Sans MS", 20)
        'x axis
        gp.DrawLine(Pens.Black, New Point(xCor(0, 0), xCor(0, 1)), New Point(xCor(1, 0), xCor(1, 1)))
        gp.DrawString("X", fn, Brushes.Black, New Point(xCor(1, 0) - 10, xCor(1, 1) + 2))

        'y axis
        gp.DrawLine(Pens.Black, New Point(yCor(0, 0), yCor(0, 1)), New Point(yCor(1, 0), yCor(1, 1)))
        gp.DrawString("Y", fn, Brushes.Black, New Point(yCor(0, 0), yCor(0, 1)))

    End Sub

    'given an array of coordinates, plot them out using the given coordinate system
    'xCor is x axis coordinates
    'yCor is y axis coordinates
    'parent is the parent control to canvas
    'r is the radius for the circle
    'coors is the coordinates for the points to be plotted
    Private Sub plot(ByVal repaint As Boolean, ByRef gp As Graphics, ByVal xCor As Double(,), ByVal yCor As Double(,), ByVal parent As Control)
        'plot the circle with r
        If Not repaint Then
            Dim temp As Double = length / R
            ratio = CInt(temp)

            If ratio > temp Then
                ratio = ratio - 1
            ElseIf ratio = 0 Then
                ratio = temp
            End If
            PlotRatio = ratio
            R = R * ratio
        End If
        'plot the coordinates
        plotCor(gp, xCor, yCor)
        Dim rCircle As Rectangle = New Rectangle(New Point((origin(0) - R), (origin(1) - R)), New Size(R * 2, R * 2))
        gp.DrawEllipse(Pens.Black, rCircle)

        'plot normalized points
        Dim labels = {"P2", "P4", "P6", "P8", "P10", "P12"}

        For index = 0 To pos.GetLength(0) - 1
            Dim x = pos(index).Coors.Xc * ratio
            Dim y = pos(index).Coors.Yc * ratio

            Dim fn As Font = New Font("Microsoft Sans MS", 16)
            'SolidBrush裡有存在圖上的顏色
            Dim solidBrush As SolidBrush = New SolidBrush(posColors(index))
            Dim pHeight As Integer = TextRenderer.MeasureText(gp, labels(index), fn).Height
            Dim pWidth As Integer = TextRenderer.MeasureText(gp, labels(index), fn).Width
            Dim rect As Rectangle = New Rectangle(New Point((origin(0) + x - 2), (origin(1) - y - 2)), New Size(10, 10))
            Dim d As Integer = 40
            Dim e As Integer = 10
            gp.FillEllipse(solidBrush, rect)

            'Text
            If pos.Length = 6 Then
                If labels(index) = "P6" Or labels(index) = "P8" Or labels(index) = "P12" Then
                    gp.DrawString(labels(index), fn, solidBrush, New System.Drawing.Point(origin(0) + x, origin(1) - y))
                    gp.DrawString(String.Format("X: {0:N1}", pos(index).Coors.Xc), fn, solidBrush, New System.Drawing.Point(origin(0) + x, origin(1) - y - pHeight * 3))
                    gp.DrawString(String.Format("Y: {0:N1}", pos(index).Coors.Yc), fn, solidBrush, New System.Drawing.Point(origin(0) + x, origin(1) - y - pHeight * 2))
                    gp.DrawString(String.Format("Z: {0:N1}", pos(index).Coors.Zc), fn, solidBrush, New System.Drawing.Point(origin(0) + x, origin(1) - y - pHeight * 1))
                Else
                    gp.DrawString(labels(index), fn, solidBrush, New System.Drawing.Point(origin(0) + x, origin(1) - y))
                    gp.DrawString(String.Format("X: {0:N1}", pos(index).Coors.Xc), fn, solidBrush, New System.Drawing.Point(origin(0) + x, origin(1) - y + pHeight - e))
                    gp.DrawString(String.Format("Y: {0:N1}", pos(index).Coors.Yc), fn, solidBrush, New System.Drawing.Point(origin(0) + x, origin(1) - y + pHeight * 2 - e))
                    gp.DrawString(String.Format("Z: {0:N1}", pos(index).Coors.Zc), fn, solidBrush, New System.Drawing.Point(origin(0) + x, origin(1) - y + pHeight * 3 - e))
                End If
            Else
                If labels(index) = "P4" Then
                    gp.DrawString(labels(index), fn, solidBrush, New System.Drawing.Point(origin(0) + x, origin(1) - y))
                    gp.DrawString(String.Format("X: {0:N1}", pos(index).Coors.Xc), fn, solidBrush, New System.Drawing.Point(origin(0) + x + pWidth - d, origin(1) - y - pHeight * 3 + e))
                    gp.DrawString(String.Format("Y: {0:N1}", pos(index).Coors.Yc), fn, solidBrush, New System.Drawing.Point(origin(0) + x + pWidth - d, origin(1) - y - pHeight * 2 + e))
                    gp.DrawString(String.Format("Z: {0:N1}", pos(index).Coors.Zc), fn, solidBrush, New System.Drawing.Point(origin(0) + x + pWidth - d, origin(1) - y - pHeight * 1 + e))
                ElseIf labels(index) = "P6" Then
                    gp.DrawString(labels(index), fn, solidBrush, New System.Drawing.Point(origin(0) + x - pWidth + d, origin(1) - y))

                    gp.DrawString(String.Format("X: {0:N1}", pos(index).Coors.Xc), fn, solidBrush, New System.Drawing.Point(origin(0) + x - 2 * d, origin(1) - y - pHeight * 3))
                    gp.DrawString(String.Format("Y: {0:N1}", pos(index).Coors.Yc), fn, solidBrush, New System.Drawing.Point(origin(0) + x - 2 * d, origin(1) - y - pHeight * 2))
                    gp.DrawString(String.Format("Z: {0:N1}", pos(index).Coors.Zc), fn, solidBrush, New System.Drawing.Point(origin(0) + x - 2 * d, origin(1) - y - pHeight * 1))
                Else
                    gp.DrawString(labels(index), fn, solidBrush, New System.Drawing.Point(origin(0) + x, origin(1) - y))
                    gp.DrawString(String.Format("X: {0:N1}", pos(index).Coors.Xc), fn, solidBrush, New System.Drawing.Point(origin(0) + x + pWidth - d, origin(1) - y + pHeight - e))
                    gp.DrawString(String.Format("Y: {0:N1}", pos(index).Coors.Yc), fn, solidBrush, New System.Drawing.Point(origin(0) + x + pWidth - d, origin(1) - y + pHeight * 2 - e))
                    gp.DrawString(String.Format("Z: {0:N1}", pos(index).Coors.Zc), fn, solidBrush, New System.Drawing.Point(origin(0) + x + pWidth - d, origin(1) - y + pHeight * 3 - e))
                End If
            End If
        Next
    End Sub

    'translate the location of point 1 from control 2's coordinates to control 3's coordinates
    'point 1 is given in var 2 coordinates
    'controls 2 and 3 are given in global terms)
    Private Function translate(ByVal p1, ByVal con2, ByVal con3)
        Dim x = p1.X + con2.X - con3.X
        Dim y = p1.Y + con2.Y - con3.Y
        Return New System.Drawing.Point(x, y)
    End Function

    'Menu Strip Functions
    'contains recording
    Public Function Save() As Boolean
        Dim saveFileDialog1 As New SaveFileDialog()
        Dim tempRun As Run_Unit = HeadRun
        saveFileDialog1.Filter = "Chao files (*.cha)|*.chao|All files (*.*)|*.*"
        saveFileDialog1.RestoreDirectory = True
        Dim outfile As StreamWriter

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            outfile = New StreamWriter(saveFileDialog1.OpenFile(), System.Text.Encoding.Unicode)
            If (outfile IsNot Nothing) Then
                'string builder for regular runs
                Dim sb As StringBuilder = New StringBuilder()
                'string builder for additional runs
                Dim A3AddSb As StringBuilder = New StringBuilder()
                'write basic data first
                sb.AppendLine(ComboBox_machine_list.Text & "," & TextBox_L.Text & "," & TextBox_r1.Text & "," & TextBox_L1.Text & "," & TextBox_L2.Text & "," & TextBox_L3.Text & "," & TextBox_r2.Text)
                For k = 0 To BasicInfoGrid.Rows.Count - 1
                    If Not k Mod 2 = 0 Then
                        If BasicInfoGrid.Rows(k).Cells(0).Value IsNot Nothing Then
                            sb.Append(BasicInfoGrid.Rows(k).Cells(0).Value.ToString())
                        End If
                        sb.AppendLine("##")
                    End If
                Next
                While tempRun IsNot Nothing
                    If tempRun.Name.Contains("Add") Or tempRun.Executed Then
                        If tempRun.GRU IsNot Nothing Then
                            If Not tempRun.GRU.NotYetAccepted Then

                                'Saveing regular runs
                                Dim tempGRU As Grid_Run_Unit = tempRun.GRU
                                'FORMAT: HEADER;SUBHEADER;LpAeq2|meter2second1,meter2second2,meter2second3,meter2second4,meter2second5...;LpAeq4|meter4second1,....;....
                                If Machine = Machines.Others Then
                                    sb.AppendLine(tempGRU.Header & ";" &
                                              tempGRU.Subheader & ";" &
                                              tempGRU.LpAeq2 & "|" & DoubleArrayToString(tempGRU.Meter2.Measurements) & ";" &
                                              tempGRU.LpAeq4 & "|" & DoubleArrayToString(tempGRU.Meter4.Measurements) & ";" &
                                              tempGRU.LpAeq6 & "|" & DoubleArrayToString(tempGRU.Meter6.Measurements) & ";" &
                                              tempGRU.LpAeq8 & "|" & DoubleArrayToString(tempGRU.Meter8.Measurements))
                                Else
                                    sb.Append(tempGRU.Header & ";" &
                                              tempGRU.Subheader & ";" &
                                              tempGRU.LpAeq2 & "|" & DoubleArrayToString(tempGRU.Meter2.Measurements) & ";" &
                                              tempGRU.LpAeq4 & "|" & DoubleArrayToString(tempGRU.Meter4.Measurements) & ";" &
                                              tempGRU.LpAeq6 & "|" & DoubleArrayToString(tempGRU.Meter6.Measurements) & ";" &
                                              tempGRU.LpAeq8 & "|" & DoubleArrayToString(tempGRU.Meter8.Measurements) & ";" &
                                              tempGRU.LpAeq10 & "|" & DoubleArrayToString(tempGRU.Meter10.Measurements) & ";" &
                                              tempGRU.LpAeq12 & "|" & DoubleArrayToString(tempGRU.Meter12.Measurements))
                                    'special case for A2 and A3 because we need to record time
                                    If tempRun.Name.Contains("A2") Then
                                        Dim tempstep As Steps = tempRun.HeadStep
                                        sb.Append("(")
                                        While tempstep IsNot Nothing
                                            sb.Append(tempstep.Time)
                                            If tempstep.HasNext Then
                                                sb.Append(",")
                                            End If
                                            tempstep = tempstep.NextStep
                                        End While
                                    ElseIf tempRun.Name.Contains("A3") Then
                                        SaveA3Time(sb, tempRun)
                                    End If
                                    sb.AppendLine()
                                End If

                                'Saving Additional runs
                                If Not tempRun.Name.Contains("A3") Then
                                    Dim addGRU As Grid_Run_Unit = tempGRU.NextGRU
                                    While addGRU IsNot Nothing
                                        If Not addGRU.NotYetAccepted Then
                                            If Machine = Machines.Others Then
                                                sb.AppendLine(addGRU.Header & ";" &
                                                          addGRU.Subheader & ";" &
                                                      addGRU.LpAeq2 & "|" & DoubleArrayToString(addGRU.Meter2.Measurements) & ";" &
                                                      addGRU.LpAeq4 & "|" & DoubleArrayToString(addGRU.Meter4.Measurements) & ";" &
                                                      addGRU.LpAeq6 & "|" & DoubleArrayToString(addGRU.Meter6.Measurements) & ";" &
                                                      addGRU.LpAeq8 & "|" & DoubleArrayToString(addGRU.Meter8.Measurements))
                                                addGRU = addGRU.NextGRU
                                            Else
                                                sb.Append(addGRU.Header & ";" &
                                                          addGRU.Subheader & ";" &
                                                      addGRU.LpAeq2 & "|" & DoubleArrayToString(addGRU.Meter2.Measurements) & ";" &
                                                      addGRU.LpAeq4 & "|" & DoubleArrayToString(addGRU.Meter4.Measurements) & ";" &
                                                      addGRU.LpAeq6 & "|" & DoubleArrayToString(addGRU.Meter6.Measurements) & ";" &
                                                      addGRU.LpAeq8 & "|" & DoubleArrayToString(addGRU.Meter8.Measurements) & ";" &
                                                      addGRU.LpAeq10 & "|" & DoubleArrayToString(addGRU.Meter10.Measurements) & ";" &
                                                      addGRU.LpAeq12 & "|" & DoubleArrayToString(addGRU.Meter12.Measurements))
                                                addGRU = addGRU.NextGRU
                                                If tempRun.Name.Contains("A2") Then
                                                    Dim tempstep As Steps = tempRun.HeadStep
                                                    sb.Append("(")
                                                    While tempstep IsNot Nothing
                                                        sb.Append(tempstep.Time)
                                                        If tempstep.HasNext Then
                                                            sb.Append(",")
                                                        End If
                                                        tempstep = tempstep.NextStep
                                                    End While
                                                End If
                                                sb.AppendLine()

                                            End If
                                        Else
                                            Exit While
                                        End If
                                    End While
                                ElseIf tempRun.Name.Contains("fwd") Then ' A3 is special case because we need to alternate additional runs after the first
                                    'additional. A3fwd_Run4.nextGRU = A3fwd_Run5, A3fwd_Run5.nextGRU = A3fwd_Run6
                                    Dim addGRU As Grid_Run_Unit = tempGRU.NextGRU
                                    Dim bkdGRU As Grid_Run_Unit = tempRun.NextUnit.GRU.NextGRU
                                    While addGRU IsNot Nothing
                                        If Not addGRU.NotYetAccepted Then

                                            A3AddSb.Append(addGRU.Header & ";" &
                                                      addGRU.Subheader & ";" &
                                                  addGRU.LpAeq2 & "|" & DoubleArrayToString(addGRU.Meter2.Measurements) & ";" &
                                                  addGRU.LpAeq4 & "|" & DoubleArrayToString(addGRU.Meter4.Measurements) & ";" &
                                                  addGRU.LpAeq6 & "|" & DoubleArrayToString(addGRU.Meter6.Measurements) & ";" &
                                                  addGRU.LpAeq8 & "|" & DoubleArrayToString(addGRU.Meter8.Measurements) & ";" &
                                                  addGRU.LpAeq10 & "|" & DoubleArrayToString(addGRU.Meter10.Measurements) & ";" &
                                                  addGRU.LpAeq12 & "|" & DoubleArrayToString(addGRU.Meter12.Measurements))
                                            Dim tempstep As Steps = tempRun.HeadStep
                                            SaveA3Time(A3AddSb, tempRun)
                                            A3AddSb.AppendLine()

                                            addGRU = addGRU.NextGRU

                                            If bkdGRU IsNot Nothing Then
                                                If Not bkdGRU.NotYetAccepted Then
                                                    A3AddSb.Append(bkdGRU.Header & ";" &
                                                          bkdGRU.Subheader & ";" &
                                                      bkdGRU.LpAeq2 & "|" & DoubleArrayToString(bkdGRU.Meter2.Measurements) & ";" &
                                                      bkdGRU.LpAeq4 & "|" & DoubleArrayToString(bkdGRU.Meter4.Measurements) & ";" &
                                                      bkdGRU.LpAeq6 & "|" & DoubleArrayToString(bkdGRU.Meter6.Measurements) & ";" &
                                                      bkdGRU.LpAeq8 & "|" & DoubleArrayToString(bkdGRU.Meter8.Measurements) & ";" &
                                                      bkdGRU.LpAeq10 & "|" & DoubleArrayToString(bkdGRU.Meter10.Measurements) & ";" &
                                                      bkdGRU.LpAeq12 & "|" & DoubleArrayToString(bkdGRU.Meter12.Measurements))
                                                    SaveA3Time(A3AddSb, tempRun)
                                                    A3AddSb.AppendLine()

                                                    bkdGRU = bkdGRU.NextGRU
                                                Else
                                                    Exit While
                                                End If
                                            Else
                                                Exit While
                                            End If

                                        Else
                                            Exit While
                                        End If
                                    End While
                                End If
                            Else
                                Exit While
                            End If

                        End If
                    End If
                    tempRun = tempRun.NextUnit
                    If tempRun IsNot Nothing Then
                        If tempRun.Name.Contains("RSS") Then
                            If A3AddSb.Length > 0 Then
                                sb.Append(A3AddSb.ToString())
                            End If
                        End If
                    End If
                End While
                outfile.WriteLine(sb.ToString())
            End If
            outfile.Close()
            DataGrid.ChangedFromLast = False
            BasicInfoDataChangedFromLast = False
            Return True
        End If
        Return False
    End Function

    'Export Function
    Public Function ExportToCSV() As Boolean

        Dim saveFileDialog1 As New SaveFileDialog()

        saveFileDialog1.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*"
        saveFileDialog1.RestoreDirectory = True
        Dim outfile As StreamWriter

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            outfile = New StreamWriter(saveFileDialog1.OpenFile(), System.Text.Encoding.Unicode)
            If (outfile IsNot Nothing) Then
                Dim sb As StringBuilder = New StringBuilder()
                'write basic data first
                sb.AppendLine(ComboBox_machine_list.Text & "," & TextBox_L.Text & "," & TextBox_r1.Text & "," & TextBox_L1.Text & "," & TextBox_L2.Text & "," & TextBox_L3.Text & "," & TextBox_r2.Text)
                For k = 0 To BasicInfoGrid.Rows.Count - 1
                    If BasicInfoGrid.Rows(k).Cells(0).Value IsNot Nothing Then
                        sb.AppendLine(BasicInfoGrid.Rows(k).Cells(0).Value.ToString())
                    End If
                Next
                'write column headers second
                For i = 0 To DataGrid.Form.Columns.Count - 1
                    sb.Append("," & DataGrid.Form.Columns(i).HeaderText)
                Next
                outfile.WriteLine(sb.ToString())
                sb.Clear()

                'write actual data now
                For j = 0 To DataGrid.Form.Rows.Count - 1
                    sb.Append(DataGrid.Form.Rows(j).HeaderCell.Value)
                    For i = 0 To DataGrid.Form.Columns.Count - 1

                        sb.Append(",")
                        If DataGrid.Form.Rows(j).Cells(i).Value IsNot Nothing Then
                            sb.Append(DataGrid.Form.Rows(j).Cells(i).Value.ToString())
                        End If
                    Next
                    sb.AppendLine()
                Next
                outfile.Write(sb.ToString())
            End If
            outfile.Close()
            Return True
        End If
        Return False
    End Function

    '讀取已存檔案
    Public Function LoadFile() As Boolean
        Dim loadFileDialog As New OpenFileDialog()
        Dim inFile As StreamReader
        Try
            If loadFileDialog.ShowDialog() = DialogResult.OK Then
                inFile = New StreamReader(loadFileDialog.OpenFile(), System.Text.Encoding.Unicode)
                If inFile IsNot Nothing Then
                    'liberate data first
                    If Change_Machine() Then

                        'load and write basic data first
                        Dim type_l_r() As String = inFile.ReadLine().Split(",")

                        'putting everything in place
                        ComboBox_machine_list.Text = type_l_r(0)
                        choice = type_l_r(0)
                        MachChosen = True
                        Dim Aes() As String = Decide_Machine().Split(",")
                        Dim AesIdx As Integer = 0

                        TextBox_L.Text = type_l_r(1)
                        TextBox_r1.Text = type_l_r(2)
                        TextBox_L1.Text = type_l_r(3)
                        TextBox_L2.Text = type_l_r(4)
                        TextBox_L3.Text = type_l_r(5)
                        TextBox_r2.Text = type_l_r(6)

                        For k = 0 To BasicInfoGrid.Rows.Count - 1
                            If Not k Mod 2 = 0 Then 'skipping even number columns to avoid titles
                                Dim tempBlob As StringBuilder = New StringBuilder(inFile.ReadLine())
                                While Not tempBlob.ToString().EndsWith("##")
                                    tempBlob.AppendLine(inFile.ReadLine())
                                End While
                                BasicInfoGrid.Rows(k).Cells(0).Value = tempBlob.Remove(tempBlob.Length - 2, 2)
                            End If
                        Next

                        'A123
                        If Not Machine = Machines.Others Then
                            A123_Prepare()
                        Else 'A4
                            A4_Prepare()
                        End If

                        _Warning = False
                        Dim halfwayAdd As Boolean = False
                        'Grid_Run_Unit_Info gives you the format as such- Background;110.9|107.5,69.2,111.9;59.5|70.9,105.7,60.5;54.7|113.2,45.9,55.7;99.3|67.6,44.9,100.3
                        Dim gruInfo() As String

                        Do
                            gruInfo = inFile.ReadLine().Split(";")
                            If Not gruInfo.Length = 0 And Not inFile.EndOfStream Then
                                'runsCount += 1
                                Dim title As String = gruInfo(0)
                                Dim subh As String = gruInfo(1)
                                'Dim time As Integer = CInt(gruInfo(gruInfo.Length - 1))

                                'This if and elseif moves the cursor to the correct place
                                'no additional after 3rd run
                                If CurRun.Name.Contains("Add") And (title.EndsWith("Run1") Or title.EndsWith("RSS")) Then

                                    If CurRun.Name.Contains("A2") Then 'A2, run1->run2->run3->RSS(or Run1)
                                        CurRun = CurRun.NextUnit.NextUnit.NextUnit
                                    ElseIf CurRun.Name.Contains("A3") Then 'A3, fwd->bkd->RSS
                                        CurRun = CurRun.NextUnit.NextUnit
                                    ElseIf CurRun.Name.Contains("A4") Then 'A4, bg->run->RSS
                                        CurRun = CurRun.NextUnit.NextUnit
                                    Else 'A1
                                        CurRun = CurRun.NextUnit 'skipping the additional Run_Unit
                                    End If
                                    'additional run (already passed the first run of additional, if there's more...)
                                ElseIf (Not CurRun.Name.Contains("Add")) And title.Contains("Run") And Not (title = "Run1" Or title = "Run2" Or title = "Run3") Then
                                    'A2
                                    If CurRun.Name.Contains("A2") Then 'A1->A2
                                        'last run has to be A1
                                        CurRun = CurRun.PrevUnit
                                    ElseIf CurRun.Name.Contains("A3") Then
                                        'if last one was A1
                                        If CurRun.PrevUnit.Name.Contains("A1") Then 'A1->A3
                                            CurRun = CurRun.PrevUnit
                                            'if last one was A2
                                        ElseIf CurRun.PrevUnit.Name.Contains("A2") Then 'A1->A2->A3
                                            CurRun = CurRun.PrevUnit.PrevUnit.PrevUnit
                                        End If
                                    ElseIf CurRun.Name.Contains("RSS") Then
                                        If CurRun.PrevUnit.Name.Contains("A2") Then 'A1->A2->RSS
                                            CurRun = CurRun.PrevUnit.PrevUnit.PrevUnit
                                        ElseIf CurRun.PrevUnit.Name.Contains("A3") Then 'A1->A2->A3->RSS or A1->A3->RSS or A1->A2->A3->A1->A3->RSS
                                            CurRun = CurRun.PrevUnit.PrevUnit
                                        ElseIf CurRun.PrevUnit.Name.Contains("A4") Then 'A4->RSS
                                            CurRun = CurRun.PrevUnit.PrevUnit
                                        End If
                                    End If
                                End If

                                'makes sure when pass an A, we increase the AesIdx to know whichA we are at
                                If CurRun.PrevUnit IsNot Nothing Then
                                    If CurRun.PrevUnit.Name.Contains("Add") And title.EndsWith("Run1") Then
                                        AesIdx += 1
                                    End If
                                End If

                                'A2
                                If CurRun.Name.Contains("A2") Then
                                    Dim dl As List(Of Double) = StringToDoubleList(gruInfo(7).Split("|")(1).Split("(")(1))
                                    Dim dlnew As List(Of Double) = New List(Of Double)

                                    For i = 0 To dl.Count
                                        If i = 0 Then
                                            dlnew.Add(-1)
                                        Else
                                            dlnew.Add(dl(i - 1))
                                        End If
                                    Next
                                    'create more gru for additional run
                                    CurRun.Set_BackColor(Color.Green)
                                    CurRun.Link.Enabled = True
                                    CurRun.Executed = True
                                    LoadInputTime(2, dlnew)

                                    If CurRun.Name.Contains("Add") Then
                                        Set_Add_GRU(2, title, subh)
                                    End If
                                    CurRun.NextUnit.Set_BackColor(Color.Green)
                                    CurRun.NextUnit.Executed = True
                                    LoadInputTime(2, dl)

                                    CurRun = CurRun.NextUnit.NextUnit
                                    CurRun.Set_BackColor(Color.Green)
                                    CurRun.Executed = True
                                    LoadInputTime(2, dl)




                                    'A1,A3,A4
                                Else
                                    CurRun.Set_BackColor(Color.Green)
                                    CurRun.Executed = True
                                    CurRun.Link.Enabled = True

                                    'create more gru for additional run
                                    'adding the right additional GRUs
                                    If CurRun.Name.Contains("Add") Then
                                        'if more than one additional run for A3 or A4
                                        'A3
                                        If CurRun.Name.Contains("A3") Then
                                            If CurRun.Name.Contains("fwd") Then
                                                halfwayAdd = True
                                            Else
                                                halfwayAdd = False
                                            End If
                                        ElseIf CurRun.Name.Contains("A4") Then
                                            'A4
                                            If CurRun.Name.Contains("A4_Mid_BG_Add") Then
                                                halfwayAdd = True
                                            Else
                                                halfwayAdd = False
                                            End If
                                        End If

                                        Set_Add_GRU(CInt(Aes(AesIdx).Substring(1)), title, subh)

                                        If CurRun.Name.Contains("bkd") Then
                                            Dim lastbkd As Grid_Run_Unit = CurRun.GRU
                                            While lastbkd.NextGRU IsNot Nothing
                                                lastbkd = lastbkd.NextGRU
                                            End While
                                            DataGrid.AddA3OverallColumn(lastbkd)
                                        End If
                                    End If

                                    'also need to load time for A1 and A3
                                    If CurRun.Name.Contains("A1") Then
                                        CurRun.Steps = New Steps(My.Resources.A1_step1, Step1, Nothing, True, A1Time)
                                        CurRun.HeadStep = CurRun.Steps
                                    ElseIf CurRun.Name.Contains("A3") Then
                                        Dim dl As List(Of Double) = StringToDoubleList(gruInfo(7).Split("|")(1).Split("(")(1))

                                        LoadInputTime(2, dl)
                                    End If
                                End If

                                ShowLoadedDataOnForm(gruInfo)
                                Dim tempGRU As Grid_Run_Unit = CurRun.GRU
                                While tempGRU.NextGRU IsNot Nothing
                                    tempGRU = tempGRU.NextGRU
                                End While
                                tempGRU.Accept()

                                If inFile.Peek() = 13 Then 'if end of file
                                    CurRun.Steps = CurRun.HeadStep
                                    Set_Second_for_Steps()
                                    If CurRun.Name.Contains("A2") Or CurRun.Name.Contains("A3") Then
                                        CurRun.CurStep = CurRun.StartStep
                                        For i = 1 To CurRun.StartStep - 1
                                            CurRun.Steps = CurRun.Steps.NextStep
                                        Next
                                        timeLeft = CurRun.Steps.Time
                                        timeLabel.Text = timeLeft & " s"
                                    End If
                                    If halfwayAdd Then
                                        MoveOnToNextRun(True, False, True)
                                    Else
                                        MoveOnToNextRun(True, True, False)
                                    End If
                                    _Warning = True
                                Else 'not end of file
                                    CurRun.Steps = Load_Steps_helper(CurRun)
                                    MoveToRun(CurRun.NextUnit)
                                End If

                            End If
                        Loop While gruInfo IsNot Nothing And Not inFile.EndOfStream
                    End If
                End If

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox("讀取檔案錯誤:" & ex.Message)
        End Try
        Return False
    End Function

    Private Sub SimulationMode()
        ButtonMeters.Enabled = True
        ButtonSim.Enabled = False
        Comm.Sim()
        PanelMeterSetup.Enabled = False
        sim = True
    End Sub

    Private Sub MetersMode()
        sim = False
        ButtonMeters.Enabled = False
        ButtonSim.Enabled = True
        Comm.Real()
        PanelMeterSetup.Enabled = True
    End Sub

    Public Function DoubleArrayToString(ByRef array As List(Of Double)) As String
        If array IsNot Nothing Then
            Dim sb As StringBuilder = New StringBuilder()
            For i = 0 To array.Count - 1
                sb.Append(array(i))
                If Not i = array.Count - 1 Then
                    sb.Append(",")
                End If
            Next
            Return sb.ToString()
        End If
        Return ""
    End Function

    Public Function StringToDoubleList(ByRef st As String) As List(Of Double)
        Dim dbArray As List(Of Double) = New List(Of Double)
        If st IsNot Nothing Then
            Dim stArray() As String = st.Split(",")

            If stArray IsNot Nothing Then
                If Not stArray.Length = 0 Then
                    For i = 0 To stArray.Length - 1
                        dbArray.Add(CDbl(stArray(i)))
                    Next
                End If
            End If
        End If
        Return dbArray
    End Function

    '處理A3要儲存時間時會遇到的特例
    Function SaveA3Time(ByRef sb As StringBuilder, ByRef tempRun As Run_Unit) As StringBuilder
        'Recording both forward and backward times
        Dim firstRun As Run_Unit
        Dim secRun As Run_Unit
        If tempRun.Name.Contains("fwd") Then
            firstRun = tempRun
            secRun = tempRun.NextUnit
        Else
            firstRun = tempRun.PrevUnit
            secRun = tempRun
        End If

        Dim tempstep As Steps = firstRun.HeadStep
        sb.Append("(")
        While tempstep IsNot Nothing
            sb.Append(tempstep.Time)
            sb.Append(",")
            tempstep = tempstep.NextStep
        End While
        tempstep = secRun.HeadStep
        While tempstep IsNot Nothing
            sb.Append(tempstep.Time)
            If tempstep.HasNext Then
                sb.Append(",")
            End If
            tempstep = tempstep.NextStep
        End While
        Return sb
    End Function

    

    Public Sub MoveToRun(ByRef Run As Run_Unit)
        CurRun = Run
    End Sub

    '位給這個function一行讀取檔裡的資料，他會將其顯示在表中
    Public Sub ShowLoadedDataOnForm(ByRef gruInfo As String())
        'precal and postcal
        If CurRun.Name.Contains("Cal") Then
            Dim num As Integer
            If CurRun.Name.Contains("2") Then
                num = 2
            ElseIf CurRun.Name.Contains("4") Then
                num = 4
            ElseIf CurRun.Name.Contains("6") Then
                num = 6
            ElseIf CurRun.Name.Contains("8") Then
                num = 8
            ElseIf CurRun.Name.Contains("10") Then
                num = 10
            ElseIf CurRun.Name.Contains("12") Then
                num = 12
            End If
            CurRun.GRU.SetM(New Meter_Measure_Unit(num, StringToDoubleList(gruInfo(num / 2 + 1).Split("|")(1)), CDbl(gruInfo(num / 2 + 1).Split("|")(0))), num)
        End If

        'everything else
        Dim tempGRU As Grid_Run_Unit = CurRun.GRU
        If tempGRU IsNot Nothing Then
            While Not IsNothing(tempGRU.NextGRU) 'for more than one additional runs
                tempGRU = tempGRU.NextGRU
            End While
        End If
        Try
            If Machine = Machines.Others Then
                tempGRU.SetMs(New Meter_Measure_Unit(2, StringToDoubleList(gruInfo(2).Split("|")(1)), CDbl(gruInfo(2).Split("|")(0))),
                                                     New Meter_Measure_Unit(4, StringToDoubleList(gruInfo(3).Split("|")(1)), CDbl(gruInfo(3).Split("|")(0))),
                                                     New Meter_Measure_Unit(6, StringToDoubleList(gruInfo(4).Split("|")(1)), CDbl(gruInfo(4).Split("|")(0))),
                                                     New Meter_Measure_Unit(8, StringToDoubleList(gruInfo(5).Split("|")(1)), CDbl(gruInfo(5).Split("|")(0))))
            Else
                tempGRU.SetMs(New Meter_Measure_Unit(2, StringToDoubleList(gruInfo(2).Split("|")(1).Split("(")(0)), CDbl(gruInfo(2).Split("|")(0))),
                                                     New Meter_Measure_Unit(4, StringToDoubleList(gruInfo(3).Split("|")(1).Split("(")(0)), CDbl(gruInfo(3).Split("|")(0))),
                                                     New Meter_Measure_Unit(6, StringToDoubleList(gruInfo(4).Split("|")(1).Split("(")(0)), CDbl(gruInfo(4).Split("|")(0))),
                                                     New Meter_Measure_Unit(8, StringToDoubleList(gruInfo(5).Split("|")(1).Split("(")(0)), CDbl(gruInfo(5).Split("|")(0))),
                                                     New Meter_Measure_Unit(10, StringToDoubleList(gruInfo(6).Split("|")(1).Split("(")(0)), CDbl(gruInfo(6).Split("|")(0))),
                                                     New Meter_Measure_Unit(12, StringToDoubleList(gruInfo(7).Split("|")(1).Split("(")(0)), CDbl(gruInfo(7).Split("|")(0))))
            End If
        Catch ex As Exception

            MsgBox(ex.Message)

        End Try


        'make sure countdown is correct
        If CurRun.Name.Contains("A1") Or CurRun.Name.Contains("A2") Or CurRun.Name.Contains("A3") Then
            Countdown = True
        Else
            Countdown = False
        End If

        DataGrid.ShowGRUonForm(tempGRU)
        If tempGRU.Header.Contains("RSS") Then
            DataGrid.ShowCalculated()
        End If
        If CurRun.Name.Contains("bkd") Then
            Dim latestFwdGRU As Grid_Run_Unit = CurRun.PrevUnit.GRU
            Dim latestBkdGRU As Grid_Run_Unit = CurRun.GRU
            While latestBkdGRU.NextGRU IsNot Nothing
                latestFwdGRU = latestFwdGRU.NextGRU
                latestBkdGRU = latestBkdGRU.NextGRU
            End While
            DataGrid.ShowA3Overall(latestFwdGRU, latestBkdGRU)
        End If
    End Sub


    Private Sub PrintPage(ByVal sender As Object, ByVal ev As PrintPageEventArgs)
        Try
            Dim pageBitmap As Bitmap = New Bitmap(GroupBox_Plot.Width, GroupBox_Plot.Height)
            Dim rect As Rectangle = New Rectangle(New Point(0, 0), New Size(GroupBox_Plot.Width, GroupBox_Plot.Height))
            GroupBox_Plot.DrawToBitmap(pageBitmap, rect)
            ev.Graphics.DrawImage(pageBitmap, New Point(0, 0))

        Catch ex As Exception
            MsgBox("列印錯誤:" & ex.Message)
        End Try

    End Sub


    Private Sub ThrowPrePostCalWarnings()
        'IF PRECAL
        If CurRun.Name.Contains("PreCal") Then
            Dim value As Integer
            If CurRun.Name.Contains("P2") Then
                value = CurRun.GRU.LpAeq2
            ElseIf CurRun.Name.Contains("P4") Then
                value = CurRun.GRU.LpAeq4
            ElseIf CurRun.Name.Contains("P6") Then
                value = CurRun.GRU.LpAeq6
            ElseIf CurRun.Name.Contains("P8") Then
                value = CurRun.GRU.LpAeq8
            ElseIf CurRun.Name.Contains("P10") Then
                value = CurRun.GRU.LpAeq10
            ElseIf CurRun.Name.Contains("P12") Then
                value = CurRun.GRU.LpAeq12
            End If
            If value > 94.7 Or value < 93.3 Then
                MsgBox("校正值超出94 ± 0.7dB範圍，建議修正!")
            End If
            'IF PostCal
        ElseIf CurRun.Name.Contains("PostCal") Then
            Dim value1 As Integer
            If CurRun.Name.Contains("P2") Then
                value1 = CurRun.GRU.LpAeq2
            ElseIf CurRun.Name.Contains("P4") Then
                value1 = CurRun.GRU.LpAeq4
            ElseIf CurRun.Name.Contains("P6") Then
                value1 = CurRun.GRU.LpAeq6
            ElseIf CurRun.Name.Contains("P8") Then
                value1 = CurRun.GRU.LpAeq8
            ElseIf CurRun.Name.Contains("P10") Then
                value1 = CurRun.GRU.LpAeq10
            ElseIf CurRun.Name.Contains("P12") Then
                value1 = CurRun.GRU.LpAeq12
            End If
            Dim tempRun As Run_Unit = CurRun
            While tempRun IsNot Nothing
                If tempRun.Name.Contains("PreCal") Then
                    Dim value2 As Integer
                    If CurRun.Name.Contains("P2") Then
                        value2 = tempRun.GRU.LpAeq2
                    ElseIf CurRun.Name.Contains("P4") Then
                        value2 = tempRun.GRU.LpAeq4
                    ElseIf CurRun.Name.Contains("P6") Then
                        value2 = tempRun.GRU.LpAeq6
                    ElseIf CurRun.Name.Contains("P8") Then
                        value2 = tempRun.GRU.LpAeq8
                    ElseIf CurRun.Name.Contains("P10") Then
                        value2 = tempRun.GRU.LpAeq10
                    ElseIf CurRun.Name.Contains("P12") Then
                        value2 = tempRun.GRU.LpAeq12
                    End If
                    If Math.Abs(value2 - value1) > 0.3 Then
                        MsgBox("前後校正值差0.3dB以上!")
                        Exit While
                    End If
                End If
                tempRun = tempRun.PrevUnit
            End While
        ElseIf CurRun.Name.Contains("RSS") Then
            If CurRun.GRU.LpAeqAvg - CurRun.GRU.Background.LpAeqAvg < 7 Then
                MsgBox("RSS-背景噪音未達7dB!!")
            End If
        End If
    End Sub

    Private Sub BackToZero()
        sum_steps = 0
        Last_timeLeft = 0
        Index_for_Setup_Time = 0
        TimerTesting = False
        MachChosen = False
        Machine = Nothing
        choice = Nothing

        Add_Test_Record = False
        startButton.Enabled = True
        stopButton.Enabled = False
        Button_Skip_Add.Enabled = False
        Test_NextButton.Enabled = False
        Test_StartButton.Enabled = True
        Test_ConfirmButton.Enabled = False
        ConnectButton.Enabled = True
        DisconnButton.Enabled = False
        SaveToolStripMenuItem.Enabled = False
        For i = 0 To 8
            array_step_s(i).Text = 0
            array_step_s(i).Enabled = False
        Next
        Button_change_machine.Enabled = False
        ComboBox_machine_list.Enabled = True
        Button_Setting_Bargraph.Enabled = False

        Null_CurRun = New Run_Unit(LinkLabel_Temp, Panel_Temp, Nothing, Nothing, 0, "Temp", 0, 0, 0)
        Temp_CurRun = Null_CurRun
        BasicInfoDataChangedFromLast = False
    End Sub
End Class
