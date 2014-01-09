﻿Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.PowerPacks

Public Class Program
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MARGINS
        Public cxLeftWidth As Integer
        Public cxRightWidth As Integer
        Public cyTopHeight As Integer
        Public cyButtomHeight As Integer
    End Structure

    <DllImport("dwmapi.dll")> Public Shared Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As MARGINS) As Integer
    End Function

    'coordinates for 6 points
    Dim pos(5) As CoorPoint
    'Dim pos(5, 2) As Double
    'lines coresponding to the points

    'coordinates for 4 points
    Dim pos2(3) As CoorPoint
    'Dim pos2(3, 2) As Double
    'lines coresponding to the points

    'Origin for the coordinates system
    Dim origin(1) As Double
    Dim length As Double
    ' Set plotting vars
    Dim canvas As New ShapeContainer
    Dim xAxis As New LineShape
    ' xCor(0) is the start point, xCor(1) is the end point
    Dim yAxis As New LineShape
    ' x y axis coordinates
    Dim xCor(1, 1) As Double
    Dim yCor(1, 1) As Double
    Dim ratio As Double
    Dim Points As ArrayList



    'determine whether we can go to the main tab
    Dim MachChosen As Boolean
    Enum Machines
        Excavator
        Loader
        Loader_Excavator
        Tractor
        Others
    End Enum
    Dim Machine As Machines

    'Graphs for the main screen
    Dim MainLineGraphs As LineGraphPanel
    Dim MainBarGraph As BarGraph

    Dim timeLeft As Integer

    'Runs throughout
    Dim HeadRun As Run_Unit
    Dim CurRun As Run_Unit

    'Steps on the Right
    Dim HeadStep As Steps
    Dim CurStep As Steps
    Dim array_step(9) As Label

    Dim Countdown As Boolean = False

    'trial represents 1st or 2nd or 3rd
    Dim trial As Integer

    Dim NoisesArray(6) As Label

    Dim sum_steps As Integer = 0

    'set a random boolean
    Dim RandGen As New Random
    Dim RandBool As Boolean

    'temporary constant
    Const seconds As Integer = 3
    Const graphW As Integer = 650
    Const graphH As Integer = 200
    Dim ASize As Size = New Size(51, 387)
    'used for first time clicking start button
    Dim NewGraph As Boolean

    Public Sub New()

        ' 此為設計工具所需的呼叫。
        InitializeComponent()
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        ' 在 InitializeComponent() 呼叫之後加入任何初始設定。

    End Sub

    ' things that need to be instantiated at startup
    Private Sub Program_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Set up the window
        Dim margins As MARGINS = New MARGINS
        margins.cxLeftWidth = 2
        margins.cxRightWidth = 2
        margins.cyTopHeight = 23
        margins.cyButtomHeight = 2
        Dim hwnd As IntPtr = Handle
        Dim result As Integer = DwmExtendFrameIntoClientArea(hwnd, margins)

        'Set up Choose Machine Page
        TextBox_L.Parent = GroupBox_A1_A2_A3
        TextBox_r1.Parent = GroupBox_A1_A2_A3
        Button_L_check.Parent = GroupBox_A1_A2_A3
        GroupBox_A1_A2_A3.Enabled = False

        TextBox_L1.Parent = GroupBox_A4
        TextBox_L2.Parent = GroupBox_A4
        TextBox_L3.Parent = GroupBox_A4
        TextBox_r2.Parent = GroupBox_A4
        Button_L1_L2_L3_check.Parent = GroupBox_A4
        GroupBox_A4.Enabled = False

        GroupBox_Plot.Parent = TabPage1
        GroupBox_Plot.AllowDrop = True
        xLabel.Parent = GroupBox_Plot
        yLabel.Parent = GroupBox_Plot
        GroupBox_Plot.BackColor = Color.Transparent

        'Set up Buttons
        startButton.Enabled = True
        stopButton.Enabled = False

        'Set up Plotting
        Points = New ArrayList
        pos(0) = New CoorPoint(p1Label)
        pos(1) = New CoorPoint(p2Label)
        pos(2) = New CoorPoint(p3Label)
        pos(3) = New CoorPoint(p4Label)
        pos(4) = New CoorPoint(p5Label)
        pos(5) = New CoorPoint(p6Label)
        pos2(0) = New CoorPoint(p7Label)
        pos2(1) = New CoorPoint(p8Label)
        pos2(2) = New CoorPoint(p9Label)
        pos2(3) = New CoorPoint(p10Label)
        For i = 0 To 5
            pos(i).Label.Parent = TabPage1
            Points.Add(pos(i))
        Next
        For i = 0 To 3
            pos2(i).Label.Parent = TabPage1
            Points.Add(pos2(i))
        Next

        GroupBox_Plot.Visible = True

        'x
        origin(0) = 200
        'y
        origin(1) = 150
        length = 150
        'xStart
        xCor(0, 0) = origin(0) - length 'x
        xCor(0, 1) = origin(1) 'y
        'xEnd
        xCor(1, 0) = origin(0) + length 'x
        xCor(1, 1) = origin(1) 'y

        'yStart
        yCor(0, 0) = origin(0) 'x
        yCor(0, 1) = origin(1) - length 'y
        'yEnd
        yCor(1, 0) = origin(0) 'x
        yCor(1, 1) = origin(1) + length 'y
        'draw coordinates
        canvas.Parent = GroupBox_Plot
        canvas.Location = New System.Drawing.Point(0, 0)
        plotCor(xCor, yCor)



        'Set invisible
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

        'location of all objects

        'LinkLabel_preCal.Location = New Point(50, 22)
        'LinkLabel_BG.Location = New Point(50, 62)
        'LinkLabel_RSS.Location = New Point(50, 262)
        'LinkLabel_postCal.Location = New Point(50, 302)


        startButton.Location = New Point(119, 126)
        Accept_Button.Location = New Point(119, 194)
        Accept_Button.Enabled = False
        timeLabel.Location = New Point(242, 126)
        stopButton.Location = New Point(367, 126)
        testButton.Location = New Point(242, 231)

        NoisesArray = {Noise1, Noise2, Noise3, Noise4, Noise5, Noise6, Noise_Avg}
        Dim noiseX = 65
        Noise1.Location = New Point(425, 654)
        Noise2.Location = New Point(425 + noiseX, 654)
        Noise3.Location = New Point(425 + noiseX * 2, 654)
        Noise4.Location = New Point(425 + noiseX * 3, 654)
        Noise5.Location = New Point(425 + noiseX * 4, 654)
        Noise6.Location = New Point(425 + noiseX * 5, 654)
        Noise_Avg.Location = New Point(425 + noiseX * 6, 654)

        Dim stepX As Integer = 900
        Dim stepY As Integer = 63
        Step1.Location = New Point(stepX, 106)
        Step2.Location = New Point(stepX, 106 + stepY * 1)
        Step3.Location = New Point(stepX, 106 + stepY * 2)
        Step4.Location = New Point(stepX, 106 + stepY * 3)
        Step5.Location = New Point(stepX, 106 + stepY * 4)
        Step6.Location = New Point(stepX, 106 + stepY * 5)
        Step7.Location = New Point(stepX, 106 + stepY * 6)
        Step8.Location = New Point(stepX, 106 + stepY * 7)
        Step9.Location = New Point(stepX, 106 + stepY * 8)
        Step10.Location = New Point(stepX, 106 + stepY * 9)
        array_step(0) = Step1
        array_step(1) = Step2
        array_step(2) = Step3
        array_step(3) = Step4
        array_step(4) = Step5
        array_step(5) = Step6
        array_step(6) = Step7
        array_step(7) = Step8
        array_step(8) = Step9
        array_step(9) = Step10


        Label1.Location = New Point(136, 329)
        Label2.Location = New Point(136, 417)
        Label3.Location = New Point(136, 509)
        Label4.Location = New Point(136, 602)

        Panel_PreCal.Location = New Point(10, 265)
        Panel_PreCal_Sub.Location = New Point(10, 295)
        Panel_Bkg.Location = New Point(10, 480)
        Panel_RSS.Location = New Point(10, 510)
        Panel_PostCal.Location = New Point(10, 540)
        Panel_PostCal_Sub.Location = New Point(10, 570)

        MachChosen = False
    End Sub




    'prevents user from clicking on tab b4 choosing machinery
    Private Sub TabControl1_IndexChanged(ByVal sender As TabControl, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If Not MachChosen Then
            If sender.SelectedIndex = 1 Then
                MessageBox.Show("還未選機具!", "My application", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                TabControl1.SelectedIndex = 0 ' go back to the machinery choose stage
            End If
        End If
    End Sub

    Private Sub ComboBox_machine_list_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_machine_list.SelectedIndexChanged

        'step1~9清空
        For index = 0 To 8
            array_step(index).Text = ""
            array_step(index).BackColor = Color.DarkGray
        Next
        Dim choice As String = ComboBox_machine_list.Text

        If choice = "開挖機(Excavator)" Then
            Picture_machine.Image = My.Resources.Resource1.小型開挖機_compact_excavator_
            Machine = Machines.Excavator
            Load_Excavator()
            MachChosen = True
        ElseIf choice = "推土機(Crawler and wheel tractor)" Then  'A1+A3
            Picture_machine.Image = My.Resources.Resource1.履帶式推土機_crawler_dozer_
            Machine = Machines.Tractor
            Load_Tractor()
            MachChosen = True
        ElseIf choice = "鐵輪壓路機(Road roller)" Then
            Picture_machine.Image = My.Resources.Resource1.壓路機_rollers_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "膠輪壓路機(Wheel roller)" Then
            Picture_machine.Image = My.Resources.Resource1.壓路機_rollers_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "振動式壓路機(Vibrating roller)" Then
            Picture_machine.Image = My.Resources.Resource1.壓路機_rollers_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "裝料機(Crawler and wheel loader)" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式裝料機_crawler_loader_
            Machine = Machines.Loader
            Load_Loader()
            MachChosen = True
        ElseIf choice = "裝料開挖機" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式開挖裝料機_crawler_backhoe_loader_
            Machine = Machines.Loader_Excavator
            Load_Loader_Excavator()
            MachChosen = True
        ElseIf choice = "履帶起重機(Crawler crane)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式推土機_wheeled_dozer_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "卡車起重機(Truck crane)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式開挖裝料機_wheeled_backhoe_loader_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "輪形起重機(Wheel crane)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式開挖機_wheeled_excavator_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "振動式樁錘(Vibrating hammer)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式裝料機_wheeled_loader_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "油壓式打樁機(Hydraulic pile driver)" Then
            Picture_machine.Image = My.Resources.Resource1.壓路機_rollers_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "拔樁機" Then
            Picture_machine.Image = My.Resources.Resource1.小型開挖機_compact_excavator_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "油壓式拔樁機" Then
            Picture_machine.Image = My.Resources.Resource1.小型膠輪式裝料機_compact_loader__wheeled_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "土壤取樣器(地鑽) (Earth auger)" Then
            Picture_machine.Image = My.Resources.Resource1.滑移裝料機_skid_steer_loader_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "全套管鑽掘機" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式開挖裝料機_crawler_backhoe_loader_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "鑽土機(Earth drill)" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式推土機_crawler_dozer_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "鑽岩機(Rock breaker)" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式裝料機_crawler_loader_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "混凝土泵車(Concrete pump)" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式開挖機_crawler_excavator_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "混凝土破碎機(Concrete breaker)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式推土機_wheeled_dozer_
            Step1.Text = My.Resources.A4_Concrete_Breaker
            Machine = Machines.Others
            Load_Others(choice)
        ElseIf choice = "瀝青混凝土舖築機(Asphalt finisher)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式開挖裝料機_wheeled_backhoe_loader_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "混凝土割切機(Concrete cutter)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式開挖機_wheeled_excavator_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "發電機(Generator)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式裝料機_wheeled_loader_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "空氣壓縮機(Compressor)" Then
            Picture_machine.Image = My.Resources.Resource1.壓路機_rollers_
            Machine = Machines.Others
            Load_Others(choice)
            MachChosen = True
        Else
            Machine = Nothing
            MachChosen = False
        End If


        'mach區分要輸入至哪個 groupbox for plotting the coordinates
        If Not IsNothing(Machine) Then
            If Not Machine = Machines.Others Then
                GroupBox_A1_A2_A3.Enabled = True
                GroupBox_A4.Enabled = False
            Else
                GroupBox_A1_A2_A3.Enabled = False
                GroupBox_A4.Enabled = True
            End If

            'if machine chose, then allow to go to the 2nd tab
            'MachChosen = True
        End If
    End Sub

    Private Sub Button_L_check_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_L_check.Click

        Dim r1 As Integer
        If String.IsNullOrWhiteSpace(TextBox_L.Text) Then
            Return
        End If

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

        plot(xCor, yCor, GroupBox_Plot, r1, pos)

    End Sub

    Private Sub Button_L1_L2_L3_check_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_L1_L2_L3_check.Click
        If String.IsNullOrWhiteSpace(TextBox_L1.Text) Or String.IsNullOrWhiteSpace(TextBox_L2.Text) Or String.IsNullOrWhiteSpace(TextBox_L3.Text) Then
            Return
        End If
        Dim L1 As Double
        Dim L2 As Double
        Dim L3 As Double
        Dim D0_2 As Double
        L1 = TextBox_L1.Text
        L2 = TextBox_L2.Text
        L3 = TextBox_L3.Text
        D0_2 = 2 * Math.Sqrt(((L1 / 2) ^ 2) + ((L2 / 2) ^ 2) + (L3 ^ 2))
        TextBox_r2.Text = D0_2
        Dim r = D0_2
        pos2(0).Coors = New ThreeDPoint(D0_2 * -0.45, D0_2 * 0.77, D0_2 * 0.45)
        pos2(1).Coors = New ThreeDPoint(D0_2 * -0.45, D0_2 * -0.77, D0_2 * 0.45)
        pos2(2).Coors = New ThreeDPoint(D0_2 * 0.89, 0, D0_2 * 0.45)
        pos2(3).Coors = New ThreeDPoint(0, 0, D0_2)

        plot(xCor, yCor, GroupBox_Plot, r, pos2)

    End Sub

    Sub Set_Panel(ByRef p As Panel, ByRef l As Label)
        If p.Name = "Panel_PreCal_1st" Then
            p.Size = New Size(33, 26)
            p.BackColor = Color.Yellow
            p.Controls.Add(l)
            l.Location = New Point(3, 5)
            l.ForeColor = Color.White
        ElseIf p.Name = "Panel_Bkg" Or p.Name = "Panel_RSS" Or p.Name = "Panel_PostCal" Then
            p.Size = New Size(85, 26)
            p.BackColor = Color.IndianRed
            p.Controls.Add(l)
            l.Location = New Point(3, 5)
            l.ForeColor = Color.White
        Else
            p.Size = New Size(33, 26)
            p.BackColor = Color.IndianRed
            p.Controls.Add(l)
            l.Location = New Point(3, 5)
            l.ForeColor = Color.White
        End If

    End Sub

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
        PanelExcavatorA1.Location = New Point(220, 265)
        PanelExcavatorA2.Size = ASize
        PanelExcavatorA2.Location = New Point(284, 265)

        'new add
        Panel_PreCal_Sub.Visible = True
        Panel_PostCal_Sub.Visible = True
        Panel_PreCal_5th.Visible = True
        Panel_PreCal_6th.Visible = True
        Panel_PostCal_5th.Visible = True
        Panel_PostCal_6th.Visible = True

        PanelExcavatorA1.Controls.Add(Panel_ExA1_Fst_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Fst_2nd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Fst_3rd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Sec_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Sec_2nd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Sec_3rd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Thd_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Thd_2nd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Thd_3rd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Add_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Add_2nd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Add_3rd)

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

        'Precal
        'Set_Panel(Panel_PreCal, LinkLabel_preCal)
        Set_Panel(Panel_PreCal_1st, LinkLabel_PreCal_1st)
        tempRun = New Run_Unit(LinkLabel_PreCal_1st, Panel_PreCal_1st, Nothing, Nothing, 0, "PreCal", 0, 0, 0)
        HeadRun = tempRun
        CurRun = HeadRun
        timeLeft = CurRun.Time
        timeLabel.Text = timeLeft & " s"

        Set_Panel(Panel_PreCal_2nd, LinkLabel_PreCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_2nd, Panel_PreCal_2nd, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_3rd, LinkLabel_PreCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_3rd, Panel_PreCal_3rd, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_4th, LinkLabel_PreCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_4th, Panel_PreCal_4th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_5th, LinkLabel_PreCal_5th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_5th, Panel_PreCal_5th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_6th, LinkLabel_PreCal_6th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_6th, Panel_PreCal_6th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        'Background
        Set_Panel(Panel_Bkg, LinkLabel_BG)
        tempRun.NextUnit = New Run_Unit(LinkLabel_BG, Panel_Bkg, Nothing, tempRun, 0, "Background", 0, 0, 0)
        tempRun = tempRun.NextUnit

        tempRun = Load_Excavator_Helper(tempRun)

        'RSS
        Set_Panel(Panel_RSS, LinkLabel_RSS)
        tempRun.NextUnit = New Run_Unit(LinkLabel_RSS, Panel_RSS, Nothing, tempRun, 0, "RSS", 0, 0, 0)
        tempRun = tempRun.NextUnit
        'PostCal
        'Set_Panel(Panel_PostCal, LinkLabel_postCal)
        Set_Panel(Panel_PostCal_1st, LinkLabel_PostCal_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_1st, Panel_PostCal_1st, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_2nd, LinkLabel_PostCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_2nd, Panel_PostCal_2nd, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_3rd, LinkLabel_PostCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_3rd, Panel_PostCal_3rd, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_4th, LinkLabel_PostCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_4th, Panel_PostCal_4th, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_5th, LinkLabel_PostCal_5th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_5th, Panel_PostCal_5th, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_6th, LinkLabel_PostCal_6th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_6th, Panel_PostCal_6th, Nothing, tempRun, 0, "PostCal_Last", 0, 0, 0)
        tempRun = tempRun.NextUnit

        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, HeadRun.Time, 97)
        MainBarGraph = New BarGraph(New Point(508, 106), New Size(390, 539), TabPage2, CGraph.Modes.A1A2A3)

    End Sub

    Function Load_Excavator_Helper(ByRef run As Run_Unit)
        'Create an object for each step
        Dim tempRun As Run_Unit = run
        ''A1
        'A1 First 1
        Set_Panel(Panel_ExA1_Fst_1st, LinkLabel_ExA1_Fst_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Fst_1st, Panel_ExA1_Fst_1st, Nothing, tempRun, 5, "ExA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 First 2
        Set_Panel(Panel_ExA1_Fst_2nd, LinkLabel_ExA1_Fst_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Fst_2nd, Panel_ExA1_Fst_2nd, Nothing, tempRun, 5, "ExA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 First 3
        Set_Panel(Panel_ExA1_Fst_3rd, LinkLabel_ExA1_Fst_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Fst_3rd, Panel_ExA1_Fst_3rd, Nothing, tempRun, 5, "ExA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Second 1
        Set_Panel(Panel_ExA1_Sec_1st, LinkLabel_ExA1_Sec_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Sec_1st, Panel_ExA1_Sec_1st, Nothing, tempRun, 5, "ExA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Second 2
        Set_Panel(Panel_ExA1_Sec_2nd, LinkLabel_ExA1_Sec_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Sec_2nd, Panel_ExA1_Sec_2nd, Nothing, tempRun, 5, "ExA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Second 3
        Set_Panel(Panel_ExA1_Sec_3rd, LinkLabel_ExA1_Sec_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Sec_3rd, Panel_ExA1_Sec_3rd, Nothing, tempRun, 5, "ExA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Third 1
        Set_Panel(Panel_ExA1_Thd_1st, LinkLabel_ExA1_Thd_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Thd_1st, Panel_ExA1_Thd_1st, Nothing, tempRun, 5, "ExA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Third 2
        Set_Panel(Panel_ExA1_Thd_2nd, LinkLabel_ExA1_Thd_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Thd_2nd, Panel_ExA1_Thd_2nd, Nothing, tempRun, 5, "ExA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Third 3
        Set_Panel(Panel_ExA1_Thd_3rd, LinkLabel_ExA1_Thd_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Thd_3rd, Panel_ExA1_Thd_3rd, Nothing, tempRun, 5, "ExA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Add 1
        Set_Panel(Panel_ExA1_Add_1st, LinkLabel_ExA1_Add_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Add_1st, Panel_ExA1_Add_1st, Nothing, tempRun, 5, "ExA1_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Add 2
        Set_Panel(Panel_ExA1_Add_2nd, LinkLabel_ExA1_Add_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Add_2nd, Panel_ExA1_Add_2nd, Nothing, tempRun, 5, "ExA1_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Add 3
        Set_Panel(Panel_ExA1_Add_3rd, LinkLabel_ExA1_Add_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA1_Add_3rd, Panel_ExA1_Add_3rd, Nothing, tempRun, 5, "ExA1_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)

        ''A2
        'A2 First 1
        Set_Panel(Panel_ExA2_Fst_1st, LinkLabel_ExA2_Fst_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Fst_1st, Panel_ExA2_Fst_1st, Nothing, tempRun, 1, "ExA2_1st", 1, 1, 9)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 First 2
        Set_Panel(Panel_ExA2_Fst_2nd, LinkLabel_ExA2_Fst_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Fst_2nd, Panel_ExA2_Fst_2nd, Nothing, tempRun, 1, "ExA2_2nd_3rd", 1, 1, 8)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 First 3
        Set_Panel(Panel_ExA2_Fst_3rd, LinkLabel_ExA2_Fst_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Fst_3rd, Panel_ExA2_Fst_3rd, Nothing, tempRun, 1, "ExA2_2nd_3rd", 1, 1, 8)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Second 1
        Set_Panel(Panel_ExA2_Sec_1st, LinkLabel_ExA2_Sec_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Sec_1st, Panel_ExA2_Sec_1st, Nothing, tempRun, 1, "ExA2_1st", 1, 1, 9)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Second 2
        Set_Panel(Panel_ExA2_Sec_2nd, LinkLabel_ExA2_Sec_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Sec_2nd, Panel_ExA2_Sec_2nd, Nothing, tempRun, 1, "ExA2_2nd_3rd", 1, 1, 8)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Second 3
        Set_Panel(Panel_ExA2_Sec_3rd, LinkLabel_ExA2_Sec_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Sec_3rd, Panel_ExA2_Sec_3rd, Nothing, tempRun, 1, "ExA2_2nd_3rd", 1, 1, 8)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Third 1
        Set_Panel(Panel_ExA2_Thd_1st, LinkLabel_ExA2_Thd_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Thd_1st, Panel_ExA2_Thd_1st, Nothing, tempRun, 1, "ExA2_1st", 1, 1, 9)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Third 2
        Set_Panel(Panel_ExA2_Thd_2nd, LinkLabel_ExA2_Thd_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Thd_2nd, Panel_ExA2_Thd_2nd, Nothing, tempRun, 1, "ExA2_2nd_3rd", 1, 1, 8)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Third 3
        Set_Panel(Panel_ExA2_Thd_3rd, LinkLabel_ExA2_Thd_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Thd_3rd, Panel_ExA2_Thd_3rd, Nothing, tempRun, 1, "ExA2_2nd_3rd", 1, 1, 8)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Add 1
        Set_Panel(Panel_ExA2_Add_1st, LinkLabel_ExA2_Add_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Add_1st, Panel_ExA2_Add_1st, Nothing, tempRun, 1, "ExA2_1st_Add", 1, 1, 9)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Add 2
        Set_Panel(Panel_ExA2_Add_2nd, LinkLabel_ExA2_Add_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Add_2nd, Panel_ExA2_Add_2nd, Nothing, tempRun, 1, "ExA2_2nd_3rd_Add", 1, 1, 8)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Add 3
        Set_Panel(Panel_ExA2_Add_3rd, LinkLabel_ExA2_Add_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_ExA2_Add_3rd, Panel_ExA2_Add_3rd, Nothing, tempRun, 1, "ExA2_2nd_3rd_Add", 1, 1, 8)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)

        Return tempRun

    End Function

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
        PanelLoaderA1.Location = New Point(220, 265)
        PanelLoaderA2.Size = ASize
        PanelLoaderA2.Location = New Point(284, 265)
        PanelLoaderA3.Size = ASize
        PanelLoaderA3.Location = New Point(341, 265)

        'new add
        Panel_PreCal_Sub.Visible = True
        Panel_PostCal_Sub.Visible = True
        Panel_PreCal_5th.Visible = True
        Panel_PreCal_6th.Visible = True
        Panel_PostCal_5th.Visible = True
        Panel_PostCal_6th.Visible = True

        PanelLoaderA1.Controls.Add(Panel_LoA1_Fst_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Fst_2nd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Fst_3rd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Sec_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Sec_2nd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Sec_3rd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Thd_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Thd_2nd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Thd_3rd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Add_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Add_2nd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Add_3rd)

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
        Set_Panel(Panel_PreCal_1st, LinkLabel_PreCal_1st)
        tempRun = New Run_Unit(LinkLabel_PreCal_1st, Panel_PreCal_1st, Nothing, Nothing, 0, "PreCal", 0, 0, 0)
        HeadRun = tempRun
        CurRun = HeadRun
        timeLeft = CurRun.Time
        timeLabel.Text = timeLeft & " s"

        Set_Panel(Panel_PreCal_2nd, LinkLabel_PreCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_2nd, Panel_PreCal_2nd, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_3rd, LinkLabel_PreCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_3rd, Panel_PreCal_3rd, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_4th, LinkLabel_PreCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_4th, Panel_PreCal_4th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_5th, LinkLabel_PreCal_5th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_5th, Panel_PreCal_5th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_6th, LinkLabel_PreCal_6th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_6th, Panel_PreCal_6th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        'Background
        Set_Panel(Panel_Bkg, LinkLabel_BG)
        tempRun.NextUnit = New Run_Unit(LinkLabel_BG, Panel_Bkg, Nothing, tempRun, 0, "Background", 0, 0, 0)
        tempRun = tempRun.NextUnit

        tempRun = Load_Loader_Helper(tempRun)

        'RSS
        Set_Panel(Panel_RSS, LinkLabel_RSS)
        tempRun.NextUnit = New Run_Unit(LinkLabel_RSS, Panel_RSS, Nothing, tempRun, 0, "RSS", 0, 0, 0)
        tempRun = tempRun.NextUnit
        'PostCal
        Set_Panel(Panel_PostCal_1st, LinkLabel_PostCal_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_1st, Panel_PostCal_1st, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_2nd, LinkLabel_PostCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_2nd, Panel_PostCal_2nd, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_3rd, LinkLabel_PostCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_3rd, Panel_PostCal_3rd, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_4th, LinkLabel_PostCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_4th, Panel_PostCal_4th, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_5th, LinkLabel_PostCal_5th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_5th, Panel_PostCal_5th, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_6th, LinkLabel_PostCal_6th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_6th, Panel_PostCal_6th, Nothing, tempRun, 0, "PostCal_Last", 0, 0, 0)
        tempRun = tempRun.NextUnit


        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, HeadRun.Time, 97)
        MainBarGraph = New BarGraph(New Point(508, 106), New Size(390, 539), TabPage2, CGraph.Modes.A1A2A3)

    End Sub

    Function Load_Loader_Helper(ByRef run As Run_Unit)
        'Create an object for each step
        Dim tempRun As Run_Unit = run
        ''A1
        'A1 First 1
        Set_Panel(Panel_LoA1_Fst_1st, LinkLabel_LoA1_Fst_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Fst_1st, Panel_LoA1_Fst_1st, Nothing, tempRun, 3, "LoA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 First 2
        Set_Panel(Panel_LoA1_Fst_2nd, LinkLabel_LoA1_Fst_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Fst_2nd, Panel_LoA1_Fst_2nd, Nothing, tempRun, 3, "LoA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 First 3
        Set_Panel(Panel_LoA1_Fst_3rd, LinkLabel_LoA1_Fst_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Fst_3rd, Panel_LoA1_Fst_3rd, Nothing, tempRun, 3, "LoA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Second 1
        Set_Panel(Panel_LoA1_Sec_1st, LinkLabel_LoA1_Sec_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Sec_1st, Panel_LoA1_Sec_1st, Nothing, tempRun, 3, "LoA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Second 2
        Set_Panel(Panel_LoA1_Sec_2nd, LinkLabel_LoA1_Sec_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Sec_2nd, Panel_LoA1_Sec_2nd, Nothing, tempRun, 3, "LoA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Second 3
        Set_Panel(Panel_LoA1_Sec_3rd, LinkLabel_LoA1_Sec_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Sec_3rd, Panel_LoA1_Sec_3rd, Nothing, tempRun, 3, "LoA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Third 1
        Set_Panel(Panel_LoA1_Thd_1st, LinkLabel_LoA1_Thd_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Thd_1st, Panel_LoA1_Thd_1st, Nothing, tempRun, 3, "LoA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Third 2
        Set_Panel(Panel_LoA1_Thd_2nd, LinkLabel_LoA1_Thd_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Thd_2nd, Panel_LoA1_Thd_2nd, Nothing, tempRun, 3, "LoA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Third 3
        Set_Panel(Panel_LoA1_Thd_3rd, LinkLabel_LoA1_Thd_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Thd_3rd, Panel_LoA1_Thd_3rd, Nothing, tempRun, 3, "LoA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Add 1
        Set_Panel(Panel_LoA1_Add_1st, LinkLabel_LoA1_Add_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Add_1st, Panel_LoA1_Add_1st, Nothing, tempRun, 3, "LoA1_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Add 2
        Set_Panel(Panel_LoA1_Add_2nd, LinkLabel_LoA1_Add_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Add_2nd, Panel_LoA1_Add_2nd, Nothing, tempRun, 3, "LoA1_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Add 3
        Set_Panel(Panel_LoA1_Add_3rd, LinkLabel_LoA1_Add_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA1_Add_3rd, Panel_LoA1_Add_3rd, Nothing, tempRun, 3, "LoA1_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)

        ''A2
        'A2 First 1
        Set_Panel(Panel_LoA2_Fst_1st, LinkLabel_LoA2_Fst_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Fst_1st, Panel_LoA2_Fst_1st, Nothing, tempRun, 1, "LoA2_1st", 1, 1, 3)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 First 2
        Set_Panel(Panel_LoA2_Fst_2nd, LinkLabel_LoA2_Fst_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Fst_2nd, Panel_LoA2_Fst_2nd, Nothing, tempRun, 1, "LoA2_2nd_3rd", 1, 1, 2)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 First 3
        Set_Panel(Panel_LoA2_Fst_3rd, LinkLabel_LoA2_Fst_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Fst_3rd, Panel_LoA2_Fst_3rd, Nothing, tempRun, 1, "LoA2_2nd_3rd", 1, 1, 2)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Second 1
        Set_Panel(Panel_LoA2_Sec_1st, LinkLabel_LoA2_Sec_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Sec_1st, Panel_LoA2_Sec_1st, Nothing, tempRun, 1, "LoA2_1st", 1, 1, 3)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Second 2
        Set_Panel(Panel_LoA2_Sec_2nd, LinkLabel_LoA2_Sec_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Sec_2nd, Panel_LoA2_Sec_2nd, Nothing, tempRun, 1, "LoA2_2nd_3rd", 1, 1, 2)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Second 3
        Set_Panel(Panel_LoA2_Sec_3rd, LinkLabel_LoA2_Sec_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Sec_3rd, Panel_LoA2_Sec_3rd, Nothing, tempRun, 1, "LoA2_2nd_3rd", 1, 1, 2)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Third 1
        Set_Panel(Panel_LoA2_Thd_1st, LinkLabel_LoA2_Thd_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Thd_1st, Panel_LoA2_Thd_1st, Nothing, tempRun, 1, "LoA2_1st", 1, 1, 3)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Third 2
        Set_Panel(Panel_LoA2_Thd_2nd, LinkLabel_LoA2_Thd_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Thd_2nd, Panel_LoA2_Thd_2nd, Nothing, tempRun, 1, "LoA2_2nd_3rd", 1, 1, 2)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Third 3
        Set_Panel(Panel_LoA2_Thd_3rd, LinkLabel_LoA2_Thd_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Thd_3rd, Panel_LoA2_Thd_3rd, Nothing, tempRun, 1, "LoA2_2nd_3rd", 1, 1, 2)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Add 1
        Set_Panel(Panel_LoA2_Add_1st, LinkLabel_LoA2_Add_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Add_1st, Panel_LoA2_Add_1st, Nothing, tempRun, 1, "LoA2_1st_Add", 1, 1, 3)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Add 2
        Set_Panel(Panel_LoA2_Add_2nd, LinkLabel_LoA2_Add_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Add_2nd, Panel_LoA2_Add_2nd, Nothing, tempRun, 1, "LoA2_2nd_3rd_Add", 1, 1, 2)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A2 Add 3
        Set_Panel(Panel_LoA2_Add_3rd, LinkLabel_LoA2_Add_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA2_Add_3rd, Panel_LoA2_Add_3rd, Nothing, tempRun, 1, "LoA2_2nd_3rd_Add", 1, 1, 2)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)

        ''A3
        'A3 First forward
        Set_Panel(Panel_LoA3_Fst_fwd, LinkLabel_LoA3_Fst_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Fst_fwd, Panel_LoA3_Fst_fwd, Nothing, tempRun, 1, "LoA3_fwd", 1, 1, 3)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A3 First backward
        Set_Panel(Panel_LoA3_Fst_bkd, LinkLabel_LoA3_Fst_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Fst_bkd, Panel_LoA3_Fst_bkd, Nothing, tempRun, 1, "LoA3_bkd", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)

        'A3 Second fwd
        Set_Panel(Panel_LoA3_Sec_fwd, LinkLabel_LoA3_Sec_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Sec_fwd, Panel_LoA3_Sec_fwd, Nothing, tempRun, 1, "LoA3_fwd", 1, 1, 3)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A3 Second backward
        Set_Panel(Panel_LoA3_Sec_bkd, LinkLabel_LoA3_Sec_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Sec_bkd, Panel_LoA3_Sec_bkd, Nothing, tempRun, 1, "LoA3_bkd", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)

        'A3 Third fwd
        Set_Panel(Panel_LoA3_Thd_fwd, LinkLabel_LoA3_Thd_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Thd_fwd, Panel_LoA3_Thd_fwd, Nothing, tempRun, 1, "LoA3_fwd", 1, 1, 3)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A3 Third bkd
        Set_Panel(Panel_LoA3_Thd_bkd, LinkLabel_LoA3_Thd_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Thd_bkd, Panel_LoA3_Thd_bkd, Nothing, tempRun, 1, "LoA3_bkd", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)

        'A3 Add fwd
        Set_Panel(Panel_LoA3_Add_fwd, LinkLabel_LoA3_Add_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Add_fwd, Panel_LoA3_Add_fwd, Nothing, tempRun, 1, "LoA3_fwd_Add", 1, 1, 3)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A3 Add bkd
        Set_Panel(Panel_LoA3_Add_bkd, LinkLabel_LoA3_Add_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_LoA3_Add_bkd, Panel_LoA3_Add_bkd, Nothing, tempRun, 1, "LoA3_bkd_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)

        Return tempRun

    End Function

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

        'new add
        Panel_PreCal_Sub.Visible = True
        Panel_PostCal_Sub.Visible = True
        Panel_PreCal_5th.Visible = True
        Panel_PreCal_6th.Visible = True
        Panel_PostCal_5th.Visible = True
        Panel_PostCal_6th.Visible = True

        PanelExcavatorA1.Size = ASize
        PanelExcavatorA1.Location = New Point(220, 265)
        PanelExcavatorA2.Size = ASize
        PanelExcavatorA2.Location = New Point(284, 265)
        PanelLoaderA1.Size = ASize
        PanelLoaderA1.Location = New Point(341, 265)
        PanelLoaderA2.Size = ASize
        PanelLoaderA2.Location = New Point(398, 265)
        PanelLoaderA3.Size = ASize
        PanelLoaderA3.Location = New Point(455, 265)

        PanelExcavatorA1.Controls.Add(Panel_ExA1_Fst_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Fst_2nd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Fst_3rd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Sec_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Sec_2nd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Sec_3rd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Thd_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Thd_2nd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Thd_3rd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Add_1st)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Add_2nd)
        PanelExcavatorA1.Controls.Add(Panel_ExA1_Add_3rd)

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
        PanelLoaderA1.Controls.Add(Panel_LoA1_Fst_2nd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Fst_3rd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Sec_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Sec_2nd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Sec_3rd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Thd_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Thd_2nd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Thd_3rd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Add_1st)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Add_2nd)
        PanelLoaderA1.Controls.Add(Panel_LoA1_Add_3rd)

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
        Set_Panel(Panel_PreCal_1st, LinkLabel_PreCal_1st)
        tempRun = New Run_Unit(LinkLabel_PreCal_1st, Panel_PreCal_1st, Nothing, Nothing, 0, "PreCal", 0, 0, 0)
        HeadRun = tempRun
        CurRun = HeadRun
        timeLeft = CurRun.Time
        timeLabel.Text = timeLeft & " s"

        Set_Panel(Panel_PreCal_2nd, LinkLabel_PreCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_2nd, Panel_PreCal_2nd, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_3rd, LinkLabel_PreCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_3rd, Panel_PreCal_3rd, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_4th, LinkLabel_PreCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_4th, Panel_PreCal_4th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_5th, LinkLabel_PreCal_5th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_5th, Panel_PreCal_5th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_6th, LinkLabel_PreCal_6th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_6th, Panel_PreCal_6th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        'Background
        Set_Panel(Panel_Bkg, LinkLabel_BG)
        tempRun.NextUnit = New Run_Unit(LinkLabel_BG, Panel_Bkg, Nothing, tempRun, 0, "Background", 0, 0, 0)
        tempRun = tempRun.NextUnit

        tempRun = Load_Excavator_Helper(tempRun)
        tempRun = Load_Loader_Helper(tempRun)

        'RSS
        Set_Panel(Panel_RSS, LinkLabel_RSS)
        tempRun.NextUnit = New Run_Unit(LinkLabel_RSS, Panel_RSS, Nothing, tempRun, 0, "RSS", 0, 0, 0)
        tempRun = tempRun.NextUnit
        'PostCal
        Set_Panel(Panel_PostCal_1st, LinkLabel_PostCal_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_1st, Panel_PostCal_1st, Nothing, tempRun, 0, "PostCal", 0, 0, 1)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_2nd, LinkLabel_PostCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_2nd, Panel_PostCal_2nd, Nothing, tempRun, 0, "PostCal", 0, 0, 1)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_3rd, LinkLabel_PostCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_3rd, Panel_PostCal_3rd, Nothing, tempRun, 0, "PostCal", 0, 0, 1)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_4th, LinkLabel_PostCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_4th, Panel_PostCal_4th, Nothing, tempRun, 0, "PostCal", 0, 0, 1)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_5th, LinkLabel_PostCal_5th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_5th, Panel_PostCal_5th, Nothing, tempRun, 0, "PostCal", 0, 0, 1)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_6th, LinkLabel_PostCal_6th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_6th, Panel_PostCal_6th, Nothing, tempRun, 0, "PostCal_Last", 0, 0, 0)
        tempRun = tempRun.NextUnit

        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, HeadRun.Time, 97)
        MainBarGraph = New BarGraph(New Point(508, 106), New Size(390, 539), TabPage2, CGraph.Modes.A1A2A3)

    End Sub

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
        PanelTractorA1.Location = New Point(220, 265)
        PanelTractorA3.Size = ASize
        PanelTractorA3.Location = New Point(284, 265)

        'new add
        Panel_PreCal_Sub.Visible = True
        Panel_PostCal_Sub.Visible = True
        Panel_PreCal_5th.Visible = True
        Panel_PreCal_6th.Visible = True
        Panel_PostCal_5th.Visible = True
        Panel_PostCal_6th.Visible = True

        PanelTractorA1.Controls.Add(Panel_TrA1_Fst_1st)
        PanelTractorA1.Controls.Add(Panel_TrA1_Fst_2nd)
        PanelTractorA1.Controls.Add(Panel_TrA1_Fst_3rd)
        PanelTractorA1.Controls.Add(Panel_TrA1_Sec_1st)
        PanelTractorA1.Controls.Add(Panel_TrA1_Sec_2nd)
        PanelTractorA1.Controls.Add(Panel_TrA1_Sec_3rd)
        PanelTractorA1.Controls.Add(Panel_TrA1_Thd_1st)
        PanelTractorA1.Controls.Add(Panel_TrA1_Thd_2nd)
        PanelTractorA1.Controls.Add(Panel_TrA1_Thd_3rd)
        PanelTractorA1.Controls.Add(Panel_TrA1_Add_1st)
        PanelTractorA1.Controls.Add(Panel_TrA1_Add_2nd)
        PanelTractorA1.Controls.Add(Panel_TrA1_Add_3rd)

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
        Set_Panel(Panel_PreCal_1st, LinkLabel_PreCal_1st)
        tempRun = New Run_Unit(LinkLabel_PreCal_1st, Panel_PreCal_1st, Nothing, Nothing, 0, "PreCal", 0, 0, 0)
        HeadRun = tempRun
        CurRun = HeadRun
        timeLeft = CurRun.Time
        timeLabel.Text = timeLeft & " s"

        Set_Panel(Panel_PreCal_2nd, LinkLabel_PreCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_2nd, Panel_PreCal_2nd, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_3rd, LinkLabel_PreCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_3rd, Panel_PreCal_3rd, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_4th, LinkLabel_PreCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_4th, Panel_PreCal_4th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_5th, LinkLabel_PreCal_5th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_5th, Panel_PreCal_5th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_6th, LinkLabel_PreCal_6th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_6th, Panel_PreCal_6th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        'Background
        Set_Panel(Panel_Bkg, LinkLabel_BG)
        tempRun.NextUnit = New Run_Unit(LinkLabel_BG, Panel_Bkg, Nothing, tempRun, 0, "Background", 0, 0, 0)
        tempRun = tempRun.NextUnit

        tempRun = Load_Tractor_Helper(tempRun)

        'RSS
        Set_Panel(Panel_RSS, LinkLabel_RSS)
        tempRun.NextUnit = New Run_Unit(LinkLabel_RSS, Panel_RSS, Nothing, tempRun, 0, "RSS", 0, 0, 0)
        tempRun = tempRun.NextUnit
        'PostCal
        Set_Panel(Panel_PostCal_1st, LinkLabel_PostCal_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_1st, Panel_PostCal_1st, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_2nd, LinkLabel_PostCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_2nd, Panel_PostCal_2nd, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_3rd, LinkLabel_PostCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_3rd, Panel_PostCal_3rd, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_4th, LinkLabel_PostCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_4th, Panel_PostCal_4th, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_5th, LinkLabel_PostCal_5th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_5th, Panel_PostCal_5th, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_6th, LinkLabel_PostCal_6th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_6th, Panel_PostCal_6th, Nothing, tempRun, 0, "PostCal_Last", 0, 0, 0)
        tempRun = tempRun.NextUnit

        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, HeadRun.Time, 97)
        MainBarGraph = New BarGraph(New Point(508, 106), New Size(390, 539), TabPage2, CGraph.Modes.A1A2A3)

    End Sub

    Function Load_Tractor_Helper(ByRef run As Run_Unit)
        'Create an object for each step
        Dim tempRun As Run_Unit = run
        ''A1
        'A1 First 1
        Set_Panel(Panel_TrA1_Fst_1st, LinkLabel_TrA1_Fst_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Fst_1st, Panel_TrA1_Fst_1st, Nothing, tempRun, 3, "TrA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 First 2
        Set_Panel(Panel_TrA1_Fst_2nd, LinkLabel_TrA1_Fst_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Fst_2nd, Panel_TrA1_Fst_2nd, Nothing, tempRun, 3, "TrA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 First 3
        Set_Panel(Panel_TrA1_Fst_3rd, LinkLabel_TrA1_Fst_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Fst_3rd, Panel_TrA1_Fst_3rd, Nothing, tempRun, 3, "TrA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Second 1
        Set_Panel(Panel_TrA1_Sec_1st, LinkLabel_TrA1_Sec_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Sec_1st, Panel_TrA1_Sec_1st, Nothing, tempRun, 3, "TrA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Second 2
        Set_Panel(Panel_TrA1_Sec_2nd, LinkLabel_TrA1_Sec_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Sec_2nd, Panel_TrA1_Sec_2nd, Nothing, tempRun, 3, "TrA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Second 3
        Set_Panel(Panel_TrA1_Sec_3rd, LinkLabel_TrA1_Sec_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Sec_3rd, Panel_TrA1_Sec_3rd, Nothing, tempRun, 3, "TrA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Third 1
        Set_Panel(Panel_TrA1_Thd_1st, LinkLabel_TrA1_Thd_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Thd_1st, Panel_TrA1_Thd_1st, Nothing, tempRun, 3, "TrA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Third 2
        Set_Panel(Panel_TrA1_Thd_2nd, LinkLabel_TrA1_Thd_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Thd_2nd, Panel_TrA1_Thd_2nd, Nothing, tempRun, 3, "TrA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Third 3
        Set_Panel(Panel_TrA1_Thd_3rd, LinkLabel_TrA1_Thd_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Thd_3rd, Panel_TrA1_Thd_3rd, Nothing, tempRun, 3, "TrA1", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Add 1
        Set_Panel(Panel_TrA1_Add_1st, LinkLabel_TrA1_Add_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Add_1st, Panel_TrA1_Add_1st, Nothing, tempRun, 3, "TrA1_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Add 2
        Set_Panel(Panel_TrA1_Add_2nd, LinkLabel_TrA1_Add_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Add_2nd, Panel_TrA1_Add_2nd, Nothing, tempRun, 3, "TrA1_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A1 Add 3
        Set_Panel(Panel_TrA1_Add_3rd, LinkLabel_TrA1_Add_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA1_Add_3rd, Panel_TrA1_Add_3rd, Nothing, tempRun, 3, "TrA1_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)


        ''A3
        'A3 First forward
        Set_Panel(Panel_TrA3_Fst_fwd, LinkLabel_TrA3_Fst_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Fst_fwd, Panel_TrA3_Fst_fwd, Nothing, tempRun, 3, "TrA3_fwd", 1, 1, 3)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A3 First backward
        Set_Panel(Panel_TrA3_Fst_bkd, LinkLabel_TrA3_Fst_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Fst_bkd, Panel_TrA3_Fst_bkd, Nothing, tempRun, 3, "TrA3_bkd", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)

        'A3 Second fwd
        Set_Panel(Panel_TrA3_Sec_fwd, LinkLabel_TrA3_Sec_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Sec_fwd, Panel_TrA3_Sec_fwd, Nothing, tempRun, 3, "TrA3_fwd", 1, 1, 3)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A3 Second backward
        Set_Panel(Panel_TrA3_Sec_bkd, LinkLabel_TrA3_Sec_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Sec_bkd, Panel_TrA3_Sec_bkd, Nothing, tempRun, 3, "TrA3_bkd", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)

        'A3 Third fwd
        Set_Panel(Panel_TrA3_Thd_fwd, LinkLabel_TrA3_Thd_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Thd_fwd, Panel_TrA3_Thd_fwd, Nothing, tempRun, 3, "TrA3_fwd", 1, 1, 3)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A3 Third bkd
        Set_Panel(Panel_TrA3_Thd_bkd, LinkLabel_TrA3_Thd_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Thd_bkd, Panel_TrA3_Thd_bkd, Nothing, tempRun, 3, "TrA3_bkd", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)

        'A3 Add fwd
        Set_Panel(Panel_TrA3_Add_fwd, LinkLabel_TrA3_Add_fwd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Add_fwd, Panel_TrA3_Add_fwd, Nothing, tempRun, 3, "TrA3_fwd_Add", 1, 1, 3)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A3 Add bkd
        Set_Panel(Panel_TrA3_Add_bkd, LinkLabel_TrA3_Add_bkd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_TrA3_Add_bkd, Panel_TrA3_Add_bkd, Nothing, tempRun, 3, "TrA3_bkd_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)

        Return tempRun

    End Function

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
        PanelA4.Location = New Point(220, 265)
        PanelA4.Controls.Add(Panel_A4_Fst)
        PanelA4.Controls.Add(Panel_A4_Sec)
        PanelA4.Controls.Add(Panel_A4_Thd)
        PanelA4.Controls.Add(Panel_A4_Add)

        'new
        Panel_PreCal_Sub.Visible = True
        Panel_PostCal_Sub.Visible = True
        Panel_PreCal_5th.Visible = False
        Panel_PreCal_6th.Visible = False
        Panel_PostCal_5th.Visible = False
        Panel_PostCal_6th.Visible = False

        'Create an object for each step
        Dim tempRun As Run_Unit
        'Precal
        Set_Panel(Panel_PreCal_1st, LinkLabel_PreCal_1st)
        tempRun = New Run_Unit(LinkLabel_PreCal_1st, Panel_PreCal_1st, Nothing, Nothing, 0, "PreCal", 0, 0, 0)
        HeadRun = tempRun
        CurRun = HeadRun
        timeLeft = CurRun.Time
        timeLabel.Text = timeLeft & " s"

        Set_Panel(Panel_PreCal_2nd, LinkLabel_PreCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_2nd, Panel_PreCal_2nd, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_3rd, LinkLabel_PreCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_3rd, Panel_PreCal_3rd, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PreCal_4th, LinkLabel_PreCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PreCal_4th, Panel_PreCal_4th, Nothing, tempRun, 0, "PreCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        'Background
        Set_Panel(Panel_Bkg, LinkLabel_BG)
        tempRun.NextUnit = New Run_Unit(LinkLabel_BG, Panel_Bkg, Nothing, tempRun, 0, "Background", 0, 0, 0)
        tempRun = tempRun.NextUnit
        'A4 1
        Set_Panel(Panel_A4_Fst, LinkLabel_A4_First)
        tempRun.NextUnit = New Run_Unit(LinkLabel_A4_First, Panel_A4_Fst, Nothing, tempRun, 0, "A4", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A4 2
        Set_Panel(Panel_A4_Sec, LinkLabel_A4_Second)
        tempRun.NextUnit = New Run_Unit(LinkLabel_A4_Second, Panel_A4_Sec, Nothing, tempRun, 0, "A4", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A4 3
        Set_Panel(Panel_A4_Thd, LinkLabel_A4_Third)
        tempRun.NextUnit = New Run_Unit(LinkLabel_A4_Second, Panel_A4_Sec, Nothing, tempRun, 0, "A4", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'A4 add
        Set_Panel(Panel_A4_Add, LinkLabel_A4_Additional)
        tempRun.NextUnit = New Run_Unit(LinkLabel_A4_Second, Panel_A4_Sec, Nothing, tempRun, 0, "A4_Add", 1, 1, 1)
        tempRun = tempRun.NextUnit
        tempRun.Steps = Load_Steps_helper(tempRun)
        'RSS
        Set_Panel(Panel_RSS, LinkLabel_RSS)
        tempRun.NextUnit = New Run_Unit(LinkLabel_RSS, Panel_RSS, Nothing, tempRun, 0, "RSS", 0, 0, 0)
        tempRun = tempRun.NextUnit
        'PostCal
        Set_Panel(Panel_PostCal_1st, LinkLabel_PostCal_1st)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_1st, Panel_PostCal_1st, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_2nd, LinkLabel_PostCal_2nd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_2nd, Panel_PostCal_2nd, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_3rd, LinkLabel_PostCal_3rd)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_3rd, Panel_PostCal_3rd, Nothing, tempRun, 0, "PostCal", 0, 0, 0)
        tempRun = tempRun.NextUnit
        Set_Panel(Panel_PostCal_4th, LinkLabel_PostCal_4th)
        tempRun.NextUnit = New Run_Unit(LinkLabel_PostCal_4th, Panel_PostCal_4th, Nothing, tempRun, 0, "PostCal_Last", 0, 0, 0)
        tempRun = tempRun.NextUnit

        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, HeadRun.Time, 97)
        MainBarGraph = New BarGraph(New Point(508, 106), New Size(390, 539), TabPage2, CGraph.Modes.A4)

        'display them on screen-set position and size
        Dim text As String
        Step1.BackColor = Color.MistyRose
        If name = "空氣壓縮機(Compressor)" Then
            text = My.Resources.A4_Compressor
        ElseIf name = "混凝土割切機(Concrete cutter)" Then
            text = My.Resources.A4_Concrete_Cutter
        ElseIf name = "瀝青混凝土舖築機(Asphalt finisher)" Then
            text = My.Resources.A4_Asphalt_Finisher
        ElseIf name = "混凝土破碎機(Concrete breaker)" Then
            text = My.Resources.A4_Concrete_Breaker
        ElseIf name = "混凝土泵車(Concrete pump)" Then
            text = My.Resources.A4_Concrete_Pump
        ElseIf name = "鑽土機(Earth drill)" Then
            text = My.Resources.A4_Auger_Drill_Driver
        ElseIf name = "全套管鑽掘機" Then
            text = My.Resources.A4_Auger_Drill_Driver
        ElseIf name = "土壤取樣器(地鑽) (Earth auger)" Then
            text = My.Resources.A4_Auger_Drill_Driver
        ElseIf name = "油壓式拔樁機" Then
            text = My.Resources.A4_Auger_Drill_Driver
        ElseIf name = "油壓式打樁機(Hydraulic pile driver)" Then
            text = My.Resources.A4_Auger_Drill_Driver
        ElseIf name = "振動式樁錘(Vibrating hammer)" Then
            text = My.Resources.A4_Vibrating_Hammer
        ElseIf name = "輪形起重機(Wheel crane)" Then
            text = My.Resources.A4_Crane
        ElseIf name = "卡車起重機(Truck crane)" Then
            text = My.Resources.A4_Crane
        ElseIf name = "履帶起重機(Crawler crane)" Then
            text = My.Resources.A4_Crane
        ElseIf name = "振動式壓路機(Vibrating roller)" Then
            text = My.Resources.A4_Roller
        ElseIf name = "膠輪壓路機(Wheel roller)" Then
            text = My.Resources.A4_Roller
        End If
        tempRun = HeadRun
        While Not IsNothing(tempRun)
            'tempRun.Steps = text
            tempRun = tempRun.NextUnit
        End While
    End Sub

    Function Load_Steps_helper(ByRef run As Run_Unit)
        Dim tempRun As Run_Unit = run
        Dim tempStep As Steps
        If tempRun.Name = "ExA1" Or tempRun.Name = "LoA1" Or tempRun.Name = "TrA1" Or tempRun.Name = "ExA1_Add" Or tempRun.Name = "LoA1_Add" Or tempRun.Name = "TrA1_Add" Then
            tempRun.Steps = New Steps(My.Resources.A1_step1, Step1, Nothing, True, 1)
            tempRun.HeadStep = tempRun.Steps
        End If
        If tempRun.Name = "ExA2_1st" Or tempRun.Name = "ExA2_1st_Add" Then
            tempRun.Steps = New Steps(My.Resources.A2_Excavator_step1, Step1, Nothing, False, 1)
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A2_Excavator_step2, Step2, Nothing, False, 1)
            tempStep = tempRun.Steps.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step3, Step3, Nothing, False, 1)
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step4, Step4, Nothing, False, 1)
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step5, Step5, Nothing, False, 1)
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step6, Step6, Nothing, False, 1)
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step7, Step7, Nothing, False, 1)
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step8, Step8, Nothing, False, 1)
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step9, Step9, Nothing, True, 1)
        End If
        If tempRun.Name = "ExA2_2nd_3rd" Or tempRun.Name = "ExA2_2nd_3rd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A2_Excavator_step2, Step1, Nothing, False, 1)
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A2_Excavator_step3, Step2, Nothing, False, 1)
            tempStep = tempRun.Steps.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step4, Step3, Nothing, False, 1)
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step5, Step4, Nothing, False, 1)
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step6, Step5, Nothing, False, 1)
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step7, Step6, Nothing, False, 1)
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step8, Step7, Nothing, False, 1)
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step9, Step8, Nothing, True, 1)
        End If
        If tempRun.Name = "LoA2_1st" Or tempRun.Name = "LoA2_1st_Add" Then
            tempRun.Steps = New Steps(My.Resources.A2_Loader_step1, Step1, Nothing, False, 2)
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A2_Loader_step2, Step2, Nothing, False, 3)
            tempStep = tempRun.Steps.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Loader_step3, Step3, Nothing, True, 4)
        End If
        If tempRun.Name = "LoA2_2nd_3rd" Or tempRun.Name = "LoA2_2nd_3rd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A2_Loader_step2, Step1, Nothing, False, 3)
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A2_Loader_step3, Step2, Nothing, True, 4)
        End If
        If tempRun.Name = "LoA3_fwd" Or tempRun.Name = "LoA3_fwd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A3_Loader_step1, Step1, Nothing, False, 2)
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A3_Loader_step2, Step2, Nothing, False, 4)
            tempStep = tempRun.Steps.NextStep
            tempStep.NextStep = New Steps(My.Resources.A3_Loader_step3, Step3, Nothing, True, 5)
        End If
        If tempRun.Name = "LoA3_bkd" Or tempRun.Name = "LoA3_bkd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A3_Loader_step4, Step1, Nothing, True, 5)
            tempRun.HeadStep = tempRun.Steps
        End If
        If tempRun.Name = "TrA3_fwd" Or tempRun.Name = "TrA3_fwd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A3_Tractor_step1, Step1, Nothing, False, 2)
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A3_Tractor_step2, Step2, Nothing, False, 4)
            tempStep = tempRun.Steps.NextStep
            tempStep.NextStep = New Steps(My.Resources.A3_Tractor_step3, Step3, Nothing, True, 5)
        End If
        If tempRun.Name = "TrA3_bkd" Or tempRun.Name = "TrA3_bkd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A3_Tractor_step4, Step1, Nothing, True, 4)
            tempRun.HeadStep = tempRun.Steps
        End If
        If tempRun.Name = "A4" Or tempRun.Name = "A4_Add" Then
            tempRun.Steps = New Steps(My.Resources.A4_Asphalt_Finisher, Step1, Nothing, True, 3)
            tempRun.HeadStep = tempRun.Steps
        End If

        Return tempRun.Steps

    End Function

    Private MouseIsDown As Boolean = False
    Private lastPos As Point

    Private Sub pLabel_MouseDown(ByVal sender As Object, ByVal e As  _
    System.Windows.Forms.MouseEventArgs) Handles p1Label.MouseDown, p2Label.MouseDown, p3Label.MouseDown, p4Label.MouseDown, p5Label.MouseDown, p6Label.MouseDown, p7Label.MouseDown, p8Label.MouseDown, p9Label.MouseDown, p10Label.MouseDown
        ' Set a flag to show that the mouse is down.
        MouseIsDown = True
        lastPos = Cursor.Position
        For Each l As CoorPoint In Points
            If l.Label.Name = sender.Name Then
                If Not IsNothing(l.Line) Then
                    l.Line.Dispose()
                End If
            End If
        Next
    End Sub

    Private Sub pLabel_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p9Label.MouseUp, p8Label.MouseUp, p7Label.MouseUp, p10Label.MouseUp, p6Label.MouseUp, p5Label.MouseUp, p4Label.MouseUp, p3Label.MouseUp, p2Label.MouseUp, p1Label.MouseUp
        MouseIsDown = False
        For Each l As CoorPoint In Points
            If l.Label.Name = sender.Name Then
                l.Line = New LineShape(canvas)
                l.Line.StartPoint = New System.Drawing.Point((origin(0) + ratio * l.Coors.Xc), (origin(1) - ratio * l.Coors.Yc))
                l.Line.EndPoint = translate(New Point(l.Label.Location.X + l.Label.Size.Width / 2, l.Label.Location.Y - l.Label.Size.Height), TabPage1.Location, GroupBox_Plot.Location + canvas.Location)
            End If
        Next
    End Sub
    Private Sub pLabel_MouseMove(ByVal sender As Label, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p9Label.MouseMove, p8Label.MouseMove, p7Label.MouseMove, p10Label.MouseMove, p6Label.MouseMove, p5Label.MouseMove, p4Label.MouseMove, p3Label.MouseMove, p2Label.MouseMove, p1Label.MouseMove
        If MouseIsDown Then
            ' Initiate dragging.
            Dim temp As Point = Cursor.Position
            sender.Location = temp - lastPos + sender.Location
            lastPos = temp
        End If
    End Sub

    Private Function GetInstantData()
        Dim num As Integer
        If Not Machine = Machines.Others Then
            num = 6
        Else
            num = 4
        End If
        Dim r = New Random()
        Dim result(num - 1) As Double
        For i = 0 To num - 1
            'result(i) = (r.Next(1, 100) / 10) + 90
            result(i) = (r.Next(1, 10) / 10) + r.Next(1, 119)
        Next
        Return result
    End Function

    Private Sub SetScreenValuesAndGraphs(ByVal vals() As Double)
        Dim num As Integer
        Dim sum As Integer = 0
        If Not Machine = Machines.Others Then
            num = 6
        Else
            num = 4
        End If
        'Set Texts
        For i = 0 To num - 1
            NoisesArray(i).Text = vals(i)
            sum += vals(i)
        Next
        NoisesArray(num).Text = Int((sum / num) * 100 + 0.5) / 100
        'Set graphs
        MainBarGraph.Update(vals)
        MainLineGraphs.Update(vals)
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Countdown = False Then
            timeLeft = timeLeft + 1
            timeLabel.Text = timeLeft & " s"

            'send values to display as text and graphs
            SetScreenValuesAndGraphs(GetInstantData())

        Else
            If timeLeft > 0 Then 'counting
                timeLeft = timeLeft - 1
                timeLabel.Text = timeLeft & " s"

                'send values to display as text and graphs
                SetScreenValuesAndGraphs(GetInstantData())
            ElseIf timeLeft = 0 Then 'time's up\
                'stop the timer
                Timer1.Stop()
                startButton.Enabled = True
                stopButton.Enabled = False

                'if HasNextStep
                'change step color , seconds 
                'do next step
                If CurRun.CurStep = CurRun.EndStep Then 'last step (not HasNextStep)

                    Accept_Button.Enabled = True
                    startButton.Enabled = False

                Else 'HasNextStep

                    Set_Step_BackColor()
                    CurRun.Steps = CurRun.Steps.NextStep 'jump to next step
                    CurRun.CurStep += 1 'curstep add 1
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                End If

                'if Not HasNextStep
                'no startbutton but acceptbutton
                'change to next Run_Unit (in AcceptButton_clicked)

            End If
        End If
        'If timeLeft > 0 Then 'counting
        '    timeLeft = timeLeft - 1
        '    timeLabel.Text = timeLeft & " s"

        '    'send values to display as text and graphs
        '    SetScreenValuesAndGraphs(GetInstantData())

        'ElseIf timeLeft = 0 Then 'time's up
        '    'SetScreenValuesAndGraphs(GetInstantData())
        '    Timer1.Stop()
        '    'System.Threading.Thread.Sleep(1000)
        '    startButton.Enabled = True
        '    stopButton.Enabled = False
        '    'determine what the next step is
        '    If CurRun.Name = "PreCal" Then
        '        CurRun = CurRun.NextUnit
        '        Accept_Button.Enabled = True
        '        startButton.Enabled = False
        '    Else
        '        If CurRun.CurStep = 0 Then
        '            If CurRun.Name = "Background" Then
        '                CurRun = CurRun.NextUnit
        '                CurStep = CurRun.Steps
        '                Accept_Button.Enabled = True
        '                startButton.Enabled = False
        '            ElseIf CurRun.Name = "PostCal" Then
        '                If CurRun.CurStep = CurRun.EndStep Then
        '                    MessageBox.Show("測試結束")
        '                    CurRun.Set_BackColor(Color.Green)
        '                    Accept_Button.Enabled = True
        '                    startButton.Enabled = False
        '                Else
        '                    CurRun = CurRun.NextUnit
        '                    CurStep = CurRun.Steps
        '                    Accept_Button.Enabled = True
        '                    startButton.Enabled = False
        '                End If
        '            Else
        '                'RSS
        '                CurRun = CurRun.NextUnit
        '                Accept_Button.Enabled = True
        '                startButton.Enabled = False
        '            End If
        '        Else
        '            If CurRun.CurStep = CurRun.EndStep And CurRun.NextUnit.EndStep > 0 Then
        '                CurRun = CurRun.NextUnit
        '                CurStep = CurRun.Steps
        '                Accept_Button.Enabled = True
        '                startButton.Enabled = False
        '            ElseIf CurRun.CurStep = CurRun.EndStep And CurRun.NextUnit.CurStep = 0 Then
        '                'A2 or A3 or A4 to RSS
        '                CurRun = CurRun.NextUnit
        '                CurStep = CurRun.Steps
        '                timeLeft = CurRun.Time
        '                timeLabel.Text = timeLeft & " s"
        '            Else
        '                Set_Step_BackColor()
        '                CurRun.Steps = CurRun.Steps.NextStep
        '                CurRun.CurStep = CurRun.CurStep + 1
        '                timeLeft = CurRun.Steps.Time
        '                timeLabel.Text = timeLeft & " s"
        '            End If
        '        End If
        '    End If

        'End If

    End Sub
    Sub Set_Panel_BackColor()
        CurRun.Set_BackColor(Color.Green)
        CurRun.NextUnit.Set_BackColor(Color.Yellow)
    End Sub
    Sub Set_Step_BackColor()
        array_step(CurRun.CurStep - 1).BackColor = Color.Green
        array_step(CurRun.CurStep).BackColor = Color.Yellow
    End Sub
    ' click event on Start Button
    Private Sub startButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles startButton.Click
        If Countdown = False Then
            startButton.Enabled = False
            Accept_Button.Enabled = False
            stopButton.Enabled = True
            Timer1.Start()
        Else
            startButton.Enabled = False
            stopButton.Enabled = True
            Timer1.Start()
        End If
        
    End Sub
    Private Sub stopButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stopButton.Click
        If Countdown = False Then
            startButton.Enabled = True
            Accept_Button.Enabled = True
            stopButton.Enabled = False
            Timer1.Stop()
        Else
            startButton.Enabled = True
            stopButton.Enabled = False
            Timer1.Stop()
        End If
    End Sub
    Private Sub AcceptButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Accept_Button.Click

        If Countdown = False Then
            If CurRun.Name = "PreCal" Then
                'Case1: now is PreCal
                ' change light
                Set_Panel_BackColor()

                'dispose old graph and create new graph
                MainLineGraphs.Dispose()
                If Machine = Machines.Others Then
                    MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)  'A4 mode
                Else
                    MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, CurRun.Time, 97) 'A1 A2 A3 mode
                End If

                'save info here

                'jump to next Run_Unit and set second to zero(should be zero here)
                CurRun = CurRun.NextUnit
                timeLeft = CurRun.Time
                timeLabel.Text = timeLeft & " s"


            ElseIf CurRun.Name = "Background" Then
                'Case2: now is Background
                ' change light
                Set_Panel_BackColor()

                'jump to next Run_Unit and change countdown False to True(for A1 A2 A3 A4 test)
                CurRun = CurRun.NextUnit
                CurStep = CurRun.HeadStep
                Countdown = True
                timeLeft = CurRun.Steps.Time
                timeLabel.Text = timeLeft & " s"

                'load A1's steps
                For index = 0 To 9
                    array_step(index).Text = ""
                    array_step(index).BackColor = Color.DarkGray
                Next
                For index = CurRun.StartStep To CurRun.EndStep
                    sum_steps += CurStep.Time
                    array_step(index - 1).Text = CurStep.step_str
                    array_step(index - 1).BackColor = Color.White
                    If CurStep.HasNext() = True Then
                        CurStep = CurStep.NextStep
                    End If
                Next
                array_step(0).BackColor = Color.Yellow

                'dispose old graph and create new graph
                MainLineGraphs.Dispose()
                If Machine = Machines.Others Then
                    MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                Else
                    MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                End If
                sum_steps = 0

            ElseIf CurRun.Name = "RSS" Then
                'Case3: now is RSS
                ' change light
                Set_Panel_BackColor()

                'dispose old graph and create new graph
                MainLineGraphs.Dispose()
                If Machine = Machines.Others Then
                    MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)  'A4 mode
                Else
                    MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, CurRun.Time, 97) 'A1 A2 A3 mode
                End If

                'save info here

                'jump to next Run_Unit and set second to zero(should be zero here)
                CurRun = CurRun.NextUnit
                timeLeft = CurRun.Time
                timeLabel.Text = timeLeft & " s"
            ElseIf CurRun.Name = "PostCal" Then
                'Case4: now is PostCal
                ' change light
                Set_Panel_BackColor()

                'dispose old graph and create new graph
                MainLineGraphs.Dispose()
                If Machine = Machines.Others Then
                    MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)  'A4 mode
                Else
                    MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, CurRun.Time, 97) 'A1 A2 A3 mode
                End If

                'save info here

                'jump to next Run_Unit and set second to zero(should be zero here)
                CurRun = CurRun.NextUnit
                timeLeft = CurRun.Time
                timeLabel.Text = timeLeft & " s"
            ElseIf CurRun.Name = "PostCal_Last" Then
                CurRun.Set_BackColor(Color.Green)
                MessageBox.Show("End")
            End If
        Else 'countdown = true
            If CurRun.Name = "ExA1" Or CurRun.Name = "TrA1" Then
                If CurRun.NextUnit.Name = "ExA1" Or CurRun.NextUnit.Name = "TrA1" Then
                    'Case1: now is ExA1 and next is also ExA1 or now is TrA1 and next is also TrA1
                    ' change light
                    Set_Panel_BackColor()

                    'jump to next Run_Unit
                    CurRun = CurRun.NextUnit
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load A1's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0

                ElseIf CurRun.NextUnit.Name = "ExA1_Add" Or CurRun.NextUnit.Name = "TrA1_Add" Then
                    'Case2: now is ExA1 and next is ExA1_Add or now is TrA1 and next is TrA1_Add
                    'have an additional test?
                    'call a function 
                    'if want_add()=true then ... elseif want_add()=false then ... endif

                    CurRun.Set_BackColor(Color.Green)
                    'jump to next Run_Unit and change light
                    RandBool = RandGen.Next(0, 2).ToString
                    If RandBool = True Then
                        'True: add test
                        CurRun = CurRun.NextUnit
                        CurRun.Set_BackColor(Color.Yellow)
                    Else
                        'False: not add test , jump to ExA2_1st or jump to TrA3_fwd
                        CurRun = CurRun.NextUnit.NextUnit.NextUnit.NextUnit
                        CurRun.Set_BackColor(Color.Yellow)
                    End If
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load A1's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                Else
                    MessageBox.Show("Error")
                End If

            ElseIf CurRun.Name = "ExA1_Add" Or CurRun.Name = "TrA1_Add" Then

                If CurRun.NextUnit.Name = "ExA1_Add" Or CurRun.NextUnit.Name = "TrA1_Add" Then
                    'Case1: now is ExA1_Add and next is also ExA1_Add or now is TrA1_Add and next is also TrA1_Add
                    ' change light
                    Set_Panel_BackColor()

                    'jump to next Run_Unit
                    CurRun = CurRun.NextUnit
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load A1's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                ElseIf CurRun.NextUnit.Name = "ExA2_1st" Or CurRun.NextUnit.Name = "TrA3_fwd" Then
                    'Case2: now is ExA1_Add and next is ExA2_1st or now is TrA1_Add and next is ExA3_fwd
                    'have an additional test?
                    'call a function 
                    'if want_add()=true then ... elseif want_add()=false then ... endif


                    'jump to next Run_Unit
                    RandBool = RandGen.Next(0, 2).ToString
                    If RandBool = True Then
                        'True: add test
                        CurRun.PrevUnit.PrevUnit.Set_BackColor(Color.Yellow)
                        CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
                        CurRun.Set_BackColor(Color.IndianRed)
                        CurRun = CurRun.PrevUnit.PrevUnit
                        CurRun.CurStep = 1
                        CurRun.NextUnit.CurStep = 1
                        CurRun.NextUnit.NextUnit.CurStep = 1
                        Load_Steps_helper(CurRun)
                        Load_Steps_helper(CurRun.NextUnit)
                        Load_Steps_helper(CurRun.NextUnit.NextUnit)
                    Else
                        'False: not add test
                        CurRun.Set_BackColor(Color.Green)
                        CurRun.NextUnit.Set_BackColor(Color.Yellow)
                        CurRun = CurRun.NextUnit
                    End If
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load A2's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                Else
                    MessageBox.Show("Error")
                End If

            ElseIf CurRun.Name = "ExA2_1st" Or CurRun.Name = "ExA2_2nd_3rd" Then
                If CurRun.NextUnit.Name = "ExA2_1st" Or CurRun.NextUnit.Name = "ExA2_2nd_3rd" Then
                    'Case: ExA2_1st to ExA2_2nd_3rd or ExA2_2nd_3rd to ExA2_1st
                    ' change light
                    Set_Panel_BackColor()

                    'jump to next Run_Unit
                    CurRun = CurRun.NextUnit
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load A2's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                ElseIf CurRun.NextUnit.Name = "ExA2_1st_Add" Then
                    'Case: ExA2_2nd_3rd to ExA2_1st_Add
                    'have an additional test?
                    'call a function 
                    'if want_add()=true then ... elseif want_add()=false then ... endif

                    CurRun.Set_BackColor(Color.Green)
                    'jump to next Run_Unit and change light
                    RandBool = RandGen.Next(0, 2).ToString
                    If RandBool = True Then
                        'True: add test
                        CurRun = CurRun.NextUnit
                        CurRun.Set_BackColor(Color.Yellow)
                        CurStep = CurRun.HeadStep
                        timeLeft = CurRun.Steps.Time
                        timeLabel.Text = timeLeft & " s"
                        'load A2's steps
                        For index = 0 To 9
                            array_step(index).Text = ""
                            array_step(index).BackColor = Color.DarkGray
                        Next
                        For index = CurRun.StartStep To CurRun.EndStep
                            sum_steps += CurStep.Time
                            array_step(index - 1).Text = CurStep.step_str
                            array_step(index - 1).BackColor = Color.White
                            If CurStep.HasNext() = True Then
                                CurStep = CurStep.NextStep
                            End If
                        Next
                        array_step(0).BackColor = Color.Yellow
                        'dispose old graph and create new graph
                        MainLineGraphs.Dispose()
                        If Machine = Machines.Others Then
                            MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                        Else
                            MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                        End If
                        sum_steps = 0
                    Else
                        'False: not add test , jump to RSS or LoA1
                        CurRun = CurRun.NextUnit.NextUnit.NextUnit.NextUnit
                        CurRun.Set_BackColor(Color.Yellow)
                        If CurRun.Name = "RSS" Then
                            timeLeft = CurRun.Time
                            timeLabel.Text = timeLeft & " s"
                            Countdown = False
                            MainLineGraphs.Dispose()
                            If Machine = Machines.Others Then
                                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)  'A4 mode
                            Else
                                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, CurRun.Time, 97) 'A1 A2 A3 mode
                            End If
                            For index = 0 To 9
                                array_step(index).Text = ""
                                array_step(index).BackColor = Color.DarkGray
                            Next
                        ElseIf CurRun.Name = "LoA1" Then
                            CurStep = CurRun.HeadStep
                            timeLeft = CurRun.Steps.Time
                            timeLabel.Text = timeLeft & " s"

                            'load LoA1's steps
                            For index = 0 To 9
                                array_step(index).Text = ""
                                array_step(index).BackColor = Color.DarkGray
                            Next
                            For index = CurRun.StartStep To CurRun.EndStep
                                sum_steps += CurStep.Time
                                array_step(index - 1).Text = CurStep.step_str
                                array_step(index - 1).BackColor = Color.White
                                If CurStep.HasNext() = True Then
                                    CurStep = CurStep.NextStep
                                End If
                            Next
                            array_step(0).BackColor = Color.Yellow
                            'dispose old graph and create new graph
                            MainLineGraphs.Dispose()
                            If Machine = Machines.Others Then
                                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                            Else
                                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                            End If
                            sum_steps = 0
                        End If
                    End If
                Else
                    MessageBox.Show("Error")
                End If
            ElseIf CurRun.Name = "ExA2_1st_Add" Or CurRun.Name = "ExA2_2nd_3rd_Add" Then
                If CurRun.NextUnit.Name = "ExA2_2nd_3rd_Add" Then
                    ' change light
                    Set_Panel_BackColor()

                    'jump to next Run_Unit
                    CurRun = CurRun.NextUnit
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load A2's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                ElseIf CurRun.NextUnit.Name = "RSS" Or CurRun.NextUnit.Name = "LoA1" Then
                    'have an additional test?
                    'call a function 
                    'if want_add()=true then ... elseif want_add()=false then ... endif

                    'jump to next Run_Unit
                    RandBool = RandGen.Next(0, 2).ToString
                    If RandBool = True Then
                        'True: add test
                        CurRun.PrevUnit.PrevUnit.Set_BackColor(Color.Yellow)
                        CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
                        CurRun.Set_BackColor(Color.IndianRed)
                        CurRun = CurRun.PrevUnit.PrevUnit
                        CurStep = CurRun.HeadStep
                        CurRun.CurStep = 1
                        CurRun.NextUnit.CurStep = 1
                        CurRun.NextUnit.NextUnit.CurStep = 1
                        Load_Steps_helper(CurRun)
                        Load_Steps_helper(CurRun.NextUnit)
                        Load_Steps_helper(CurRun.NextUnit.NextUnit)
                        timeLeft = CurRun.Steps.Time
                        timeLabel.Text = timeLeft & " s"
                        'load A2's steps
                        For index = 0 To 9
                            array_step(index).Text = ""
                            array_step(index).BackColor = Color.DarkGray
                        Next
                        For index = CurRun.StartStep To CurRun.EndStep
                            sum_steps += CurStep.Time
                            array_step(index - 1).Text = CurStep.step_str
                            array_step(index - 1).BackColor = Color.White
                            If CurStep.HasNext() = True Then
                                CurStep = CurStep.NextStep
                            End If
                        Next
                        array_step(0).BackColor = Color.Yellow
                        'dispose old graph and create new graph
                        MainLineGraphs.Dispose()
                        If Machine = Machines.Others Then
                            MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                        Else
                            MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                        End If
                        sum_steps = 0
                    Else
                        'False: not add test ,  jump to RSS or LoA1
                        CurRun.Set_BackColor(Color.Green)
                        CurRun.NextUnit.Set_BackColor(Color.Yellow)
                        CurRun = CurRun.NextUnit
                        If CurRun.NextUnit.Name = "RSS" Then

                            timeLeft = CurRun.Time
                            timeLabel.Text = timeLeft & " s"
                            Countdown = False
                            MainLineGraphs.Dispose()
                            If Machine = Machines.Others Then
                                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)  'A4 mode
                            Else
                                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, CurRun.Time, 97) 'A1 A2 A3 mode
                            End If
                            For index = 0 To 9
                                array_step(index).Text = ""
                                array_step(index).BackColor = Color.DarkGray
                            Next
                        ElseIf CurRun.NextUnit.Name = "LoA1" Then
                            timeLeft = CurRun.Steps.Time
                            timeLabel.Text = timeLeft & " s"

                            'load LoA1's steps
                            For index = 0 To 9
                                array_step(index).Text = ""
                                array_step(index).BackColor = Color.DarkGray
                            Next
                            For index = CurRun.StartStep To CurRun.EndStep
                                sum_steps += CurStep.Time
                                array_step(index - 1).Text = CurStep.step_str
                                array_step(index - 1).BackColor = Color.White
                                If CurStep.HasNext() = True Then
                                    CurStep = CurStep.NextStep
                                End If
                            Next
                            array_step(0).BackColor = Color.Yellow
                            'dispose old graph and create new graph
                            MainLineGraphs.Dispose()
                            If Machine = Machines.Others Then
                                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)  'A4 mode
                            Else
                                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, CurRun.Time, 97) 'A1 A2 A3 mode
                            End If
                        End If
                    End If
                Else
                    MessageBox.Show("Error")
                End If
            ElseIf CurRun.Name = "LoA1" Then
                If CurRun.NextUnit.Name = "LoA1" Then
                    'Case1: now is ExA1 and next is also ExA1
                    ' change light
                    Set_Panel_BackColor()

                    'jump to next Run_Unit
                    CurRun = CurRun.NextUnit
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load LoA1's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0

                ElseIf CurRun.NextUnit.Name = "LoA1_Add" Then
                    'Case2: now is ExA1 and next is ExA1_Add
                    'have an additional test?
                    'call a function 
                    'if want_add()=true then ... elseif want_add()=false then ... endif

                    CurRun.Set_BackColor(Color.Green)
                    'jump to next Run_Unit and change light
                    RandBool = RandGen.Next(0, 2).ToString
                    If RandBool = True Then
                        'True: add test
                        CurRun = CurRun.NextUnit
                        CurRun.Set_BackColor(Color.Yellow)
                    Else
                        'False: not add test , jump to LoA2_1st
                        CurRun = CurRun.NextUnit.NextUnit.NextUnit.NextUnit
                        CurRun.Set_BackColor(Color.Yellow)
                    End If
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load LoA1 or LoA2's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                Else
                    MessageBox.Show("Error")
                End If
            ElseIf CurRun.Name = "LoA1_Add" Then
                If CurRun.NextUnit.Name = "LoA1_Add" Then
                    'Case1: now is LoA1_Add and next is also LoA1_Add
                    ' change light
                    Set_Panel_BackColor()

                    'jump to next Run_Unit
                    CurRun = CurRun.NextUnit
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load LoA1's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                ElseIf CurRun.NextUnit.Name = "LoA2_1st" Then
                    'Case2: now is LoA1_Add and next is also LoA2_1st
                    'have an additional test?
                    'call a function 
                    'if want_add()=true then ... elseif want_add()=false then ... endif


                    'jump to next Run_Unit
                    RandBool = RandGen.Next(0, 2).ToString
                    If RandBool = True Then
                        'True: add test
                        CurRun.PrevUnit.PrevUnit.Set_BackColor(Color.Yellow)
                        CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
                        CurRun.Set_BackColor(Color.IndianRed)
                        CurRun = CurRun.PrevUnit.PrevUnit
                        CurRun.CurStep = 1
                        CurRun.NextUnit.CurStep = 1
                        CurRun.NextUnit.NextUnit.CurStep = 1
                        Load_Steps_helper(CurRun)
                        Load_Steps_helper(CurRun.NextUnit)
                        Load_Steps_helper(CurRun.NextUnit.NextUnit)
                    Else
                        'False: not add test
                        CurRun.Set_BackColor(Color.Green)
                        CurRun.NextUnit.Set_BackColor(Color.Yellow)
                        CurRun = CurRun.NextUnit
                    End If
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load LoA1_Add or LoA2_1st's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                Else
                    MessageBox.Show("Error")
                End If
            ElseIf CurRun.Name = "LoA2_1st" Or CurRun.Name = "LoA2_2nd_3rd" Then
                If CurRun.NextUnit.Name = "LoA2_1st" Or CurRun.NextUnit.Name = "LoA2_2nd_3rd" Then
                    'Case: LoA2_1st to LoA2_2nd_3rd or LoA2_2nd_3rd to LoA2_1st
                    ' change light
                    Set_Panel_BackColor()

                    'jump to next Run_Unit
                    CurRun = CurRun.NextUnit
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load LoA2's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                ElseIf CurRun.NextUnit.Name = "LoA2_1st_Add" Then
                    'Case: LoA2_2nd_3rd to LoA2_1st_Add
                    'have an additional test?
                    'call a function 
                    'if want_add()=true then ... elseif want_add()=false then ... endif

                    CurRun.Set_BackColor(Color.Green)
                    'jump to next Run_Unit and change light
                    RandBool = RandGen.Next(0, 2).ToString
                    If RandBool = True Then
                        'True: add test
                        CurRun = CurRun.NextUnit
                        CurRun.Set_BackColor(Color.Yellow)
                    Else
                        'False: not add test , jump to LoA3_fwd
                        CurRun = CurRun.NextUnit.NextUnit.NextUnit.NextUnit
                        CurRun.Set_BackColor(Color.Yellow)
                    End If
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load LoA2_1st_Add or LoA3_fwd's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                Else
                    MessageBox.Show("Error")
                End If
            ElseIf CurRun.Name = "LoA2_1st_Add" Or CurRun.Name = "LoA2_2nd_3rd_Add" Then
                If CurRun.NextUnit.Name = "LoA2_2nd_3rd_Add" Then
                    'Case1: now is LoA2_1st_Add or LoA2_2nd_3rd_Add and next is LoA2_2nd_3rd_Add
                    ' change light
                    Set_Panel_BackColor()

                    'jump to next Run_Unit
                    CurRun = CurRun.NextUnit
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load LoA1's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                ElseIf CurRun.NextUnit.Name = "LoA3_fwd" Then
                    'Case2: now is LoA2_2nd_3rd_Add and next is also LoA3_fwd
                    'have an additional test?
                    'call a function 
                    'if want_add()=true then ... elseif want_add()=false then ... endif


                    'jump to next Run_Unit
                    RandBool = RandGen.Next(0, 2).ToString
                    If RandBool = True Then
                        'True: add test
                        CurRun.PrevUnit.PrevUnit.Set_BackColor(Color.Yellow)
                        CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
                        CurRun.Set_BackColor(Color.IndianRed)
                        CurRun = CurRun.PrevUnit.PrevUnit
                        CurRun.CurStep = 1
                        CurRun.NextUnit.CurStep = 1
                        CurRun.NextUnit.NextUnit.CurStep = 1
                        Load_Steps_helper(CurRun)
                        Load_Steps_helper(CurRun.NextUnit)
                        Load_Steps_helper(CurRun.NextUnit.NextUnit)
                    Else
                        'False: not add test
                        CurRun.Set_BackColor(Color.Green)
                        CurRun.NextUnit.Set_BackColor(Color.Yellow)
                        CurRun = CurRun.NextUnit
                    End If
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load LoA2_1st_Add or LoA3_fwd's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                Else
                    MessageBox.Show("Error")
                End If
            ElseIf CurRun.Name = "LoA3_fwd" Or CurRun.Name = "LoA3_bkd" Or CurRun.Name = "TrA3_fwd" Or CurRun.Name = "TrA3_bkd" Then
                If CurRun.NextUnit.Name = "LoA3_fwd" Or CurRun.NextUnit.Name = "LoA3_bkd" Or CurRun.NextUnit.Name = "TrA3_fwd" Or CurRun.NextUnit.Name = "TrA3_bkd" Then
                    'Case: LoA3_fwd to LoA3_bkd or LoA3_bkd to LoA3_fwd or TrA3_bkd to TrA3_fwd or TrA3_bkd to TrA3_fwd
                    ' change light
                    Set_Panel_BackColor()

                    'jump to next Run_Unit
                    CurRun = CurRun.NextUnit
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load LoA3 or TrA3's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                ElseIf CurRun.NextUnit.Name = "LoA3_fwd_Add" Or CurRun.NextUnit.Name = "TrA3_fwd_Add" Then
                    'Case: LoA3_bkd to LoA3_fwd_Add or TrA3_bkd to TrA3_fwd_Add
                    'have an additional test?
                    'call a function 
                    'if want_add()=true then ... elseif want_add()=false then ... endif

                    CurRun.Set_BackColor(Color.Green)
                    'jump to next Run_Unit and change light
                    RandBool = RandGen.Next(0, 2).ToString
                    If RandBool = True Then
                        'True: add test
                        CurRun = CurRun.NextUnit
                        CurRun.Set_BackColor(Color.Yellow)
                        CurStep = CurRun.HeadStep
                        timeLeft = CurRun.Steps.Time
                        timeLabel.Text = timeLeft & " s"
                        'load LoA3_fwd_Add or TrA3_fwd_Add's steps
                        For index = 0 To 9
                            array_step(index).Text = ""
                            array_step(index).BackColor = Color.DarkGray
                        Next
                        For index = CurRun.StartStep To CurRun.EndStep
                            sum_steps += CurStep.Time
                            array_step(index - 1).Text = CurStep.step_str
                            array_step(index - 1).BackColor = Color.White
                            If CurStep.HasNext() = True Then
                                CurStep = CurStep.NextStep
                            End If
                        Next
                        array_step(0).BackColor = Color.Yellow

                        'dispose old graph and create new graph
                        MainLineGraphs.Dispose()
                        If Machine = Machines.Others Then
                            MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                        Else
                            MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                        End If
                        sum_steps = 0
                    Else
                        'False: not add test , jump to RSS
                        CurRun = CurRun.NextUnit.NextUnit.NextUnit
                        CurRun.Set_BackColor(Color.Yellow)
                        timeLeft = CurRun.Time
                        timeLabel.Text = timeLeft & " s"
                        Countdown = False
                        For index = 0 To 9
                            array_step(index).Text = ""
                            array_step(index).BackColor = Color.DarkGray
                        Next
                        'dispose old graph and create new graph
                        MainLineGraphs.Dispose()
                        If Machine = Machines.Others Then
                            MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                        Else
                            MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                        End If
                    End If
                Else
                    MessageBox.Show("Error")
                End If
            ElseIf CurRun.Name = "LoA3_fwd_Add" Or CurRun.Name = "LoA3_bkd_Add" Or CurRun.Name = "TrA3_fwd_Add" Or CurRun.Name = "TrA3_bkd_Add" Then
                If CurRun.NextUnit.Name = "LoA3_bkd_Add" Or CurRun.NextUnit.Name = "TrA3_bkd_Add" Then
                    'Case: LoA3_fwd_Add to LoA3_bkd_Add 
                    ' change light
                    Set_Panel_BackColor()

                    'jump to next Run_Unit
                    CurRun = CurRun.NextUnit
                    CurStep = CurRun.HeadStep
                    timeLeft = CurRun.Steps.Time
                    timeLabel.Text = timeLeft & " s"

                    'load LoA3 or TrA3's steps
                    For index = 0 To 9
                        array_step(index).Text = ""
                        array_step(index).BackColor = Color.DarkGray
                    Next
                    For index = CurRun.StartStep To CurRun.EndStep
                        sum_steps += CurStep.Time
                        array_step(index - 1).Text = CurStep.step_str
                        array_step(index - 1).BackColor = Color.White
                        If CurStep.HasNext() = True Then
                            CurStep = CurStep.NextStep
                        End If
                    Next
                    array_step(0).BackColor = Color.Yellow

                    'dispose old graph and create new graph
                    MainLineGraphs.Dispose()
                    If Machine = Machines.Others Then
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                    Else
                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                    End If
                    sum_steps = 0
                ElseIf CurRun.NextUnit.Name = "RSS" Then
                    'Case: LoA3_bkd_Add to RSS or TrA3_bkd_Add to RSS
                    'have an additional test?
                    'call a function 
                    'if want_add()=true then ... elseif want_add()=false then ... endif

                    RandBool = RandGen.Next(0, 2).ToString
                    If RandBool = True Then
                        'True: add test , jump to LoA3_fwd_Add or jump to TrA3_fwd_Add
                        'True: add test
                        CurRun.PrevUnit.Set_BackColor(Color.Yellow)
                        CurRun.Set_BackColor(Color.IndianRed)
                        CurRun = CurRun.PrevUnit
                        CurStep = CurRun.HeadStep
                        CurRun.CurStep = 1
                        CurRun.NextUnit.CurStep = 1
                        Load_Steps_helper(CurRun)
                        Load_Steps_helper(CurRun.NextUnit)
                        timeLeft = CurRun.Steps.Time
                        timeLabel.Text = timeLeft & " s"

                        'load LoA3_fwd or TrA3_fwd's steps
                        For index = 0 To 9
                            array_step(index).Text = ""
                            array_step(index).BackColor = Color.DarkGray
                        Next
                        For index = CurRun.StartStep To CurRun.EndStep
                            sum_steps += CurStep.Time
                            array_step(index - 1).Text = CurStep.step_str
                            array_step(index - 1).BackColor = Color.White
                            If CurStep.HasNext() = True Then
                                CurStep = CurStep.NextStep
                            End If
                        Next
                        array_step(0).BackColor = Color.Yellow

                        'dispose old graph and create new graph
                        MainLineGraphs.Dispose()
                        If Machine = Machines.Others Then
                            MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                        Else
                            MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                        End If
                        sum_steps = 0
                    Else
                        'False: not add test , jump to RSS
                        CurRun.Set_BackColor(Color.Green)
                        CurRun = CurRun.NextUnit
                        CurRun.Set_BackColor(Color.Yellow)
                        timeLeft = CurRun.Time
                        timeLabel.Text = timeLeft & " s"
                        Countdown = False
                        For index = 0 To 9
                            array_step(index).Text = ""
                            array_step(index).BackColor = Color.DarkGray
                        Next
                        'dispose old graph and create new graph
                        MainLineGraphs.Dispose()
                        If Machine = Machines.Others Then
                            MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, sum_steps, 97)  'A4 mode
                        Else
                            MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97) 'A1 A2 A3 mode
                        End If
                    End If
                Else
                    MessageBox.Show("Error")
                End If

            End If
        End If
        Accept_Button.Enabled = False
        startButton.Enabled = True

        'If MessageBox.Show("此步驟數據已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
        'MessageBox.Show("此機具所有測量步驟已結束")
        'End If
        'If CurRun.PrevUnit.Name = "PreCal" Then
        '    MainLineGraphs.Dispose()
        '    If Machine = Machines.Others Then
        '        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)
        '    Else
        '        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, CurRun.Time, 97)
        '    End If
        '    Set_Panel_BackColor()
        '    timeLeft = CurRun.Time
        '    timeLabel.Text = timeLeft & " s"
        'Else
        '    If CurRun.PrevUnit.CurStep = 0 Then
        '        If CurRun.PrevUnit.Name = "Background" Then
        '            MainLineGraphs.Dispose()
        '            If Machine = Machines.Others Then
        '                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)
        '            Else
        '                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, CurRun.Time, 97)
        '            End If
        '            Set_Panel_BackColor()
        '            For index = 0 To 9
        '                array_step(index).Text = ""
        '                array_step(index).BackColor = Color.DarkGray
        '            Next
        '            For index = CurRun.StartStep To CurRun.EndStep
        '                sum_steps += CurStep.Time
        '                array_step(index - 1).Text = CurStep.step_str
        '                array_step(index - 1).BackColor = Color.Yellow
        '                If CurStep.HasNext() = True Then
        '                    CurStep = CurStep.NextStep
        '                End If
        '            Next
        '            timeLeft = CurStep.Time
        '            timeLabel.Text = timeLeft & " s"
        '            sum_steps = 0
        '        ElseIf CurRun.PrevUnit.Name = "PostCal" Then
        '            MainLineGraphs.Dispose()
        '            If Machine = Machines.Others Then
        '                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)
        '            Else
        '                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, CurRun.Time, 97)
        '            End If
        '            Set_Panel_BackColor()
        '            timeLeft = CurRun.Time
        '            timeLabel.Text = timeLeft & " s"
        '        Else
        '            MainLineGraphs.Dispose()
        '            If Machine = Machines.Others Then
        '                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)
        '            Else
        '                MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, CurRun.Time, 97)
        '            End If
        '            Set_Panel_BackColor()
        '            timeLeft = CurRun.Time
        '            timeLabel.Text = timeLeft & " s"
        '        End If
        '    Else
        '        If CurRun.PrevUnit.CurStep = CurRun.PrevUnit.EndStep And CurRun.EndStep > 0 Then
        '            If CurRun.Name = "ExA1_Add" Or CurRun.Name = "ExA2_1st_Add" Or CurRun.Name = "LoA1_Add" Or CurRun.Name = "TrA1_Add" Or CurRun.Name = "LoA2_1st_Add" Or CurRun.Name = "LoA3_fwd_Add" Or CurRun.Name = "TrA3_fwd_Add" Or CurRun.Name = "A4_Add" Then
        '                If MessageBox.Show("是否要加測?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
        '                    CurRun.PrevUnit.Set_BackColor(Color.Green)
        '                    If CurRun.Name = "ExA1_Add" Or CurRun.Name = "LoA1_Add" Or CurRun.Name = "TrA1_Add" Or CurRun.Name = "LoA2_1st_Add" Or CurRun.NextUnit.NextUnit.NextUnit.Name = "LoA1" Then
        '                        CurRun = CurRun.NextUnit.NextUnit.NextUnit
        '                        CurStep = CurRun.Steps
        '                        CurRun.Set_BackColor(Color.Yellow)
        '                        MainLineGraphs.Dispose()
        '                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97)
        '                        For index = 0 To 9
        '                            array_step(index).Text = ""
        '                            array_step(index).BackColor = Color.DarkGray
        '                        Next
        '                        For index = CurRun.StartStep To CurRun.EndStep
        '                            sum_steps += CurStep.Time
        '                            array_step(index - 1).Text = CurStep.step_str
        '                            array_step(index - 1).BackColor = Color.White
        '                            If CurStep.HasNext() = True Then
        '                                CurStep = CurStep.NextStep
        '                            End If
        '                        Next
        '                        timeLeft = CurRun.Steps.Time
        '                        timeLabel.Text = timeLeft & " s"
        '                        array_step(0).BackColor = Color.Yellow
        '                        sum_steps = 0
        '                    ElseIf CurRun.Name = "A4_Add" Then
        '                        For index = 0 To 9
        '                            array_step(index).Text = ""
        '                            array_step(index).BackColor = Color.DarkGray
        '                        Next
        '                        CurRun = CurRun.NextUnit
        '                        CurStep = CurRun.Steps
        '                        CurRun.Set_BackColor(Color.Yellow)
        '                        MainLineGraphs.Dispose()
        '                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)
        '                        timeLeft = CurRun.Time
        '                        timeLabel.Text = timeLeft & " s"
        '                    ElseIf CurRun.Name = "ExA2_1st_Add" Then
        '                        For index = 0 To 9
        '                            array_step(index).Text = ""
        '                            array_step(index).BackColor = Color.DarkGray
        '                        Next
        '                        CurRun = CurRun.NextUnit.NextUnit.NextUnit
        '                        CurStep = CurRun.Steps
        '                        CurRun.Set_BackColor(Color.Yellow)
        '                        MainLineGraphs.Dispose()
        '                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97)
        '                        timeLeft = CurRun.Time
        '                        timeLabel.Text = timeLeft & " s"
        '                    ElseIf CurRun.Name = "LoA3_fwd_Add" Or CurRun.Name = "TrA3_fwd_Add" Then
        '                        For index = 0 To 9
        '                            array_step(index).Text = ""
        '                            array_step(index).BackColor = Color.DarkGray
        '                        Next
        '                        CurRun = CurRun.NextUnit.NextUnit
        '                        CurStep = CurRun.Steps
        '                        CurRun.Set_BackColor(Color.Yellow)
        '                        MainLineGraphs.Dispose()
        '                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, sum_steps, 97)
        '                        timeLeft = CurRun.Time
        '                        timeLabel.Text = timeLeft & " s"
        '                    End If
        '                Else
        '                    MainLineGraphs.Dispose()
        '                    If Machine = Machines.Others Then
        '                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)
        '                    Else
        '                        MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, CurRun.Time, 97)
        '                    End If
        '                    Set_Panel_BackColor()
        '                    For index = 0 To 9
        '                        array_step(index).Text = ""
        '                        array_step(index).BackColor = Color.DarkGray
        '                    Next
        '                    For index = CurRun.StartStep To CurRun.EndStep
        '                        sum_steps += CurStep.Time
        '                        array_step(index - 1).Text = CurStep.step_str
        '                        array_step(index - 1).BackColor = Color.White
        '                        If CurStep.HasNext() = True Then
        '                            CurStep = CurStep.NextStep
        '                        End If
        '                    Next
        '                    timeLeft = CurRun.Steps.Time
        '                    timeLabel.Text = timeLeft & " s"
        '                    array_step(0).BackColor = Color.Yellow
        '                    sum_steps = 0
        '                End If
        '            Else
        '                MainLineGraphs.Dispose()
        '                If Machine = Machines.Others Then
        '                    MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A4, 1, CurRun.Time, 97)
        '                Else
        '                    MainLineGraphs = New LineGraphPanel(New Point(10, 3), New Size(1219, 97), TabPage2, CGraph.Modes.A1A2A3, 1, CurRun.Time, 97)
        '                End If
        '                Set_Panel_BackColor()
        '                For index = 0 To 9
        '                    array_step(index).Text = ""
        '                    array_step(index).BackColor = Color.DarkGray
        '                Next
        '                For index = CurRun.StartStep To CurRun.EndStep
        '                    sum_steps += CurStep.Time
        '                    array_step(index - 1).Text = CurStep.step_str
        '                    array_step(index - 1).BackColor = Color.White
        '                    If CurStep.HasNext() = True Then
        '                        CurStep = CurStep.NextStep
        '                    End If
        '                Next
        '                timeLeft = CurRun.Steps.Time
        '                timeLabel.Text = timeLeft & " s"
        '                array_step(0).BackColor = Color.Yellow
        '                sum_steps = 0
        '            End If

        '        End If
        '    End If
        'End If
    End Sub




    'plot the coordinate system on startup
    Private Sub plotCor(ByVal xCor, ByVal yCor)

        'x axis
        xAxis = New LineShape(xCor(0, 0), xCor(0, 1), xCor(1, 0), xCor(1, 1))
        xAxis.Parent = canvas
        xLabel.Text = "x"
        xLabel.Size = New System.Drawing.Size(10, 10)
        xLabel.Location = New System.Drawing.Point(xCor(1, 0), xCor(1, 1))

        'y axis
        yAxis = New LineShape(yCor(0, 0), yCor(0, 1), yCor(1, 0), yCor(1, 1))
        yAxis.Parent = canvas
        yLabel.Text = "y"
        yLabel.Location = New System.Drawing.Point(yCor(0, 0), yCor(0, 1))
    End Sub

    'given an array of coordinates, plot them out using the given coordinate system
    'xCor is x axis coordinates
    'yCor is y axis coordinates
    'parent is the parent control to canvas
    'r is the radius for the circle
    'coors is the coordinates for the points to be plotted
    Private Sub plot(ByVal xCor As Double(,), ByVal yCor As Double(,), ByVal parent As Control, ByVal r As Double, ByVal coors As CoorPoint())
        'clear canvas first
        If Not IsNothing(canvas) Then
            canvas.Dispose()
        End If
        'clear labels
        For Each cp As CoorPoint In pos
            cp.Label.Text = ""
        Next
        For Each cp As CoorPoint In pos2
            cp.Label.Text = ""
        Next

        canvas = New ShapeContainer()
        canvas.Parent = parent
        canvas.BackColor = Color.Transparent
        'plot the coordinates
        plotCor(xCor, yCor)

        'plot the circle with r
        Dim temp As Double = length / r
        ratio = CInt(temp)

        If ratio > temp Then
            ratio = ratio - 1
        ElseIf ratio = 0 Then
            ratio = temp
        End If
        r = r * ratio
        Dim rCircle = New OvalShape((origin(0) - r), (origin(1) - r), 2 * r, 2 * r)
        rCircle.Parent = canvas
        'plot normalized points

        For index = 0 To coors.GetLength(0) - 1
            Dim x = coors(index).Coors.Xc * ratio
            Dim y = coors(index).Coors.Yc * ratio
            Dim rPoint = New OvalShape((origin(0) + x - 2), (origin(1) - y - 2), 2, 2)

            rPoint.Parent = canvas
            Dim xText = Format(coors(index).Coors.Xc, "#.##")
            Dim yText = Format(coors(index).Coors.Yc, "#.##")
            Dim zText = Format(coors(index).Coors.Zc, "#.##")
            If xText = "" Then
                xText = "0"
            End If
            If yText = "" Then
                yText = "0"
            End If
            If zText = "" Then
                zText = "0"
            End If
            coors(index).Label.Text = "(" + xText + " , " + yText + " , " + zText + ")"
            coors(index).Label.Location = translate(New System.Drawing.Point(origin(0) + x, origin(1) - y), GroupBox_Plot.Location, TabPage1.Location)
            coors(index).Label.BringToFront()
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

End Class
