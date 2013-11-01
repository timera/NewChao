Imports System.Runtime.InteropServices
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

    'Graphs for the main screen
    Dim MainLineGraphs As LineGraphPanel
    Dim MainBarGraph As BarGraph

    Dim MachChosen As Boolean
    Dim timeLeft As Integer
    Dim status As Integer  'status=1==>A1+A2  ,status=2==>A1+A3  ,status=3==>A1+A2+A3  ,status=4==>A4
    Dim state As Integer   'state表示現在正在測的是A1 or A2 or A3 or A4

    '手動設秒數 = false  ,  自動設30s = true
    Dim auto As Boolean

    'sStep==>循環開始(start)的步驟 ,eStep==>循環的最後一個(end)步驟
    Dim sStep As Integer
    Dim eStep As Integer

    'pass 為是否接受數據還是要加測
    Dim pass As Integer

    Dim TimeofStep(1, 8) As Double
    Dim array_step(8) As Label

    '紀錄a2 a3跑到哪個步驟
    Dim A2_step As Integer
    Dim A3_step As Integer


    Public Sub New()

        ' 此為設計工具所需的呼叫。
        InitializeComponent()
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        ' 在 InitializeComponent() 呼叫之後加入任何初始設定。

    End Sub
    Private Sub Program_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim margins As MARGINS = New MARGINS
        margins.cxLeftWidth = 2
        margins.cxRightWidth = 2
        margins.cyTopHeight = 23
        margins.cyButtomHeight = 2
        Dim hwnd As IntPtr = Handle
        Dim result As Integer = DwmExtendFrameIntoClientArea(hwnd, margins)

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

        light_preCal.BackColor = Color.Red
        light_BG.BackColor = Color.Red
        light_1st.BackColor = Color.Red
        light_2nd.BackColor = Color.Red
        light_3rd.BackColor = Color.Red
        light_Additional.BackColor = Color.Red
        light_RSS.BackColor = Color.Red
        light_postCal.BackColor = Color.Red

        LinkLabel_A1.Enabled = False
        LinkLabel_A2.Enabled = False
        LinkLabel_A3.Enabled = False
        LinkLabel_A4.Enabled = False

        startButton.Enabled = True
        stopButton.Enabled = False
        continueButton.Enabled = False

        array_step(0) = Step1
        array_step(1) = Step2
        array_step(2) = Step3
        array_step(3) = Step4
        array_step(4) = Step5
        array_step(5) = Step6
        array_step(6) = Step7
        array_step(7) = Step8
        array_step(8) = Step9

        A2_step = 0
        A3_step = 0

        Noise1.Text = 100
        Noise2.Text = 105
        Noise3.Text = 103.8


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



        'location of all objects
        light_preCal.Location = New Point(19, 17)
        light_BG.Location = New Point(19, 57)
        light_1st.Location = New Point(19, 97)
        light_2nd.Location = New Point(19, 137)
        light_3rd.Location = New Point(19, 177)
        light_Additional.Location = New Point(19, 217)
        light_RSS.Location = New Point(19, 257)
        light_postCal.Location = New Point(19, 297)
        light_A1.Location = New Point(19, 508)
        light_A2.Location = New Point(19, 548)
        light_A3.Location = New Point(19, 588)
        light_A4.Location = New Point(19, 628)

        LinkLabel_preCal.Location = New Point(50, 22)
        LinkLabel_BG.Location = New Point(50, 62)
        LinkLabel_1st.Location = New Point(50, 102)
        LinkLabel_2nd.Location = New Point(50, 142)
        LinkLabel_3rd.Location = New Point(50, 182)
        LinkLabel_Additional.Location = New Point(50, 222)
        LinkLabel_RSS.Location = New Point(50, 262)
        LinkLabel_postCal.Location = New Point(50, 302)
        LinkLabel_A1.Location = New Point(50, 513)
        LinkLabel_A2.Location = New Point(50, 553)
        LinkLabel_A3.Location = New Point(50, 593)
        LinkLabel_A4.Location = New Point(50, 633)

        startButton.Location = New Point(9, 354)
        timeLabel.Location = New Point(9, 380)
        stopButton.Location = New Point(9, 461)
        continueButton.Location = New Point(90, 461)

        Noise1.Location = New Point(246, 580)
        Noise2.Location = New Point(319, 580)
        Noise3.Location = New Point(392, 580)
        Noise4.Location = New Point(465, 580)
        Noise5.Location = New Point(538, 580)
        Noise6.Location = New Point(611, 580)

        Step1.Location = New Point(703, 22)
        Step2.Location = New Point(703, 84)
        Step3.Location = New Point(703, 146)
        Step4.Location = New Point(703, 208)
        Step5.Location = New Point(703, 270)
        Step6.Location = New Point(703, 332)
        Step7.Location = New Point(703, 394)
        Step8.Location = New Point(703, 456)
        Step9.Location = New Point(703, 518)

        MachChosen = False

        'Graphing for the main screen
        MainLineGraphs = New LineGraphPanel(New Point(300, 100), New Size(800, 200), TabPage2, CGraph.Modes.A1A2A3, 3, {30, 30, 30})
        MainBarGraph = New BarGraph(New Point(300, 300), New Size(800, 200), TabPage2, CGraph.Modes.A1A2A3)
    End Sub




    Private Sub TabControl1_IndexChanged(ByVal sender As TabControl, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If Not MachChosen Then
            If sender.SelectedIndex = 1 Then
                MessageBox.Show("還未選機具!", "My application", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                TabControl1.SelectedIndex = 0 ' go back to the machinery choose stage
            End If
        End If
    End Sub

    Private Sub ComboBox_machine_list_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_machine_list.SelectedIndexChanged
        Dim mach As Integer = 0

        'step1~9清空
        Step1.Text = ""
        Step2.Text = ""
        Step3.Text = ""
        Step4.Text = ""
        Step5.Text = ""
        Step6.Text = ""
        Step7.Text = ""
        Step8.Text = ""
        Step9.Text = ""
        Step1.BackColor = Color.DarkGray
        Step2.BackColor = Color.DarkGray
        Step3.BackColor = Color.DarkGray
        Step4.BackColor = Color.DarkGray
        Step5.BackColor = Color.DarkGray
        Step6.BackColor = Color.DarkGray
        Step7.BackColor = Color.DarkGray
        Step8.BackColor = Color.DarkGray
        Step9.BackColor = Color.DarkGray

        'A1+A2
        If ComboBox_machine_list.Text = "開挖機(Excavator)" Then
            Picture_machine.Image = My.Resources.Resource1.小型開挖機_compact_excavator_
            mach = 1
            status = 1
        End If

        'A1+A3
        If ComboBox_machine_list.Text = "推土機(Crawler and wheel tractor)" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式推土機_crawler_dozer_
            mach = 1
            status = 2
        End If
        If ComboBox_machine_list.Text = "鐵輪壓路機(Road roller)" Then
            Picture_machine.Image = My.Resources.Resource1.壓路機_rollers_
            mach = 1
            status = 2
        End If
        If ComboBox_machine_list.Text = "膠輪壓路機(Wheel roller)" Then
            Picture_machine.Image = My.Resources.Resource1.壓路機_rollers_
            mach = 1
            status = 2
        End If
        If ComboBox_machine_list.Text = "振動式壓路機(Vibrating roller)" Then
            Picture_machine.Image = My.Resources.Resource1.壓路機_rollers_
            mach = 1
            status = 2
        End If

        'A1+A2+A3
        If ComboBox_machine_list.Text = "裝料機(Crawler and wheel loader)" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式裝料機_crawler_loader_
            mach = 1
            status = 3
        End If
        If ComboBox_machine_list.Text = "裝料開挖機" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式開挖裝料機_crawler_backhoe_loader_
            mach = 1
            status = 3
        End If

        'A4
        If ComboBox_machine_list.Text = "履帶起重機(Crawler crane)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式推土機_wheeled_dozer_
            mach = 2
            Step1.Text = My.Resources.A4_Crane
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "卡車起重機(Truck crane)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式開挖裝料機_wheeled_backhoe_loader_
            mach = 2
            Step1.Text = My.Resources.A4_Crane
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "輪形起重機(Wheel crane)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式開挖機_wheeled_excavator_
            mach = 2
            Step1.Text = My.Resources.A4_Crane
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "振動式樁錘(Vibrating hammer)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式裝料機_wheeled_loader_
            mach = 2
            Step1.Text = My.Resources.A4_Vibrating_Hammer
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "油壓式打樁機(Hydraulic pile driver)" Then
            Picture_machine.Image = My.Resources.Resource1.壓路機_rollers_
            mach = 2
            Step1.Text = My.Resources.A4_Auger_Drill_Driver
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "拔樁機" Then
            Picture_machine.Image = My.Resources.Resource1.小型開挖機_compact_excavator_
            mach = 2
            Step1.Text = My.Resources.A4_Auger_Drill_Driver
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "油壓式拔樁機" Then
            Picture_machine.Image = My.Resources.Resource1.小型膠輪式裝料機_compact_loader__wheeled_
            mach = 2
            Step1.Text = My.Resources.A4_Auger_Drill_Driver
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "土壤取樣器(地鑽) (Earth auger)" Then
            Picture_machine.Image = My.Resources.Resource1.滑移裝料機_skid_steer_loader_
            mach = 2
            Step1.Text = My.Resources.A4_Auger_Drill_Driver
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "全套管鑽掘機" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式開挖裝料機_crawler_backhoe_loader_
            mach = 2
            Step1.Text = My.Resources.A4_Auger_Drill_Driver
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "鑽土機(Earth drill)" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式推土機_crawler_dozer_
            mach = 2
            Step1.Text = My.Resources.A4_Auger_Drill_Driver
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "鑽岩機(Rock breaker)" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式裝料機_crawler_loader_
            mach = 2
            Step1.Text = My.Resources.A4_Auger_Drill_Driver
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "混凝土泵車(Concrete pump)" Then
            Picture_machine.Image = My.Resources.Resource1.履帶式開挖機_crawler_excavator_
            mach = 2
            Step1.Text = My.Resources.A4_Concrete_Pump
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "混凝土破碎機(Concrete breaker)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式推土機_wheeled_dozer_
            mach = 2
            Step1.Text = My.Resources.A4_Concrete_Breaker
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "瀝青混凝土舖築機(Asphalt finisher)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式開挖裝料機_wheeled_backhoe_loader_
            mach = 2
            Step1.Text = My.Resources.A4_Asphalt_Finisher
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "混凝土割切機(Concrete cutter)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式開挖機_wheeled_excavator_
            mach = 2
            Step1.Text = My.Resources.A4_Concrete_Cutter
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "發電機(Generator)" Then
            Picture_machine.Image = My.Resources.Resource1.膠輪式裝料機_wheeled_loader_
            mach = 2
            Step1.Text = My.Resources.A4_Generator
            Step1.BackColor = Color.MistyRose
        End If
        If ComboBox_machine_list.Text = "空氣壓縮機(Compressor)" Then
            Picture_machine.Image = My.Resources.Resource1.壓路機_rollers_
            mach = 2
            Step1.Text = My.Resources.A4_Compressor
            Step1.BackColor = Color.MistyRose
        End If

        If Not mach = 0 Then
            MachChosen = True
        End If
        'mach區分要輸入至哪個groupbox
        If mach = 1 Then
            GroupBox_A1_A2_A3.Enabled = True
            GroupBox_A4.Enabled = False
        End If
        If mach = 2 Then
            GroupBox_A1_A2_A3.Enabled = False
            GroupBox_A4.Enabled = True
            status = 0
            auto = True

            'light A1 A2 A3和linklabel A1 A2 A3 要 disable
            light_A1.BackColor = Color.DarkGray
            light_A2.BackColor = Color.DarkGray
            light_A3.BackColor = Color.DarkGray
            light_A4.BackColor = Color.Red

            LinkLabel_A1.Enabled = False
            LinkLabel_A2.Enabled = False
            LinkLabel_A3.Enabled = False
            LinkLabel_A4.Enabled = True
            'time of step
            TimeofStep(0, 0) = 30
            TimeofStep(1, 0) = -1 'timeofstep(1,0)=-1 用來判別表A4狀態

        End If

        'status 1,2,3 從A1開始做
        If status = 1 Or status = 2 Or status = 3 Then
            'step1~9清空
            Step1.Text = ""
            Step2.Text = ""
            Step3.Text = ""
            Step4.Text = ""
            Step5.Text = ""
            Step6.Text = ""
            Step7.Text = ""
            Step8.Text = ""
            Step9.Text = ""
            'load A1_step1
            Step1.Text = My.Resources.A1_step1
            Step1.BackColor = Color.Orange
            state = 1
            'time of step
            TimeofStep(0, 0) = 30
            TimeofStep(1, 0) = 3
            Step2.Text = "循環剩餘次數 : " & TimeofStep(1, 0) & " 次"
            'A1固定30s  auto改true 
            auto = True
            'A4 disable A1 enable
            light_A4.BackColor = Color.DarkGray
            LinkLabel_A4.Enabled = False
            light_A1.BackColor = Color.Red
            LinkLabel_A1.Enabled = True
            If status = 1 Then
                light_A3.BackColor = Color.DarkGray
                LinkLabel_A3.Enabled = False
                light_A2.BackColor = Color.Red
                LinkLabel_A2.Enabled = True
            End If
            If status = 2 Then
                light_A2.BackColor = Color.DarkGray
                LinkLabel_A2.Enabled = False
                light_A3.BackColor = Color.Red
                LinkLabel_A3.Enabled = True
            End If
            If status = 3 Then
                light_A3.BackColor = Color.Red
                LinkLabel_A3.Enabled = True
                light_A2.BackColor = Color.Red
                LinkLabel_A2.Enabled = True
            End If
        End If

        If ComboBox_machine_list.Text = "A1+A2" Or ComboBox_machine_list.Text = "A1+A3" Or ComboBox_machine_list.Text = "A1+A2+A3" Or ComboBox_machine_list.Text = "A4" Then
            Step1.Text = ""
            Step2.Text = ""
            Step3.Text = ""
            Step4.Text = ""
            Step5.Text = ""
            Step6.Text = ""
            Step7.Text = ""
            Step8.Text = ""
            Step9.Text = ""

            light_A1.BackColor = Color.DarkGray
            light_A2.BackColor = Color.DarkGray
            light_A3.BackColor = Color.DarkGray
            light_A4.BackColor = Color.DarkGray

            LinkLabel_A1.Enabled = False
            LinkLabel_A2.Enabled = False
            LinkLabel_A3.Enabled = False
            LinkLabel_A4.Enabled = False
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

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If timeLeft > 0 Then
            timeLeft = timeLeft - 1
            timeLabel.Text = timeLeft & " s"
        End If
        If timeLeft = 0 Then
            Timer1.Stop()
            timeLabel.Text = "Time's up!"
            startButton.Enabled = True
            If A2_step > 0 Then
                A2_step += 1
                If A2_step > eStep Then
                    A2_step = sStep
                End If
            End If
            If A3_step > 0 Then
                A3_step += 1
                If A3_step > eStep Then
                    A3_step = sStep
                End If
            End If
            'A1改燈號
            If TimeofStep(0, 0) = 30 And TimeofStep(1, 0) = 0 And state = 1 Then
                If light_1st.BackColor = Color.Red Then
                    TimeofStep(1, 0) = 3
                    light_1st.BackColor = Color.DarkGreen
                    light_2nd.BackColor = Color.LightGreen
                ElseIf light_2nd.BackColor = Color.LightGreen Then
                    TimeofStep(1, 0) = 3
                    light_2nd.BackColor = Color.DarkGreen
                    light_3rd.BackColor = Color.LightGreen
                ElseIf light_3rd.BackColor = Color.LightGreen Then
                    TimeofStep(0, 0) = 0
                    light_3rd.BackColor = Color.DarkGreen
                    '這裡計算是否要加測or接收數據
                    pass = 1  '目前假設不加測
                    'if 加測==>additional
                    'else 接收數據==>判斷下一步是A2還是A3
                    '判斷完設定狀態及載入步驟
                    If pass = 1 Then
                        '下一步是A2
                        '設定參數
                        light_A1.BackColor = Color.DarkGreen
                        If status = 1 Or status = 3 Then
                            A2_step = 1
                            'if A1+A2===>excavator
                            If status = 1 Then
                                '清空所有steps
                                Step1.Text = ""
                                Step2.Text = ""
                                Step3.Text = ""
                                Step4.Text = ""
                                Step5.Text = ""
                                Step6.Text = ""
                                Step7.Text = ""
                                Step8.Text = ""
                                Step9.Text = ""

                                '改手動輸入秒數
                                auto = False
                                '設定該循環的步驟
                                sStep = 2
                                eStep = 9

                                '載入步驟
                                Step1.Text = My.Resources.A2_Excavator_step1
                                Step1.BackColor = Color.MistyRose
                                Step2.Text = My.Resources.A2_Excavator_step2
                                Step2.BackColor = Color.Orange
                                Step3.Text = My.Resources.A2_Excavator_step3
                                Step3.BackColor = Color.Orange
                                Step4.Text = My.Resources.A2_Excavator_step4
                                Step4.BackColor = Color.Orange
                                Step5.Text = My.Resources.A2_Excavator_step5
                                Step5.BackColor = Color.Orange
                                Step6.Text = My.Resources.A2_Excavator_step6
                                Step6.BackColor = Color.Orange
                                Step7.Text = My.Resources.A2_Excavator_step7
                                Step7.BackColor = Color.Orange
                                Step8.Text = My.Resources.A2_Excavator_step8
                                Step8.BackColor = Color.Orange
                                Step9.Text = My.Resources.A2_Excavator_step9
                                Step9.BackColor = Color.Orange
                                state = 2
                                '設定timeofstep array 中的值
                                For index = 0 To eStep - 1
                                    TimeofStep(0, index) = 5
                                    If index > sStep - 2 Then
                                        TimeofStep(1, index) = 3
                                    Else
                                        TimeofStep(1, index) = 0
                                    End If
                                Next
                                '1st 2nd 3rd 燈號變換
                                light_1st.BackColor = Color.LightGreen
                                light_2nd.BackColor = Color.Red
                                light_3rd.BackColor = Color.Red

                            End If
                            'if A1+A2+A3==>
                            If status = 3 Then
                                '清空所有steps
                                Step1.Text = ""
                                Step2.Text = ""
                                Step3.Text = ""
                                Step4.Text = ""
                                Step5.Text = ""
                                Step6.Text = ""
                                Step7.Text = ""
                                Step8.Text = ""
                                Step9.Text = ""

                                '改手動輸入秒數
                                auto = False
                                '設定該循環的步驟
                                sStep = 2
                                eStep = 3

                                '載入步驟
                                Step1.Text = My.Resources.A2_Loader_step1
                                Step1.BackColor = Color.MistyRose
                                Step2.Text = My.Resources.A2_Loader_step2
                                Step2.BackColor = Color.Orange
                                Step3.Text = My.Resources.A2_Loader_step3
                                Step3.BackColor = Color.Orange
                                Step4.BackColor = Color.DarkGray
                                Step5.BackColor = Color.DarkGray
                                Step6.BackColor = Color.DarkGray
                                Step7.BackColor = Color.DarkGray
                                Step8.BackColor = Color.DarkGray
                                Step9.BackColor = Color.DarkGray
                                state = 2
                                '設定timeofstep array 中的值
                                For i = 0 To eStep - 1
                                    TimeofStep(0, i) = 5
                                Next
                                For j = sStep - 1 To eStep - 1
                                    TimeofStep(0, j) = 3
                                Next
                                '1st 2nd 3rd 燈號變換
                                light_1st.BackColor = Color.LightGreen
                                light_2nd.BackColor = Color.Red
                                light_3rd.BackColor = Color.Red
                            End If
                            MessageBox.Show("A2測量開始")
                            light_A2.BackColor = Color.LightGreen
                        End If
                        '下一步是A3
                        '設定參數
                        'if A1+A3==>tractor
                        If status = 2 Then
                            A3_step = 1
                            '清空所有steps
                            Step1.Text = ""
                            Step2.Text = ""
                            Step3.Text = ""
                            Step4.Text = ""
                            Step5.Text = ""
                            Step6.Text = ""
                            Step7.Text = ""
                            Step8.Text = ""
                            Step9.Text = ""

                            '改手動輸入秒數
                            auto = False
                            '設定該循環的步驟
                            sStep = 3
                            eStep = 4

                            '載入步驟
                            Step1.Text = My.Resources.A3_Tractor_step1
                            Step1.BackColor = Color.MistyRose
                            Step2.Text = My.Resources.A3_Tractor_step2
                            Step2.BackColor = Color.MistyRose
                            Step3.Text = My.Resources.A3_Tractor_step3
                            Step3.BackColor = Color.Orange
                            Step4.Text = My.Resources.A3_Tractor_step4
                            Step4.BackColor = Color.Orange
                            Step5.BackColor = Color.DarkGray
                            Step6.BackColor = Color.DarkGray
                            Step7.BackColor = Color.DarkGray
                            Step8.BackColor = Color.DarkGray
                            Step9.BackColor = Color.DarkGray
                            state = 3
                            '設定timeofstep array 中的值
                            For i = 0 To eStep - 1
                                TimeofStep(0, i) = 5
                            Next
                            For j = sStep - 1 To eStep - 1
                                TimeofStep(0, j) = 3
                            Next
                            '1st 2nd 3rd 燈號變換
                            light_1st.BackColor = Color.LightGreen
                            light_2nd.BackColor = Color.Red
                            light_3rd.BackColor = Color.Red
                            MessageBox.Show("A3測量開始")
                            light_A3.BackColor = Color.LightGreen
                        End If
                    End If
                End If
            End If
            'A2改燈號
            If state = 2 And status = 1 And TimeofStep(1, 8) = 0 Then
                If light_1st.BackColor = Color.LightGreen Then
                    For index = sStep - 1 To eStep - 1
                        TimeofStep(1, index) = 3
                    Next
                    light_1st.BackColor = Color.DarkGreen
                    light_2nd.BackColor = Color.LightGreen
                    A2_step = 1
                ElseIf light_2nd.BackColor = Color.LightGreen Then
                    For index = sStep - 1 To eStep - 1
                        TimeofStep(1, index) = 3
                    Next
                    light_2nd.BackColor = Color.DarkGreen
                    light_3rd.BackColor = Color.LightGreen
                    A2_step = 1
                ElseIf light_3rd.BackColor = Color.LightGreen Then
                    TimeofStep(0, 0) = 0
                    light_3rd.BackColor = Color.DarkGreen
                    MessageBox.Show("測量結束")
                    light_A2.BackColor = Color.DarkGreen
                End If
            End If
            If state = 2 And status = 3 And TimeofStep(1, 2) = 0 Then

            End If
            'A3改燈號
            If state = 3 And status = 2 And TimeofStep(1, 3) = 0 Then

            End If
            If state = 3 And status = 3 And TimeofStep(1, 3) = 0 Then

            End If
            'A4改燈號
            If TimeofStep(0, 0) = 30 And TimeofStep(1, 0) = -2 And state = 4 Then
                If light_1st.BackColor = Color.Red Then
                    TimeofStep(1, 0) = -1
                    light_1st.BackColor = Color.DarkGreen
                    light_2nd.BackColor = Color.LightGreen
                ElseIf light_2nd.BackColor = Color.LightGreen Then
                    TimeofStep(1, 0) = -1
                    light_2nd.BackColor = Color.DarkGreen
                    light_3rd.BackColor = Color.LightGreen
                ElseIf light_3rd.BackColor = Color.LightGreen Then
                    TimeofStep(0, 0) = 0
                    light_3rd.BackColor = Color.DarkGreen
                    '這裡計算是否要加測or接收數據
                    'if 加測==>additional ==>加測完還要再計算
                    'else 接收數據==>A4亮綠燈
                End If
            End If
        End If

        Noise1.Text = Noise1.Text - 1
        Noise2.Text = Noise2.Text - 2
        Noise3.Text = Noise3.Text - 1

        MainBarGraph.Update({Noise1.Text - 1, Noise2.Text - 2, Noise3.Text - 1, 100, 100, 95})
        MainLineGraphs.CurrentLineGraph.Value.Update({Noise1.Text - 1, Noise2.Text - 2, Noise3.Text - 1, 100, 100, 95})
    End Sub

    Private Sub startButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles startButton.Click
        startButton.Enabled = False
        stopButton.Enabled = True
        continueButton.Enabled = False

        'case A1
        If TimeofStep(0, 0) = 30 And TimeofStep(1, 0) > 0 And state = 1 Then
            timeLeft = TimeofStep(0, 0)
            TimeofStep(1, 0) = TimeofStep(1, 0) - 1
            Step1.ForeColor = Color.Red
            Step2.Text = "循環剩餘次數 : " & TimeofStep(1, 0) & " 次"
            timeLabel.Text = timeLeft & " s"
            Timer1.Start()
        End If
        'case A2
        If state = 2 And status = 1 Then
            timeLeft = TimeofStep(0, A2_step - 1)
            If A2_step + 1 > sStep Then
                TimeofStep(1, A2_step - 1) = TimeofStep(1, A2_step - 1) - 1
            End If
            If A2_step + 1 > eStep Then
                array_step(A2_step - 1).ForeColor = Color.Red
                array_step(A2_step - 2).ForeColor = Color.Black
                A2_step = sStep - 1
            Else
                If A2_step - 1 > 0 Then
                    array_step(A2_step - 2).ForeColor = Color.Black
                End If
                array_step(A2_step - 1).ForeColor = Color.Red
            End If
            timeLabel.Text = timeLeft & " s"
            Timer1.Start()
        End If
        If state = 2 And status = 3 Then

        End If
        'case A3
        If state = 3 And status = 2 Then

        End If
        If state = 3 And status = 3 Then

        End If
        'case A4
        If TimeofStep(0, 0) = 30 And TimeofStep(1, 0) = -1 And state = 4 Then
            timeLeft = TimeofStep(0, 0)
            TimeofStep(1, 0) = TimeofStep(1, 0) - 1
            Step1.ForeColor = Color.Red
            timeLabel.Text = timeLeft & " s"
            Timer1.Start()
        End If
    End Sub

    Private Sub stopButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stopButton.Click
        startButton.Enabled = True
        stopButton.Enabled = False
        continueButton.Enabled = True
        Timer1.Stop()
    End Sub

    Private Sub continueButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles continueButton.Click
        startButton.Enabled = False
        stopButton.Enabled = True
        continueButton.Enabled = False
        Timer1.Start()
    End Sub


    Private Sub LinkLabel_preCal_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_preCal.LinkClicked
        If MessageBox.Show("前校正已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            light_preCal.BackColor = Color.DarkGreen
            light_BG.BackColor = Color.LightGreen
        End If
    End Sub
    Private Sub LinkLabel_BG_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_BG.LinkClicked
        If MessageBox.Show("背景噪音已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            light_BG.BackColor = Color.DarkGreen
            light_1st.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub LinkLabel_1st_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_1st.LinkClicked
        If MessageBox.Show("第一組已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            light_1st.BackColor = Color.DarkGreen
            light_2nd.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub LinkLabel_2nd_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_2nd.LinkClicked
        If MessageBox.Show("第二組已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            light_2nd.BackColor = Color.DarkGreen
            light_3rd.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub LinkLabel_3rd_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_3rd.LinkClicked
        If MessageBox.Show("第三組已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            light_3rd.BackColor = Color.DarkGreen
            light_Additional.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub LinkLabel_Additional_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_Additional.LinkClicked
        If MessageBox.Show("Additional已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            light_Additional.BackColor = Color.DarkGreen
            light_RSS.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub LinkLabel_RSS_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_RSS.LinkClicked
        If MessageBox.Show("背景噪音已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            light_RSS.BackColor = Color.DarkGreen
            light_postCal.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub LinkLabel_postCal_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_postCal.LinkClicked
        If MessageBox.Show("後校正已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            light_postCal.BackColor = Color.DarkGreen
        End If
    End Sub

    Private Sub LinkLabel_A1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_A1.LinkClicked
        If MessageBox.Show("A1所有數據已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            light_A1.BackColor = Color.DarkGreen
            If status = 1 Or status = 3 Then
                'if A1+A2===>excavator
                If status = 1 Then
                    '清空所有steps
                    Step1.Text = ""
                    Step2.Text = ""
                    Step3.Text = ""
                    Step4.Text = ""
                    Step5.Text = ""
                    Step6.Text = ""
                    Step7.Text = ""
                    Step8.Text = ""
                    Step9.Text = ""

                    '改手動輸入秒數
                    auto = False
                    '設定該循環的步驟
                    sStep = 2
                    eStep = 9

                    '載入步驟
                    Step1.Text = My.Resources.A2_Excavator_step1
                    Step1.BackColor = Color.MistyRose
                    Step2.Text = My.Resources.A2_Excavator_step2
                    Step2.BackColor = Color.Orange
                    Step3.Text = My.Resources.A2_Excavator_step3
                    Step3.BackColor = Color.Orange
                    Step4.Text = My.Resources.A2_Excavator_step4
                    Step4.BackColor = Color.Orange
                    Step5.Text = My.Resources.A2_Excavator_step5
                    Step5.BackColor = Color.Orange
                    Step6.Text = My.Resources.A2_Excavator_step6
                    Step6.BackColor = Color.Orange
                    Step7.Text = My.Resources.A2_Excavator_step7
                    Step7.BackColor = Color.Orange
                    Step8.Text = My.Resources.A2_Excavator_step8
                    Step8.BackColor = Color.Orange
                    Step9.Text = My.Resources.A2_Excavator_step9
                    Step9.BackColor = Color.Orange
                End If
                'if A1+A2+A3==>
                If status = 3 Then
                    '清空所有steps
                    Step1.Text = ""
                    Step2.Text = ""
                    Step3.Text = ""
                    Step4.Text = ""
                    Step5.Text = ""
                    Step6.Text = ""
                    Step7.Text = ""
                    Step8.Text = ""
                    Step9.Text = ""

                    '改手動輸入秒數
                    auto = False
                    '設定該循環的步驟
                    sStep = 2
                    eStep = 3

                    '載入步驟
                    Step1.Text = My.Resources.A2_Loader_step1
                    Step1.BackColor = Color.MistyRose
                    Step2.Text = My.Resources.A2_Loader_step2
                    Step2.BackColor = Color.Orange
                    Step3.Text = My.Resources.A2_Loader_step3
                    Step3.BackColor = Color.Orange
                    Step4.BackColor = Color.DarkGray
                    Step5.BackColor = Color.DarkGray
                    Step6.BackColor = Color.DarkGray
                    Step7.BackColor = Color.DarkGray
                    Step8.BackColor = Color.DarkGray
                    Step9.BackColor = Color.DarkGray
                End If
                MessageBox.Show("A2測量開始")
                light_A2.BackColor = Color.LightGreen
            End If
            'if A1+A3==>tractor
            If status = 2 Then
                '清空所有steps
                Step1.Text = ""
                Step2.Text = ""
                Step3.Text = ""
                Step4.Text = ""
                Step5.Text = ""
                Step6.Text = ""
                Step7.Text = ""
                Step8.Text = ""
                Step9.Text = ""

                '改手動輸入秒數
                auto = False
                '設定該循環的步驟
                sStep = 3
                eStep = 4

                '載入步驟
                Step1.Text = My.Resources.A3_Tractor_step1
                Step1.BackColor = Color.MistyRose
                Step2.Text = My.Resources.A3_Tractor_step2
                Step2.BackColor = Color.MistyRose
                Step3.Text = My.Resources.A3_Tractor_step3
                Step3.BackColor = Color.Orange
                Step4.Text = My.Resources.A3_Tractor_step4
                Step4.BackColor = Color.Orange
                Step5.BackColor = Color.DarkGray
                Step6.BackColor = Color.DarkGray
                Step7.BackColor = Color.DarkGray
                Step8.BackColor = Color.DarkGray
                Step9.BackColor = Color.DarkGray
                MessageBox.Show("A3測量開始")
                light_A3.BackColor = Color.LightGreen
            End If
        End If
    End Sub

    Private Sub LinkLabel_A2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_A2.LinkClicked
        If MessageBox.Show("A2所有數據已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            light_A2.BackColor = Color.DarkGreen
        End If
        If status = 3 Then
            '清空所有steps
            Step1.Text = ""
            Step2.Text = ""
            Step3.Text = ""
            Step4.Text = ""
            Step5.Text = ""
            Step6.Text = ""
            Step7.Text = ""
            Step8.Text = ""
            Step9.Text = ""

            '改手動輸入秒數
            auto = False
            '設定該循環的步驟
            sStep = 3
            eStep = 4

            '載入步驟
            Step1.Text = My.Resources.A3_Loader_step1
            Step1.BackColor = Color.MistyRose
            Step2.Text = My.Resources.A3_Loader_step2
            Step2.BackColor = Color.MistyRose
            Step3.Text = My.Resources.A3_Loader_step3
            Step3.BackColor = Color.Orange
            Step4.Text = My.Resources.A3_Loader_step4
            Step4.BackColor = Color.Orange
            Step5.BackColor = Color.DarkGray
            Step6.BackColor = Color.DarkGray
            Step7.BackColor = Color.DarkGray
            Step8.BackColor = Color.DarkGray
            Step9.BackColor = Color.DarkGray
            MessageBox.Show("A3測量開始")
            light_A3.BackColor = Color.LightGreen
        End If
        If status = 1 Then
            MessageBox.Show("此機具所有測量步驟已結束")
        End If
    End Sub

    Private Sub LinkLabel_A3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_A3.LinkClicked
        If MessageBox.Show("A3所有數據已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            light_A3.BackColor = Color.DarkGreen
            MessageBox.Show("此機具所有測量步驟已結束")
        End If
    End Sub

    Private Sub LinkLabel_A4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_A4.LinkClicked
        If MessageBox.Show("A4所有數據已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            light_A4.BackColor = Color.DarkGreen
            MessageBox.Show("此機具所有測量步驟已結束")
        End If
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

'helper classes for keeping track of labels and points
Public Class CoorPoint
    Public Label As Label
    Public Coors As ThreeDPoint
    Public Line As LineShape

    Public Sub New(ByVal lab As Label)
        Label = lab
        Coors = New ThreeDPoint
    End Sub

    Public Sub New(ByVal lab As Label, ByVal p As ThreeDPoint, ByVal l As LineShape)
        Label = lab
        Coors = p
        Line = l
    End Sub

End Class

Public Class ThreeDPoint
    Public Xc As Double
    Public Yc As Double
    Public Zc As Double
    Public Sub New()
    End Sub

    Public Sub New(ByVal x As Double, ByVal y As Double, ByVal z As Double)
        Xc = x
        Yc = y
        Zc = z
    End Sub

End Class