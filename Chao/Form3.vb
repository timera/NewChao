Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.PowerPacks
Imports System.IO
Imports System.Text
Imports System.Drawing.Printing
Imports System.Threading
Imports System.Net.Sockets
Imports System.Net
Imports System.Text.RegularExpressions

Partial Public Class Program
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MARGINS
        Public cxLeftWidth As Integer
        Public cxRightWidth As Integer
        Public cyTopHeight As Integer
        Public cyButtomHeight As Integer
    End Structure

    <DllImport("dwmapi.dll")> Public Shared Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As MARGINS) As Integer
    End Function

    Public _Warning As Boolean = True

    'coordinates for 6 points
    Dim pos() As CoorPoint
    'Dim pos(5, 2) As Double
    Dim posColors() As Color = {Color.DarkRed, Color.DarkOliveGreen, Color.DarkMagenta, Color.DeepPink, Color.DarkBlue, Color.Chocolate}

    'Origin for the coordinates system
    Dim origin(1) As Double
    Dim length As Double
    ' Set plotting vars
    Dim xAxis As New LineShape
    ' xCor(0) is the start point, xCor(1) is the end point
    Dim yAxis As New LineShape
    ' x y axis coordinates
    Dim xCor(1, 1) As Double
    Dim yCor(1, 1) As Double
    Dim ratio As Double
    Dim Points As ArrayList
    Dim GP As Graphics
    Dim R As Integer
    Dim PlotRatio As Double



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

    Public A4_step_text As String

    'Graphs for the main screen
    Dim MainLineGraph As LineGraph
    Dim MainBarGraph As BarGraph

    Public timeLeft As Integer

    'Runs throughout
    Public HeadRun As Run_Unit
    Public CurRun As Run_Unit

    'Steps on the Right
    Dim HeadStep As Steps
    Public CurStep As Steps
    Dim array_step(9) As Label
    Public array_step_s(8) As TextBox
    Dim array_light(8) As Label
    Public array_step_display(8) As Label

    Public Countdown As Boolean = False

    'trial represents 1st or 2nd or 3rd
    Dim trial As Integer

    Dim NoisesArray(6) As Label

    Dim sum_steps As Integer

    Dim choice As String
    Dim TimerTesting As Boolean 'if set to true, it's currently timer testing
    'temporary constant
    Const seconds As Integer = 3
    Const graphW As Integer = 650
    Const graphH As Integer = 200
    Dim ASize As Size = New Size(50, 447)
    'used for first time clicking start button
    Dim NewGraph As Boolean

    'doing the test for input s of steps
    'Dim Test_S As Boolean = False
    Dim Last_timeLeft As Integer
    Dim Index_for_Setup_Time As Integer

    'dim PreCal, BG, RSS, PostCal jump pointer
    Public PreCal_1st As Run_Unit
    Public PreCal_2nd As Run_Unit
    Public PreCal_3rd As Run_Unit
    Public PreCal_4th As Run_Unit
    Public PreCal_5th As Run_Unit
    Public PreCal_6th As Run_Unit
    Public BG As Run_Unit
    Public RSS As Run_Unit
    Public PostCal_1st As Run_Unit
    Public PostCal_2nd As Run_Unit
    Public PostCal_3rd As Run_Unit
    Public PostCal_4th As Run_Unit
    Public PostCal_5th As Run_Unit
    Public PostCal_6th As Run_Unit

    'dim ExA1 and ExA2 jump pointer
    Public ExA1_Fst As Run_Unit
    Public ExA1_Sec As Run_Unit
    Public ExA1_Thd As Run_Unit
    Public ExA1_Add As Run_Unit

    Public ExA2_Fst As Run_Unit
    Public ExA2_Sec As Run_Unit
    Public ExA2_Thd As Run_Unit
    Public ExA2_Add As Run_Unit

    'dim LoA1 and LoA2 and LoA3 jump pointer
    Public LoA1_Fst As Run_Unit
    Public LoA1_Sec As Run_Unit
    Public LoA1_Thd As Run_Unit
    Public LoA1_Add As Run_Unit

    Public LoA2_Fst As Run_Unit
    Public LoA2_Sec As Run_Unit
    Public LoA2_Thd As Run_Unit
    Public LoA2_Add As Run_Unit

    Public LoA3_Fst_fwd As Run_Unit
    Public LoA3_Fst_bkd As Run_Unit
    Public LoA3_Sec_fwd As Run_Unit
    Public LoA3_Sec_bkd As Run_Unit
    Public LoA3_Thd_fwd As Run_Unit
    Public LoA3_Thd_bkd As Run_Unit
    Public LoA3_Add_fwd As Run_Unit
    Public LoA3_Add_bkd As Run_Unit

    'dim TrA1 and TrA3 jump pointer
    Public TrA1_Fst As Run_Unit
    Public TrA1_Sec As Run_Unit
    Public TrA1_Thd As Run_Unit
    Public TrA1_Add As Run_Unit

    Public TrA3_Fst_fwd As Run_Unit
    Public TrA3_Fst_bkd As Run_Unit
    Public TrA3_Sec_fwd As Run_Unit
    Public TrA3_Sec_bkd As Run_Unit
    Public TrA3_Thd_fwd As Run_Unit
    Public TrA3_Thd_bkd As Run_Unit
    Public TrA3_Add_fwd As Run_Unit
    Public TrA3_Add_bkd As Run_Unit

    'dim A4 jump pointer
    Public A4_Fst As Run_Unit
    Public A4_Sec_Mid As Run_Unit
    Public A4_Sec As Run_Unit
    Public A4_Thd_Mid As Run_Unit
    Public A4_Thd As Run_Unit
    Public A4_Add_Mid As Run_Unit
    Public A4_Add As Run_Unit

    'save currun
    Public Null_CurRun As Run_Unit
    Public Temp_CurRun As Run_Unit
    Public Temp_Countdown As Boolean

    Dim Result

    Public array_time(8) As Integer

    Public Comm As Communication

    'if it is have additional test => true
    Dim Add_Test_Record As Boolean

    Dim ADictionary As New Dictionary(Of String, String)

    Dim A1Time As Integer = 5

    Private PhoneSocket As Socket
    Public sim As Boolean = True
    Private DataGrid As Grid
    Private WithEvents BasicInfoGrid As DataGridView
    Private A_Unit_Size As Size = New Size(1250, 450)

    Private BasicInfoGridFields() As String

    Public Sub New()

        ' 此為設計工具所需的呼叫。
        InitializeComponent()
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        ' 在 InitializeComponent() 呼叫之後加入任何初始設定。

    End Sub


    ' 基本程式開時要準備的東西
    Private Sub Program_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '##A4 Dictionary 給選機具以後可以查他的簡寫，以便WIFI傳輸 for Android
        ADictionary.Add("瀝青混凝土舖築機(Asphalt finisher)", "A4-Asp_Fin")
        ADictionary.Add("鑽土機(Earth drill)", "A4-Aug_Dri_Dri")
        ADictionary.Add("全套管鑽掘機", "A4-Aug_Dri_Dri")
        ADictionary.Add("土壤取樣器(地鑽) (Earth auger)", "A4-Aug_Dri_Dri")
        ADictionary.Add("油壓式拔樁機", "A4-Aug_Dri_Dri")
        ADictionary.Add("油壓式打樁機(Hydraulic pile driver)", "A4-Aug_Dri_Dri")
        ADictionary.Add("拔樁機", "A4-Aug_Dri_Dri")
        ADictionary.Add("鑽岩機(Rock breaker)", "A4-Aug_Dri_Dri")
        ADictionary.Add("空氣壓縮機(Compressor)", "A4-Com")
        ADictionary.Add("混凝土破碎機(Concrete breaker)", "A4-Con_Bre")
        ADictionary.Add("混凝土割切機(Concrete cutter)", "A4-Con_Cut")
        ADictionary.Add("混凝土泵車(Concrete pump)", "A4-Con_Pum")
        ADictionary.Add("履帶起重機(Crawler crane)", "A4-Cra")
        ADictionary.Add("卡車起重機(Truck crane)", "A4-Cra")
        ADictionary.Add("輪形起重機(Wheel crane)", "A4-Cra")
        ADictionary.Add("發電機(Generator)", "A4-Gen")
        ADictionary.Add("鐵輪壓路機(Road roller)", "A4-Rol")
        ADictionary.Add("膠輪壓路機(Wheel roller)", "A4-Rol")
        ADictionary.Add("振動式壓路機(Vibrating roller)", "A4-Rol")
        ADictionary.Add("振動式樁錘(Vibrating hammer)", "A4-Vib_Ham")
        'other machines
        ADictionary.Add("開挖機(Excavator)", "ex")
        ADictionary.Add("推土機(Crawler and wheel tractor)", "Tr")
        ADictionary.Add("裝料機(Crawler and wheel loader)", "Lo")
        ADictionary.Add("裝料開挖機", "loex")

        '加到主選單中
        Dim machListforComboBox As New List(Of String)
        machListforComboBox.Add("A1+A2")
        machListforComboBox.Add("開挖機(Excavator)")
        machListforComboBox.Add("A1+A3")
        machListforComboBox.Add("推土機(Crawler and wheel tractor)")
        machListforComboBox.Add("A1+A2+A3")
        machListforComboBox.Add("裝料機(Crawler and wheel loader)")
        machListforComboBox.Add("A1+A2 A1+A2+A3")
        machListforComboBox.Add("裝料開挖機")
        machListforComboBox.Add("A4")
        For Each key As String In ADictionary.Keys
            If ADictionary(key).Contains("A4") Then
                machListforComboBox.Add(key)
            End If
        Next
        ComboBox_machine_list.Items.AddRange(machListforComboBox.ToArray())

        BasicInfoGridFields = {
            "測試編號：",
            "測試日期：",
            "機具名稱：",
            "廠牌：",
            "型號：",
            "規格：",
            "附屬配備：",
            "減音裝置：",
            "馬力：",
            "最大轉速：",
            "基本尺寸：",
            "量測位置（量測點及其高度、聲音感應器高度等、與施工機具音源相對位置）：",
            "測試環境反射面的實際描述：",
            "周圍地形的位置、周圍之情況（周圍之建築物、地形、地貌、防音設施等）："
        }
        

        '##COMMUNICATION with Noise Meters
        Comm = New Communication()
        SimulationMode()

        Last_timeLeft = 0
        Index_for_Setup_Time = 0
        TimerTesting = False
        Add_Test_Record = False

        'Set up the window
        Dim margins As MARGINS = New MARGINS
        margins.cxLeftWidth = 2
        margins.cxRightWidth = 2
        margins.cyTopHeight = 23
        margins.cyButtomHeight = 2
        Dim hwnd As IntPtr = Handle
        Dim result As Integer = DwmExtendFrameIntoClientArea(hwnd, margins)

        TabControl1.Location = New Point(2, 102)
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
        GroupBox_Plot.BackColor = Color.Transparent

        BasicInfoGrid = New DataGridView()
        BasicInfoGrid.Parent = TabPage1

        Me.TabPageTimer.Enabled = False

        '##Set up Buttons
        startButton.Enabled = True
        stopButton.Enabled = False
        Button_Skip_Add.Enabled = False
        Test_NextButton.Enabled = False
        Test_StartButton.Enabled = True
        Test_ConfirmButton.Enabled = False
        ConnectButton.Enabled = True
        DisconnButton.Enabled = False
        SaveToolStripMenuItem.Enabled = False





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

        '###SIZE of OBJECTS
        Me.Size = New Size(1200, 700)
        startButton.Size = New Size(75, 50)
        Accept_Button.Size = New Size(75, 25)
        stopButton.Size = New Size(75, 50)
        Button_Skip_Add.Size = New Size(75, 23)
        TabControl2.Size = New Size(400, 540)
        Label_Ex_A1.Size = New Size(29, 16)
        Label_Ex_A2.Size = New Size(29, 16)
        Label_Lo_A1.Size = New Size(29, 16)
        Label_Lo_A2.Size = New Size(29, 16)
        Label_Lo_A3.Size = New Size(29, 16)
        Label_Tr_A1.Size = New Size(29, 16)
        Label_Tr_A3.Size = New Size(29, 16)
        Label_A4.Size = New Size(29, 16)
        Panel_PreCal.Size = New Size(82, 40)
        LinkLabel_preCal.Size = New Size(56, 15)
        LinkLabel_postCal.Size = New Size(56, 15)
        Panel_PreCal_Sub.Size = New Size(80, 166)
        Panel_PostCal_Sub.Size = New Size(80, 166)
        Panel_PreCal.Size = New Size(80, 40)
        Panel_PostCal.Size = New Size(80, 40)
        Label1.Size = New Size(17, 16)
        Label2.Size = New Size(17, 16)
        Label3.Size = New Size(17, 16)
        Label4.Size = New Size(34, 15)
        Test_Input_S_Panel.Size = New Size(140, 362)
        Test_StartButton.Size = New Size(75, 70)
        Test_StopButton.Size = New Size(75, 70)
        Test_NextButton.Size = New Size(75, 40)
        Test_ConfirmButton.Size = New Size(75, 60)
        S1.Size = New Size(12, 12)
        S2.Size = New Size(12, 12)
        S3.Size = New Size(12, 12)
        S4.Size = New Size(12, 12)
        S5.Size = New Size(12, 12)
        S6.Size = New Size(12, 12)
        S7.Size = New Size(12, 12)
        S8.Size = New Size(12, 12)
        S9.Size = New Size(12, 12)
        timeLabel.Size = New Size(120, 60)
        Panel_Setting_Bargraph_Max_Min.Size = New Size(300, 225)
        Setting_Bargraph_Title.Size = New Size(293, 43)
        Setting_Bargraph_Max.Size = New Size(73, 33)
        Setting_Bargraph_Min.Size = New Size(73, 33)
        TextBox_Setting_Bargraph_Max.Size = New Size(142, 20)
        TextBox_Setting_Bargraph_Min.Size = New Size(142, 20)
        Button_Setting_Bargraph.Size = New Size(75, 50)

        PanelMobile.Size = New Size(300, 100)
        ButtonConnectMobile.Size = New Size(75, 23)
        LabelMobileIP.Size = New Size(93, 15)

        '###LOCATION of OBJECTS

        startButton.Location = New Point(129, 0)
        Accept_Button.Location = New Point(219, 0)
        Button_Skip_Add.Location = New Point(219, 27)
        Accept_Button.Enabled = False
        timeLabel.Location = New Point(5, 5)
        stopButton.Location = New Point(309, 0)
        Test_StartButton.Location = New Point(60, 20)
        Test_StopButton.Location = New Point(60, 100)
        Test_NextButton.Location = New Point(60, 180)
        Test_ConfirmButton.Location = New Point(60, 322)
        ButtonExport.Location = New Point(20, 500)
        Test_Input_S_Panel.Location = New Point(155, 20)
        TabControl2.Location = New Point(timeLabel.Left, timeLabel.Top + timeLabel.Height + 3)

        NoisesArray = {Noise1, Noise2, Noise3, Noise4, Noise5, Noise6, Noise_Avg}
        Dim meterFigWidth = 47
        Dim meterX = 868
        Dim metery = 560
        Noise1.Location = New Point(meterX, metery)
        Noise2.Location = New Point(meterX + meterFigWidth, metery)
        Noise3.Location = New Point(meterX + meterFigWidth * 2, metery)
        Noise4.Location = New Point(meterX + meterFigWidth * 3, metery)
        Noise5.Location = New Point(meterX + meterFigWidth * 4, metery)
        Noise6.Location = New Point(meterX + meterFigWidth * 5, metery)
        Noise_Avg.Location = New Point(meterX + meterFigWidth * 6, metery)
        For i = 0 To 6
            NoisesArray(i).Parent = TabPage2
            NoisesArray(i).Visible = False
            If Not i = 6 Then
                NoisesArray(i).Size = New Size(46, 39)
            Else
                NoisesArray(i).Size = New Size(60, 39)
            End If
        Next
        Dim stepX As Integer = 450
        Dim stepY As Integer = 55
        Dim stepY_start As Integer = 93
        Step1.Location = New Point(stepX, stepY_start)
        Step2.Location = New Point(stepX, stepY_start + stepY * 1)
        Step3.Location = New Point(stepX, stepY_start + stepY * 2)
        Step4.Location = New Point(stepX, stepY_start + stepY * 3)
        Step5.Location = New Point(stepX, stepY_start + stepY * 4)
        Step6.Location = New Point(stepX, stepY_start + stepY * 5)
        Step7.Location = New Point(stepX, stepY_start + stepY * 6)
        Step8.Location = New Point(stepX, stepY_start + stepY * 7)
        Step9.Location = New Point(stepX, stepY_start + stepY * 8)
        Step10.Location = New Point(stepX, stepY_start + stepY * 9)
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

        For i = 0 To 9
            array_step(i).Size = New Size(305, 52)
        Next

        stepX = 410
        Label_step1_second.Location = New Point(stepX, stepY_start)
        Label_step2_second.Location = New Point(stepX, stepY_start + stepY * 1)
        Label_step3_second.Location = New Point(stepX, stepY_start + stepY * 2)
        Label_step4_second.Location = New Point(stepX, stepY_start + stepY * 3)
        Label_step5_second.Location = New Point(stepX, stepY_start + stepY * 4)
        Label_step6_second.Location = New Point(stepX, stepY_start + stepY * 5)
        Label_step7_second.Location = New Point(stepX, stepY_start + stepY * 6)
        Label_step8_second.Location = New Point(stepX, stepY_start + stepY * 7)
        Label_step9_second.Location = New Point(stepX, stepY_start + stepY * 8)

        array_light(0) = Light1
        array_light(1) = Light2
        array_light(2) = Light3
        array_light(3) = Light4
        array_light(4) = Light5
        array_light(5) = Light6
        array_light(6) = Light7
        array_light(7) = Light8
        array_light(8) = Light9
        For i = 0 To 8
            array_light(i).Location = New Point(310, (i * 40) + 30)
        Next

        array_step_s(0) = Input_S_Step1
        array_step_s(1) = Input_S_Step2
        array_step_s(2) = Input_S_Step3
        array_step_s(3) = Input_S_Step4
        array_step_s(4) = Input_S_Step5
        array_step_s(5) = Input_S_Step6
        array_step_s(6) = Input_S_Step7
        array_step_s(7) = Input_S_Step8
        array_step_s(8) = Input_S_Step9
        For i = 0 To 8
            array_step_s(i).Text = 0
            array_step_s(i).Enabled = False
            array_step_s(i).Size = New Size(85, 22)
            array_step_s(i).Location = New Point(20, (i * 40) + 10)
        Next

        S1.Location = New Point(110, 15)
        S2.Location = New Point(110, 55)
        S3.Location = New Point(110, 95)
        S4.Location = New Point(110, 135)
        S5.Location = New Point(110, 175)
        S6.Location = New Point(110, 215)
        S7.Location = New Point(110, 255)
        S8.Location = New Point(110, 295)
        S9.Location = New Point(110, 335)

        array_step_display(0) = Label_step1_second
        array_step_display(1) = Label_step2_second
        array_step_display(2) = Label_step3_second
        array_step_display(3) = Label_step4_second
        array_step_display(4) = Label_step5_second
        array_step_display(5) = Label_step6_second
        array_step_display(6) = Label_step7_second
        array_step_display(7) = Label_step8_second
        array_step_display(8) = Label_step9_second
        For i = 0 To 9
            array_step_display(0).Size = New Size(35, 52)
        Next

        Label1.Location = New Point(103, 120)
        Label2.Location = New Point(103, 228)
        Label3.Location = New Point(103, 336)
        Label4.Location = New Point(100, 444)


        Panel_PreCal.Location = New Point(10, 30)
        LinkLabel_preCal.Location = New Point(3, 13)
        Panel_PreCal_Sub.Location = New Point(10, 60)
        Panel_Bkg.Location = New Point(10, 245)
        Panel_RSS.Location = New Point(10, 273)
        Panel_PostCal.Location = New Point(10, 305)
        Panel_PostCal_Sub.Location = New Point(10, 335)
        LinkLabel_postCal.Location = New Point(3, 13)

        Panel_PreCal_1st.Location = New Point(0, 0)
        Panel_PreCal_2nd.Location = New Point(0, 28)
        Panel_PreCal_3rd.Location = New Point(0, 56)
        Panel_PreCal_4th.Location = New Point(0, 84)
        Panel_PreCal_5th.Location = New Point(0, 112)
        Panel_PreCal_6th.Location = New Point(0, 140)

        Panel_PostCal_1st.Location = New Point(0, 0)
        Panel_PostCal_2nd.Location = New Point(0, 28)
        Panel_PostCal_3rd.Location = New Point(0, 56)
        Panel_PostCal_4th.Location = New Point(0, 84)
        Panel_PostCal_5th.Location = New Point(0, 112)
        Panel_PostCal_6th.Location = New Point(0, 140)

        Panel_ExA1_Fst_1st.Location = New Point(0, 16)
        Panel_ExA1_Sec_1st.Location = New Point(0, 124)
        Panel_ExA1_Thd_1st.Location = New Point(0, 232)
        Panel_ExA1_Add_1st.Location = New Point(0, 340)

        Panel_LoA1_Fst_1st.Location = New Point(0, 16)
        Panel_LoA1_Sec_1st.Location = New Point(0, 124)
        Panel_LoA1_Thd_1st.Location = New Point(0, 232)
        Panel_LoA1_Add_1st.Location = New Point(0, 340)

        Panel_TrA1_Fst_1st.Location = New Point(0, 16)
        Panel_TrA1_Sec_1st.Location = New Point(0, 124)
        Panel_TrA1_Thd_1st.Location = New Point(0, 232)
        Panel_TrA1_Add_1st.Location = New Point(0, 340)

        Panel_ExA2_Fst_1st.Location = New Point(2, 16)
        Panel_ExA2_Fst_2nd.Location = New Point(2, 52)
        Panel_ExA2_Fst_3rd.Location = New Point(2, 88)
        Panel_ExA2_Sec_1st.Location = New Point(2, 124)
        Panel_ExA2_Sec_2nd.Location = New Point(2, 160)
        Panel_ExA2_Sec_3rd.Location = New Point(2, 196)
        Panel_ExA2_Thd_1st.Location = New Point(2, 232)
        Panel_ExA2_Thd_2nd.Location = New Point(2, 268)
        Panel_ExA2_Thd_3rd.Location = New Point(2, 304)
        Panel_ExA2_Add_1st.Location = New Point(2, 340)
        Panel_ExA2_Add_2nd.Location = New Point(2, 376)
        Panel_ExA2_Add_3rd.Location = New Point(2, 412)

        Panel_LoA2_Fst_1st.Location = New Point(2, 16)
        Panel_LoA2_Fst_2nd.Location = New Point(2, 52)
        Panel_LoA2_Fst_3rd.Location = New Point(2, 88)
        Panel_LoA2_Sec_1st.Location = New Point(2, 124)
        Panel_LoA2_Sec_2nd.Location = New Point(2, 160)
        Panel_LoA2_Sec_3rd.Location = New Point(2, 196)
        Panel_LoA2_Thd_1st.Location = New Point(2, 232)
        Panel_LoA2_Thd_2nd.Location = New Point(2, 268)
        Panel_LoA2_Thd_3rd.Location = New Point(2, 304)
        Panel_LoA2_Add_1st.Location = New Point(2, 340)
        Panel_LoA2_Add_2nd.Location = New Point(2, 376)
        Panel_LoA2_Add_3rd.Location = New Point(2, 412)

        Panel_LoA3_Fst_fwd.Location = New Point(2, 16)
        Panel_LoA3_Fst_bkd.Location = New Point(2, 70)
        Panel_LoA3_Sec_fwd.Location = New Point(2, 124)
        Panel_LoA3_Sec_bkd.Location = New Point(2, 178)
        Panel_LoA3_Thd_fwd.Location = New Point(2, 232)
        Panel_LoA3_Thd_bkd.Location = New Point(2, 286)
        Panel_LoA3_Add_fwd.Location = New Point(2, 340)
        Panel_LoA3_Add_bkd.Location = New Point(2, 394)

        Panel_TrA3_Fst_fwd.Location = New Point(2, 16)
        Panel_TrA3_Fst_bkd.Location = New Point(2, 70)
        Panel_TrA3_Sec_fwd.Location = New Point(2, 124)
        Panel_TrA3_Sec_bkd.Location = New Point(2, 178)
        Panel_TrA3_Thd_fwd.Location = New Point(2, 232)
        Panel_TrA3_Thd_bkd.Location = New Point(2, 286)
        Panel_TrA3_Add_fwd.Location = New Point(2, 340)
        Panel_TrA3_Add_bkd.Location = New Point(2, 394)

        Panel_A4_Fst.Location = New Point(2, 16)
        Panel_A4_Sec_Mid_Background.Location = New Point(2, 124)
        Panel_A4_Sec.Location = New Point(2, 178)
        Panel_A4_Thd_Mid_Background.Location = New Point(2, 232)
        Panel_A4_Thd.Location = New Point(2, 286)
        Panel_A4_Add_Mid_Background.Location = New Point(2, 340)
        Panel_A4_Add.Location = New Point(2, 394)

        Panel_Setting_Bargraph_Max_Min.Location = New Point(17, 20)
        Setting_Bargraph_Title.Location = New Point(3, 11)
        Setting_Bargraph_Max.Location = New Point(3, 83)
        Setting_Bargraph_Min.Location = New Point(3, 135)
        TextBox_Setting_Bargraph_Max.Location = New Point(93, 83)
        TextBox_Setting_Bargraph_Min.Location = New Point(93, 135)
        Button_Setting_Bargraph.Location = New Point(215, 160)

        PanelMobile.Location = New Point(Panel_Setting_Bargraph_Max_Min.Location.X, Panel_Setting_Bargraph_Max_Min.Location.Y + 300)
        LabelMobileIP.Location = New Point(5, 5)
        MaskedTextBoxIPAddress.Location = New Point(LabelMobileIP.Location.X + LabelMobileIP.Width + 10, LabelMobileIP.Location.Y)
        ButtonConnectMobile.Location = New Point(MaskedTextBoxIPAddress.Location.X + MaskedTextBoxIPAddress.Width + 10, MaskedTextBoxIPAddress.Location.Y)


        Button_change_machine.Enabled = False

        MachChosen = False

        Null_CurRun = New Run_Unit(LinkLabel_Temp, Panel_Temp, Nothing, Nothing, 0, "Temp", 0, 0, 0)
        Temp_CurRun = Null_CurRun

        '''SETTINGS TAB
        Button_Setting_Bargraph.Enabled = False

        'all size and location in machine choose 
        ComboBox_machine_list.Location = New Point(10, 10)
        ComboBox_machine_list.Size = New Size(280, 20)
        Button_change_machine.Location = New Point(ComboBox_machine_list.Location.X + ComboBox_machine_list.Width, 8)
        Button_change_machine.Size = New Size(80, 27)
        Label_machine_pic.Location = New Point(10, 40)
        Label_machine_pic.Size = New Size(74, 21)
        Label_machine_pic.Visible = False
        Picture_machine.Location = New Point(10, 40)
        Picture_machine.Size = New Size(280, 190)


        Dim GroupBoxWidth As Integer = 280
        Dim GroupBoxX As Integer = 10
        GroupBox_A1_A2_A3.Location = New Point(GroupBoxX, Picture_machine.Location.Y + Picture_machine.Height + 8)
        GroupBox_A1_A2_A3.Size = New Size(GroupBoxWidth, 80)

        Label_input_L.Location = New Point(6, 18)
        Label_input_L.Size = New Size(75, 26)
        Label_r1.Location = New Point(46, 46)
        Label_r1.Size = New Size(30, 24)
        TextBox_L.Location = New Point(105, 22)
        TextBox_L.Size = New Size(99, 22)
        TextBox_r1.Location = New Point(105, 52)
        TextBox_r1.Size = New Size(99, 22)
        Button_L_check.Location = New Point(210, 20)
        Button_L_check.Size = New Size(50, 30)

        GroupBox_A4.Location = New Point(GroupBoxX, GroupBox_A1_A2_A3.Location.Y + GroupBox_A1_A2_A3.Height + 8)
        GroupBox_A4.Size = New Size(GroupBoxWidth, 143)

        Label_input_L1.Location = New Point(6, 18)
        Label_input_L1.Size = New Size(88, 26)
        Label_input_L2.Location = New Point(6, 47)
        Label_input_L2.Size = New Size(88, 26)
        Label_input_L3.Location = New Point(6, 75)
        Label_input_L3.Size = New Size(88, 26)
        Label_r2.Location = New Point(57, 105)
        Label_r2.Size = New Size(30, 26)
        TextBox_L1.Location = New Point(105, 20)
        TextBox_L1.Size = New Size(99, 22)
        TextBox_L2.Location = New Point(105, 50)
        TextBox_L2.Size = New Size(99, 22)
        TextBox_L3.Location = New Point(105, 80)
        TextBox_L3.Size = New Size(99, 22)
        TextBox_r2.Location = New Point(105, 110)
        TextBox_r2.Size = New Size(99, 22)
        Button_L1_L2_L3_check.Location = New Point(210, 80)
        Button_L1_L2_L3_check.Size = New Size(50, 30)

        GroupBox1.Location = New Point(GroupBoxX, GroupBox_A4.Location.Y + GroupBox_A4.Height + 8)
        GroupBox1.Size = New Size(GroupBoxWidth, 125)

        PanelMeterSetup.Location = New Point(16, 21)
        PanelMeterSetup.Size = New Size(252, 68)
        ComboBoxComs.Location = New Point(9, 3)
        ComboBoxComs.Size = New Size(168, 20)
        ButtonComRefresh.Location = New Point(183, 3)
        ButtonComRefresh.Size = New Size(64, 23)
        ConnectButton.Location = New Point(9, 30)
        ConnectButton.Size = New Size(110, 37)
        DisconnButton.Location = New Point(125, 30)
        DisconnButton.Size = New Size(110, 37)
        ButtonSim.Location = New Point(25, 94)
        ButtonSim.Size = New Size(110, 23)
        ButtonMeters.Location = New Point(141, 94)
        ButtonMeters.Size = New Size(110, 23)

        'A4 Hints
        Dim A4HintHeight As Integer = Picture_machine.Location.Y
        Label_A4_Hint1.Location = New Point(GroupBoxX + Picture_machine.Width + 10, A4HintHeight)
        Label_A4_Hint1.Size = New Size(32, 273)
        Label_A4_Hint2.Location = New Point(Label_A4_Hint1.Location.X + Label_A4_Hint1.Width + 10, A4HintHeight)
        Label_A4_Hint2.Size = New Size(35, 164)

        GroupBox_Plot.Location = New Point(365, 26)
        GroupBox_Plot.Size = New Size(576, 576)
        GroupBox_Plot.Visible = True
        GP = GroupBox_Plot.CreateGraphics()

        BasicInfoGrid.Location = New Point(GroupBox_Plot.Location.X + GroupBox_Plot.Size.Width + 20, GroupBox_Plot.Location.Y + 8)
        BasicInfoGrid.Size = New Size(270, GroupBox_Plot.Size.Height - 8)

        'COOR PLOT
        'x
        origin(0) = GroupBox_Plot.Width / 2
        'y
        origin(1) = GroupBox_Plot.Height / 2
        length = 275
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


        plotCor(GroupBox_Plot.CreateGraphics, xCor, yCor)

        Picture_machine.BorderStyle = BorderStyle.None

        BasicInfoGrid.BorderStyle = BorderStyle.None
        BasicInfoGrid.RowHeadersVisible = False
        BasicInfoGrid.Columns.Add("Data", "")
        BasicInfoGrid.Rows.Add(26)
        BasicInfoGrid.ColumnHeadersVisible = False
        For i = 0 To 12
            BasicInfoGrid.Rows(i * 2).Cells(0).Value = BasicInfoGridFields(i)
            BasicInfoGrid.Rows(i * 2).Cells(0).ReadOnly = True
            BasicInfoGrid.Rows(i * 2).Cells(0).Style.BackColor = Color.Aqua

        Next
        BasicInfoGrid.AllowUserToAddRows = False
        BasicInfoGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill


    End Sub

    '程式要關之前要做的事
    Private Sub Program_Close(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If SaveToolStripMenuItem.Enabled Then
            Dim answer As DialogResult = MessageBox.Show("尚未儲存資料，是否儲存?", "Save?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If answer = DialogResult.Yes Then
                If Save() Then
                    SaveToolStripMenuItem.Enabled = False
                Else
                    e.Cancel = True
                End If
            ElseIf answer = DialogResult.Cancel Then
                e.Cancel = True
            End If
        End If
    End Sub

    'prevents user from clicking on tab before choosing machinery
    Private Sub TabControl1_IndexChanged(ByVal sender As TabControl, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If Not MachChosen Then
            If sender.SelectedIndex = 1 Or sender.SelectedIndex = 2 Then
                MessageBox.Show("還未選機具!", "My application", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                TabControl1.SelectedIndex = 0 ' go back to the machinery choose stage
            End If
        End If
    End Sub




    '機具選單中變換
    Private Sub ComboBox_machine_list_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_machine_list.SelectedIndexChanged

        'steps text 清空
        Clear_Steps()

        choice = ComboBox_machine_list.Text
        Dim decided As String = Decide_Machine()

        'mach區分要輸入至哪個 groupbox for plotting the coordinates
        If Not decided = "" Then
            If Not Machine = Machines.Others Then 'A123
                GroupBox_A1_A2_A3.Enabled = True
                TextBox_L.Enabled = True
                Button_L_check.Enabled = True
                GroupBox_A4.Enabled = False
            Else
                GroupBox_A1_A2_A3.Enabled = False 'A4
                GroupBox_A4.Enabled = True
                TextBox_L1.Enabled = True
                TextBox_L2.Enabled = True
                TextBox_L3.Enabled = True
                TextBox_r2.Enabled = True
                TextBox_r2.ReadOnly = False
                Button_L1_L2_L3_check.Enabled = True
            End If
        Else
            GroupBox_A1_A2_A3.Enabled = False
            GroupBox_A4.Enabled = False
        End If
    End Sub


    '開始計時候每秒跳要做的事
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Countdown = False Then '如果是正數
            timeLeft = timeLeft + 1
            timeLabel.Text = timeLeft & " s"

            'precal,postcal
            If timeLeft = 1 And (CurRun.Name.Contains("PreCal") Or CurRun.Name.Contains("PostCal")) Then
                stopButton.Enabled = True
            ElseIf timeLeft = 2 Then            'A4,RSS,Background
                stopButton.Enabled = True
            End If

            'send values to display as text and graphs
            SetScreenValuesAndGraphs(GetInstantData())


        Else '倒數
            stopButton.Enabled = True
            If timeLeft > 0 Then 'counting
                If Not timeLeft = 1 Then
                    timeLeft = timeLeft - 1
                    timeLabel.Text = timeLeft & " s"

                    'send values to display as text and graphs
                    SetScreenValuesAndGraphs(GetInstantData())
                ElseIf timeLeft = 1 Then 'time's up\

                    timeLeft = timeLeft - 1
                    timeLabel.Text = timeLeft & " s"
                    'send values to display as text and graphs

                    SetScreenValuesAndGraphs(GetInstantData())

                    'if HasNextStep
                    'change step color , seconds 
                    'do next step
                    If CurRun.CurStep = CurRun.EndStep Then 'last step (not HasNextStep)
                        'stop the timer

                        Timer1.Stop()
                        'Continuing on for A2's unfinished runs
                        If CurRun.NextUnit.Name.Contains("_2nd_3rd") Then
                            Comm.PauseMeasure(True)
                            Accept_Button.Enabled = False
                            startButton.Enabled = True
                            stopButton.Enabled = False
                            array_step(CurRun.CurStep - 1).BackColor = Color.Green
                            Set_Panel_BackColor()
                            CurRun.Executed = True
                            MoveToRun(CurRun.NextUnit)
                            Set_Run_Unit()

                            'load A2's steps
                            Clear_Steps()
                            Load_Steps()

                            'dispose old graph and create new graph
                            'A2三個round的圖要連續所以不load新的圖
                            'Load_New_Graph_CD_True()

                        Else 'Countdown finally stop for entire run
                            If CurRun.Name.Contains("_2nd_3rd") Then
                                Comm.PauseMeasure(False)
                            End If
                            Comm.StopMeasure()
                            Accept_Button.Enabled = True
                            Button_Skip_Add.Enabled = False
                            startButton.Enabled = False
                            stopButton.Enabled = False
                            array_step(CurRun.CurStep - 1).BackColor = Color.Green
                            updateFinalBarGraph()
                            ShowResultsOnForm()
                        End If

                        'Cur_A_Unit.AddGrid_Run_Unit()
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
        End If

        'if connected to Android phone then send signal to phone
        SendNonChangeSignalToMobile()
    End Sub

    '##BUTTON CLICKS

    '變更機具按鈕按下
    Private Sub Button_change_machine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_change_machine.Click
        Change_Machine()
    End Sub



    '除了A4的機具按下確認鍵
    Private Sub Button_L_check_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_L_check.Click
        If String.IsNullOrWhiteSpace(TextBox_L.Text) Then
            Return
        End If
        A123_Prepare()

    End Sub


    'A4機具按下確認鍵
    Private Sub Button_L1_L2_L3_check_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_L1_L2_L3_check.Click
        If String.IsNullOrWhiteSpace(TextBox_L1.Text) Or String.IsNullOrWhiteSpace(TextBox_L2.Text) Or String.IsNullOrWhiteSpace(TextBox_L3.Text) Then
            Return
        End If
        A4_Prepare()
    End Sub

    ' 開始測量鈕按下
    Private Sub startButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles startButton.Click
        If Comm.StartMeasure() Then
            All_Panel_Disable()
            stopButton.Focus()
            If Countdown = False Then
                startButton.Enabled = False
                Accept_Button.Enabled = False
                stopButton.Enabled = False
                'If CurRun.Name.Contains("A4") Or CurRun.Name.Contains("RSS") Or CurRun.Name.Contains("Background") Then
                '    stopButton.Enabled = False
                'Else
                '    stopButton.Enabled = True
                'End If

                Timer1.Start()
            Else
                startButton.Enabled = False
                stopButton.Enabled = True
                Timer1.Start()
            End If
        End If
    End Sub

    '停止鍵按下
    Private Sub stopButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stopButton.Click
        Timer1.Stop()
        If Countdown = False Then '正數
            'startButton.Enabled = True
            Accept_Button.Enabled = True
            Button_Skip_Add.Enabled = False
            stopButton.Enabled = False

            'final Leq
            updateFinalBarGraph()
            'Tell meters to stop measuring
            Comm.StopMeasure()
            ShowResultsOnForm()

            ThrowPrePostCalWarnings()

        Else '倒數
            'final Leq
            updateFinalBarGraph()
            'Tell meters to stop measuring
            Comm.StopMeasure()

            startButton.Enabled = True
            stopButton.Enabled = False
            All_Panel_Enable()

            Dim isJumpStep As Boolean = False
            If Temp_CurRun Is Nothing Then
                isJumpStep = True
            ElseIf Not Temp_CurRun.Link.Name = "LinkLabel_Temp" Then
                isJumpStep = True
            End If

            If Not isJumpStep Then 'if not jump step
                '這些是倒數的步驟，當遇到按STOP，往前跳一格
                If CurRun.Name.Contains("A1") Or CurRun.Name = "A4" Or CurRun.Name = "A4_Add" Or CurRun.Name.Contains("A2_1st") Or CurRun.Name.Contains("A3") Then
                    'dispose data

                    'reset run_unit
                    Set_Run_Unit()
                    'load A1's steps
                    Clear_Steps()
                    Load_Steps()
                    'dispose old graph and create new graph
                    Load_New_Graph_CD_True()
                    'bounce back to previous step
                ElseIf CurRun.Name.Contains("A2_2nd_3rd") Then 'A2要往回跳兩到三格
                    If CurRun.PrevUnit.Name.Contains("A2_1st") Then
                        'dispose data

                        'reset run_unit
                        Restart_from_1st_Previous_Run_Unit()
                        'load LoA3_fwd or TrA3_fwd's steps
                        Clear_Steps()
                        Load_Steps()
                        'dispose old graph and create new graph
                        Load_New_Graph_CD_True()
                    Else 'jump back two steps
                        'dispose data

                        'reset run_unit
                        Restart_from_2nd_Previous_Run_Unit()
                        'load LoA3_fwd or TrA3_fwd's steps
                        Clear_Steps()
                        Load_Steps()
                        'dispose old graph and create new graph
                        Load_New_Graph_CD_True()
                    End If
                End If
            Else 'if jump step
                If CurRun.Name.Contains("A2_1st") Then
                    CurRun.NextUnit.Set_BackColor(Color.Green)
                    CurRun.NextUnit.NextUnit.Set_BackColor(Color.Green)
                ElseIf CurRun.Name.Contains("A2_2nd_3rd") Then
                    If CurRun.NextUnit.Name.Contains("_2nd_3rd") Then
                        CurRun.NextUnit.Set_BackColor(Color.Green)
                    End If
                End If
                'set the jump run green and the one we wanna jump back to yellow
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
                If CurRun Is Nothing Then
                    BackToEnd()
                End If
            End If
        End If
    End Sub


    '測時間的開始健按下
    Private Sub Test_StartButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Test_StartButton.Click

        Load_New_Graph_CD_True()
        Test_NextButton.Enabled = False
        Test_StopButton.Enabled = True
        Test_StartButton.Enabled = False
        Timer1.Start()
    End Sub

    '測時間的停止鍵按下
    Private Sub Test_StopButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Test_StopButton.Click
        Test_NextButton.Enabled = True
        Test_StopButton.Enabled = False
        Test_StartButton.Enabled = True
        Timer1.Stop()
        CurRun.Steps.Time = timeLeft
        array_step_s(Index_for_Setup_Time).Text = CurRun.Steps.Time
        array_step_s(Index_for_Setup_Time).Enabled = True
        timeLeft = 0

        If CurRun.CurStep = CurRun.EndStep Then 'last step (not HasNextStep)
            Test_ConfirmButton.Enabled = True
            Test_NextButton.Enabled = False
        End If
    End Sub

    '測時間的下一步鍵按下
    Private Sub Test_NextButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Test_NextButton.Click
        timeLabel.Text = timeLeft & " s"
        Test_NextButton.Enabled = False
        If CurRun.CurStep = CurRun.EndStep Then 'last step (not HasNextStep)
            array_step(CurRun.CurStep - 1).BackColor = Color.Green
        Else 'HasNextStep
            Set_Step_BackColor()
        End If
        Index_for_Setup_Time += 1
        CurRun.Steps = CurRun.Steps.NextStep 'jump to next step
        CurRun.CurStep += 1 'curstep add 1
        array_step_s(Index_for_Setup_Time).Focus()
    End Sub

    '測時間的CONFIRM鍵按下
    Private Sub Test_ConfirmButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Test_ConfirmButton.Click
        Dim Input_S_Apply As Boolean = True

        CurStep = Load_Steps_helper(CurRun)
        If Not (CurRun.Name = "TrA3_bkd" Or CurRun.Name = "LoA3_bkd") Then
            For i = 0 To CurRun.EndStep - 1
                If Not CurStep.Time = -1 Then
                    If array_step_s(i).Text = "" Or array_step_s(i).Text = "0" Then
                        MessageBox.Show("Step" + (i + 1).ToString + "尚未輸入時間")
                        Input_S_Apply = False
                        Test_ConfirmButton.Enabled = True
                    End If
                    CurStep = CurStep.NextStep
                Else
                    CurStep = CurStep.NextStep
                End If
            Next
        ElseIf array_step_s(CurRun.PrevUnit.EndStep).Text = "" Or array_step_s(CurRun.PrevUnit.EndStep).Text = "0" Then
            MessageBox.Show("Step" + (CurRun.PrevUnit.EndStep).ToString + "尚未輸入時間")
            Input_S_Apply = False
            Test_ConfirmButton.Enabled = True
        End If

        If Input_S_Apply = True Then
            LoadInputTime(1, Nothing)

            If CurRun.NextUnit.Name = "LoA3_bkd" Or CurRun.NextUnit.Name = "TrA3_bkd" Then
                'back to initial test condition
                For i = 0 To CurRun.EndStep - 1
                    array_step_s(i).Enabled = False
                Next
                Test_StartButton.Enabled = True
                Test_NextButton.Enabled = False
                Test_ConfirmButton.Enabled = False
                ' change light
                Set_Panel_BackColor()
                Index_for_Setup_Time += 1
                'jump to next Run_Unit
                MoveToRun(CurRun.NextUnit)
                Set_Run_Unit()
                timeLeft = 0

                'load A1's steps
                Clear_Steps()
                Load_Steps()
                array_light(0).BackColor = Color.Green
                array_light(2).BackColor = Color.Green
                array_light(3).BackColor = Color.Yellow

                'dispose old graph and create new graph
                Load_New_Graph_CD_True()
            Else
                TimerTesting = False
                SetTimerTabOn(TimerTesting)
                Countdown = True
                Index_for_Setup_Time = 0

                If CurRun.Name = "LoA3_bkd" Or CurRun.Name = "TrA3_bkd" Then
                    Restart_from_1st_Previous_Run_Unit()
                Else
                    'don't jump to next Run_Unit
                    Set_Run_Unit()
                End If

                'load steps
                Clear_Steps()
                Load_Steps()

                'dispose old graph and create new graph
                Load_New_Graph_CD_True()
            End If
        End If
        Set_Second_for_Steps()
        SendNonChangeSignalToMobile()
    End Sub

    '測完數據會需要按的接受鍵
    Private Sub AcceptButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Accept_Button.Click

        Accept_Button.Enabled = False
        startButton.Enabled = True

        Dim keyGRU As Grid_Run_Unit = CurRun.GRU
        While Not IsNothing(keyGRU.NextGRU) 'for more than one additional runs
            keyGRU = keyGRU.NextGRU
        End While

        Result = MessageBox.Show("此步驟數據已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        Dim isJumpStep As Boolean = False
        If Temp_CurRun Is Nothing Then
            isJumpStep = True 'because it jumps from END
        ElseIf Not Temp_CurRun.Link.Name = "LinkLabel_Temp" Then
            isJumpStep = True
        End If

        If isJumpStep Then ' jump step
            If Result = DialogResult.Yes Then
                'save currun data
                keyGRU.Accept()
                EnableSave()
                CurRun.Executed = True
                If CurRun.Name.Contains("bkd") Then
                    keyGRU.OverallGRU.Accept()
                End If
                startButton.Focus()
            End If
            If CurRun.Name.Contains("PreCal") Or CurRun.Name = "Background" Or CurRun.Name = "RSS" Or IsNothing(CurRun.NextUnit) Or CurRun.Name.Contains("PostCal") Or CurRun.Name = "A4_Mid_BG" Or CurRun.Name = "A4_Mid_BG_Add" Then
                Jump_Back_Countdown_False(Result)
            Else
                Jump_Back_Countdown_True(Result)
            End If
            If CurRun Is Nothing Then
                BackToEnd()
            End If
        Else
            If Result = DialogResult.Yes Then
                keyGRU.Accept()
                EnableSave()
                CurRun.Executed = True
                If CurRun.Name.Contains("bkd") Then
                    keyGRU.OverallGRU.Accept()
                End If
                MoveOnToNextRun(True, True, False)
                startButton.Focus()
            ElseIf Result = DialogResult.No Then
                If CurRun.Name.Contains("Add") Then
                    Button_Skip_Add.Enabled = True
                End If

                If CurRun.Name = "ExA2_1st" Or CurRun.Name = "ExA2_2nd_3rd" Or CurRun.Name = "LoA2_1st" Or CurRun.Name = "LoA2_2nd_3rd" Then
                    All_Panel_Enable()

                    'reset run_unit
                    Restart_from_2nd_Previous_Run_Unit()
                    'load steps
                    Clear_Steps()
                    Load_Steps()
                    'dispose old graph and create new graph
                    Load_New_Graph_CD_True()
                ElseIf CurRun.Name = "ExA2_1st_Add" Or CurRun.Name = "ExA2_2nd_3rd_Add" Or CurRun.Name = "LoA2_1st_Add" Or CurRun.Name = "LoA2_2nd_3rd_Add" Then
                    All_Panel_Enable()
                    'dispose data

                    'reset run_unit
                    Restart_from_2nd_Previous_Run_Unit()
                    'load steps
                    Clear_Steps()
                    Load_Steps()

                    'dispose old graph and create new graph
                    Load_New_Graph_CD_True()
                Else
                    Accept_No()
                End If
                keyGRU.Deny()
                DataGrid.ClearColumn(keyGRU)
                If CurRun.Name.Contains("bkd") Then
                    keyGRU.OverallGRU.Deny()
                    DataGrid.ClearColumn(keyGRU.OverallGRU)
                End If
            ElseIf Result = DialogResult.Cancel Then
                Accept_Cancel()
            End If
        End If
        If CurRun IsNot Nothing Then
            SendNonChangeSignalToMobile()
        End If
    End Sub


    '畫出機具座標
    Private Sub GroupBox_Plot_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles GroupBox_Plot.Paint
        If R = 0 Or pos Is Nothing Then
            Return
        End If
        plot(True, e.Graphics, xCor, yCor, GroupBox_Plot)
    End Sub


    '連結到噪音計的鍵按下
    Private Sub ConnectButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConnectButton.Click
        If Comm.Open() Then
            If Comm.SetupServer() Then
                ConnectButton.Enabled = False
                DisconnButton.Enabled = True

            Else
                MsgBox("Cannot Set up Server Correctly!")
            End If
        End If
    End Sub

    '中斷到噪音計的連結
    Private Sub DisconnButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DisconnButton.Click
        If Comm.Close() Then
            DisconnButton.Enabled = False
            ConnectButton.Enabled = True
        End If
    End Sub

    '存檔鍵按下
    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        If Save() Then
            SaveToolStripMenuItem.Enabled = False
        End If
    End Sub

    '模擬鍵按下
    Private Sub ButtonSim_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSim.Click
        SimulationMode()
    End Sub

    '正常噪音計鍵按下
    Private Sub ButtonMeters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMeters.Click
        MetersMode()
    End Sub

    '重新跟噪音計連結鍵按下
    Private Sub ButtonComRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonComRefresh.Click
        ComboBoxComs.Items.Clear()
        For Each com In Communication1_1.GetComs()
            ComboBoxComs.Items.Add(com)
        Next
    End Sub

    'Bargraph設定確定鍵按下
    Private Sub Button_Setting_Bargraph_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Setting_Bargraph.Click
        If String.IsNullOrWhiteSpace(TextBox_Setting_Bargraph_Max.Text) Then
            MessageBox.Show("請輸入最大值!")
            Return
        ElseIf String.IsNullOrWhiteSpace(TextBox_Setting_Bargraph_Min.Text) Then
            MessageBox.Show("請輸入最小值!")
            Return
        ElseIf Convert.ToInt32(TextBox_Setting_Bargraph_Min.Text) > Convert.ToInt32(TextBox_Setting_Bargraph_Max.Text) Then
            MessageBox.Show("最大值不能小於最小值!")
            Return
        End If
        MainBarGraph.chart.ChartAreas(0).AxisY.Crossing = TextBox_Setting_Bargraph_Min.Text
        MainBarGraph.chart.ChartAreas(0).AxisY.Maximum = TextBox_Setting_Bargraph_Max.Text
        MainBarGraph.chart.ChartAreas(0).AxisY.Minimum = TextBox_Setting_Bargraph_Min.Text
        Button_Setting_Bargraph.Enabled = False
    End Sub

    'Bargraph設定輸入格有輸入
    Private Sub TextBox_Setting_Bargraph_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Setting_Bargraph_Max.TextChanged, TextBox_Setting_Bargraph_Min.TextChanged
        Button_Setting_Bargraph.Enabled = True
    End Sub

    '跳過加測鍵按下
    Private Sub Button_Skip_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Skip_Add.Click
        Timer1.Stop()
        Button_Skip_Add.Enabled = False
        startButton.Enabled = True
        stopButton.Enabled = False
        Accept_Button.Enabled = False

        'for change run_unit back_color
        If Add_Test_Record = True Then
            If CurRun.Name.Contains("A1") Then
                CurRun.Set_BackColor(Color.Green)
            ElseIf CurRun.Name.Contains("A3_fwd_Add") Or CurRun.Name = "A4_Mid_BG_Add" Then
                CurRun.Set_BackColor(Color.Green)
                CurRun.NextUnit.Set_BackColor(Color.Green)
            ElseIf CurRun.Name.Contains("A3_bkd_Add") Or CurRun.Name = "A4_Add" Then
                CurRun.PrevUnit.Set_BackColor(Color.Green)
                CurRun.Set_BackColor(Color.Green)
            ElseIf CurRun.Name.Contains("A2_1st") Then
                CurRun.Set_BackColor(Color.Green)
                CurRun.NextUnit.Set_BackColor(Color.Green)
                CurRun.NextUnit.NextUnit.Set_BackColor(Color.Green)
            ElseIf CurRun.Name.Contains("A2_2nd") And CurRun.PrevUnit.Name.Contains("A2_1st") Then
                CurRun.PrevUnit.Set_BackColor(Color.Green)
                CurRun.Set_BackColor(Color.Green)
                CurRun.NextUnit.Set_BackColor(Color.Green)
            ElseIf CurRun.Name.Contains("A2_2nd") And CurRun.PrevUnit.Name.Contains("A2_2nd") Then
                CurRun.PrevUnit.PrevUnit.Set_BackColor(Color.Green)
                CurRun.PrevUnit.Set_BackColor(Color.Green)
                CurRun.Set_BackColor(Color.Green)
            End If
            Add_Test_Record = False
        Else
            If CurRun.Name.Contains("A1") Then
                CurRun.Set_BackColor(Color.IndianRed)
            ElseIf CurRun.Name.Contains("A3_fwd_Add") Or CurRun.Name = "A4_Mid_BG_Add" Then
                CurRun.Set_BackColor(Color.IndianRed)
                CurRun.NextUnit.Set_BackColor(Color.IndianRed)
            ElseIf CurRun.Name.Contains("A3_bkd_Add") Or CurRun.Name = "A4_Add" Then
                CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
                CurRun.Set_BackColor(Color.IndianRed)
            ElseIf CurRun.Name.Contains("A2_1st") Then
                CurRun.Set_BackColor(Color.IndianRed)
                CurRun.NextUnit.Set_BackColor(Color.IndianRed)
                CurRun.NextUnit.NextUnit.Set_BackColor(Color.IndianRed)
            ElseIf CurRun.Name.Contains("A2_2nd") And CurRun.PrevUnit.Name.Contains("A2_1st") Then
                CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
                CurRun.Set_BackColor(Color.IndianRed)
                CurRun.NextUnit.Set_BackColor(Color.IndianRed)
            ElseIf CurRun.Name.Contains("A2_2nd") And CurRun.PrevUnit.Name.Contains("A2_2nd") Then
                CurRun.PrevUnit.PrevUnit.Set_BackColor(Color.IndianRed)
                CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
                CurRun.Set_BackColor(Color.IndianRed)
            End If
        End If

        If CurRun.Name.Contains("A4_Add") Or CurRun.Name.Contains("A1") Or CurRun.Name.Contains("A3_bkd") Or CurRun.NextUnit.Name = "RSS" Or CurRun.NextUnit.Name = "LoA1" Or CurRun.NextUnit.Name = "LoA3_fwd" Then
            'skip 1 unit (include self)
            MoveToRun(CurRun.NextUnit)
        ElseIf CurRun.Name.Contains("A3_fwd") Or CurRun.PrevUnit.Name.Contains("A2_1st_Add") Or CurRun.Name = "A4_Mid_BG_Add" Then
            'skip 2 unit (include self)
            MoveToRun(CurRun.NextUnit.NextUnit)
        ElseIf CurRun.Name.Contains("A2_1st_Add") Then
            'skip 3 unit (include self)
            MoveToRun(CurRun.NextUnit.NextUnit.NextUnit)
        End If


        If CurRun.Name = "RSS" Then
            CurRun.Steps = Load_Steps_helper(CurRun)
            timeLeft = CurRun.Time
            timeLabel.Text = timeLeft & " s"
            Countdown = False
            Load_New_Graph_CD_False()
            Clear_Steps()
            For index = 0 To 8
                array_step_display(index).Text = ""
            Next
        Else
            CurRun.Set_BackColor(Color.Yellow)
            If CurRun.Name.Contains("A2") Or CurRun.Name.Contains("A3") Then
                Reset_Test_Time()
            Else
                Set_Run_Unit()
            End If
            Clear_Steps()
            Load_Steps()
            'dispose old graph and create new graph
            Load_New_Graph_CD_True()
        End If
    End Sub


    '表格輸出按下
    Private Sub ButtonExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExport.Click
        If ExportToCSV() Then
            ButtonExport.Enabled = False
        End If
    End Sub

    '開檔鍵按下
    Private Sub LoadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadToolStripMenuItem.Click
        LoadFile()
    End Sub

    '機具基本資料改變(助於存檔鍵的enable)
    Private BasicInfoDataChangedFromLast As Boolean = False
    Public Sub BasicInfoDataChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BasicInfoGrid.CellValueChanged
        BasicInfoDataChangedFromLast = True
        EnableSave()
    End Sub

    '列印鍵按下
    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim pd As PrintDocument = New PrintDocument()
        AddHandler pd.PrintPage, AddressOf Me.PrintPage
        Try
            pd.Print()
            pd.Dispose()
            pd = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    '連到ANDROID鍵按下
    Private Sub ButtonConnectMobile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonConnectMobile.Click
        Dim ip As String = MaskedTextBoxIPAddress.Text
        ip = ip.Replace("_", "")
        Dim s As Socket = CommunicationWithAndroid.ConnectSocket(ip, "60000")
        If s IsNot Nothing Then
            PhoneSocket = s
            LabelConnectMobile.Text = "Connected"
        Else
            LabelConnectMobile.Text = "Not Connected"
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
End Class
