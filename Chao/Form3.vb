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
    Dim pos() As CoorPoint
    'Dim pos(5, 2) As Double
    Dim posColors() As Color = {Color.Red, Color.Orange, Color.Yellow, Color.Green, Color.Blue, Color.Cyan}

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

    Dim sum_steps As Integer = 0

    Dim choice As String
    Dim TimerTesting As Boolean = False 'if set to true, it's currently timer testing
    'temporary constant
    Const seconds As Integer = 3
    Const graphW As Integer = 650
    Const graphH As Integer = 200
    Dim ASize As Size = New Size(50, 447)
    'used for first time clicking start button
    Dim NewGraph As Boolean

    'doing the test for input s of steps
    'Dim Test_S As Boolean = False
    Dim Last_timeLeft As Integer = 0
    Dim Index_for_Setup_Time = 0
    Dim Test_Change_Second As Boolean
    Dim Original_Second_Dispose As Boolean = False

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

    Public array_ExA2_time(8) As Integer
    Public array_LoA2_time(2) As Integer
    Public array_LoA3_time(3) As Integer
    Public array_TrA3_time(3) As Integer

    Public Comm As Communication

    Public Sub New()

        ' 此為設計工具所需的呼叫。
        InitializeComponent()
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        ' 在 InitializeComponent() 呼叫之後加入任何初始設定。

    End Sub


    ' things that need to be instantiated at startup
    Private Sub Program_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '##COMMUNICATION
        Comm = New Communication()
        SimulationMode()

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
        SaveButton.Enabled = False





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
        Me.Size = New Size(1150, 700)
        startButton.Size = New Size(75, 50)
        Accept_Button.Size = New Size(75, 25)
        stopButton.Size = New Size(75, 50)
        Button_Skip_Add.Size = New Size(75, 23)
        TabControl2.Size = New Size(400, 535)
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
        timeLabel.Size = New Size(80, 80)
        Panel_Setting_Bargraph_Max_Min.Size = New Size(300, 225)
        Setting_Bargraph_Title.Size = New Size(293, 43)
        Setting_Bargraph_Max.Size = New Size(73, 33)
        Setting_Bargraph_Min.Size = New Size(73, 33)
        TextBox_Setting_Bargraph_Max.Size = New Size(142, 20)
        TextBox_Setting_Bargraph_Min.Size = New Size(142, 20)
        Button_Setting_Bargraph.Size = New Size(75, 50)


        '###LOCATION of OBJECTS

        'LinkLabel_preCal.Location = New Point(50, 22)
        'LinkLabel_BG.Location = New Point(50, 62)
        'LinkLabel_RSS.Location = New Point(50, 262)
        'LinkLabel_postCal.Location = New Point(50, 302)

        'Timer Tab
        'Dim input_step_gap As Integer = 40
        'Dim input_step_x As Integer = 18
        'Dim input_step_size As Size = New Size(85, 22)
        'Dim input_light_size As Size = New Size(100, 30)

        'Input_S_Step1.Location = New Point(input_step_x, 23)
        'Input_S_Step1.Size = input_step_size

        'Light1.Parent = Test_Input_S_Panel
        'Light1.Location = New Point(input_step_x - 5, 23 - 5)
        'Light1.Size = input_light_size


        'Input_S_Step2.Location = New Point(input_step_x, 23 + input_step_gap)
        'Input_S_Step2.Size = input_step_size

        'Light2.Parent = Test_Input_S_Panel
        'Light2.Location = New Point(input_step_x - 5, 23 + input_step_gap - 5)
        'Light2.Size = input_light_size

        'Input_S_Step3.Location = New Point(input_step_x, 23 + input_step_gap * 2)
        'Input_S_Step3.Size = input_step_size

        'Light3.Parent = Test_Input_S_Panel
        'Light3.Location = New Point(input_step_x - 5, 23 + input_step_gap * 2 - 5)
        'Light3.Size = input_light_size

        'Input_S_Step4.Location = New Point(input_step_x, 23 + input_step_gap * 3)
        'Input_S_Step4.Size = input_step_size

        'Light4.Parent = Test_Input_S_Panel
        'Light4.Location = New Point(input_step_x - 5, 23 + input_step_gap * 3 - 5)
        'Light4.Size = input_light_size

        'Input_S_Step5.Location = New Point(input_step_x, 23 + input_step_gap * 4)
        'Input_S_Step5.Size = input_step_size

        'Light5.Parent = Test_Input_S_Panel
        'Light5.Location = New Point(input_step_x - 5, 23 + input_step_gap * 4 - 5)
        'Light5.Size = input_light_size

        'Input_S_Step6.Location = New Point(input_step_x, 23 + input_step_gap * 5)
        'Input_S_Step6.Size = input_step_size

        'Light6.Parent = Test_Input_S_Panel
        'Light6.Location = New Point(input_step_x - 5, 23 + input_step_gap * 5 - 5)
        'Light6.Size = input_light_size

        'Input_S_Step7.Location = New Point(input_step_x, 23 + input_step_gap * 6)
        'Input_S_Step7.Size = input_step_size

        'Light7.Parent = Test_Input_S_Panel
        'Light7.Location = New Point(input_step_x - 5, 23 + input_step_gap * 6 - 5)
        'Light7.Size = input_light_size

        'Input_S_Step8.Location = New Point(input_step_x, 23 + input_step_gap * 7)
        'Input_S_Step8.Size = input_step_size

        'Light8.Parent = Test_Input_S_Panel
        'Light8.Location = New Point(input_step_x - 5, 23 + input_step_gap * 7 - 5)
        'Light8.Size = input_light_size

        'Input_S_Step9.Location = New Point(input_step_x, 23 + input_step_gap * 8)
        'Input_S_Step9.Size = input_step_size

        'Light9.Parent = Test_Input_S_Panel
        'Light9.Location = New Point(input_step_x - 5, 23 + input_step_gap * 8 - 5)
        'Light9.Size = input_light_size


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
        Test_Input_S_Panel.Location = New Point(155, 20)
        TabControl2.Location = New Point(timeLabel.Left, timeLabel.Top + timeLabel.Height + 3)

        NoisesArray = {Noise1, Noise2, Noise3, Noise4, Noise5, Noise6, Noise_Avg}
        Dim meterFigWidth = 47
        Dim meterX = 790
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




        Button_change_machine.Enabled = False

        MachChosen = False

        Null_CurRun = New Run_Unit(LinkLabel_Temp, Panel_Temp, Nothing, Nothing, 0, "Temp", 0, 0, 0)
        Temp_CurRun = Null_CurRun

        '''SETTINGS TAB
        Button_Setting_Bargraph.Enabled = False

        'all size and location in machine choose 
        ComboBox_machine_list.Location = New Point(10, 10)
        ComboBox_machine_list.Size = New Size(280, 20)
        Button_change_machine.Location = New Point(300, 10)
        Button_change_machine.Size = New Size(90, 30)
        Label_machine_pic.Location = New Point(10, 40)
        Label_machine_pic.Size = New Size(74, 21)
        Picture_machine.Location = New Point(10, 65)
        Picture_machine.Size = New Size(320, 190)

        GroupBox_A1_A2_A3.Location = New Point(10, 260)
        GroupBox_A1_A2_A3.Size = New Size(275, 80)
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

        GroupBox_A4.Location = New Point(10, 350)
        GroupBox_A4.Size = New Size(275, 145)
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

        GroupBox1.Location = New Point(10, 500)
        GroupBox1.Size = New Size(275, 145)
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

        GroupBox_Plot.Location = New Point(365, 40)
        GroupBox_Plot.Size = New Size(580, 580)
        GroupBox_Plot.Visible = True
        GP = GroupBox_Plot.CreateGraphics()
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
        'draw coordinates
        'canvas.Parent = GroupBox_Plot
        'canvas.Location = New System.Drawing.Point(0, 0)

        plotCor(GroupBox_Plot.CreateGraphics, xCor, yCor)

        RichTextBox_Info.Size = New Size(200, GroupBox_Plot.Height)
        RichTextBox_Info.Location = New Point(GroupBox_Plot.Left + GroupBox_Plot.Width + 10, GroupBox_Plot.Top)

        Label_A4_Hint1.Location = New Point(288, 328)
        Label_A4_Hint1.Size = New Size(32, 273)
        Label_A4_Hint2.Location = New Point(323, 328)
        Label_A4_Hint2.Size = New Size(35, 164)
    End Sub

    Private Sub Program_Close(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If SaveButton.Enabled Then
            Dim answer As DialogResult = MessageBox.Show("尚未儲存資料，是否儲存?", "Save?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If answer = DialogResult.Yes Then
                If DataGrid.Save() Then
                    SaveButton.Enabled = False
                Else
                    e.Cancel = True
                End If
            ElseIf answer = DialogResult.Cancel Then
                e.Cancel = True
            End If
        End If
    End Sub

    'prevents user from clicking on tab b4 choosing machinery
    Private Sub TabControl1_IndexChanged(ByVal sender As TabControl, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If Not MachChosen Then
            If sender.SelectedIndex = 1 Or sender.SelectedIndex = 2 Then
                MessageBox.Show("還未選機具!", "My application", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                TabControl1.SelectedIndex = 0 ' go back to the machinery choose stage
            End If
        End If
    End Sub

    Sub SetTimerTabOn(ByVal index As Boolean)
        'Tabcontrol_Changed(index=true)=>jump to tabpagetimer(enable) and tabpage3(disable),use for input seconds
        'Tabcontrol_Changed(index=false)=>jump to tabpageProcedure(enable) and tabpage4(disable),use for button confirm
        Dim tabpage As TabPage
        CurStep = Load_Steps_helper(CurRun)
        If index = True Then
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
        ElseIf index = False Then
            Me.TabControl2.SelectedIndex = 0
            For Each tabpage In TabControl2.TabPages
                If tabpage.Name = "TabPageTimer" Then
                    tabpage.Enabled = False
                ElseIf tabpage.Name = "TabPageProcedure" Then
                    tabpage.Enabled = True
                End If
            Next
            For i = 0 To 8
                array_light(i).BackColor = Color.DarkGray
            Next
        End If
    End Sub

    Private Sub Button_change_machine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_change_machine.Click
        If SaveButton.Enabled Then
            Dim answer As DialogResult = MessageBox.Show("尚未儲存資料，是否儲存?", "Save?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If answer = DialogResult.Yes Then
                If DataGrid.Save() Then
                    SaveButton.Enabled = False
                Else
                    Return
                End If
            ElseIf answer = DialogResult.Cancel Then
                Return
            End If
            'can choose machine again
            ComboBox_machine_list.Enabled = True
            Button_change_machine.Enabled = False
            MachChosen = False
            Machine = Nothing
            choice = Nothing

            'dispose graph
            MainLineGraph.Dispose()
            MainBarGraph.Dispose()

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

            For i = 0 To 8
                array_step_s(i).Text = 0
                array_step_s(i).Enabled = False
            Next

            A4_step_text = Nothing

            Label_A4_Hint1.Visible = False
            Label_A4_Hint2.Visible = False
        End If

        'If MessageBox.Show("是否放棄目前測量數據?", "My application", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
        'End If
    End Sub

    Private Sub ComboBox_machine_list_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_machine_list.SelectedIndexChanged

        'steps text 清空
        Clear_Steps()

        Dim inner_choice As String = ComboBox_machine_list.Text

        If inner_choice = "開挖機(Excavator)" Then
            Picture_machine.Image = My.Resources.小型開挖機_compact_excavator_
            Machine = Machines.Excavator
        ElseIf inner_choice = "推土機(Crawler and wheel tractor)" Then  'A1+A3
            Picture_machine.Image = My.Resources.履帶式推土機_crawler_dozer_
            Machine = Machines.Tractor
        ElseIf inner_choice = "鐵輪壓路機(Road roller)" Then
            Picture_machine.Image = My.Resources.壓路機_rollers_
            Machine = Machines.Others
        ElseIf inner_choice = "膠輪壓路機(Wheel roller)" Then
            Picture_machine.Image = My.Resources.壓路機_rollers_
            Machine = Machines.Others
        ElseIf inner_choice = "振動式壓路機(Vibrating roller)" Then
            Picture_machine.Image = My.Resources.壓路機_rollers_
            Machine = Machines.Others
        ElseIf inner_choice = "裝料機(Crawler and wheel loader)" Then
            Picture_machine.Image = My.Resources.履帶式裝料機_crawler_loader_
            Machine = Machines.Loader
        ElseIf inner_choice = "裝料開挖機" Then
            Picture_machine.Image = My.Resources.履帶式開挖裝料機_crawler_backhoe_loader_
            Machine = Machines.Loader_Excavator
        ElseIf inner_choice = "履帶起重機(Crawler crane)" Then
            Picture_machine.Image = My.Resources.履帶起重機
            Machine = Machines.Others
        ElseIf inner_choice = "卡車起重機(Truck crane)" Then
            Picture_machine.Image = My.Resources.卡車起重機
            Machine = Machines.Others
        ElseIf inner_choice = "輪形起重機(Wheel crane)" Then
            Picture_machine.Image = My.Resources.輪式起重機
            Machine = Machines.Others
        ElseIf inner_choice = "振動式樁錘(Vibrating hammer)" Then
            Picture_machine.Image = My.Resources.vibrating_hammer
            Machine = Machines.Others
        ElseIf inner_choice = "油壓式打樁機(Hydraulic pile driver)" Then
            Picture_machine.Image = My.Resources.油壓式打樁機_Hydraulic_pile_driver_
            Machine = Machines.Others
        ElseIf inner_choice = "拔樁機" Then
            Picture_machine.Image = My.Resources.拔樁機
            Machine = Machines.Others
        ElseIf inner_choice = "油壓式拔樁機" Then
            Picture_machine.Image = My.Resources.油壓式打樁機_Hydraulic_pile_driver_
            Machine = Machines.Others
        ElseIf inner_choice = "土壤取樣器(地鑽) (Earth auger)" Then
            Picture_machine.Image = My.Resources.土壤取樣器_地鑽___Earth_auger_
            Machine = Machines.Others
        ElseIf inner_choice = "全套管鑽掘機" Then
            Picture_machine.Image = My.Resources.全套管鑽掘機
            Machine = Machines.Others
        ElseIf inner_choice = "鑽土機(Earth drill)" Then
            Picture_machine.Image = My.Resources.鑽土機_Earth_drill_
            Machine = Machines.Others
        ElseIf inner_choice = "鑽岩機(Rock breaker)" Then
            Picture_machine.Image = My.Resources.鑽岩機_Rock_breaker_
            Machine = Machines.Others
        ElseIf inner_choice = "混凝土泵車(Concrete pump)" Then
            Picture_machine.Image = My.Resources.混凝土泵車_Concrete_pump_
            Machine = Machines.Others
        ElseIf inner_choice = "混凝土破碎機(Concrete breaker)" Then
            Picture_machine.Image = My.Resources.混凝土破碎機_Concrete_breaker_
            Machine = Machines.Others
        ElseIf inner_choice = "瀝青混凝土舖築機(Asphalt finisher)" Then
            Picture_machine.Image = My.Resources.瀝青混凝土舖築機_Asphalt_finisher_
            Machine = Machines.Others
        ElseIf inner_choice = "混凝土割切機(Concrete cutter)" Then
            Picture_machine.Image = My.Resources.混凝土割切機_Concrete_cutter_
            Machine = Machines.Others
        ElseIf inner_choice = "發電機(Generator)" Then
            Picture_machine.Image = My.Resources.發電機_Generator_
            Machine = Machines.Others
        ElseIf inner_choice = "空氣壓縮機(Compressor)" Then
            Picture_machine.Image = My.Resources.空氣壓縮機_Compressor_
            Machine = Machines.Others
        Else
            Machine = Nothing
        End If

        choice = inner_choice

        'mach區分要輸入至哪個 groupbox for plotting the coordinates
        If Not IsNothing(Machine) Then
            If Not Machine = Machines.Others Then
                If inner_choice.Contains("A1") Or inner_choice.Contains("A4") Then
                    GroupBox_A1_A2_A3.Enabled = False
                    GroupBox_A4.Enabled = False
                Else
                    GroupBox_A1_A2_A3.Enabled = True
                    TextBox_L.Enabled = True
                    Button_L_check.Enabled = True
                    GroupBox_A4.Enabled = False
                End If
            Else
                GroupBox_A1_A2_A3.Enabled = False
                GroupBox_A4.Enabled = True
                TextBox_L1.Enabled = True
                TextBox_L2.Enabled = True
                TextBox_L3.Enabled = True
                TextBox_r2.Enabled = True
                Button_L1_L2_L3_check.Enabled = True
            End If

            'if machine chose, then allow to go to the 2nd tab
            'MachChosen = True
        End If
    End Sub

    Private DataGrid As Grid
    Private A_Unit_Size As Size = New Size(1250, 450)
    'Create charts for chosen machine
    Private Sub CreateChart(ByVal r As Double)
        DataGrid = New Grid(TabPageCharts, A_Unit_Size, New Point(10, 10), r, Machine, HeadRun)
    End Sub

    Private Sub DisposeChart()
        If Not IsNothing(DataGrid) Then
            DataGrid.Form.Dispose()
            DataGrid = Nothing
        End If
    End Sub

    Private Sub Button_L_check_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_L_check.Click
        'Set up Plotting
        'Points = New ArrayList
        pos = {New CoorPoint(), New CoorPoint(), New CoorPoint(), New CoorPoint(), New CoorPoint(), New CoorPoint()}

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
        R = r1
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
        Else
            MachChosen = False
        End If

        DisposeChart()
        CreateChart(r1)
        '選擇機具後可更換
        If MachChosen = True Then
            Button_change_machine.Enabled = True
            ComboBox_machine_list.Enabled = False
            Button_L_check.Enabled = False
        End If
        TextBox_L.Enabled = False
    End Sub

    Private Sub Button_L1_L2_L3_check_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_L1_L2_L3_check.Click

        Label_A4_Hint1.Visible = True
        Label_A4_Hint2.Visible = True
        'Set up Plotting
        pos = {New CoorPoint(), New CoorPoint(), New CoorPoint(), New CoorPoint()}
        Dim L1 As Double
        Dim L2 As Double
        Dim L3 As Double
        Dim r2 As Double
        If String.IsNullOrWhiteSpace(TextBox_L1.Text) Or String.IsNullOrWhiteSpace(TextBox_L2.Text) Or String.IsNullOrWhiteSpace(TextBox_L3.Text) Then
            If String.IsNullOrWhiteSpace(TextBox_r2.Text) Then
                Return
            Else
                r2 = TextBox_r2.Text
            End If
        Else
            L1 = TextBox_L1.Text
            L2 = TextBox_L2.Text
            L3 = TextBox_L3.Text
            r2 = Math.Ceiling(2 * Math.Sqrt(((L1 / 2) ^ 2) + ((L2 / 2) ^ 2) + (L3 ^ 2)))
            TextBox_r2.Text = r2
        End If



        pos(0).Coors = New ThreeDPoint(r2 * -0.45, r2 * 0.77, r2 * 0.45)
        pos(1).Coors = New ThreeDPoint(r2 * -0.45, r2 * -0.77, r2 * 0.45)
        pos(2).Coors = New ThreeDPoint(r2 * 0.89, 0, r2 * 0.45)
        pos(3).Coors = New ThreeDPoint(0, 0, r2)

        R = r2
        plot(False, GP, xCor, yCor, GroupBox_Plot)
        If choice = "鐵輪壓路機(Road roller)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "膠輪壓路機(Wheel roller)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "振動式壓路機(Vibrating roller)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "履帶起重機(Crawler crane)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "卡車起重機(Truck crane)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "輪形起重機(Wheel crane)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "振動式樁錘(Vibrating hammer)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "油壓式打樁機(Hydraulic pile driver)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "拔樁機" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "油壓式拔樁機" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "土壤取樣器(地鑽) (Earth auger)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "全套管鑽掘機" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "鑽土機(Earth drill)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "鑽岩機(Rock breaker)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "混凝土泵車(Concrete pump)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "混凝土破碎機(Concrete breaker)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "瀝青混凝土舖築機(Asphalt finisher)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "混凝土割切機(Concrete cutter)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "發電機(Generator)" Then
            Load_Others(choice)
            MachChosen = True
        ElseIf choice = "空氣壓縮機(Compressor)" Then
            Load_Others(choice)
            MachChosen = True
        Else
            MachChosen = False
        End If

        DisposeChart()
        CreateChart(r2)

        '選擇機具後可更換
        If MachChosen = True Then
            Button_change_machine.Enabled = True
            ComboBox_machine_list.Enabled = False
            Button_L1_L2_L3_check.Enabled = False
        End If
        TextBox_L1.Enabled = False
        TextBox_L2.Enabled = False
        TextBox_L3.Enabled = False
    End Sub

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

        MainLineGraph = New LineGraph(New Point(90, 2), New Size(1055, 83), TabPage2, CGraph.Modes.A1A2A3, HeadRun.Time)
        MainBarGraph = New BarGraph(New Point(765, 93), New Size(380, 450), TabPage2, CGraph.Modes.A1A2A3)
        For i = 0 To 6
            NoisesArray(i).Visible = True
        Next
    End Sub

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


        MainLineGraph = New LineGraph(New Point(90, 2), New Size(1055, 83), TabPage2, CGraph.Modes.A1A2A3, HeadRun.Time)
        MainBarGraph = New BarGraph(New Point(765, 93), New Size(380, 450), TabPage2, CGraph.Modes.A1A2A3)
        For i = 0 To 6
            NoisesArray(i).Visible = True
        Next
    End Sub

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

        MainLineGraph = New LineGraph(New Point(90, 2), New Size(1055, 83), TabPage2, CGraph.Modes.A1A2A3, HeadRun.Time)
        MainBarGraph = New BarGraph(New Point(765, 93), New Size(380, 450), TabPage2, CGraph.Modes.A1A2A3)
        For i = 0 To 6
            NoisesArray(i).Visible = True
        Next
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

        MainLineGraph = New LineGraph(New Point(90, 2), New Size(1055, 83), TabPage2, CGraph.Modes.A1A2A3, HeadRun.Time)
        MainBarGraph = New BarGraph(New Point(765, 93), New Size(380, 450), TabPage2, CGraph.Modes.A1A2A3)
        For i = 0 To 6
            NoisesArray(i).Visible = True
        Next
    End Sub

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

        MainLineGraph = New LineGraph(New Point(90, 2), New Size(1055, 83), TabPage2, CGraph.Modes.A4, HeadRun.Time)
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

    Function Load_Steps_helper(ByRef run As Run_Unit)
        Dim tempRun As Run_Unit = run
        Dim tempStep As Steps
        If tempRun.Name = "ExA1" Or tempRun.Name = "LoA1" Or tempRun.Name = "TrA1" Or tempRun.Name = "ExA1_Add" Or tempRun.Name = "LoA1_Add" Or tempRun.Name = "TrA1_Add" Then
            tempRun.Steps = New Steps(My.Resources.A1_step1, Step1, Nothing, True, 5)
            tempRun.HeadStep = tempRun.Steps
        ElseIf tempRun.Name = "ExA2_1st" Or tempRun.Name = "ExA2_1st_Add" Or tempRun.Name = "ExA2_2nd_3rd" Or tempRun.Name = "ExA2_2nd_3rd_Add" Then
            If tempRun.Name = "ExA2_1st" Or tempRun.Name = "ExA2_1st_Add" Then
                tempRun.Steps = New Steps(My.Resources.A2_Excavator_step1, Step1, Nothing, False, -1)
                tempRun.HeadStep = tempRun.Steps
                tempRun.Steps.NextStep = New Steps(My.Resources.A2_Excavator_step2, Step2, Nothing, False, -1)
                tempStep = tempRun.Steps.NextStep
                tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step3, Step3, Nothing, False, array_ExA2_time(2))
                tempStep = tempStep.NextStep
            ElseIf tempRun.Name = "ExA2_2nd_3rd" Or tempRun.Name = "ExA2_2nd_3rd_Add" Then
                tempRun.Steps = New Steps(My.Resources.A2_Excavator_step2, Step1, Nothing, False, -1)
                tempRun.HeadStep = tempRun.Steps
                tempRun.Steps.NextStep = New Steps(My.Resources.A2_Excavator_step3, Step2, Nothing, False, array_ExA2_time(2))
                tempStep = tempRun.Steps.NextStep
            End If
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step4, Step4, Nothing, False, array_ExA2_time(3))
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step5, Step5, Nothing, False, array_ExA2_time(4))
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step6, Step6, Nothing, False, array_ExA2_time(5))
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step7, Step7, Nothing, False, array_ExA2_time(6))
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step8, Step8, Nothing, False, array_ExA2_time(7))
            tempStep = tempStep.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Excavator_step9, Step9, Nothing, True, array_ExA2_time(8))
        ElseIf tempRun.Name = "LoA2_1st" Or tempRun.Name = "LoA2_1st_Add" Then
            tempRun.Steps = New Steps(My.Resources.A2_Loader_step1, Step1, Nothing, False, -1)
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A2_Loader_step2, Step2, Nothing, False, array_LoA2_time(1))
            tempStep = tempRun.Steps.NextStep
            tempStep.NextStep = New Steps(My.Resources.A2_Loader_step3, Step3, Nothing, True, array_LoA2_time(2))
        ElseIf tempRun.Name = "LoA2_2nd_3rd" Or tempRun.Name = "LoA2_2nd_3rd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A2_Loader_step2, Step1, Nothing, False, array_LoA2_time(1))
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A2_Loader_step3, Step2, Nothing, True, array_LoA2_time(2))
        ElseIf tempRun.Name = "LoA3_fwd" Or tempRun.Name = "LoA3_fwd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A3_Loader_step1, Step1, Nothing, False, -1)
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A3_Loader_step2, Step2, Nothing, False, -1)
            tempStep = tempRun.Steps.NextStep
            tempStep.NextStep = New Steps(My.Resources.A3_Loader_step3, Step3, Nothing, True, array_LoA3_time(2))
        ElseIf tempRun.Name = "LoA3_bkd" Or tempRun.Name = "LoA3_bkd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A3_Loader_step4, Step1, Nothing, True, array_LoA3_time(3))
            tempRun.HeadStep = tempRun.Steps
        ElseIf tempRun.Name = "TrA3_fwd" Or tempRun.Name = "TrA3_fwd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A3_Tractor_step1, Step1, Nothing, False, -1)
            tempRun.HeadStep = tempRun.Steps
            tempRun.Steps.NextStep = New Steps(My.Resources.A3_Tractor_step2, Step2, Nothing, False, -1)
            tempStep = tempRun.Steps.NextStep
            tempStep.NextStep = New Steps(My.Resources.A3_Tractor_step3, Step3, Nothing, True, array_TrA3_time(2))
        ElseIf tempRun.Name = "TrA3_bkd" Or tempRun.Name = "TrA3_bkd_Add" Then
            tempRun.Steps = New Steps(My.Resources.A3_Tractor_step4, Step1, Nothing, True, array_TrA3_time(3))
            tempRun.HeadStep = tempRun.Steps
        ElseIf tempRun.Name = "A4" Or tempRun.Name = "A4_Add" Then
            tempRun.Steps = New Steps(A4_step_text, Step1, Nothing, True, 0)
            tempRun.HeadStep = tempRun.Steps
        End If

        Return tempRun.Steps

    End Function



    Private Function GetInstantData()
        Dim result(6 - 1) As Double
        Dim temp() As String =
        Comm.GetMeasurementsFromMeters(Communication.Measurements.Lp)
        'TEMP
        For i = 0 To temp.Length - 1

            result(i) = temp(i)
        Next
        For i = temp.Length To result.Length - 1
            result(i) = "0"
        Next
        Return result
    End Function

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
        For i = 0 To 6 - 1
            If Not onlyCal Then
                NoisesArray(i).Text = vals(i)
                sum += 10 ^ (0.1 * vals(i))
            ElseIf i = meter Then
                NoisesArray(i).Text = vals(i)
                sum = vals(i)
            Else
                vals(i) = 0
            End If
        Next

        Dim valsAndAvg() As Double = New Double(6) {}
        Array.Copy(vals, valsAndAvg, 6)

        Dim avg As Double
        If Not onlyCal Then
            avg = 10 * (Math.Log10(sum / num))
        Else
            avg = sum
        End If
        NoisesArray(6).Text = CStr(Math.Round(avg, 1))
        valsAndAvg(6) = avg
        'Set graphs
        MainBarGraph.Update(valsAndAvg)
        MainLineGraph.Update(vals)
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Countdown = False Then
            timeLeft = timeLeft + 1
            timeLabel.Text = timeLeft & " s"
            If timeLeft = 30 Then
                stopButton.Enabled = True
            End If
            'send values to display as text and graphs
            SetScreenValuesAndGraphs(GetInstantData())
        Else
            If timeLeft > 0 Then 'counting
                If Not timeLeft = 1 Then
                    timeLeft = timeLeft - 1
                    timeLabel.Text = timeLeft & " s"

                    'send values to display as text and graphs
                    SetScreenValuesAndGraphs(GetInstantData())
                ElseIf timeLeft = 1 Then 'time's up\
                    Comm.StopMeasure()
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
                            Accept_Button.Enabled = False
                            startButton.Enabled = True
                            stopButton.Enabled = False
                            array_step(CurRun.CurStep - 1).BackColor = Color.Green
                            Set_Panel_BackColor()
                            CurRun = CurRun.NextUnit
                            Set_Run_Unit()

                            'load A2's steps
                            Clear_Steps()
                            Load_Steps()

                            'dispose old graph and create new graph
                            'A2三個round的圖要連續所以不load新的圖
                            'Load_New_Graph_CD_True()

                        Else 'Countdown finally stop for entire run
                            Accept_Button.Enabled = True
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
            MainLineGraph = New LineGraph(New Point(90, 2), New Size(1055, 83), TabPage2, CGraph.Modes.A4, CurRun.Time)  'A4 mode
        Else
            MainLineGraph = New LineGraph(New Point(90, 2), New Size(1055, 83), TabPage2, CGraph.Modes.A1A2A3, CurRun.Time) 'A1 A2 A3 mode
        End If
    End Sub
    Sub Load_New_Graph_CD_True()
        'load new graph(when variable countdown = true)
        MainLineGraph.Dispose()
        If Machine = Machines.Others Then
            MainLineGraph = New LineGraph(New Point(90, 2), New Size(1055, 83), TabPage2, CGraph.Modes.A4, sum_steps)  'A4 mode
        Else
            MainLineGraph = New LineGraph(New Point(90, 2), New Size(1055, 83), TabPage2, CGraph.Modes.A1A2A3, sum_steps) 'A1 A2 A3 mode
        End If
        sum_steps = 0
    End Sub
    ' click event on Start Button
    Private Sub startButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles startButton.Click
        All_Panel_Disable()
        Comm.StartMeasure()
        If Countdown = False Then
            startButton.Enabled = False
            Accept_Button.Enabled = False
            If CurRun.Name.Contains("A4") Or CurRun.Name.Contains("RSS") Or CurRun.Name.Contains("Background") Then
                stopButton.Enabled = False
            Else
                stopButton.Enabled = True
            End If

            Timer1.Start()
        Else
            startButton.Enabled = False
            stopButton.Enabled = True
            Timer1.Start()
        End If

    End Sub

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
        Dim temps() As String = Comm.GetMeasurementsFromBuffer(Communication.Measurements.Leq)
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
        MainBarGraph.Update(Leqs)

    End Sub

    Private Sub ShowResultsOnForm()

        'Precal,postcal,RSS,background
        Dim series As List(Of DataVisualization.Charting.Series) = MainLineGraph.GetSeries()

        Dim Leqpoints As DataVisualization.Charting.DataPointCollection = MainBarGraph.GetSeries(0).Points()

        'precal and postcal
        If CurRun.Name.Contains("Cal") Then
            If CurRun.Name.Contains("2") Then
                CurRun.GRU.SetM(Meter_Measure_Unit.SeriesToMMU(series(0), Leqpoints(0).YValues(0)), 2)
            ElseIf CurRun.Name.Contains("4") Then
                CurRun.GRU.SetM(Meter_Measure_Unit.SeriesToMMU(series(1), Leqpoints(1).YValues(0)), 4)
            ElseIf CurRun.Name.Contains("6") Then
                CurRun.GRU.SetM(Meter_Measure_Unit.SeriesToMMU(series(2), Leqpoints(2).YValues(0)), 6)
            ElseIf CurRun.Name.Contains("8") Then
                CurRun.GRU.SetM(Meter_Measure_Unit.SeriesToMMU(series(3), Leqpoints(3).YValues(0)), 8)
            ElseIf CurRun.Name.Contains("10") Then
                CurRun.GRU.SetM(Meter_Measure_Unit.SeriesToMMU(series(4), Leqpoints(4).YValues(0)), 10)
            ElseIf CurRun.Name.Contains("12") Then
                CurRun.GRU.SetM(Meter_Measure_Unit.SeriesToMMU(series(5), Leqpoints(5).YValues(0)), 12)

            End If
        End If


        'everything else
        Dim tempGRU As Grid_Run_Unit = CurRun.GRU
        While Not IsNothing(tempGRU.NextGRU) 'for more than one additional runs
            tempGRU = tempGRU.NextGRU
        End While
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

    '##BUTTON CLICKS
    Private Sub stopButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stopButton.Click

        If Countdown = False Then
            Timer1.Stop()
            'startButton.Enabled = True
            Accept_Button.Enabled = True
            stopButton.Enabled = False

            'final Leq
            updateFinalBarGraph()
            'Tell meters to stop measuring
            Comm.StopMeasure()
            ShowResultsOnForm()

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
        Else
            Timer1.Stop()
            'final Leq
            updateFinalBarGraph()
            'Tell meters to stop measuring
            Comm.StopMeasure()

            startButton.Enabled = True
            stopButton.Enabled = False
            All_Panel_Enable()
            If Temp_CurRun.Link.Name = "LinkLabel_Temp" Then 'if not jump step
                'bounce back
                If CurRun.Name = "ExA1" Or CurRun.Name = "LoA1" Or CurRun.Name = "TrA1" Or CurRun.Name = "A4" Or CurRun.Name = "ExA1_Add" Or CurRun.Name = "LoA1_Add" Or CurRun.Name = "TrA1_Add" Or CurRun.Name = "A4_Add" Or CurRun.Name = "ExA2_1st" Or CurRun.Name = "ExA2_1st_Add" Or CurRun.Name = "LoA2_1st" Or CurRun.Name = "LoA2_1st_Add" Or CurRun.Name = "LoA3_fwd" Or CurRun.Name = "LoA3_bkd" Or CurRun.Name = "TrA3_fwd" Or CurRun.Name = "TrA3_bkd" Or CurRun.Name = "LoA3_fwd_Add" Or CurRun.Name = "LoA3_bkd_Add" Or CurRun.Name = "TrA3_fwd_Add" Or CurRun.Name = "TrA3_bkd_Add" Then
                    'dispose data

                    'reset run_unit
                    Set_Run_Unit()
                    'load A1's steps
                    Clear_Steps()
                    Load_Steps()
                    'dispose old graph and create new graph
                    Load_New_Graph_CD_True()
                    'bounce back to previous step
                ElseIf CurRun.Name = "ExA2_2nd_3rd" Or CurRun.Name = "LoA2_2nd_3rd" Or CurRun.Name = "ExA2_2nd_3rd_Add" Or CurRun.Name = "LoA2_2nd_3rd_Add" Then
                    If CurRun.PrevUnit.Name = "ExA2_1st" Or CurRun.PrevUnit.Name = "ExA2_1st_Add" Or CurRun.PrevUnit.Name = "LoA2_1st" Or CurRun.PrevUnit.Name = "LoA2_1st_Add" Then
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
            Else
                If CurRun.Name = "ExA2_1st" Or CurRun.Name = "LoA2_1st" Or CurRun.Name = "ExA2_1st_Add" Or CurRun.Name = "LoA2_1st_Add" Then
                    CurRun.NextUnit.Set_BackColor(Color.Green)
                    CurRun.NextUnit.NextUnit.Set_BackColor(Color.Green)
                ElseIf CurRun.Name = "ExA2_2nd_3rd" Or CurRun.Name = "LoA2_2nd_3rd" Or CurRun.Name = "ExA2_2nd_3rd_Add" Or CurRun.Name = "LoA2_2nd_3rd_Add" Then
                    If CurRun.NextUnit.Name.Contains("_2nd_3rd") Then
                        CurRun.NextUnit.Set_BackColor(Color.Green)
                    End If
                End If
                CurRun.Set_BackColor(Color.Green)
                Temp_CurRun.Set_BackColor(Color.Yellow)

                CurRun = Temp_CurRun
                Temp_CurRun = Null_CurRun

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
                timeLabel.Text = timeLeft & " s"

            End If
        End If
    End Sub

    'two approaches, if there's already a GRU existent, add as next GRU, if not, attach to the RU
    Sub Set_Add_GRU(ByVal whichA As Integer, ByVal colName As String, ByVal subHeader As String)
        Dim tempGRU = New Grid_Run_Unit(colName) 'the final GRU added
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

            While Not IsNothing(lastGRU.NextGRU)
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


    Sub Set_Run_Unit()
        CurRun.Steps = Load_Steps_helper(CurRun)
        CurStep = CurRun.HeadStep
        CurRun.CurStep = 1
        For index = CurRun.StartStep To CurRun.EndStep
            If CurRun.Steps.Time = -1 Then
                CurRun.Steps = CurRun.Steps.NextStep
                CurRun.CurStep += 1
            End If
        Next
        timeLeft = CurRun.Steps.Time
        timeLabel.Text = timeLeft & " s"
    End Sub
    Sub Restart_from_1st_Previous_Run_Unit()
        CurRun.PrevUnit.Set_BackColor(Color.Yellow)
        CurRun.Set_BackColor(Color.IndianRed)
        CurRun = CurRun.PrevUnit
        CurRun.Steps = Load_Steps_helper(CurRun)
        CurStep = CurRun.HeadStep
        CurRun.CurStep = 1
        For index = CurRun.StartStep To CurRun.EndStep
            If CurRun.Steps.Time = -1 Then
                CurRun.Steps = CurRun.Steps.NextStep
                CurRun.CurStep += 1
            End If
        Next
        CurRun.NextUnit.CurStep = 1
        timeLeft = CurRun.Steps.Time
        timeLabel.Text = timeLeft & " s"
    End Sub
    Sub Restart_from_2nd_Previous_Run_Unit()
        CurRun.PrevUnit.PrevUnit.Set_BackColor(Color.Yellow)
        CurRun.PrevUnit.Set_BackColor(Color.IndianRed)
        CurRun.Set_BackColor(Color.IndianRed)
        CurRun = CurRun.PrevUnit.PrevUnit
        CurRun.Steps = Load_Steps_helper(CurRun)
        CurStep = CurRun.HeadStep
        CurRun.CurStep = 1
        For index = CurRun.StartStep To CurRun.EndStep
            If CurRun.Steps.Time = -1 Then
                CurRun.Steps = CurRun.Steps.NextStep
                CurRun.CurStep += 1
            End If
        Next
        CurRun.NextUnit.CurStep = 1
        CurRun.NextUnit.NextUnit.CurStep = 1
        timeLeft = CurRun.Steps.Time
        timeLabel.Text = timeLeft & " s"
    End Sub

    Private Sub Test_StartButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Test_StartButton.Click

        Load_New_Graph_CD_True()
        Test_NextButton.Enabled = False
        Test_StopButton.Enabled = True
        Test_StartButton.Enabled = False
        Timer1.Start()
    End Sub
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
    End Sub
    Private Sub Test_ConfirmButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Test_ConfirmButton.Click
        Dim Input_S_Apply As Boolean = True
        Dim array_time As Array

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
            CurStep = Load_Steps_helper(CurRun)
            If Not (CurRun.Name = "TrA3_bkd" Or CurRun.Name = "LoA3_bkd") Then
                If CurRun.Name = "ExA2_1st" Then
                    array_time = array_ExA2_time
                ElseIf CurRun.Name = "LoA2_1st" Then
                    array_time = array_LoA2_time
                ElseIf CurRun.Name = "LoA3_fwd" Then
                    array_time = array_LoA3_time
                ElseIf CurRun.Name = "TrA3_fwd" Then
                    array_time = array_TrA3_time
                End If
                For i = 0 To CurRun.EndStep - 1
                    If Not CurStep.Time = -1 Then
                        array_time(i) = array_step_s(i).Text
                        CurStep.Time = array_step_s(i).Text
                        CurStep = CurStep.NextStep
                    Else
                        array_time(i) = -1
                        CurStep = CurStep.NextStep
                    End If
                Next
            Else
                'bkd 的時間和 fwd的時間都放在同一個array裡
                If CurRun.Name = "TrA3_bkd" Then
                    array_TrA3_time(CurRun.PrevUnit.EndStep) = array_step_s(CurRun.PrevUnit.EndStep).Text
                ElseIf CurRun.Name = "LoA3_bkd" Then
                    array_LoA3_time(CurRun.PrevUnit.EndStep) = array_step_s(CurRun.PrevUnit.EndStep).Text
                End If
            End If

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
                CurRun = CurRun.NextUnit
                CurRun.Steps = Load_Steps_helper(CurRun)
                CurStep = CurRun.HeadStep
                timeLeft = 0
                timeLabel.Text = timeLeft & " s"
                'Load_Steps_helper(CurRun)

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
    End Sub
    Sub Set_Second_for_Steps()
        If CurRun.Name.Contains("ExA2_1st") Then
            If Not Temp_CurRun.Link.Name = "LinkLabel_Temp" Then
                For index = 8 To 1
                    array_step_display(index).Text = array_step_display(index - 1).Text
                Next
                array_step_display(0).Text = "0"
            Else
                For index = 0 To 8
                    array_step_display(index).Text = ""
                Next
                For index = 0 To 8
                    array_step_display(index).Text = array_step_s(index).Text
                Next
            End If
        ElseIf CurRun.Name.Contains("ExA2_2nd") Then
            If Not Temp_CurRun.Link.Name = "LinkLabel_Temp" Then
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
            If Not Temp_CurRun.Link.Name = "LinkLabel_Temp" Then
                For index = 2 To 1
                    array_step_display(index).Text = array_step_display(index - 1).Text
                Next
                array_step_display(0).Text = "0"
            Else
                For index = 0 To 8
                    array_step_display(index).Text = ""
                Next
                For index = 0 To 2
                    array_step_display(index).Text = array_step_s(index).Text
                Next
            End If

        ElseIf CurRun.Name.Contains("LoA2_2nd") Then
            If Not Temp_CurRun.Link.Name = "LinkLabel_Temp" Then
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
            End If
        Next
        timeLeft = 0
        timeLabel.Text = timeLeft & " s"
    End Sub

    Sub Jump_Back_Countdown_True()
        If CurRun.Name = "ExA2_1st" Or CurRun.Name = "ExA2_1st_Add" Or CurRun.Name = "LoA2_1st" Or CurRun.Name = "LoA2_1st_Add" Or CurRun.Name = "ExA2_2nd_3rd" And CurRun.NextUnit.Name = "ExA2_2nd_3rd" Or CurRun.Name = "ExA2_2nd_3rd_Add" And CurRun.NextUnit.Name = "ExA2_2nd_3rd_Add" Or CurRun.Name = "LoA2_2nd_3rd" And CurRun.NextUnit.Name = "LoA2_2nd_3rd" Or CurRun.Name = "LoA2_2nd_3rd_Add" And CurRun.NextUnit.Name = "LoA2_2nd_3rd_Add" Then
            'MessageBox.Show("here")
            Set_Panel_BackColor()
            CurRun = CurRun.NextUnit
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
            Result = MessageBox.Show("此步驟數據已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
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
                Temp_CurRun.Set_BackColor(Color.Yellow)

                CurRun = Temp_CurRun
                Temp_CurRun = Null_CurRun

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
                Temp_CurRun.Set_BackColor(Color.Yellow)

                CurRun = Temp_CurRun
                Temp_CurRun = Null_CurRun

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
                timeLabel.Text = timeLeft & " s"

            ElseIf Result = DialogResult.Cancel Then
                Accept_Cancel()
            End If
        End If

    End Sub
    Sub Jump_Back_Countdown_False()
        Dim keyGRU As Grid_Run_Unit = CurRun.GRU
        While Not IsNothing(keyGRU.NextGRU) 'for more than one additional runs
            keyGRU = keyGRU.NextGRU
        End While
        Result = MessageBox.Show("此步驟數據已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
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
            Temp_CurRun.Set_BackColor(Color.Yellow)
            CurRun = Temp_CurRun
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
            Temp_CurRun.Set_BackColor(Color.Yellow)
            CurRun = Temp_CurRun
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
    End Sub

    Private Sub AcceptButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Accept_Button.Click

        Accept_Button.Enabled = False
        startButton.Enabled = True

        Dim keyGRU As Grid_Run_Unit = CurRun.GRU
        While Not IsNothing(keyGRU.NextGRU) 'for more than one additional runs
            keyGRU = keyGRU.NextGRU
        End While

        If Temp_CurRun.Link.Name = "LinkLabel_Temp" Then
            Result = MessageBox.Show("此步驟數據已測量完畢且接受此數據?", "My application", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If Result = DialogResult.Yes Then
                keyGRU.Accept()
                If CurRun.Name.Contains("bkd") Then
                    keyGRU.OverallGRU.Accept()
                End If
                If Countdown = False Then
                    If CurRun.Name.Contains("PreCal") Then
                        'didn't move
                        All_Panel_Enable()
                        'Case1: now is PreCal
                        ' change light
                        Set_Panel_BackColor()
                        CurRun.Link.Enabled = True
                        'dispose old graph and create new graph
                        Load_New_Graph_CD_False()

                        'save info here

                        'jump to next Run_Unit and set second to zero(should be zero here)
                        CurRun = CurRun.NextUnit
                        timeLeft = CurRun.Time
                        timeLabel.Text = timeLeft & " s"


                    ElseIf CurRun.Name = "Background" Then
                        All_Panel_Enable()
                        'Case2: now is Background
                        ' change light
                        Set_Panel_BackColor()
                        CurRun.Link.Enabled = True
                        'jump to next Run_Unit and change countdown False to True(for A1 A2 A3 A4 test)
                        CurRun = CurRun.NextUnit
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
                        All_Panel_Enable()
                        'Case3: now is RSS
                        ' change light
                        Set_Panel_BackColor()
                        CurRun.Link.Enabled = True
                        'dispose old graph and create new graph
                        Load_New_Graph_CD_False()

                        'Show calculated results for chart
                        DataGrid.ShowCalculated()

                        'save info here
                        'jump to next Run_Unit and set second to zero(should be zero here)
                        CurRun = CurRun.NextUnit
                        timeLeft = CurRun.Time
                        timeLabel.Text = timeLeft & " s"

                    ElseIf IsNothing(CurRun.NextUnit) Then
                        All_Panel_Enable()
                        CurRun.Set_BackColor(Color.Green)
                        CurRun.Link.Enabled = True
                        MessageBox.Show("End")
                        startButton.Enabled = False
                        stopButton.Enabled = False
                        Accept_Button.Enabled = False

                    ElseIf CurRun.Name.Contains("PostCal") Then
                        All_Panel_Enable()
                        'Case4: now is PostCal
                        ' change light
                        Set_Panel_BackColor()
                        CurRun.Link.Enabled = True
                        'dispose old graph and create new graph
                        Load_New_Graph_CD_False()

                        'save info here

                        'jump to next Run_Unit and set second to zero(should be zero here)
                        CurRun = CurRun.NextUnit
                        timeLeft = CurRun.Time
                        timeLabel.Text = timeLeft & " s"


                    ElseIf CurRun.Name = "A4" Then
                        All_Panel_Enable()
                        If CurRun.NextUnit.Name = "A4_Mid_BG" Then
                            'Case: now is A4 and next is A4_Mid_BG
                            ' change light
                            Set_Panel_BackColor()
                            CurRun.Link.Enabled = True
                            'jump to next Run_Unit
                            CurRun = CurRun.NextUnit
                            timeLeft = CurRun.Time
                            timeLabel.Text = timeLeft & " s"
                            Clear_Steps()
                            'dispose old graph and create new graph
                            Load_New_Graph_CD_False()
                        ElseIf CurRun.NextUnit.Name = "A4_Mid_BG_Add" Then
                            'have an additional test?
                            'call a function 
                            'if want_add()=true then ... elseif want_add()=false then ... endif

                            CurRun.Set_BackColor(Color.Green)
                            CurRun.Link.Enabled = True
                            'CurRun.Link.Enabled = True
                            'jump to next Run_Unit and change light
                            'Add?
                            Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                            list.Add(CurRun.GRU) 'adding test 3 
                            list.Add(CurRun.PrevUnit.PrevUnit.GRU) 'adding testing 2
                            list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding test 1
                            If DataGrid.NeedAdd(list) Then
                                'True: add test
                                CurRun = CurRun.NextUnit
                                Set_Add_GRU(0, "Background", "")    'True: add test
                                ' change light
                                CurRun.Set_BackColor(Color.Yellow)
                                timeLeft = CurRun.Time
                                timeLabel.Text = timeLeft & " s"
                                Clear_Steps()
                                'dispose old graph and create new graph
                                Load_New_Graph_CD_False()
                                'skip button enable
                                Button_Skip_Add.Enabled = True
                            Else
                                'False: not add test , jump to RSS
                                CurRun = CurRun.NextUnit.NextUnit.NextUnit
                                CurRun.Set_BackColor(Color.Yellow)
                                timeLeft = CurRun.Time
                                timeLabel.Text = timeLeft & " s"
                                Clear_Steps()
                                For index = 0 To 8
                                    array_step_display(index).Text = ""
                                Next
                                'dispose old graph and create new graph
                                Load_New_Graph_CD_False()
                            End If
                        End If

                    ElseIf CurRun.Name = "A4_Add" Then
                        All_Panel_Enable()
                        'have an additional test?
                        'call a function 
                        'if want_add()=true then ... elseif want_add()=false then ... endif
                        CurRun.Link.Enabled = True
                        'CurRun.Link.Enabled = True
                        'Add?
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
                        If DataGrid.NeedAdd(list) Then
                            CurRun.PrevUnit.Set_BackColor(Color.Yellow)
                            CurRun.Set_BackColor(Color.IndianRed)
                            CurRun = CurRun.PrevUnit
                            'True: add test
                            Set_Add_GRU(4, "Background", "")    'True: add test
                            'True: add test , jump to A4_Mid_BG_Add

                            timeLeft = CurRun.Time
                            timeLabel.Text = timeLeft & " s"

                            Clear_Steps()

                            'dispose old graph and create new graph
                            Load_New_Graph_CD_True()
                        Else
                            'False: not add test , jump to RSS
                            Set_Panel_BackColor()
                            CurRun = CurRun.NextUnit
                            CurRun.Steps = Load_Steps_helper(CurRun)
                            timeLeft = CurRun.Time
                            timeLabel.Text = timeLeft & " s"
                            Clear_Steps()
                            For index = 0 To 8
                                array_step_display(index).Text = ""
                            Next
                            'dispose old graph and create new graph
                            Load_New_Graph_CD_True()
                        End If

                    ElseIf CurRun.Name = "A4_Mid_BG" Or CurRun.Name = "A4_Mid_BG_Add" Then
                        All_Panel_Enable()
                        'Case: now is A4_Mid_BG or A4_Mid_BG_Add and next is A4 or A4_Add
                        ' change light
                        Set_Panel_BackColor()
                        CurRun.Link.Enabled = True

                        CurRun = CurRun.NextUnit
                        'jump to next Run_Unit
                        'case: next is A4_Add
                        If CurRun.Name.Contains("Add") Then
                            Dim tempGRU = CurRun.GRU
                            Dim i As Integer = 4
                            While Not IsNothing(tempGRU)
                                tempGRU = tempGRU.NextGRU
                                i += 1
                            End While

                            Set_Add_GRU(4, "Run " & i, "")
                        End If
                        Set_Run_Unit()
                        'load A4's steps
                        Clear_Steps()
                        Load_Steps()

                        'dispose old graph and create new graph
                        Load_New_Graph_CD_True()


                    End If
                Else 'countdown = true
                    If CurRun.Name = "ExA1" Or CurRun.Name = "TrA1" Or CurRun.Name = "LoA1" Then
                        All_Panel_Enable()
                        If CurRun.NextUnit.Name = "ExA1" Or CurRun.NextUnit.Name = "TrA1" Or CurRun.NextUnit.Name = "LoA1" Then
                            'Case1: now is ExA1 and next is also ExA1 or now is TrA1 and next is also TrA1 or now is LoA1 and next is also LoA1
                            ' change light
                            Set_Panel_BackColor()
                            CurRun.Link.Enabled = True
                            'jump to next Run_Unit
                            CurRun = CurRun.NextUnit
                            Set_Run_Unit()

                            'load A1's steps
                            Clear_Steps()
                            Load_Steps()

                            'dispose old graph and create new graph
                            Load_New_Graph_CD_True()

                        ElseIf CurRun.NextUnit.Name.Contains("A1_Add") Then
                            'Case2: now is ExA1 and next is ExA1_Add or now is TrA1 and next is TrA1_Add or now is LoA1 and next is LoA1_Add
                            'have an additional test?
                            'call a function 

                            CurRun.Set_BackColor(Color.Green)
                            CurRun.Link.Enabled = True
                            'jump to next Run_Unit and change light

                            'Add?
                            Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                            list.Add(CurRun.GRU) 'adding test 3
                            list.Add(CurRun.PrevUnit.GRU) 'adding testing 2
                            list.Add(CurRun.PrevUnit.PrevUnit.GRU) 'adding test 1
                            If DataGrid.NeedAdd(list) Then
                                'True: add test
                                CurRun = CurRun.NextUnit
                                Set_Add_GRU(0, "Run 4", "")
                                CurRun.Set_BackColor(Color.Yellow)
                                Set_Run_Unit()
                            Else
                                'False: not add test , jump to ExA2_1st or jump to TrA3_fwd or jump to LoA2_1st
                                CurRun = CurRun.NextUnit.NextUnit
                                CurRun.Set_BackColor(Color.Yellow)
                                Reset_Test_Time()
                            End If
                            'load steps
                            Clear_Steps()
                            Load_Steps()

                            'dispose old graph and create new graph
                            Load_New_Graph_CD_True()
                        End If

                    ElseIf CurRun.Name = "ExA1_Add" Or CurRun.Name = "TrA1_Add" Or CurRun.Name = "LoA1_Add" Then
                        All_Panel_Enable()
                        'Case2: now is ExA1_Add and next is ExA2_1st or now is TrA1_Add and next is TrA3_fwd or now is LoA1_Add and next is LoA2_1st
                        'have an additional test?
                        'call a function 
                        'if want_add()=true then ... elseif want_add()=false then ... endif


                        'CurRun.Link.Enabled = True
                        'Add?
                        Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                        list.Add(CurRun.PrevUnit.GRU) 'adding test 3
                        list.Add(CurRun.PrevUnit.PrevUnit.GRU) 'adding testing 2
                        list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding test 1
                        Dim tempGRU = CurRun.GRU
                        Dim i As Integer = 4
                        While Not IsNothing(tempGRU)
                            list.Add(tempGRU)
                            tempGRU = tempGRU.NextGRU
                            i += 1
                        End While
                        If DataGrid.NeedAdd(list) Then
                            'True: add test
                            CurRun.Set_BackColor(Color.Yellow)
                            Set_Add_GRU(0, "Run " & i, "")
                            Set_Run_Unit()
                        Else
                            'False: not add test
                            Set_Panel_BackColor()
                            CurRun = CurRun.NextUnit
                            Reset_Test_Time()
                        End If
                        'load steps
                        Clear_Steps()
                        Load_Steps()

                        'dispose old graph and create new graph
                        Load_New_Graph_CD_True()

                    ElseIf CurRun.Name = "ExA2_1st" Or CurRun.Name = "ExA2_2nd_3rd" Or CurRun.Name = "LoA2_1st" Or CurRun.Name = "LoA2_2nd_3rd" Then
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
                            Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                            list.Add(CurRun.GRU) 'adding test 3
                            list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding testing 2
                            list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding test 1
                            If DataGrid.NeedAdd(list) Then
                                'True: add test
                                CurRun = CurRun.NextUnit
                                Set_Add_GRU(2, "Run 4", "")    'True: add test
                                Set_Run_Unit()
                                CurRun.Set_BackColor(Color.Yellow)
                                'load A2's steps
                                Clear_Steps()
                                Load_Steps()
                                'dispose old graph and create new graph
                                Load_New_Graph_CD_True()
                            Else
                                'False: not add test , jump to RSS(Ex) or LoA1(Ex) or LoA3_fwd(Lo)
                                CurRun = CurRun.NextUnit.NextUnit.NextUnit.NextUnit
                                CurRun.Set_BackColor(Color.Yellow)
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
                                ElseIf CurRun.Name = "LoA1" Then
                                    Set_Run_Unit()
                                    'load steps
                                    Clear_Steps()
                                    Load_Steps()
                                    'dispose old graph and create new graph
                                    Load_New_Graph_CD_True()
                                ElseIf CurRun.Name = "LoA3_fwd" Then
                                    Reset_Test_Time()
                                    Clear_Steps()
                                    Load_Steps()
                                    'dispose old graph and create new graph
                                    Load_New_Graph_CD_True()
                                End If
                            End If
                        ElseIf CurRun.NextUnit.Name = "ExA2_1st" Or CurRun.NextUnit.Name = "LoA2_1st" Then
                            Set_Panel_BackColor()
                            CurRun = CurRun.NextUnit
                            Set_Run_Unit()
                            'load A2's steps
                            Clear_Steps()
                            Load_Steps()

                            'dispose old graph and create new graph
                            Load_New_Graph_CD_True()
                        End If



                    ElseIf CurRun.Name = "ExA2_1st_Add" Or CurRun.Name = "ExA2_2nd_3rd_Add" Or CurRun.Name = "LoA2_1st_Add" Or CurRun.Name = "LoA2_2nd_3rd_Add" Then
                        All_Panel_Enable()
                        'CurRun.PrevUnit.PrevUnit.Link.Enabled = True
                        If CurRun.NextUnit.Name = "RSS" Or CurRun.NextUnit.Name = "LoA1" Or CurRun.NextUnit.Name = "LoA3_fwd" Then
                            'have an additional test?
                            'call a function 
                            'if want_add()=true then ... elseif want_add()=false then ... endif

                            'jump to next Run_Unit
                            'Add?
                            Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                            list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding test 3
                            list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding testing 2
                            list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU) 'adding test 1
                            Dim tempGRU = CurRun.GRU
                            Dim i As Integer = 4
                            While Not IsNothing(tempGRU)
                                list.Add(tempGRU)
                                tempGRU = tempGRU.NextGRU
                                i += 1
                            End While
                            If DataGrid.NeedAdd(list) Then
                                'True: add test
                                Restart_from_2nd_Previous_Run_Unit()
                                Set_Add_GRU(2, "Run " & i, "")
                                'load A2's steps
                                Clear_Steps()
                                Load_Steps()

                                'dispose old graph and create new graph
                                Load_New_Graph_CD_True()
                            Else
                                'False: not add test ,  jump to RSS or LoA1
                                Set_Panel_BackColor()
                                CurRun = CurRun.NextUnit

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
                                ElseIf CurRun.Name = "LoA1" Then
                                    Set_Run_Unit()
                                    'load LoA1's steps
                                    Clear_Steps()
                                    Load_Steps()
                                    'dispose old graph and create new graph
                                    Load_New_Graph_CD_True()
                                ElseIf CurRun.Name = "LoA3_fwd" Then
                                    Reset_Test_Time()
                                    Clear_Steps()
                                    Load_Steps()
                                    'dispose old graph and create new graph
                                    Load_New_Graph_CD_True()
                                End If
                            End If
                        End If

                    ElseIf CurRun.Name = "LoA3_fwd" Or CurRun.Name = "LoA3_bkd" Or CurRun.Name = "TrA3_fwd" Or CurRun.Name = "TrA3_bkd" Then
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
                            CurRun = CurRun.NextUnit
                            Set_Run_Unit()
                            'load LoA3 or TrA3's steps
                            Clear_Steps()
                            Load_Steps()

                            'dispose old graph and create new graph
                            Load_New_Graph_CD_True()
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
                            Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                            list.Add(CurRun.GRU.NextGRU) 'adding test 3 overall
                            list.Add(CurRun.PrevUnit.PrevUnit.GRU.NextGRU) 'adding testing 2 overall
                            list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU.NextGRU) 'adding test 1
                            If DataGrid.NeedAdd(list) Then
                                'True: add test
                                CurRun = CurRun.NextUnit
                                Set_Add_GRU(3, "Run 4", "前進")    'True: add test

                                Set_Run_Unit()
                                CurRun.Set_BackColor(Color.Yellow)
                                'load LoA3_fwd_Add or TrA3_fwd_Add's steps
                                Clear_Steps()
                                Load_Steps()

                                'dispose old graph and create new graph
                                Load_New_Graph_CD_True()
                            Else
                                'False: not add test , jump to RSS
                                CurRun = CurRun.NextUnit.NextUnit.NextUnit
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
                        All_Panel_Enable()
                        If CurRun.NextUnit.Name = "LoA3_bkd_Add" Or CurRun.NextUnit.Name = "TrA3_bkd_Add" Then
                            'Case: LoA3_fwd_Add to LoA3_bkd_Add or TrA3_fwd_Add to TrA3_bkd_Add
                            ' change light
                            Set_Panel_BackColor()
                            'CurRun.Link.Enabled = True
                            'jump to next Run_Unit
                            CurRun = CurRun.NextUnit

                            If Not IsNothing(CurRun.GRU) Then
                                Dim i As Integer = 5
                                Dim tempGRU As Grid_Run_Unit = CurRun.GRU
                                While Not IsNothing(tempGRU.NextGRU)
                                    tempGRU = tempGRU.NextGRU
                                    i += 1
                                End While
                                Set_Add_GRU(3, "Run " & i, "後退")
                                tempGRU = tempGRU.NextGRU
                                DataGrid.AddA3OverallColumn(tempGRU)
                            Else
                                Set_Add_GRU(3, "Run 4", "後退")
                                DataGrid.AddA3OverallColumn(CurRun.GRU)
                            End If

                            Set_Run_Unit()
                            'load LoA3 or TrA3's steps
                            Clear_Steps()
                            Load_Steps()

                            'dispose old graph and create new graph
                            Load_New_Graph_CD_True()
                        ElseIf CurRun.NextUnit.Name = "RSS" Then
                            'Case: LoA3_bkd_Add to RSS or TrA3_bkd_Add to RSS
                            'have an additional test?
                            'call a function 
                            'if want_add()=true then ... elseif want_add()=false then ... endif

                            'CurRun.Link.Enabled = True
                            'Add?
                            Dim list As List(Of Grid_Run_Unit) = New List(Of Grid_Run_Unit)
                            list.Add(CurRun.PrevUnit.PrevUnit.GRU.OverallGRU) 'adding test 3
                            list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU.OverallGRU) 'adding testing 2
                            list.Add(CurRun.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.PrevUnit.GRU.OverallGRU) 'adding test 1
                            Dim tempGRU = CurRun.GRU
                            'DataGrid.AddA3Overall(CurRun.PrevUnit.GRU, CurRun.GRU) 'adding overall gru and column
                            Dim i As Integer = 4
                            While Not IsNothing(tempGRU)
                                list.Add(tempGRU)
                                tempGRU = tempGRU.NextGRU
                                i += 1
                            End While
                            If DataGrid.NeedAdd(list) Then
                                Restart_from_1st_Previous_Run_Unit()
                                'True: add test
                                Set_Add_GRU(3, "Run " & i, "前進")

                                'load LoA3_fwd or TrA3_fwd's steps
                                Clear_Steps()
                                Load_Steps()

                                'dispose old graph and create new graph
                                Load_New_Graph_CD_True()
                            Else
                                'False: not add test , jump to RSS
                                Set_Panel_BackColor()
                                CurRun = CurRun.NextUnit
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



                    End If
                End If

            ElseIf Result = DialogResult.No Then

                If CurRun.Name = "ExA2_1st" Or CurRun.Name = "ExA2_2nd_3rd" Or CurRun.Name = "LoA2_1st" Or CurRun.Name = "LoA2_2nd_3rd" Then
                    All_Panel_Enable()
                    'dispose data

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
        ElseIf CurRun.Name.Contains("PreCal") Or CurRun.Name = "Background" Or CurRun.Name = "RSS" Or IsNothing(CurRun.NextUnit) Or CurRun.Name.Contains("PostCal") Or CurRun.Name = "A4_Mid_BG" Or CurRun.Name = "A4_Mid_BG_Add" Then
            Jump_Back_Countdown_False()
        Else
            Jump_Back_Countdown_True()
        End If
    End Sub




    'plot the coordinate system on startup
    Private Sub plotCor(ByRef gp As Graphics, ByVal xCor As Double(,), ByVal yCor As Double(,))
        Dim fn As Font = New Font("Microsoft Sans MS", 20)
        'x axis
        gp.DrawLine(Pens.Black, New Point(xCor(0, 0), xCor(0, 1)), New Point(xCor(1, 0), xCor(1, 1)))
        gp.DrawString("X", fn, Brushes.Black, New Point(xCor(1, 0) + 3, xCor(1, 1)))

        'y axis
        gp.DrawLine(Pens.Black, New Point(yCor(0, 0), yCor(0, 1)), New Point(yCor(1, 0), yCor(1, 1)))
        gp.DrawString("Y", fn, Brushes.Black, New Point(yCor(0, 0), yCor(0, 1)))

        'xAxis = New LineShape(xCor(0, 0), xCor(0, 1), xCor(1, 0), xCor(1, 1))
        'xAxis.Parent = canvas
        'xLabel.Text = "x"
        'xLabel.Size = New System.Drawing.Size(20, 20)
        'xLabel.Location = New System.Drawing.Point(xCor(1, 0), xCor(1, 1))

        'y axis
        'yAxis = New LineShape(yCor(0, 0), yCor(0, 1), yCor(1, 0), yCor(1, 1))
        'yAxis.Parent = canvas
        'yLabel.Text = "y"
        'yLabel.Location = New System.Drawing.Point(yCor(0, 0), yCor(0, 1))
    End Sub

    'given an array of coordinates, plot them out using the given coordinate system
    'xCor is x axis coordinates
    'yCor is y axis coordinates
    'parent is the parent control to canvas
    'r is the radius for the circle
    'coors is the coordinates for the points to be plotted
    Private Sub plot(ByRef repaint As Boolean, ByRef gp As Graphics, ByVal xCor As Double(,), ByVal yCor As Double(,), ByVal parent As Control)
        ''clear canvas first
        'If canvas IsNot Nothing Then
        '    canvas.Dispose()
        'End If

        'canvas = New ShapeContainer()
        'canvas.Parent = parent
        'canvas.BackColor = Color.Transparent
        ''plot the coordinates
        plotCor(gp, xCor, yCor)

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
        Dim rCircle As Rectangle = New Rectangle(New Point((origin(0) - R), (origin(1) - R)), New Size(R * 2, R * 2))
        gp.DrawEllipse(Pens.Black, rCircle)
        'rCircle.Parent = canvas
        'plot normalized points
        Dim labels = {"P2", "P4", "P6", "P8", "P10", "P12"}

        For index = 0 To pos.GetLength(0) - 1
            Dim x = pos(index).Coors.Xc * ratio
            Dim y = pos(index).Coors.Yc * ratio

            Dim fn As Font = New Font("Microsoft Sans MS", 20)
            Dim solidBrush As SolidBrush = New SolidBrush(posColors(index))
            Dim bBrush As SolidBrush = New SolidBrush(Color.Blue)
            Dim rBrush As SolidBrush = New SolidBrush(Color.Red)
            Dim gBrush As SolidBrush = New SolidBrush(Color.DarkGreen)
            Dim pHeight As Integer = TextRenderer.MeasureText(gp, labels(index), fn).Height
            Dim pWidth As Integer = TextRenderer.MeasureText(gp, labels(index), fn).Width
            Dim rect As Rectangle = New Rectangle(New Point((origin(0) + x - 2), (origin(1) - y - 2)), New Size(5, 5))
            Dim d As Integer = 40
            Dim e As Integer = 10
            gp.FillEllipse(solidBrush, rect)

            'Text
            If labels(index) = "P6" Or labels(index) = "P8" Or labels(index) = "P12" Then
                gp.DrawString(labels(index), fn, solidBrush, New System.Drawing.Point(origin(0) + x, origin(1) - y - pHeight))
                gp.DrawString(String.Format("X: {0:N1}", pos(index).Coors.Xc), fn, bBrush, New System.Drawing.Point(origin(0) + x + pWidth - d, origin(1) - y - e))
                gp.DrawString(String.Format("Y: {0:N1}", pos(index).Coors.Yc), fn, rBrush, New System.Drawing.Point(origin(0) + x + pWidth - d, origin(1) - y + pHeight - e))
                gp.DrawString(String.Format("Z: {0:N1}", pos(index).Coors.Zc), fn, gBrush, New System.Drawing.Point(origin(0) + x + pWidth - d, origin(1) - y + pHeight * 2 - e))
            Else
                gp.DrawString(labels(index), fn, solidBrush, New System.Drawing.Point(origin(0) + x, origin(1) - y))
                gp.DrawString(String.Format("X: {0:N1}", pos(index).Coors.Xc), fn, bBrush, New System.Drawing.Point(origin(0) + x + pWidth - d, origin(1) - y + pHeight - e))
                gp.DrawString(String.Format("Y: {0:N1}", pos(index).Coors.Yc), fn, rBrush, New System.Drawing.Point(origin(0) + x + pWidth - d, origin(1) - y + pHeight * 2 - e))
                gp.DrawString(String.Format("Z: {0:N1}", pos(index).Coors.Zc), fn, gBrush, New System.Drawing.Point(origin(0) + x + pWidth - d, origin(1) - y + pHeight * 3 - e))
            End If
            'Dim rPoint = New OvalShape((origin(0) + x - 2), (origin(1) - y - 2), 5, 5)


            'rPoint.Parent = canvas
            'rPoint.BorderColor = posColors(index)
            'rPoint.BackColor = posColors(index)
            'rPoint.BackStyle = BackStyle.Opaque
            'coors(index).Label.Size = New Size(170, 100)
            'coors(index).Label.P = labels(index)
            'coors(index).Label.X = coors(index).Coors.Xc
            'coors(index).Label.Y = coors(index).Coors.Yc
            'coors(index).Label.Z = coors(index).Coors.Zc
            'coors(index).Label.Location = translate(New System.Drawing.Point(origin(0) + x, origin(1) - y), GroupBox_Plot.Location, TabPage1.Location)
            'coors(index).Label.BringToFront()
        Next
    End Sub

    Private Sub GroupBox_Plot_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles GroupBox_Plot.Paint
        If R = 0 Or pos Is Nothing Then
            Return
        End If
        plot(True, e.Graphics, xCor, yCor, GroupBox_Plot)
    End Sub

    'translate the location of point 1 from control 2's coordinates to control 3's coordinates
    'point 1 is given in var 2 coordinates
    'controls 2 and 3 are given in global terms)
    Private Function translate(ByVal p1, ByVal con2, ByVal con3)
        Dim x = p1.X + con2.X - con3.X
        Dim y = p1.Y + con2.Y - con3.Y
        Return New System.Drawing.Point(x, y)
    End Function

    'Private MouseIsDown As Boolean = False
    'Private lastPos As Point

    'Private Sub pLabel_MouseDown(ByVal sender As Object, ByVal e As  _
    'System.Windows.Forms.MouseEventArgs) Handles p8Label.MouseDown, p2Label.MouseDown, p10Label.MouseDown, p4Label.MouseDown, p12Label.MouseDown, p6Label.MouseDown
    '    ' Set a flag to show that the mouse is down.
    '    MouseIsDown = True
    '    lastPos = Cursor.Position
    '    For Each l As CoorPoint In Points
    '        If l.Label.Name = sender.Name Then
    '            If Not IsNothing(l.Line) Then
    '                l.Line.Dispose()
    '            End If
    '        End If
    '    Next
    'End Sub

    'Private Sub pLabel_MouseUp(ByVal sender As ColorLabel, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p6Label.MouseUp, p12Label.MouseUp, p4Label.MouseUp, p10Label.MouseUp, p2Label.MouseUp, p8Label.MouseUp
    '    MouseIsDown = False
    '    For Each l As CoorPoint In Points
    '        If l.Label.Name = sender.Name Then
    '            l.Line = New LineShape(canvas)
    '            l.Line.BorderColor = sender.ForeColor
    '            l.Line.StartPoint = New System.Drawing.Point((origin(0) + ratio * l.Coors.Xc), (origin(1) - ratio * l.Coors.Yc))
    '            l.Line.EndPoint = translate(New Point(l.Label.Location.X + l.Label.Size.Width / 2, l.Label.Location.Y), TabPage1.Location, GroupBox_Plot.Location + canvas.Location)
    '        End If
    '    Next
    'End Sub
    'Private Sub pLabel_MouseMove(ByVal sender As ColorLabel, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p6Label.MouseMove, p12Label.MouseMove, p4Label.MouseMove, p10Label.MouseMove, p2Label.MouseMove, p8Label.MouseMove
    '    If MouseIsDown Then
    '        ' Initiate dragging.
    '        Dim temp As Point = Cursor.Position
    '        sender.Location = temp - lastPos + sender.Location
    '        lastPos = temp
    '    End If
    'End Sub

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

    Private Sub DisconnButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DisconnButton.Click
        If Comm.Close() Then
            DisconnButton.Enabled = False
            ConnectButton.Enabled = True
        End If
    End Sub

    Private Sub SaveButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveButton.Click
        If DataGrid.Save() Then
            SaveButton.Enabled = False
        End If
    End Sub

    Public sim As Boolean = True

    Private Sub SimulationMode()
        ButtonMeters.Enabled = True
        ButtonSim.Enabled = False
        Comm.Sim()
        PanelMeterSetup.Enabled = False
        sim = True
    End Sub

    Private Sub ButtonSim_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSim.Click
        SimulationMode()
    End Sub


    Private Sub ButtonMeters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMeters.Click
        sim = False
        ButtonMeters.Enabled = False
        ButtonSim.Enabled = True
        PanelMeterSetup.Enabled = True
    End Sub

    Private Sub ButtonComRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonComRefresh.Click
        ComboBoxComs.Items.Clear()
        For Each com In Communication.GetComs()
            ComboBoxComs.Items.Add(com)
        Next
    End Sub

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

    Private Sub TextBox_Setting_Bargraph_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Setting_Bargraph_Max.TextChanged, TextBox_Setting_Bargraph_Min.TextChanged
        Button_Setting_Bargraph.Enabled = True
    End Sub

End Class
