﻿Imports System.IO
Imports System.Text


Public Class Grid
    'contains a grid
    Public WithEvents Form As DataGridView
    Private _R As Double
    Private _Machine As Program.Machines
    Private _Parent As Control

    Private _Background As Grid_Run_Unit
    Public ReadOnly Property Background As Grid_Run_Unit
        Get
            Return _Background
        End Get
    End Property

    Public ListAsGRUs As List(Of List(Of Grid_Run_Unit))

    Public Sub New(ByRef Parent As Control, ByVal Size As Size, ByVal Position As Point, ByVal R As Double, ByVal Machine As Program.Machines, ByVal headRU As Run_Unit)
        If IsNothing(headRU) Then
            MsgBox("Cannot create chart because given head run_unit is null!")
            Return
        End If
        Form = New DataGridView()
        _Parent = Parent
        Form.Parent = Parent
        Form.Size = Size
        Form.Location = Position
        Dim col As New DataGridViewTextBoxColumn
        Form.Columns.Add(col)
        'need to add a column before rows
        _R = R
        _Machine = Machine
        Form.Rows.Add(16)
        Form.Rows(1).HeaderCell.Value = "LpAeq2"
        Form.Rows(2).HeaderCell.Value = "LpAeq4"
        Form.Rows(3).HeaderCell.Value = "LpAeq6"
        Form.Rows(4).HeaderCell.Value = "LpAeq8"
        Form.Rows(5).HeaderCell.Value = "LpAeq10"
        Form.Rows(6).HeaderCell.Value = "LpAeq12"
        Form.Rows(7).HeaderCell.Value = "LpAeq avg"
        Form.Rows(8).HeaderCell.Value = "Time(sec)"
        Form.Rows(9).HeaderCell.Value = "deltaLA"
        Form.Rows(10).HeaderCell.Value = "K1A"
        Form.Rows(11).HeaderCell.Value = "L*W"
        Form.Rows(12).HeaderCell.Value = "LW"
        Form.Rows(13).HeaderCell.Value = "K2A"
        Form.Rows(14).HeaderCell.Value = "LWA"
        Form.Rows(15).HeaderCell.Value = "LWA 採用"

        ListAsGRUs = New List(Of List(Of Grid_Run_Unit))

        'Set A sequence
        Dim tempRU As Run_Unit = headRU

        If Machine = Program.Machines.Excavator Then 'A1 A2
            _LWA_Final = New Integer(1) {}

            For i = 0 To 11
                Form.Columns.Add(New DataGridViewTextBoxColumn())
            Next

            tempRU = PreCalConnect(tempRU)

            Dim tempListGRUs As New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A1's GRUs
            Form.Columns(2).HeaderText = "A1"
            AddLabelColumn(2)

            tempRU = ConnectGRU_RU_COL("Run1", "", 3, tempRU)
            tempRU = ConnectGRU_RU_COL("Run2", "", 4, tempRU)
            tempRU = ConnectGRU_RU_COL("Run3", "", 5, tempRU)
            tempRU = tempRU.NextUnit ' skipping additional


            tempListGRUs = New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A2's GRUs
            Form.Columns(6).HeaderText = "A2"
            AddLabelColumn(6)

            tempRU = tempRU.NextUnit.NextUnit 'skipping the first two small runs
            tempRU = ConnectGRU_RU_COL("Run1", "", 7, tempRU)

            tempRU = tempRU.NextUnit.NextUnit 'skipping the first two small runs
            tempRU = ConnectGRU_RU_COL("Run2", "", 8, tempRU)

            tempRU = tempRU.NextUnit.NextUnit 'skipping the first two small runs
            tempRU = ConnectGRU_RU_COL("Run3", "", 9, tempRU)
            tempRU = tempRU.NextUnit.NextUnit.NextUnit 'skipping add units

            tempRU = ConnectGRU_RU_COL("RSS", "", 10, tempRU)

            tempRU = PostCalConnect(11, tempRU)

        ElseIf Machine = Program.Machines.Loader Then 'A1 A2 A3
            'A_s = 3
            _LWA_Final = New Integer(2) {}
            For i = 0 To 21
                Form.Columns.Add(New DataGridViewTextBoxColumn())
            Next

            tempRU = PreCalConnect(tempRU)

            Dim tempListGRUs As New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A1's GRUs
            Form.Columns(2).HeaderText = "A1"
            AddLabelColumn(2)

            tempRU = ConnectGRU_RU_COL("Run1", "", 3, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "", 4, tempRU)

            tempRU = ConnectGRU_RU_COL("Run3", "", 5, tempRU)
            tempRU = tempRU.NextUnit 'skipping add unit

            tempListGRUs = New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A2's GRUs
            Form.Columns(6).HeaderText = "A2"
            AddLabelColumn(6)

            tempRU = tempRU.NextUnit.NextUnit
            tempRU = ConnectGRU_RU_COL("Run1", "", 7, tempRU)

            tempRU = tempRU.NextUnit.NextUnit
            tempRU = ConnectGRU_RU_COL("Run2", "", 8, tempRU)

            tempRU = tempRU.NextUnit.NextUnit
            tempRU = ConnectGRU_RU_COL("Run3", "", 9, tempRU)

            tempRU = tempRU.NextUnit.NextUnit.NextUnit 'skipping A2 add

            tempListGRUs = New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A3's GRUs
            Form.Columns(10).HeaderText = "A3"
            AddLabelColumn(10)

            tempRU = ConnectGRU_RU_COL("Run1", "前進", 11, tempRU)

            tempRU = ConnectGRU_RU_COL("Run1", "後退", 12, tempRU)

            tempRU = ConnectGRU_RU_COL("Run1", "平均", 13, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "前進", 14, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "後退", 15, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "平均", 16, tempRU)


            tempRU = ConnectGRU_RU_COL("Run3", "前進", 17, tempRU)

            tempRU = ConnectGRU_RU_COL("Run3", "後退", 18, tempRU)

            tempRU = ConnectGRU_RU_COL("Run3", "平均", 19, tempRU)

            tempRU = tempRU.NextUnit.NextUnit 'skipping add

            tempRU = ConnectGRU_RU_COL("RSS", "", 20, tempRU)

            tempRU = PostCalConnect(21, tempRU)

        ElseIf Machine = Program.Machines.Tractor Then 'A1 A3
            _LWA_Final = New Integer(1) {}
            'A_s = 2
            For i = 0 To 17
                Form.Columns.Add(New DataGridViewTextBoxColumn())
            Next
            tempRU = PreCalConnect(tempRU)

            Dim tempListGRUs As New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A1's GRUs
            Form.Columns(2).HeaderText = "A1"
            AddLabelColumn(2)

            tempRU = ConnectGRU_RU_COL("Run1", "", 3, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "", 4, tempRU)

            tempRU = ConnectGRU_RU_COL("Run3", "", 5, tempRU)
            tempRU = tempRU.NextUnit 'skipping additional

            tempListGRUs = New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A3's GRUs
            Form.Columns(6).HeaderText = "A3"
            AddLabelColumn(6)

            tempRU = ConnectGRU_RU_COL("Run1", "前進", 7, tempRU)

            tempRU = ConnectGRU_RU_COL("Run1", "後退", 8, tempRU)

            tempRU = ConnectGRU_RU_COL("Run1", "平均", 9, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "前進", 10, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "後退", 11, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "平均", 12, tempRU)

            tempRU = ConnectGRU_RU_COL("Run3", "前進", 13, tempRU)

            tempRU = ConnectGRU_RU_COL("Run3", "後退", 14, tempRU)

            tempRU = ConnectGRU_RU_COL("Run3", "平均", 15, tempRU)

            tempRU = tempRU.NextUnit.NextUnit 'skipping add

            tempRU = ConnectGRU_RU_COL("RSS", "", 16, tempRU)
            tempRU = PostCalConnect(17, tempRU)

        ElseIf Machine = Program.Machines.Loader_Excavator Then 'A1 A2, A1 A2 A3
            _LWA_Final = New Integer(4) {}
            'A_s = 5
            For i = 0 To 28
                Form.Columns.Add(New DataGridViewTextBoxColumn())
            Next
            tempRU = PreCalConnect(tempRU)

            Dim tempListGRUs As New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A1's GRUs
            Form.Columns(2).HeaderText = "A1"
            AddLabelColumn(2)

            tempRU = ConnectGRU_RU_COL("Run1", "", 3, tempRU)
            tempRU = ConnectGRU_RU_COL("Run2", "", 4, tempRU)

            tempRU = ConnectGRU_RU_COL("Run3", "", 5, tempRU)
            tempRU = tempRU.NextUnit 'skipping add

            tempListGRUs = New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A2's GRUs
            Form.Columns(6).HeaderText = "A2"
            AddLabelColumn(6)

            tempRU = tempRU.NextUnit.NextUnit 'skipping first two
            tempRU = ConnectGRU_RU_COL("Run1", "", 7, tempRU)

            tempRU = tempRU.NextUnit.NextUnit 'skipping first two
            tempRU = ConnectGRU_RU_COL("Run2", "", 8, tempRU)

            tempRU = tempRU.NextUnit.NextUnit 'skipping first two
            tempRU = ConnectGRU_RU_COL("Run3", "", 9, tempRU)
            tempRU = tempRU.NextUnit.NextUnit.NextUnit 'skipping A2 add

            tempListGRUs = New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A1's GRUs
            Form.Columns(10).HeaderText = "A1"
            AddLabelColumn(10)

            tempRU = ConnectGRU_RU_COL("Run1", "", 11, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "", 12, tempRU)

            tempRU = ConnectGRU_RU_COL("Run3", "", 13, tempRU)
            tempRU = tempRU.NextUnit 'skipping A1 add

            tempListGRUs = New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A2's GRUs
            Form.Columns(14).HeaderText = "A2"
            AddLabelColumn(14)

            tempRU = tempRU.NextUnit.NextUnit 'skipping first two
            tempRU = ConnectGRU_RU_COL("Run1", "", 15, tempRU)

            tempRU = tempRU.NextUnit.NextUnit 'skipping first two
            tempRU = ConnectGRU_RU_COL("Run2", "", 16, tempRU)

            tempRU = tempRU.NextUnit.NextUnit 'skipping first two
            tempRU = ConnectGRU_RU_COL("Run3", "", 17, tempRU)
            tempRU = tempRU.NextUnit.NextUnit.NextUnit 'skipping A2 add

            tempListGRUs = New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A3's GRUs
            Form.Columns(18).HeaderText = "A3"
            AddLabelColumn(18)

            tempRU = ConnectGRU_RU_COL("Run1", "前進", 19, tempRU)

            tempRU = ConnectGRU_RU_COL("Run1", "後退", 20, tempRU)

            tempRU = ConnectGRU_RU_COL("Run1", "平均", 21, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "前進", 22, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "後退", 23, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "平均", 24, tempRU)

            tempRU = ConnectGRU_RU_COL("Run3", "前進", 25, tempRU)

            tempRU = ConnectGRU_RU_COL("Run3", "後退", 26, tempRU)
            tempRU = ConnectGRU_RU_COL("Run3", "平均", 27, tempRU)
            tempRU = tempRU.NextUnit.NextUnit 'skipping add



            tempRU = ConnectGRU_RU_COL("RSS", "", 28, tempRU)

            tempRU = PostCalConnect(29, tempRU)

        ElseIf Machine = Program.Machines.Others Then 'A4
            _LWA_Final = New Integer(0) {}
            'A_s = 1
            For i = 0 To 8
                Form.Columns.Add(New DataGridViewTextBoxColumn())
            Next
            Form.Columns(0).HeaderText = "PreCal"
            'p2-p12
            For i = 1 To 4
                tempRU.GRU = New Grid_Run_Unit("PreCal_P" & i * 2)
                tempRU.GRU.Column = Form.Columns(0)
                tempRU.GRU.ParentRU = tempRU
                tempRU = tempRU.NextUnit
            Next

            Dim tempListGRUs As New List(Of Grid_Run_Unit)
            ListAsGRUs.Add(tempListGRUs) ' for A4's GRUs
            tempRU = ConnectGRU_RU_COL("Background", "", 1, tempRU)

            tempRU = ConnectGRU_RU_COL("Run1", "", 2, tempRU)

            tempRU = ConnectGRU_RU_COL("Background", "", 3, tempRU)

            tempRU = ConnectGRU_RU_COL("Run2", "", 4, tempRU)

            tempRU = ConnectGRU_RU_COL("Background", "", 5, tempRU)

            tempRU = ConnectGRU_RU_COL("Run3", "", 6, tempRU)

            tempRU = tempRU.NextUnit.NextUnit 'skipping add

            tempRU = ConnectGRU_RU_COL("RSS", "", 7, tempRU)

            Form.Columns(8).HeaderText = "PostCal"
            'p2
            Dim postcalGRU = New Grid_Run_Unit("PostCal")
            'p2-p12
            For i = 1 To 4
                tempRU.GRU = New Grid_Run_Unit("PostCal_P" & i * 2)
                tempRU.GRU.Column = Form.Columns(8)
                tempRU.GRU.ParentRU = tempRU
                tempRU = tempRU.NextUnit
            Next
        End If
    End Sub

    Public Shared Function Round(ByVal value As Double) As Double
        Dim diff As Double = value * 100 Mod 10
        Dim low As Double = (value * 100 - diff) / 100
        If diff >= 5 Then
            Return low + 0.1
        Else
            Return low
        End If
    End Function

    '將Grid_RUN_UNIT and Run_Unit 作連結
    Function ConnectGRU_RU_COL(ByVal colName As String, ByVal subHeader As String, ByVal colNum As Integer, ByRef tempRU As Run_Unit) As Run_Unit

        Dim A3overall As Boolean = False
        If subHeader.Contains("平均") Then
            A3overall = True
        End If
        Form.Columns(colNum).HeaderText = colName
        Form.Rows(0).Cells(colNum).Value = subHeader
        Dim tempGRU As Grid_Run_Unit = New Grid_Run_Unit(colName)
        tempGRU.Subheader = subHeader
        If (tempGRU.isRegular) And Not subHeader.Contains("前進") And Not subHeader.Contains("平均") Then
            ListAsGRUs(ListAsGRUs.Count - 1).Add(tempGRU)
        End If
        If A3overall Then
            tempRU.PrevUnit.GRU.OverallGRU = tempGRU
            tempGRU.Column = Form.Columns(colNum)
            tempGRU.Background = _Background
            Return tempRU
        End If
        tempRU.GRU = tempGRU
        tempRU.GRU.Column = Form.Columns(colNum)
        tempRU.GRU.ParentRU = tempRU
        If colName.Contains("Background") Or subHeader.Contains("Background") Then
            _Background = tempRU.GRU
        Else
            tempRU.GRU.Background = _Background
        End If
        tempRU = tempRU.NextUnit
        Return tempRU
    End Function

    Function PreCalConnect(ByRef tempRU As Run_Unit) As Run_Unit
        Form.Columns(0).HeaderText = "PreCal"
        'p2-p12
        For i = 1 To 6
            tempRU.GRU = New Grid_Run_Unit("PreCal_P" & i * 2)
            tempRU.GRU.Column = Form.Columns(0)
            tempRU.GRU.ParentRU = tempRU
            tempRU = tempRU.NextUnit
        Next

        tempRU = ConnectGRU_RU_COL("Background", "", 1, tempRU)
        Return tempRU
    End Function

    Function PostCalConnect(ByVal colNum As Integer, ByRef tempRU As Run_Unit) As Run_Unit
        Form.Columns(colNum).HeaderText = "PostCal"
        'p2-p12
        For i = 1 To 6
            tempRU.GRU = New Grid_Run_Unit("PostCal_P" & i * 2)
            tempRU.GRU.Column = Form.Columns(colNum)
            tempRU.GRU.ParentRU = tempRU
            tempRU = tempRU.NextUnit
        Next

        Return tempRU
    End Function


    Private Sub AddLabelColumn(ByRef colInd As Integer)
        Form.Rows(1).Cells(colInd).Value = "LpAeq2"
        Form.Rows(2).Cells(colInd).Value = "LpAeq4"
        Form.Rows(3).Cells(colInd).Value = "LpAeq6"
        Form.Rows(4).Cells(colInd).Value = "LpAeq8"
        Form.Rows(5).Cells(colInd).Value = "LpAeq10"
        Form.Rows(6).Cells(colInd).Value = "LpAeq12"
        Form.Rows(7).Cells(colInd).Value = "LpAeq avg"
        Form.Rows(8).Cells(colInd).Value = "Time(sec)"
        Form.Rows(9).Cells(colInd).Value = "deltaLA"
        Form.Rows(10).Cells(colInd).Value = "K1A"
        Form.Rows(11).Cells(colInd).Value = "L*W"
        Form.Rows(12).Cells(colInd).Value = "LW"
        Form.Rows(13).Cells(colInd).Value = "K2A"
        Form.Rows(14).Cells(colInd).Value = "LWA"
        Form.Rows(15).Cells(colInd).Value = "LWA 採用"
    End Sub

    'contains grid_run_units
    'Private GRUList As New List(Of Grid_Run_Unit)
    'Private GRUMap As New Dictionary(Of String, Grid_Run_Unit)
    Private RSS As Grid_Run_Unit

    Public Enum Run_Modes
        Background
        Regular
        RSS
    End Enum

    'link the grid_run_unit's column to the column indexed in form
    Public Sub LinkGrid_Run_Unit(ByRef gru As Grid_Run_Unit, ByVal index As Integer)
        gru.Column = Form.Columns(index)
    End Sub



    'show GRU on the Form so we can see figures
    Public Sub ShowGRUonForm(ByRef Run As Grid_Run_Unit)

        If Run.Header = Run_Modes.RSS.ToString() Then
            RSS = Run
        End If

        'for PreCal and PostCal
        Dim curColInd As Integer = Run.Column.Index

        If Run.Header.Contains("P2") Then
            Form.Rows(1).Cells(curColInd).Value = Run.LpAeq2
            Return
        ElseIf Run.Header.Contains("P4") Then
            Form.Rows(2).Cells(curColInd).Value = Run.LpAeq4
            Return
        ElseIf Run.Header.Contains("P6") Then
            Form.Rows(3).Cells(curColInd).Value = Run.LpAeq6
            Return
        ElseIf Run.Header.Contains("P8") Then
            Form.Rows(4).Cells(curColInd).Value = Run.LpAeq8
            Return
        ElseIf Run.Header.Contains("P10") Then
            Form.Rows(5).Cells(curColInd).Value = Run.LpAeq10
            Return
        ElseIf Run.Header.Contains("P12") Then
            Form.Rows(6).Cells(curColInd).Value = Run.LpAeq12
            Return
        End If
        
        'meter 2
        Form.Rows(1).Cells(curColInd).Value = Run.LpAeq2
        'meter 4
        Form.Rows(2).Cells(curColInd).Value = Run.LpAeq4
        'meter 6
        Form.Rows(3).Cells(curColInd).Value = Run.LpAeq6
        'meter 8
        Form.Rows(4).Cells(curColInd).Value = Run.LpAeq8
        'meter 10
        Form.Rows(5).Cells(curColInd).Value = Run.LpAeq10
        'meter 12
        Form.Rows(6).Cells(curColInd).Value = Run.LpAeq12
        'meters average
        Form.Rows(7).Cells(curColInd).Value = Run.LpAeqAvg
        'time
        Form.Rows(8).Cells(curColInd).Value = Run.Time
        If Run.isRegular Then
            'deltaA
            Form.Rows(9).Cells(curColInd).Value = Run.deltaLA
            'K1A
            Form.Rows(10).Cells(curColInd).Value = Run.K1A
        End If



        'add next column for next record
        'Dim col As New DataGridViewTextBoxColumn
        'col.DataPropertyName = "Run"
        'col.Name = Run.Header & Run.Subheader
        'Form.Columns.Add(col)
        'GRUMap.Add(Run.Column.HeaderText & Form.Rows(0).Cells(curColInd).Value, Run)

    End Sub

    Private Sub ShowFinal(ByRef gru As Grid_Run_Unit, ByVal i As Integer)
        Dim curColInd As Integer = gru.Column.Index
        'meter 2
        Form.Rows(1).Cells(curColInd).Value = gru.LpAeq2
        'meter 4
        Form.Rows(2).Cells(curColInd).Value = gru.LpAeq4
        'meter 6
        Form.Rows(3).Cells(curColInd).Value = gru.LpAeq6
        'meter 8
        Form.Rows(4).Cells(curColInd).Value = gru.LpAeq8
        'meter 10
        Form.Rows(5).Cells(curColInd).Value = gru.LpAeq10
        'meter 12
        Form.Rows(6).Cells(curColInd).Value = gru.LpAeq12
        'meters average
        Form.Rows(7).Cells(curColInd).Value = gru.LpAeqAvg
        'time
        Form.Rows(8).Cells(curColInd).Value = gru.Time

        'deltaA
        Form.Rows(9).Cells(curColInd).Value = gru.deltaLA
        'K1A
        Form.Rows(10).Cells(curColInd).Value = gru.K1A

        'L*W

        Form.Rows(11).Cells(curColInd).Value = LsW

        Form.Rows(12).Cells(curColInd).Value = Lwr

        Form.Rows(13).Cells(curColInd).Value = K2A

        Form.Rows(14).Cells(curColInd).Value = gru.LWA

        Form.Rows(15).Cells(curColInd).Value = LWA_Final(i)

    End Sub

    'this shows the A3 overall column during regular runs
    Public Sub ShowA3Overall(ByRef runFwd As Grid_Run_Unit, ByRef runBkd As Grid_Run_Unit)
        ' the only way to access the overall GRU is through the previous backward GRU
        Dim gru As Grid_Run_Unit = runBkd.OverallGRU
        'Try
        If gru Is Nothing Then
            gru = New Grid_Run_Unit(runBkd.Header)
            runBkd.OverallGRU = gru
        End If
        gru.SetMs(runFwd.Meter2 + runBkd.Meter2,
                runFwd.Meter4 + runBkd.Meter4,
                runFwd.Meter6 + runBkd.Meter6,
                runFwd.Meter8 + runBkd.Meter8,
                runFwd.Meter10 + runBkd.Meter10,
                runFwd.Meter12 + runBkd.Meter12)
        ShowGRUonForm(gru)
        'Catch ex As Exception
        '    MsgBox("In ShowA3Overall: " & ex.Message)
        'End Try

    End Sub

    'function that adds the overall column after runFwdard and runBwdard in additional runs
    Public Sub AddA3OverallColumn(ByRef runBkd As Grid_Run_Unit)

        Dim col As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn()
        col.HeaderText = runBkd.Column.HeaderText

        Dim gru As Grid_Run_Unit = New Grid_Run_Unit(runBkd.Column.HeaderText)
        runBkd.OverallGRU = gru
        gru.Column = col
        gru.Background = _Background
        Me.Form.Columns.Insert(runBkd.Column.Index + 1, col)
        Me.Form.Rows(0).Cells(col.Index).Value = "平均"
    End Sub

    'call this after RSS has been done
    Public Sub ShowCalculated()
        Calculate()
        Dim i = 0
        For Each l In ListAsGRUs
            ShowLWAs(l)
            ShowFinal(l(l.Count - 2), i)
            i += 1
        Next
    End Sub

    Private Sub ShowLWAs(ByRef listGRUs As List(Of Grid_Run_Unit))
        For Each gru In listGRUs
            Form.Rows(14).Cells(gru.Column.Index).Value = gru.LWA
        Next
    End Sub

    Public Sub ClearColumn(ByRef gru As Grid_Run_Unit)
        Dim curColInd As Integer = gru.Column.Index

        If gru.Header.Contains("P2") Then
            Form.Rows(1).Cells(curColInd).Value = ""
            Return
        ElseIf gru.Header.Contains("P4") Then
            Form.Rows(2).Cells(curColInd).Value = ""
            Return
        ElseIf gru.Header.Contains("P6") Then
            Form.Rows(3).Cells(curColInd).Value = ""
            Return
        ElseIf gru.Header.Contains("P8") Then
            Form.Rows(4).Cells(curColInd).Value = ""
            Return
        ElseIf gru.Header.Contains("P10") Then
            Form.Rows(5).Cells(curColInd).Value = ""
            Return
        ElseIf gru.Header.Contains("P12") Then
            Form.Rows(6).Cells(curColInd).Value = ""
            Return
        End If

        For i = 1 To Form.Rows.Count - 1
            Form.Rows(i).Cells(curColInd).Value = ""
        Next
    End Sub

    'contains calculations- NeedAdd should be called before (Calc_LWAs)Calculate
    Public Sub Calculate()
        If RSS IsNot Nothing Then
            Calc_LsW()
            Calc_K2A()
            Dim j = 0
            Dim howManyAs As Integer = ListAsGRUs.Count
            For i = 0 To howManyAs - 1
                Dim l As List(Of Grid_Run_Unit) = ListAsGRUs(i)
                'adding additional runs
                If l(0).ParentRU.Name.Contains("A1") Then 'case for A1
                    If l(2).ParentRU.NextUnit.GRU IsNot Nothing Then
                        Dim add As Grid_Run_Unit = l(2).ParentRU.NextUnit.GRU
                        While add IsNot Nothing
                            If l.IndexOf(add) < 0 Then
                                l.Add(add)

                            End If
                            add = add.NextGRU
                        End While
                    End If

                ElseIf l(0).ParentRU.Name.Contains("A2") Then 'case for A2
                    If l(2).ParentRU.NextUnit.NextUnit.NextUnit.GRU IsNot Nothing Then
                        Dim add As Grid_Run_Unit = l(2).ParentRU.NextUnit.NextUnit.NextUnit.GRU
                        While add IsNot Nothing
                            If l.IndexOf(add) < 0 Then
                                l.Add(add)

                            End If
                            add = add.NextGRU
                        End While
                    End If

                ElseIf l(0).ParentRU.Name.Contains("A3") Then 'case for A3
                    If l(2).ParentRU.NextUnit.NextUnit.GRU IsNot Nothing Then
                        Dim addbkd As Grid_Run_Unit = l(2).ParentRU.NextUnit.NextUnit.GRU

                        While addbkd IsNot Nothing
                            If l.IndexOf(addbkd) < 0 Then
                                l.Add(addbkd.OverallGRU)

                            End If
                            addbkd = addbkd.NextGRU
                        End While
                    End If

                ElseIf l(0).ParentRU.Name.Contains("A4") Then 'case for A4
                    If l(2).ParentRU.NextUnit.NextUnit.GRU IsNot Nothing Then
                        Dim add As Grid_Run_Unit = l(2).ParentRU.NextUnit.NextUnit.GRU
                        While add IsNot Nothing
                            If l.IndexOf(add) Then
                                l.Add(add)

                            End If
                            add = add.NextGRU
                        End While
                    End If
                End If
                'actually calculating the LWAs
                Calc_LWAs(l, j)
                j += 1
            Next
        End If
    End Sub

    Private _LsW As Double
    Public Sub Calc_LsW()
        Try
            If Not IsNothing(RSS) Then
                _LsW = RSS.LpAeqAvg
            End If
        Catch ex As Exception
            MsgBox("in Calc_LsW: " & ex.Message)
            _LsW = -1
        End Try
    End Sub

    Public ReadOnly Property LsW()
        Get
            Return _LsW
        End Get
    End Property

    Private _Lwr As Double = 100.0
    Public ReadOnly Property Lwr()
        Get
            Return _Lwr
        End Get
    End Property

    Private _K2A As Double
    Public Sub Calc_K2A()
        _K2A = Round(_LsW - _Lwr)
    End Sub

    Public ReadOnly Property K2A()
        Get
            Return _K2A
        End Get
    End Property

    Private _LWA_Final() As Integer
    Public Sub Calc_LWAs(ByRef grus As List(Of Grid_Run_Unit), ByVal index As Integer)
        'calculating all the LWAs
        Dim topTwo As Double() = {0.0, 0.0}
        If grus.Count > 1 Then
            For i = 0 To grus.Count - 1
                Dim cur As Grid_Run_Unit = grus(i)
                cur.Calc_LWA(_K2A, _R)
                If cur.LWA > topTwo(0) Then
                    topTwo(1) = topTwo(0)
                    topTwo(0) = cur.LWA
                ElseIf cur.LWA > topTwo(1) Then
                    topTwo(1) = cur.LWA
                End If
            Next
        End If

        'calculating LWA 採用值
        Dim orig As Double = (topTwo(0) + topTwo(1)) / 2
        If Int(orig + 0.5) > Int(orig) Then
            _LWA_Final(index) = Int(orig) + 1
        Else
            _LWA_Final(index) = Int(orig)
        End If
    End Sub

    Public ReadOnly Property LWA_Final()
        Get
            Return _LWA_Final
        End Get
    End Property



    'given a list of GRUs, determine whether needing additional measuring or not
    Public Function NeedAdd(ByRef grus As List(Of Grid_Run_Unit)) As Boolean
        Dim r = New Random()
        If r.Next(2) Then
            If Program._Warning Then
                MsgBox("即將多增加一次測試，因為前幾次差距大於1!")
            End If
            Return True
        End If
        Return False
        'Dim result As Boolean = True
        'For i = 0 To grus.Count - 1
        '    Dim cur = grus(i)
        '    For j = 0 To grus.Count - 1
        '        Dim temp = grus(j)
        '        If temp IsNot cur Then
        '            If (Math.Abs(LeqK1A(temp) - LeqK1A(cur)) <= 1) Then
        '                cur.Considered = True
        '                temp.Considered = True
        '                result = False
        '            End If
        '        End If
        '    Next
        'Next
        'if result then
        'msgbox("即將多增加一次測試，因為前幾次差距大於1!")
        'end if
        '
        'Return result
    End Function

    Private Function LeqK1A(ByRef Run As Grid_Run_Unit) As Decimal
        Return Decimal.op_Explicit(Round(Run.LpAeqAvg + Run.K1A))
    End Function

    Public ChangedFromLast As Boolean = False

    Private Sub Form_CellValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Form.CellValueChanged
        ChangedFromLast = True
    End Sub

End Class


'each Meter_Measure_Unit contains measurements for one meter over one run
Public Class Meter_Measure_Unit

    Public Shared Function SeriesToMMU(ByRef series As DataVisualization.Charting.Series, ByVal Leq As Double) As Meter_Measure_Unit
        Dim mNum As Integer = CInt(series.Name)
        Dim ms As New List(Of Double)
        Dim enu = series.Points.GetEnumerator()
        'ms.Add(enu.Current.YValues(0))
        While enu.MoveNext()
            ms.Add(enu.Current.YValues(0))
        End While
        Return New Meter_Measure_Unit(mNum, ms, Leq)
    End Function

    Public Sub New(ByVal MeterNum As Integer, ByVal Measurements As List(Of Double), ByVal Leq As Double)
        _MeterNum = MeterNum
        _Measurements = Measurements
        _Leq = Leq
        If Not IsNothing(Measurements) Then
            _Time = Measurements.Count
        End If
    End Sub

    Public Shared Operator +(ByVal m1 As Meter_Measure_Unit,
                        ByVal m2 As Meter_Measure_Unit) As Meter_Measure_Unit
        Dim num As Integer = m1.MeterNum
        Dim list As List(Of Double) = New List(Of Double)
        Dim enum1 = m1.Measurements.GetEnumerator()
        list.Add(enum1.Current)
        While enum1.MoveNext
            list.Add(enum1.Current)
        End While
        enum1 = m2.Measurements.GetEnumerator()
        list.Add(enum1.Current)
        While enum1.MoveNext
            list.Add(enum1.Current)
        End While
        Dim t1 = m1.Measurements.Count
        Dim t2 = m2.Measurements.Count
        Dim leq = Grid.Round(10 * Math.Log10((t1 * 10 ^ (0.1 * m1.Leq) + (t2 * 10 ^ (0.1 * m2.Leq))) / (t1 + t2)))
        Return New Meter_Measure_Unit(num, list, leq)
    End Operator

    Private _Leq As Double
    Public Property Leq() As Double
        Get
            Return _Leq
        End Get
        Set(ByVal value As Double)
            _Leq = Grid.Round(value)
        End Set
    End Property

    Private _MeterNum As Integer
    Public Property MeterNum() As Integer
        Get
            Return _MeterNum
        End Get
        Set(ByVal value As Integer)
            _MeterNum = value
        End Set
    End Property

    Private _Measurements As List(Of Double)
    Public Property Measurements() As List(Of Double)
        Get
            Return _Measurements
        End Get
        Set(ByVal value As List(Of Double))
            _Measurements = value
        End Set
    End Property

    Private _Time As Integer
    Public Property Time() As Integer
        Get
            Return _Time
        End Get
        Set(ByVal value As Integer)
            _Time = value
        End Set
    End Property

End Class


'這個class代表的是表格上每一列的所有資訊，包含看得到的和看不到的，每個RUN_UNIT不見得搭載一個GRU，因為像A2要重複的三步驟就都聚集在同一個GRU裡面
Public Class Grid_Run_Unit
    Public Considered As Boolean = False
    Public Column As DataGridViewTextBoxColumn
    Public ParentRU As Run_Unit
    Public NextGRU As Grid_Run_Unit
    Public OverallGRU As Grid_Run_Unit 'only used in A3
    Public NotYetAccepted As Boolean = True

    Public Sub New(ByVal Header As String)
        _Header = Header
        If _Header.Contains("Background") Or _Header.Contains("RSS") Or _Header.Contains("Cal") Then
            _isRegular = False
        Else
            _isRegular = True
        End If
    End Sub

    Private Sub ClearMeterList()
        MeterList.Clear()
    End Sub

    Public Sub SetMs(ByRef Meter2 As Meter_Measure_Unit,
                   ByRef Meter4 As Meter_Measure_Unit,
                   ByRef Meter6 As Meter_Measure_Unit,
                   ByRef Meter8 As Meter_Measure_Unit,
                   ByRef Meter10 As Meter_Measure_Unit,
                   ByRef Meter12 As Meter_Measure_Unit
                   )
        If IsNothing(Meter2) Or IsNothing(Meter4) Or IsNothing(Meter6) Or IsNothing(Meter8) Or IsNothing(Meter10) Or IsNothing(Meter12) Then
            _isRegular = False
        End If
        ClearMeterList()
        _Meter2 = Meter2
        MeterList.Add(_Meter2)
        _Meter4 = Meter4
        MeterList.Add(_Meter4)
        _Meter6 = Meter6
        MeterList.Add(_Meter6)
        _Meter8 = Meter8
        MeterList.Add(_Meter8)
        _Meter10 = Meter10
        MeterList.Add(_Meter10)
        _Meter12 = Meter12
        MeterList.Add(_Meter12)
        A4 = False
        If Not IsNothing(_Meter2) Then
            _Time = _Meter2.Time
        ElseIf Not IsNothing(_Meter4) Then
            _Time = _Meter4.Time
        ElseIf Not IsNothing(_Meter6) Then
            _Time = _Meter6.Time
        ElseIf Not IsNothing(_Meter8) Then
            _Time = _Meter8.Time
        ElseIf Not IsNothing(_Meter10) Then
            _Time = _Meter10.Time
        ElseIf Not IsNothing(_Meter12) Then
            _Time = _Meter12.Time
        End If
        If _isRegular Then
            Calculate()
        Else
            Calc_LpAeqAvg()
        End If
        NotYetAccepted = True
    End Sub

    Public Sub SetMs(ByRef Meter2 As Meter_Measure_Unit,
               ByRef Meter4 As Meter_Measure_Unit,
               ByRef Meter6 As Meter_Measure_Unit,
               ByRef Meter8 As Meter_Measure_Unit
               )
        If IsNothing(Meter2) Or IsNothing(Meter4) Or IsNothing(Meter6) Or IsNothing(Meter8) Then
            _isRegular = False
        End If
        ClearMeterList()
        _Meter2 = Meter2
        MeterList.Add(_Meter2)
        _Meter4 = Meter4
        MeterList.Add(_Meter4)
        _Meter6 = Meter6
        MeterList.Add(_Meter6)
        _Meter8 = Meter8
        MeterList.Add(_Meter8)
        A4 = True
        If Not IsNothing(_Meter2) Then
            _Time = _Meter2.Time
        ElseIf Not IsNothing(_Meter4) Then
            _Time = _Meter4.Time
        ElseIf Not IsNothing(_Meter6) Then
            _Time = _Meter6.Time
        ElseIf Not IsNothing(_Meter8) Then
            _Time = _Meter8.Time
        End If
        If _isRegular Then
            Calculate()
        Else
            Calc_LpAeqAvg()
        End If
        NotYetAccepted = True
    End Sub

    Public Sub SetM(ByRef Meter As Meter_Measure_Unit, ByVal num As Integer)
        If num = 2 Then
            _Meter2 = Meter
        ElseIf num = 4 Then
            _Meter4 = Meter
        ElseIf num = 6 Then
            _Meter6 = Meter
        ElseIf num = 8 Then
            _Meter8 = Meter
        ElseIf num = 10 Then
            _Meter10 = Meter
        ElseIf num = 12 Then
            _Meter12 = Meter
        End If
        NotYetAccepted = True
    End Sub
    Public Sub Deny()
        NotYetAccepted = False
    End Sub
    Public Sub Accept()
        NotYetAccepted = False
        If _Meter2 IsNot Nothing Then
            _RealMeter2 = _Meter2
        End If
        If _Meter4 IsNot Nothing Then
            _RealMeter4 = _Meter4
        End If
        If _Meter6 IsNot Nothing Then
            _RealMeter6 = _Meter6
        End If
        If _Meter8 IsNot Nothing Then
            _RealMeter8 = _Meter8
        End If
        If _Meter10 IsNot Nothing Then
            _RealMeter10 = _Meter10
        End If
        If _Meter12 IsNot Nothing Then
            _RealMeter12 = _Meter12
        End If
        _RealdeltaLA = _deltaLA
        _RealK1A = _K1A
        _RealLpAeqAvg = _LpAeqAvg
        _RealTime = _Time
        If Me.OverallGRU IsNot Nothing Then
            OverallGRU.Accept()
        End If
    End Sub

    Private _Background As Grid_Run_Unit
    Public Property Background() As Grid_Run_Unit
        Get
            Return _Background
        End Get
        Set(ByVal value As Grid_Run_Unit)
            _Background = value
        End Set
    End Property

    Private MeterList As New List(Of Meter_Measure_Unit)
    Private RealMeterList As New List(Of Meter_Measure_Unit)

    Public A4 As Boolean

    Private _Header As String
    Public Property Header() As String
        Get
            Return _Header
        End Get
        Set(ByVal value As String)
            _Header = value
        End Set
    End Property

    Private _Subheader As String
    Public Property Subheader() As String
        Get
            Return _Subheader
        End Get
        Set(ByVal value As String)
            _Subheader = value
        End Set
    End Property

    'Real meter 和 meter 的差別在於real meter 是存已被接受過的數據，因為有時候測試過的數據看完後，不接受的話，要跳回原來的數據
    Private _Meter2 As Meter_Measure_Unit
    Private _RealMeter2 As Meter_Measure_Unit
    Public ReadOnly Property Meter2() As Meter_Measure_Unit
        Get
            If NotYetAccepted Then
                Return _Meter2
            Else
                Return _RealMeter2
            End If
        End Get
    End Property
    Public ReadOnly Property LpAeq2() As Double
        Get
            If NotYetAccepted Then
                If _Meter2 IsNot Nothing Then
                    Return _Meter2.Leq
                End If
                Return Nothing
            Else
                If _RealMeter2 IsNot Nothing Then
                    Return _RealMeter2.Leq
                End If
                Return Nothing
            End If
        End Get
    End Property


    Private _Meter4 As Meter_Measure_Unit
    Private _RealMeter4 As Meter_Measure_Unit
    Public ReadOnly Property Meter4() As Meter_Measure_Unit
        Get
            If NotYetAccepted Then
                Return _Meter4
            Else
                Return _RealMeter4
            End If
        End Get
    End Property
    Public ReadOnly Property LpAeq4() As Double
        Get
            If NotYetAccepted Then
                If _Meter4 IsNot Nothing Then
                    Return _Meter4.Leq
                End If
                Return Nothing
            Else
                If _RealMeter4 IsNot Nothing Then
                    Return _RealMeter4.Leq
                End If
                Return Nothing
            End If
        End Get
    End Property

    Private _Meter6 As Meter_Measure_Unit
    Private _RealMeter6 As Meter_Measure_Unit
    Public ReadOnly Property Meter6() As Meter_Measure_Unit
        Get
            If NotYetAccepted Then
                Return _Meter6
            Else
                Return _RealMeter6
            End If
        End Get
    End Property
    Public ReadOnly Property LpAeq6() As Double
        Get
            If NotYetAccepted Then
                If _Meter6 IsNot Nothing Then
                    Return _Meter6.Leq
                End If
                Return Nothing
            Else
                If _RealMeter6 IsNot Nothing Then
                    Return _RealMeter6.Leq
                End If
                Return Nothing
            End If
        End Get
        
    End Property

    Private _Meter8 As Meter_Measure_Unit
    Private _RealMeter8 As Meter_Measure_Unit
    Public ReadOnly Property Meter8() As Meter_Measure_Unit
        Get
            If NotYetAccepted Then
                Return _Meter8
            Else
                Return _RealMeter8
            End If
        End Get
    End Property
    Public ReadOnly Property LpAeq8() As Double
        Get
            If NotYetAccepted Then
                If _Meter8 IsNot Nothing Then
                    Return _Meter8.Leq
                End If
                Return Nothing
            Else
                If _RealMeter8 IsNot Nothing Then
                    Return _RealMeter8.Leq
                End If
                Return Nothing
            End If
        End Get
    End Property

    Private _Meter10 As Meter_Measure_Unit
    Private _RealMeter10 As Meter_Measure_Unit
    Public ReadOnly Property Meter10() As Meter_Measure_Unit
        Get
            If NotYetAccepted Then
                Return _Meter10
            Else
                Return _RealMeter10
            End If
        End Get
    End Property
    Public ReadOnly Property LpAeq10() As Double
        Get
            If NotYetAccepted Then
                If _Meter10 IsNot Nothing Then
                    Return _Meter10.Leq
                End If
                Return Nothing
            Else
                If _RealMeter10 IsNot Nothing Then
                    Return _RealMeter10.Leq
                End If
                Return Nothing
            End If
        End Get
    End Property

    Private _Meter12 As Meter_Measure_Unit
    Private _RealMeter12 As Meter_Measure_Unit
    Public ReadOnly Property Meter12() As Meter_Measure_Unit
        Get
            If NotYetAccepted Then
                Return _Meter12
            Else
                Return _RealMeter12
            End If
        End Get
    End Property
    Public ReadOnly Property LpAeq12() As Double
        Get
            If NotYetAccepted Then
                If _Meter12 IsNot Nothing Then
                    Return _Meter12.Leq
                End If
                Return Nothing
            Else
                If _RealMeter12 IsNot Nothing Then
                    Return _RealMeter12.Leq
                End If
                Return Nothing
            End If
        End Get
    End Property

    Private Sub Calculate()
        Calc_LpAeqAvg()
        Calc_deltaLA()
        Calc_K1A()
    End Sub

    Public Sub Calc_LpAeqAvg()
        Try
            Dim num = 6
            If A4 Then
                num = 4
            End If
            Dim total As Double = 0.0
            For i = 0 To num - 1
                total += 10 ^ (0.1 * MeterList(i).Leq)
            Next
            _LpAeqAvg = Grid.Round(10 * Math.Log10(total / num))

        Catch ex As Exception
            MsgBox("In Calc_LpAeqAvg():" & ex.Message)
            _LpAeqAvg = -1
        End Try
    End Sub

    Private _LpAeqAvg As Double
    Private _RealLpAeqAvg As Double
    Public ReadOnly Property LpAeqAvg() As Double
        Get
            If NotYetAccepted Then
                Return _LpAeqAvg
            Else
                Return _RealLpAeqAvg
            End If
        End Get
    End Property

    Private _deltaLA As Double
    Private _RealdeltaLA As Double
    Public ReadOnly Property deltaLA() As Double
        Get
            If NotYetAccepted Then
                Return _deltaLA
            Else
                Return _RealdeltaLA
            End If
        End Get
    End Property

    Public Sub Calc_deltaLA()
        Try
            If IsNothing(_Background) Then
                _deltaLA = 0
                Return
            End If
            If _Background.LpAeqAvg = 0 Then
                _deltaLA = 0
                Return
            End If
            _deltaLA = Grid.Round(Me._LpAeqAvg - _Background.LpAeqAvg)
        Catch ex As Exception
            MsgBox("In Calc_deltaLA():" & ex.Message)
            _deltaLA = -1
        End Try
    End Sub

   


    Public Sub Calc_K1A()
        If isRegular Then

            If _deltaLA = 0 Or _deltaLA > 10 Then
                _K1A = 0
                Return
            ElseIf _deltaLA < 3 Then
                If Program._Warning Then
                    MsgBox("ΔLA < 3dB 所以背景噪音修正值K1A無法計算!")
                End If
                Return
            End If
            Try
                _K1A = Grid.Round(-10 * Math.Log10(1 - 10 ^ (-0.1 * _deltaLA)))
            Catch ex As Exception
                MsgBox("In Calc_K1A():" & ex.Message)
                _K1A = -1
            End Try
        End If
    End Sub

    Private _K1A As Double
    Private _RealK1A As Double
    Public ReadOnly Property K1A() As Double
        Get
            If NotYetAccepted Then
                Return _K1A
            Else
                Return _RealK1A
            End If
        End Get
    End Property

    Private _isRegular As Boolean
    Public ReadOnly Property isRegular() As Boolean
        Get
            Return _isRegular
        End Get
    End Property

    Private _Time As Integer
    Private _RealTime As Integer
    Public ReadOnly Property Time() As Integer
        Get
            If NotYetAccepted Then
                Return _Time
            Else
                Return _RealTime
            End If
        End Get
    End Property

    Private _LWA As Double
    Public ReadOnly Property LWA() As Double
        Get
            Return _LWA
        End Get
    End Property

    Public Function Calc_LWA(ByVal K2A As Double, ByVal r As Double)
        Try
            If isRegular Then
                _LWA = Grid.Round(LpAeqAvg - K1A - K2A + 10 * Math.Log10(2 * Math.PI * r * r))
                Return _LWA
            End If
        Catch ex As Exception
            MsgBox("In Calc_LWA():" & ex.Message)
        End Try
        Return -1
    End Function

End Class