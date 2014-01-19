﻿Imports System.IO
Imports System.Text


Public Class Grid
    'contains a grid
    Public Form As DataGridView
    Private _R As Double
    Private _Machine As Program.Machines
    'Public Cur_Column

    'Private A_s As Integer 'keep track of which A's this one has
    'Private As_Last_Indices() As Integer ' index of last column of each A run
    'Enum Excavator_As
    '    A1
    '    A2
    'End Enum

    'Enum Loader_As
    '    A1
    '    A2
    '    A3
    'End Enum

    'Enum Tractor_As
    '    A1
    '    A3
    'End Enum

    'Enum LoadExcavator_As
    '    A1e
    '    A2e
    '    A1l
    '    A2l
    '    A3l
    'End Enum

    Public Sub New(ByRef Parent As Control, ByVal Size As Size, ByVal Position As Point, ByVal R As Double, ByVal Machine As Program.Machines)
        Form = New DataGridView()
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

        'Grid.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)

        'Set A sequence
        If Machine = Program.Machines.Excavator Then 'A1 A2
            'A_s = 2
            'As_Last_Indices = {3, 7}
            For i = 0 To 9
                Form.Columns.Add(New DataGridViewTextBoxColumn())
            Next
            Form.Columns(0).HeaderText = "Background"
            Form.Columns(1).HeaderText = "A1"
            AddLabelColumn(1)
            Form.Columns(2).HeaderText = "Run1"
            Form.Columns(3).HeaderText = "Run2"
            Form.Columns(4).HeaderText = "Run3"

            Form.Columns(5).HeaderText = "A2"
            AddLabelColumn(5)
            Form.Columns(6).HeaderText = "Run1"
            Form.Columns(7).HeaderText = "Run2"
            Form.Columns(8).HeaderText = "Run3"
            Form.Columns(9).HeaderText = "RSS"
            
        ElseIf Machine = Program.Machines.Loader Then 'A1 A2 A3
            'A_s = 3

            For i = 0 To 19
                Form.Columns.Add(New DataGridViewTextBoxColumn())
            Next
            Form.Columns(0).HeaderText = "Background"
            Form.Columns(1).HeaderText = "A1"
            AddLabelColumn(1)
            Form.Columns(2).HeaderText = "Run1"
            Form.Columns(3).HeaderText = "Run2"
            Form.Columns(4).HeaderText = "Run3"

            Form.Columns(5).HeaderText = "A2"
            AddLabelColumn(5)
            Form.Columns(6).HeaderText = "Run1"
            Form.Columns(7).HeaderText = "Run2"
            Form.Columns(8).HeaderText = "Run3"

            Form.Columns(9).HeaderText = "A3"
            AddLabelColumn(9)
            Form.Columns(10).HeaderText = "Run1"
            Form.Rows(0).Cells(10).Value = "前進"
            Form.Columns(11).HeaderText = "Run1"
            Form.Rows(0).Cells(11).Value = "後退"
            Form.Columns(12).HeaderText = "Run1"
            Form.Rows(0).Cells(12).Value = "平均"

            Form.Columns(13).HeaderText = "Run2"
            Form.Rows(0).Cells(13).Value = "前進"
            Form.Columns(14).HeaderText = "Run2"
            Form.Rows(0).Cells(14).Value = "後退"
            Form.Columns(15).HeaderText = "Run2"
            Form.Rows(0).Cells(15).Value = "平均"

            Form.Columns(16).HeaderText = "Run3"
            Form.Rows(0).Cells(16).Value = "前進"
            Form.Columns(17).HeaderText = "Run3"
            Form.Rows(0).Cells(17).Value = "後退"
            Form.Columns(18).HeaderText = "Run3"
            Form.Rows(0).Cells(18).Value = "平均"
            Form.Columns(19).HeaderText = "RSS"
        ElseIf Machine = Program.Machines.Tractor Then 'A1 A3
            'A_s = 2
            For i = 0 To 15
                Form.Columns.Add(New DataGridViewTextBoxColumn())
            Next
            Form.Columns(0).HeaderText = "Background"
            Form.Columns(1).HeaderText = "A1"
            AddLabelColumn(1)
            Form.Columns(2).HeaderText = "Run1"
            Form.Columns(3).HeaderText = "Run2"
            Form.Columns(4).HeaderText = "Run3"

            Form.Columns(5).HeaderText = "A3"
            AddLabelColumn(5)
            Form.Columns(6).HeaderText = "Run1"
            Form.Rows(0).Cells(6).Value = "前進"
            Form.Columns(7).HeaderText = "Run1"
            Form.Rows(0).Cells(7).Value = "後退"
            Form.Columns(8).HeaderText = "Run1"
            Form.Rows(0).Cells(8).Value = "平均"

            Form.Columns(9).HeaderText = "Run2"
            Form.Rows(0).Cells(9).Value = "前進"
            Form.Columns(10).HeaderText = "Run2"
            Form.Rows(0).Cells(10).Value = "後退"
            Form.Columns(11).HeaderText = "Run2"
            Form.Rows(0).Cells(11).Value = "平均"

            Form.Columns(12).HeaderText = "Run3"
            Form.Rows(0).Cells(12).Value = "前進"
            Form.Columns(13).HeaderText = "Run3"
            Form.Rows(0).Cells(13).Value = "後退"
            Form.Columns(14).HeaderText = "Run3"
            Form.Rows(0).Cells(14).Value = "平均"
            Form.Columns(15).HeaderText = "RSS"
        ElseIf Machine = Program.Machines.Loader_Excavator Then 'A1 A2, A1 A2 A3
            'A_s = 5
            For i = 0 To 28
                Form.Columns.Add(New DataGridViewTextBoxColumn())
            Next
            Form.Columns(0).HeaderText = "Background"
            Form.Columns(1).HeaderText = "A1"
            AddLabelColumn(1)
            Form.Columns(2).HeaderText = "Run1"
            Form.Columns(3).HeaderText = "Run2"
            Form.Columns(4).HeaderText = "Run3"

            Form.Columns(5).HeaderText = "A2"
            AddLabelColumn(5)
            Form.Columns(6).HeaderText = "Run1"
            Form.Columns(7).HeaderText = "Run2"
            Form.Columns(8).HeaderText = "Run3"

            Form.Columns(9).HeaderText = "A1"
            AddLabelColumn(9)
            Form.Columns(10).HeaderText = "Run1"
            Form.Columns(11).HeaderText = "Run2"
            Form.Columns(12).HeaderText = "Run3"

            Form.Columns(13).HeaderText = "A2"
            AddLabelColumn(13)
            Form.Columns(14).HeaderText = "Run1"
            Form.Columns(15).HeaderText = "Run2"
            Form.Columns(16).HeaderText = "Run3"

            Form.Columns(17).HeaderText = "A3"
            AddLabelColumn(17)
            Form.Columns(18).HeaderText = "Run1"
            Form.Rows(0).Cells(19).Value = "前進"
            Form.Columns(19).HeaderText = "Run1"
            Form.Rows(0).Cells(20).Value = "後退"
            Form.Columns(21).HeaderText = "Run1"
            Form.Rows(0).Cells(21).Value = "平均"

            Form.Columns(22).HeaderText = "Run2"
            Form.Rows(0).Cells(22).Value = "前進"
            Form.Columns(23).HeaderText = "Run2"
            Form.Rows(0).Cells(23).Value = "後退"
            Form.Columns(24).HeaderText = "Run2"
            Form.Rows(0).Cells(24).Value = "平均"

            Form.Columns(25).HeaderText = "Run3"
            Form.Rows(0).Cells(25).Value = "前進"
            Form.Columns(26).HeaderText = "Run3"
            Form.Rows(0).Cells(26).Value = "後退"
            Form.Columns(27).HeaderText = "Run3"
            Form.Rows(0).Cells(27).Value = "平均"
            Form.Columns(28).HeaderText = "RSS"
        ElseIf Machine = Program.Machines.Others Then 'A4
            'A_s = 1
            For i = 0 To 6
                Form.Columns.Add(New DataGridViewTextBoxColumn())
            Next
            Form.Columns(0).HeaderText = "Background"
            Form.Columns(1).HeaderText = "Run1"
            Form.Columns(2).HeaderText = "Background"
            Form.Columns(3).HeaderText = "Run2"
            Form.Columns(4).HeaderText = "Background"
            Form.Columns(5).HeaderText = "Run3"
            Form.Columns(6).HeaderText = "RSS"
        End If
    End Sub

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
    Private GRUList As New List(Of Grid_Run_Unit)
    Private RSS As Grid_Run_Unit

    Public Enum Run_Modes
        Background
        Regular
        RSS
    End Enum

    'link the grid_run_unit's column to the column indexed in form
    Public Sub LinkGrid_Run_Unit(ByRef run As Grid_Run_Unit, ByVal index As Integer)
        run.Column = Form.Columns(index)
    End Sub

    'show GRU on the Form so we can see figures
    Public Sub ShowGRUonForm(ByRef Run As Grid_Run_Unit)

        If Run.Header = Run_Modes.RSS.ToString() Then
            RSS = Run
        End If
        'If the first added is not background then throw error message and return
        If GRUList.Count = 0 And Run.isRegular Then
            MsgBox("Trying to add a regular record to the first column where it should be a background record!")
            Return
        End If
        'If this run is not background
        If Run.isRegular Then
            'see if previous was background then use it as this one's background for A4
            If Not GRUList.Count = 0 Then
                If GRUList(GRUList.Count - 1).isRegular Then
                    Run.Background = GRUList(GRUList.Count - 1)
                End If
            Else 'else we set the background to be the first in line for A1,A2,A3
                Run.Background = GRUList(0)
            End If
        End If

        Dim curColInd As Integer = Run.Column.Index
        'Form.Columns(curColInd).HeaderText = Run.Header
        ''subHeader
        'Form.Rows(0).Cells(curColInd).Value = Run.Subheader

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
        GRUList.Add(Run)

        '##TODO
        Dim value = Form.Rows(0).Cells(Run.Column.Index).Value
        If Not IsNothing(value) Then
            If value.Contains("後退") Then
                AddA3Overall(Run)
            End If
        End If

    End Sub

    'function that adds the overall column after forward and backward measurements are in place
    Private Sub AddA3Overall(ByRef Run As Grid_Run_Unit)
        If Not IsNothing(GRUList) Then
            If GRUList.Count > 3 Then
                Dim forw As Grid_Run_Unit = GRUList(GRUList.Count - 2)
                Dim backw As Grid_Run_Unit = GRUList(GRUList.Count - 1)
                Dim subheader = "平均"
                Dim gru As Grid_Run_Unit = New Grid_Run_Unit(forw.Meter2 + backw.Meter2,
                                                             forw.Meter4 + backw.Meter4,
                                                             forw.Meter6 + backw.Meter6,
                                                             forw.Meter8 + backw.Meter8,
                                                             forw.Meter10 + backw.Meter10,
                                                             forw.Meter12 + backw.Meter12,
                                                             forw.Header, forw.Subheader)

                Dim curColInd As Integer = Run.Column.Index
                'Form.Columns(curColInd).HeaderText = gru.Header
                ''subHeader
                'Form.Rows(0).Cells(curColInd).Value = gru.Subheader

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
                'add next column for next record
                'Dim col As New DataGridViewTextBoxColumn
                'col.Name = gru.Header & gru.Subheader
                'Form.Columns.Add(col)
                GRUList.Add(gru)
            End If
        End If
    End Sub

    'call this after RSS has been done
    Public Sub ShowCalculated(ByVal StartIndex As Integer, ByVal Length As Integer)
        Calculate()
        'A1 and A2
        'LsW row
        Form.Rows(11).Cells(Form.Columns.Count - 1).Value = _LsW
        'Lw row
        Form.Rows(12).Cells(Form.Columns.Count - 1).Value = _Lwr
        'K2A row
        Form.Rows(13).Cells(Form.Columns.Count - 1).Value = _K2A
        'LWA choice row
        Form.Rows(15).Cells(Form.Columns.Count - 1).Value = _LWA_Final

        'LWA 
        'A1, A2
        'If A < 3 Then
        '    If GRUList.Count > 1 Then
        '        For i = 1 To Form.Columns.Count - 2 'all the regulars
        '            Form.Rows(14).Cells(i).Value = GRUList(i).LWA
        '        Next
        '    End If
        'End If

        ''A3
        'If A = 3 Then
        '    If GRUList.Count > 1 Then
        '        For i = 3 To GRUList.Count - 2 Step 3
        '            Form.Rows(14).Cells(i).Value = GRUList(i).LWA
        '        Next
        '    End If

        'End If

        ''A4
        'If A = 4 Then
        '    If GRUList.Count > 1 Then
        '        For i = 1 To GRUList.Count - 2 Step 2
        '            Form.Rows(14).Cells(i).Value = GRUList(i).LWA
        '        Next
        '    End If
        'End If
    End Sub

    'contains calculations- NeedAdd should be called before (Calc_LWAs)Calculate
    Public Sub Calculate()
        If Not IsNothing(RSS) Then
            Calc_LsW()
            Calc_K2A()
            Calc_LWAs()
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
        _K2A = _LsW - _Lwr
    End Sub

    Public ReadOnly Property K2A()
        Get
            Return _K2A
        End Get
    End Property

    Private _LWA_Final As Integer
    Public Sub Calc_LWAs()
        'calculating all the LWAs
        Dim topTwo As Double() = {0.0, 0.0}
        If Not GRUList.Count > 1 Then
            For i = 1 To GRUList.Count - 2
                Dim cur As Grid_Run_Unit = GRUList(i)
                cur.Calc_LWA(_K2A, _R)
                If cur.Considered Then
                    If cur.LWA > topTwo(0) Then
                        topTwo(0) = cur.LWA
                    ElseIf cur.LWA > topTwo(1) Then
                        topTwo(1) = cur.LWA
                    End If
                End If
            Next
        End If

        'calculating LWA 採用值
        Dim orig As Double = (topTwo(0) + topTwo(1)) / 2
        If Int(orig + 0.5) > Int(orig) Then
            _LWA_Final = Int(orig) + 1
        Else
            _LWA_Final = Int(orig)
        End If
    End Sub

    Public ReadOnly Property LWA_Final()
        Get
            Return _LWA_Final
        End Get
    End Property



    'determine whether needing additional measuring or not
    Public Function NeedAdd()
        Dim result As Boolean = True
        For i = 0 To GRUList.Count - 1
            Dim cur = GRUList(i)
            If cur.isRegular Then
                'If Not A = 3 Then 'A1, A2, and A4
                '    For j = 1 To GRUList.Count - 1
                '        Dim temp = GRUList(j)
                '        If temp IsNot cur And temp.isRegular Then
                '            If (Math.Abs(LeqK1A(temp) - LeqK1A(cur)) <= 1) Then
                '                cur.Considered = True
                '                temp.Considered = True
                '                result = False
                '            End If
                '        End If
                '    Next
                'Else 'for A3 because it has forward and backward and overall
                '    For j = 3 To GRUList.Count - 1 Step 3
                '        Dim temp = GRUList(j)
                '        If temp IsNot cur And temp.isRegular Then
                '            If (Math.Abs(LeqK1A(temp) - LeqK1A(cur)) <= 1) Then
                '                cur.Considered = True
                '                temp.Considered = True
                '                result = False
                '            End If
                '        End If
                '    Next
                'End If
            End If
        Next
        Return result
    End Function

    Private Function LeqK1A(ByRef Run As Grid_Run_Unit) As Decimal
        Return Decimal.op_Explicit(Run.LpAeqAvg + Run.K1A)
    End Function

    'contains recording
    Public Sub Save()

        Dim saveFileDialog1 As New SaveFileDialog()

        saveFileDialog1.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*"
        saveFileDialog1.RestoreDirectory = True
        Dim outfile As StreamWriter

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            outfile = New StreamWriter(saveFileDialog1.OpenFile(), System.Text.Encoding.Unicode)
            If (outfile IsNot Nothing) Then
                'write column headers first
                Dim sb As StringBuilder = New StringBuilder()
                For i = 0 To Form.Columns.Count - 1
                    sb.Append("," & Form.Columns(i).HeaderText)
                Next
                outfile.WriteLine(sb.ToString())
                sb.Clear()

                'write actual data now
                For j = 0 To Form.Rows.Count - 1
                    sb.Append(Form.Rows(j).HeaderCell.Value)
                    For i = 0 To Form.Columns.Count - 1
                        sb.Append("," & Form.Rows(j).Cells(i).Value)
                    Next
                    sb.AppendLine()
                Next
                outfile.Write(sb.ToString())
            End If
            outfile.Close()
        End If

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
        Dim leq = 10 * Math.Log10((t1 * 10 ^ (0.1 * m1.Leq) + (t2 * 10 ^ (0.1 * m2.Leq))) / (t1 + t2))
        Return New Meter_Measure_Unit(num, list, leq)
    End Operator

    Private _Leq As Double
    Public Property Leq() As Double
        Get
            Return _Leq
        End Get
        Set(ByVal value As Double)
            _Leq = value
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



Public Class Grid_Run_Unit
    Public Considered As Boolean = False
    Public Column As DataGridViewTextBoxColumn

    '
    Public Sub New(ByVal Col As DataGridViewTextBoxColumn, ByVal isReg As Boolean)
        Column = Col
        _isRegular = isReg
    End Sub

    Public Sub New(ByRef Meter2 As Meter_Measure_Unit,
                   ByRef Meter4 As Meter_Measure_Unit,
                   ByRef Meter6 As Meter_Measure_Unit,
                   ByRef Meter8 As Meter_Measure_Unit,
                   ByRef Meter10 As Meter_Measure_Unit,
                   ByRef Meter12 As Meter_Measure_Unit,
                   ByVal Header As String,
                   ByVal Subheader As String
                   )
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

        _Header = Header
        _Subheader = Subheader
        If Not IsNothing(_Meter2) Then
            _Time = _Meter2.Time
        End If
        _isRegular = True
        If _Header = Grid.Run_Modes.Background.ToString() Or _Subheader = Grid.Run_Modes.Background.ToString() _
        Or _Header = Grid.Run_Modes.RSS.ToString() Or _Header = Grid.Run_Modes.RSS.ToString() Then
            _isRegular = False
        End If

        Calculate()
    End Sub

    Public Sub New(ByRef Meter2 As Meter_Measure_Unit,
               ByRef Meter4 As Meter_Measure_Unit,
               ByRef Meter6 As Meter_Measure_Unit,
               ByRef Meter8 As Meter_Measure_Unit,
               ByVal Header As String,
               ByVal Subheader As String
               )
        _Meter2 = Meter2
        MeterList.Add(_Meter2)
        _Meter4 = Meter4
        MeterList.Add(_Meter4)
        _Meter6 = Meter6
        MeterList.Add(_Meter6)
        _Meter8 = Meter8
        MeterList.Add(_Meter8)
        A4 = True

        _Header = Header
        _Subheader = Subheader
        If Not IsNothing(_Meter2) Then
            _Time = _Meter2.Time
        End If
        _isRegular = True
        If _Header = Grid.Run_Modes.Background.ToString() Or _Subheader = Grid.Run_Modes.Background.ToString() _
        Or _Header = Grid.Run_Modes.RSS.ToString() Or _Header = Grid.Run_Modes.RSS.ToString() Then
            _isRegular = False
        End If
        Calculate()
    End Sub

    Public Sub SetMs(ByRef Meter2 As Meter_Measure_Unit,
                   ByRef Meter4 As Meter_Measure_Unit,
                   ByRef Meter6 As Meter_Measure_Unit,
                   ByRef Meter8 As Meter_Measure_Unit,
                   ByRef Meter10 As Meter_Measure_Unit,
                   ByRef Meter12 As Meter_Measure_Unit,
                   ByVal Header As String,
                   ByVal Subheader As String
                   )
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
        End If
        _Header = Header
        _Subheader = Subheader
    End Sub

    Public Sub SetMs(ByRef Meter2 As Meter_Measure_Unit,
               ByRef Meter4 As Meter_Measure_Unit,
               ByRef Meter6 As Meter_Measure_Unit,
               ByRef Meter8 As Meter_Measure_Unit,
               ByVal Header As String,
               ByVal Subheader As String
               )
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
        End If
        _Header = Header
        _Subheader = Subheader
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

    Public A4 As Boolean

    Private _isRegular As Boolean
    Public ReadOnly Property isRegular() As Boolean
        Get
            Return _isRegular
        End Get
    End Property

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

    Private _Meter2 As Meter_Measure_Unit
    Public ReadOnly Property Meter2() As Meter_Measure_Unit
        Get
            Return _Meter2
        End Get
    End Property
    Public Property LpAeq2() As Double
        Get
            Return _Meter2.Leq
        End Get
        Set(ByVal value As Double)
            _Meter2.Leq = value
        End Set
    End Property


    Private _Meter4 As Meter_Measure_Unit
    Public ReadOnly Property Meter4() As Meter_Measure_Unit
        Get
            Return _Meter4
        End Get
    End Property
    Public Property LpAeq4() As Double
        Get
            Return _Meter4.Leq
        End Get
        Set(ByVal value As Double)
            _Meter4.Leq = value
        End Set
    End Property

    Private _Meter6 As Meter_Measure_Unit
    Public ReadOnly Property Meter6() As Meter_Measure_Unit
        Get
            Return _Meter6
        End Get
    End Property
    Public Property LpAeq6() As Double
        Get
            Return _Meter6.Leq
        End Get
        Set(ByVal value As Double)
            _Meter6.Leq = value
        End Set
    End Property

    Private _Meter8 As Meter_Measure_Unit
    Public ReadOnly Property Meter8() As Meter_Measure_Unit
        Get
            Return _Meter8
        End Get
    End Property
    Public Property LpAeq8() As Double
        Get
            Return _Meter8.Leq
        End Get
        Set(ByVal value As Double)
            _Meter8.Leq = value
        End Set
    End Property

    Private _Meter10 As Meter_Measure_Unit
    Public ReadOnly Property Meter10() As Meter_Measure_Unit
        Get
            Return _Meter10
        End Get
    End Property
    Public Property LpAeq10() As Double
        Get
            Return _Meter10.Leq
        End Get
        Set(ByVal value As Double)
            _Meter10.Leq = value
        End Set
    End Property

    Private _Meter12 As Meter_Measure_Unit
    Public ReadOnly Property Meter12() As Meter_Measure_Unit
        Get
            Return _Meter12
        End Get
    End Property
    Public Property LpAeq12() As Double
        Get
            Return _Meter12.Leq
        End Get
        Set(ByVal value As Double)
            _Meter12.Leq = value
        End Set
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
            _LpAeqAvg = 10 * Math.Log10(total / num)

        Catch ex As Exception
            MsgBox("In Calc_LpAeqAvg():" & ex.Message)
            _LpAeqAvg = -1
        End Try
    End Sub

    Private _LpAeqAvg As Double
    Public ReadOnly Property LpAeqAvg() As Double
        Get
            Return _LpAeqAvg
        End Get
    End Property

    Private _deltaLA As Double
    Public Sub Calc_deltaLA()
        If isRegular Then
            Try
                If IsNothing(_Background) Then
                    _deltaLA = 0
                    Return
                End If
                If _Background._LpAeqAvg = 0 Then
                    _deltaLA = 0
                    Return
                End If
                _deltaLA = Me._LpAeqAvg - _Background._LpAeqAvg
            Catch ex As Exception
                MsgBox("In Calc_deltaLA():" & ex.Message)
                _deltaLA = -1
            End Try
        End If
    End Sub

    Public ReadOnly Property deltaLA() As Double
        Get
            Return _deltaLA
        End Get
    End Property


    Public Sub Calc_K1A()
        If isRegular Then
            Try
                If _deltaLA = 0 Or _deltaLA > 10 Then
                    _K1A = 0
                    Return
                ElseIf _deltaLA < 3 Then
                    MsgBox("deltaLA < 3dB 所以背景噪音修正值K1A無法計算!")
                    Return
                End If
                _K1A = -10 * Math.Log10(1 - 10 ^ (-0.1 * _deltaLA))
            Catch ex As Exception
                MsgBox("In Calc_K1A():" & ex.Message)
                _K1A = -1
            End Try
        End If
    End Sub

    Private _K1A As Double
    Public ReadOnly Property K1A() As Double
        Get
            Return _K1A
        End Get
    End Property


    Private _LWA As Double
    Public Function Calc_LWA(ByVal K2A As Double, ByVal r As Double)
        Try
            If isRegular Then
                _LWA = _LpAeqAvg - _K1A - K2A + 10 * Math.Log10(2 * Math.PI * r * r)
                Return _LWA
            End If
        Catch ex As Exception
            MsgBox("In Calc_LWA():" & ex.Message)
            Return -1
        End Try
    End Function

    Public ReadOnly Property LWA() As Double
        Get
            Return _LWA
        End Get
    End Property

    Private _Time As Integer
    Public ReadOnly Property Time() As Integer
        Get
            Return _Time
        End Get
    End Property

End Class