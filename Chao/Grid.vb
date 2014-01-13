Public Class Grid

End Class

Public Class A_Unit
    'contains a grid
    Public Grid As DataGridView

    Public Sub New(ByRef Parent As Control, ByVal Size As Size, ByVal Position As Point)
        Grid = New DataGridView()
        Grid.Parent = Parent
        Grid.Size = Size
        Grid.Location = Position
        Dim col As New DataGridViewTextBoxColumn
        Grid.Columns.Add(col)
        'need to add a column before rows

        Grid.Rows.Add(16)
        Grid.Rows(1).HeaderCell.Value = "LpAeq2"
        Grid.Rows(2).HeaderCell.Value = "LpAeq4"
        Grid.Rows(3).HeaderCell.Value = "LpAeq6"
        Grid.Rows(4).HeaderCell.Value = "LpAeq8"
        Grid.Rows(5).HeaderCell.Value = "LpAeq10"
        Grid.Rows(6).HeaderCell.Value = "LpAeq12"
        Grid.Rows(7).HeaderCell.Value = "LpAeq avg"
        Grid.Rows(8).HeaderCell.Value = "Time(sec)"
        Grid.Rows(9).HeaderCell.Value = "deltaLA"
        Grid.Rows(10).HeaderCell.Value = "K1A"
        Grid.Rows(11).HeaderCell.Value = "L*W"
        Grid.Rows(12).HeaderCell.Value = "LW"
        Grid.Rows(13).HeaderCell.Value = "K2A"
        Grid.Rows(14).HeaderCell.Value = "LWA"
        Grid.Rows(15).HeaderCell.Value = "LWA 採用"
        'Grid.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)
    End Sub

    'contains grid_run_units
    Private GRUList As New List(Of Grid_Run_Unit)
    Private RSS As Grid_Run_Unit

    Public Enum Run_Modes
        Background
        Regular
        RSS
    End Enum

    Public Sub AddGrid_Run_Unit(ByRef Run As Grid_Run_Unit)

        If Run.Header = Run_Modes.RSS.ToString() Then
            RSS = Run
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

        Dim curColInd As Integer = Grid.Columns.Count - 1
        Grid.Columns(curColInd).HeaderText = Run.Header
        'subHeader
        Grid.Rows(0).Cells(curColInd).Value = Run.Subheader

        'meter 2
        Grid.Rows(1).Cells(curColInd).Value = Run.LpAeq2
        'meter 4
        Grid.Rows(2).Cells(curColInd).Value = Run.LpAeq4
        'meter 6
        Grid.Rows(3).Cells(curColInd).Value = Run.LpAeq6
        'meter 8
        Grid.Rows(4).Cells(curColInd).Value = Run.LpAeq8
        'meter 10
        Grid.Rows(5).Cells(curColInd).Value = Run.LpAeq10
        'meter 12
        Grid.Rows(6).Cells(curColInd).Value = Run.LpAeq12
        'meters average
        Grid.Rows(7).Cells(curColInd).Value = Run.LpAeqAvg
        'time
        Grid.Rows(8).Cells(curColInd).Value = Run.Time
        If Run.isRegular Then
            'deltaA
            Grid.Rows(9).Cells(curColInd).Value = Run.deltaLA
            'K1A
            Grid.Rows(10).Cells(curColInd).Value = Run.K1A
        End If

        'add next column for next record
        Dim col As New DataGridViewTextBoxColumn
        'col.DataPropertyName = "Run"
        'col.Name = "colWhateverName"
        Grid.Columns.Add(col)
        GRUList.Add(Run)
    End Sub

    'contains calculations
    Public Sub Calculate()
        Calc_LsW()
        Calc_K2A()
        Calc_LWA_Final()
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

    '###TODO
    Private _LWA_Final As Double
    Public Sub Calc_LWA_Final()
        _LWA_Final = 1
    End Sub

    Public ReadOnly Property LWA_Final()
        Get
            Return _LWA_Final
        End Get
    End Property



    'determine whether needing additional measuring or not
    Public Function NeedAdd()
        For i = 0 To GRUList.Count - 1
            Dim cur = GRUList(i)
            If cur.isRegular Then
                For j = 0 To GRUList.Count - 1
                    Dim temp = GRUList(j)
                    If temp IsNot cur And temp.isRegular Then
                        If (Math.Abs(LeqK1A(temp) - LeqK1A(cur)) <= 1) Then

                            Return False
                        End If
                    End If
                Next
            End If
        Next
        Return True
    End Function

    Private Function LeqK1A(ByRef Run As Grid_Run_Unit) As Decimal
        Return Decimal.op_Explicit(Run.LpAeqAvg + Run.K1A)
    End Function

    'contains recording


End Class


'each Meter_Measure_Unit contains measurements for one meter over one run
Public Class Meter_Measure_Unit

    Public Sub New(ByVal MeterNum As Integer, ByVal Measurements As List(Of Double), ByVal Leq As Double)
        _MeterNum = MeterNum
        _Measurements = Measurements
        _Leq = Leq
        If Not IsNothing(Measurements) Then
            _Time = Measurements.Count
        End If
    End Sub

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

    Public Sub New(ByVal Meter2 As Meter_Measure_Unit,
                   ByVal Meter4 As Meter_Measure_Unit,
                   ByVal Meter6 As Meter_Measure_Unit,
                   ByVal Meter8 As Meter_Measure_Unit,
                   ByVal Meter10 As Meter_Measure_Unit,
                   ByVal Meter12 As Meter_Measure_Unit,
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
        If _Header = A_Unit.Run_Modes.Background.ToString() Or _Subheader = A_Unit.Run_Modes.Background.ToString() _
        Or _Header = A_Unit.Run_Modes.RSS.ToString() Or _Header = A_Unit.Run_Modes.RSS.ToString() Then
            _isRegular = False
        End If
        
        Calculate()
    End Sub

    Public Sub New(ByVal Meter2 As Meter_Measure_Unit,
               ByVal Meter4 As Meter_Measure_Unit,
               ByVal Meter6 As Meter_Measure_Unit,
               ByVal Meter8 As Meter_Measure_Unit,
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
        If _Header = A_Unit.Run_Modes.Background.ToString() Or _Subheader = A_Unit.Run_Modes.Background.ToString() _
        Or _Header = A_Unit.Run_Modes.RSS.ToString() Or _Header = A_Unit.Run_Modes.RSS.ToString() Then
            _isRegular = False
        End If
        Calculate()
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
    Public Property LpAeq2() As Double
        Get
            Return _Meter2.Leq
        End Get
        Set(ByVal value As Double)
            _Meter2.Leq = value
        End Set
    End Property

    Private _Meter4 As Meter_Measure_Unit
    Public Property LpAeq4() As Double
        Get
            Return _Meter4.Leq
        End Get
        Set(ByVal value As Double)
            _Meter4.Leq = value
        End Set
    End Property

    Private _Meter6 As Meter_Measure_Unit
    Public Property LpAeq6() As Double
        Get
            Return _Meter6.Leq
        End Get
        Set(ByVal value As Double)
            _Meter6.Leq = value
        End Set
    End Property

    Private _Meter8 As Meter_Measure_Unit
    Public Property LpAeq8() As Double
        Get
            Return _Meter8.Leq
        End Get
        Set(ByVal value As Double)
            _Meter8.Leq = value
        End Set
    End Property

    Private _Meter10 As Meter_Measure_Unit
    Public Property LpAeq10() As Double
        Get
            Return _Meter10.Leq
        End Get
        Set(ByVal value As Double)
            _Meter10.Leq = value
        End Set
    End Property

    Private _Meter12 As Meter_Measure_Unit
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