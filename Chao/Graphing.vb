Imports Microsoft.VisualBasic
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.VisualBasic.PowerPacks

Public Class CGraph
    'protected global variables
    Protected chart As Chart
    Protected series As System.Collections.Generic.List(Of Series)
    Protected legend As Legend
    Protected TestMode As Modes
    Protected NumOfMics As Integer
    Protected Panel As Panel
    Protected SecTextFontSize As Integer = 12
    Protected SecTextBoxSize As Size = New Size(50, 20)
    Protected SetTimeButtonFontSize As Integer = 12
    Protected SetTimeButtonSize As Size = New Size(50, 40)

    Enum Modes
        A1A2A3
        A4
    End Enum

    Public Sub New()
    End Sub
    'Constructor
    'loc = Location
    'siz = Size
    'parent = Parent control
    'sDimensions = series dimensions
    'type = chart type
    Public Sub New(ByVal loc As Point,
                   ByVal size As Size,
                   ByVal parent As Control,
                   ByVal mode As Modes
                   )
        'Plot controls
        chart = New Chart
        Panel = New Panel()
        parent.Controls.Add(Panel)
        Panel.Controls.Add(chart)
        'Panel.Name = "Panel"
        Panel.Size = size
        Panel.Location = loc

        series = New List(Of Series)

        chart.Location = New Point(0, 0) 'so it hugs the border
        chart.Size = size
        chart.ChartAreas.Add(New ChartArea("ChartArea1"))
        chart.Legends.Add(New Legend("Legend1"))




        TestMode = mode
        If mode = Modes.A1A2A3 Then
            NumOfMics = 6
        ElseIf mode = Modes.A4 Then
            NumOfMics = 4
        End If


    End Sub
    'Update function
    Public Overridable Sub Update(ByVal newVal() As Double)
    End Sub

End Class

Public Class LineGraph
    Inherits CGraph
    'private global vars
    'Dim SecTextBox As MaskedTextBox
    Dim TimeInSec As Integer
    'Dim SetTimeButton As Button

    'constructor
    Public Sub New(ByVal loc As Point,
                   ByVal size As Size,
                   ByVal parent As Control,
                   ByVal mode As Modes,
                   ByVal time As Integer
                   )
        MyBase.New(loc, size, parent, mode)
        'Offset the chart a bit to the right to allow room for the button and textbox
        'chart.Location = New Point(SetTimeButtonSize.Width, 0)
        'chart.Size = New Size(size.Width - SecTextBoxSize.Width, size.Height)
        chart.Location = New Point(0, 0)
        chart.Size = New Size(size.Width, size.Height)
        'Set Time TextBox
        'SecTextBox = New MaskedTextBox()
        'Panel.Controls.Add(SecTextBox)
        'SecTextBox.Font = New Font(SecTextBox.Font.FontFamily, SecTextFontSize)
        'SecTextBox.Size = SecTextBoxSize
        'SecTextBox.Location = New Point(0, size.Height / 4)
        'If time > 0 Then
        '    SecTextBox.Text = time.ToString()
        '    SecTextBox.Enabled = False
        '    SetTime(time)
        'End If
        'SecTextBox.Mask = "999" 'digits including empty spaces

        ''Set Time button
        'SetTimeButton = New Button()
        'Panel.Controls.Add(SetTimeButton)
        'SetTimeButton.Font = New Font(SetTimeButton.Font.FontFamily, SecTextFontSize)
        'SetTimeButton.Size = SetTimeButtonSize
        'SetTimeButton.Location = New Point(0, size.Height / 4 + SecTextBox.Size.Height)
        'SetTimeButton.Text = "Set"

        'Set Chart zoomability
        chart.ChartAreas(0).CursorY.IsUserSelectionEnabled = True
        chart.ChartAreas(0).AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
        chart.ChartAreas(0).AxisY.IsLabelAutoFit = False
        'chart.ChartAreas(0).AxisX.Interval = Int(time / 15) + 1
        chart.ChartAreas(0).AxisX.IsMarginVisible = True
        'set up series
            For i As Integer = 2 To NumOfMics * 2 Step 2
                Dim s As Series = New Series(i.ToString())
                series.Add(s)
                s.ChartType = SeriesChartType.Line
                s.IsVisibleInLegend = False
                s.IsValueShownAsLabel = False
                chart.Series.Add(s)
            Next

    End Sub

    Public Sub ClearChart()
        For i = 0 To chart.Series.Count - 1
            chart.Series(i).Points.Clear()
        Next
    End Sub

    'get the time for chart
    Public Function GetTime()
        Return TimeInSec
    End Function
    'Show initial chart
    Public Sub ShowInitChart(ByVal width As Integer, ByVal height As Integer)
        chart.ChartAreas(0).AxisX.Maximum = width
        chart.ChartAreas(0).AxisY.Maximum = height
    End Sub

    'Set time from the textbox 
    'Public Sub SetTimeFromBox()
    '    If Integer.TryParse(SecTextBox.Text, TimeInSec) Then
    '        chart.ChartAreas(0).AxisX.Maximum = TimeInSec
    '    End If
    'End Sub
    'Set Time
    Public Sub SetTime(ByVal time As Integer)
        TimeInSec = time
        chart.ChartAreas(0).AxisX.Maximum = TimeInSec
    End Sub

    'Update function
    Public Overrides Sub Update(ByVal newVal() As Double)
        If Not newVal.Length = NumOfMics Then
            Return
        End If
        For i As Integer = 0 To newVal.Length - 1
            Dim tag = (i + 1) * 2
            If Not IsNothing(series(i)) Then
                series(i).Points.Add(newVal(i))
            End If
        Next
    End Sub

    Public Sub Dispose()
        'SecTextBox.Dispose()
        'SetTimeButton.Dispose()
        chart.Dispose()
        Panel.Dispose()
    End Sub

End Class



Public Class BarGraph
    Inherits CGraph
    'private global vars
    Dim Labels As List(Of Label)
    Dim listOfNames As List(Of String)

    'constructor
    Public Sub New(ByVal loc As Point,
                   ByVal size As Size,
                   ByVal parent As Control,
                   ByVal mode As Modes)

        MyBase.New(loc, size, parent, mode)
        Labels = New List(Of Label)

        'set up series
        Dim s As Series = New Series()
        series.Add(s)
        s.ChartType = SeriesChartType.Column
        s.IsVisibleInLegend = False
        s.IsValueShownAsLabel = False
        chart.Series.Add(s)
        listOfNames = New List(Of String)
        Dim listOfPoints = New List(Of Double)

        For i As Integer = 0 To NumOfMics - 1
            Dim tag As String = "P" + CStr((i + 1) * 2)
            listOfNames.Add(tag)
            listOfPoints.Add(0)
        Next
        listOfNames.Add("Average")
        listOfPoints.Add(0)
        series(0).Points.DataBindXY(listOfNames, listOfPoints)

    End Sub

    'Update function
    Public Overrides Sub Update(ByVal newVal() As Double)
        If Not newVal.Length = NumOfMics Then
            Return
        End If
        'adding average datapoint
        Dim temp(newVal.Length) As Double
        Dim sum As Double = 0
        For i = 0 To newVal.Length - 1
            temp(i) = newVal(i)
            sum += newVal(i)
        Next
        temp(newVal.Length) = sum / newVal.Length
        series(0).Points.DataBindXY(listOfNames, temp)
    End Sub

    Public Sub Dispose()
        chart.Dispose()
        Panel.Dispose()
    End Sub
End Class

Public Class LineGraphPanel
    Public LineGraphPanel As Panel
    Public LineGraphs
    Public CurrentLineGraph As Integer
    Private cMode As CGraph.Modes
    Private cNumSteps As Integer
    Private lTime As Integer

    Public Sub New(ByVal loc As Point, ByVal size As Size, ByVal parent As Control, ByVal mode As CGraph.Modes, ByVal numSteps As Integer, ByVal time As Integer, ByVal offset As Integer)
        LineGraphPanel = New Panel()
        parent.Controls.Add(LineGraphPanel)
        LineGraphPanel.Size = size
        LineGraphPanel.Location = loc
        LineGraphPanel.AutoScroll = True
        LineGraphs = New LinkedList(Of LineGraph)
        cMode = mode
        cNumSteps = numSteps
        lTime = time
        Dim temp(numSteps - 1) As LineGraph
        CurrentLineGraph = 0
        For i = 0 To numSteps - 1
            temp(i) = New LineGraph(New Point(0, i * offset), New Size(1119, 97), LineGraphPanel, mode, time)
        Next
        LineGraphs = temp
    End Sub

    Public Sub MoveToNextGraph()
        If CurrentLineGraph < cNumSteps - 1 Then
            CurrentLineGraph += 1
            'Else ' refresh old graphs to hold new data
            '    For i = 0 To LineGraphs.Length - 1
            '        LineGraphs(i).ClearChart()
            '    Next
            '    CurrentLineGraph = 0
        End If
    End Sub

    Public Function HasNextGraph()
        If CurrentLineGraph < cNumSteps - 1 Then
            Return True
        End If
        Return False
    End Function

    'mockup for saving data directly from graphs
    Public Sub SaveData()

    End Sub

    Public Sub Update(ByVal newVal() As Double)
        LineGraphs(CurrentLineGraph).Update(newVal)
    End Sub

    Public Sub Dispose()
        For i = 0 To LineGraphs.Length - 1
            LineGraphs(i).Dispose()
        Next
        LineGraphPanel.Dispose()
    End Sub
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