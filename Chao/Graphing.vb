﻿Imports Microsoft.VisualBasic
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.VisualBasic.PowerPacks

Public Class CGraph
    'protected global variables
    Public chart As Chart
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
    Public Sub New(
                   ByVal mode As Modes
                   )
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

    Public Function GetSeries() As List(Of Series)
        Return series
    End Function
End Class

Public Class LineGraph
    Inherits CGraph

    Dim _DefaultLineGraphSize As Size = New Size(1200, 83)
    Dim _DefaultLineGraphPos As Point = New Point(90, 2)
    'private global vars
    'Dim SecTextBox As MaskedTextBox
    Dim TimeInSec As Integer
    'Dim SetTimeButton As Button

    'constructor
    Public Sub New(
                   ByVal parent As Control,
                   ByVal mode As Modes,
                   ByVal time As Integer
                   )
        MyBase.New(mode)
        'Plot controls
        chart = New Chart
        Panel = New Panel()
        parent.Controls.Add(Panel)
        Panel.Controls.Add(chart)
        'Panel.Name = "Panel"
        Panel.Size = _DefaultLineGraphSize
        Panel.Location = _DefaultLineGraphPos
        Panel.BackColor = Color.DarkGray

        series = New List(Of Series)
        chart.Location = New Point(0, -5)
        chart.Size = New Size(_DefaultLineGraphSize.Width, _DefaultLineGraphSize.Height + 15)

        chart.ChartAreas.Add(New ChartArea("ChartArea1"))
        chart.Legends.Add(New Legend("Legend1"))
        chart.BackColor = Color.DarkGray

        'Set Chart zoomability
        chart.ChartAreas(0).CursorY.IsUserSelectionEnabled = True
        chart.ChartAreas(0).AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
        chart.ChartAreas(0).AxisY.IsLabelAutoFit = False
        'chart.ChartAreas(0).AxisX.Interval = Int(time / 15) + 1
        chart.ChartAreas(0).AxisX.IsMarginVisible = True
        chart.ChartAreas(0).AxisX.Minimum = 1
        chart.ChartAreas(0).AxisX.Crossing = 1
        chart.ChartAreas(0).AxisY.Maximum = 120
        chart.ChartAreas(0).AxisY.Minimum = 40
        'set up series
        Dim s As Series
        For i As Integer = 2 To NumOfMics * 2 Step 2
            s = New Series(i.ToString())
            series.Add(s)
            s.ChartType = SeriesChartType.Line
            s.IsVisibleInLegend = False
            s.IsValueShownAsLabel = False
            chart.Series.Add(s)
        Next
        s = New Series("Leq")
        series.Add(s)
        s.ChartType = SeriesChartType.Line
        s.IsVisibleInLegend = False
        s.IsValueShownAsLabel = False
        chart.Series.Add(s)

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

    'Set Time
    Public Sub SetTime(ByVal time As Integer)
        TimeInSec = time
        chart.ChartAreas(0).AxisX.Maximum = TimeInSec
    End Sub

    'Update function newval must contain the avg
    Public Overrides Sub Update(ByVal newVal() As Double)
        For i As Integer = 0 To NumOfMics - 1
            If Not IsNothing(series(i)) Then
                series(i).Points.Add(newVal(i))
            End If
        Next
    End Sub

    Public Sub Dispose()
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

        MyBase.New(mode)
        'Plot controls
        chart = New Chart
        Panel = New Panel()
        parent.Controls.Add(Panel)
        Panel.Controls.Add(chart)
        'Panel.Name = "Panel"
        Panel.Size = New Size(size.Width + 20, size.Height)
        Panel.Location = loc

        series = New List(Of Series)
        chart.Location = New Point(-20, 20) 'so it hugs the border
        chart.Size = New Size(size.Width + 80, size.Height)

        chart.ChartAreas.Add(New ChartArea("ChartArea1"))
        chart.Legends.Add(New Legend("Legend1"))
        chart.BackColor = Color.DarkGray


        Labels = New List(Of Label)

        'set up series
        Dim s As Series = New Series()
        series.Add(s)
        s.ChartType = SeriesChartType.Column
        s.IsVisibleInLegend = False
        s.IsValueShownAsLabel = False
        's("PointWidth") = "1.2"
        chart.Series.Add(s)
        chart.ChartAreas(0).AxisY.Crossing = 40
        chart.ChartAreas(0).AxisY.Maximum = 120
        chart.ChartAreas(0).AxisY.Minimum = 40
        listOfNames = New List(Of String)
        Dim listOfPoints = New List(Of Double)


        For i As Integer = 0 To 6 - 1
            Dim tag As String = "P" + CStr((i + 1) * 2)
            listOfNames.Add(tag)
            listOfPoints.Add(0)
        Next
        listOfNames.Add("Average")
        listOfPoints.Add(0)
        series(0).Points.DataBindXY(listOfNames, listOfPoints)

    End Sub



    'Update, feed in values, including an avg
    Public Overrides Sub Update(ByVal newVal() As Double)
        series(0).Points.DataBindXY(listOfNames, newVal)
    End Sub

    Public Sub Dispose()
        chart.Dispose()
        Panel.Dispose()
    End Sub

End Class

'helper classes for keeping track of labels and points
Public Class CoorPoint
    'Public Label As ColorLabel
    Public Coors As ThreeDPoint
    Public Line As LineShape

    Public Sub New()
        'Label = lab
        Coors = New ThreeDPoint
    End Sub

    Public Sub New(ByVal p As ThreeDPoint, ByVal l As LineShape)
        'Label = lab
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