Imports System.IO.Ports
Imports System.Threading

Public Class Communication
    Enum Meters
        p2
        p4
        p6
        p8
        p10
        p12
    End Enum

    Enum Measurements
        Lp
        Leq
        Le
        Lmax
        Lmin
        add
        Ln1
        Ln2
        Ln3
        Ln4
        Ln5
        subLp
        overload
        underrange
    End Enum

    Private buffer(5) As String

    Public WithEvents port As SerialPort = New  _
    System.IO.Ports.SerialPort("COM4",
                            9600,
                            Parity.None,
                            8,
                            StopBits.One)

    Public Sub New()

    End Sub

    Public Function Open() As Boolean
        'Try
        '    If Not port.IsOpen Then
        '        port.Open()
        '    End If
        '    If port.IsOpen Then
        '        Return True
        '    End If
        '    Return False
        'Catch ex As Exception
        '    MsgBox("Can't Connect to Meter" & vbCrLf & ex.Message)
        '    Return False
        'End Try
        Return True
    End Function

    Public Function Close() As Boolean
        'Try
        '    If Not port.IsOpen Then
        '        Return True
        '    End If
        '    If port.IsOpen Then
        '        port.Close()
        '        Return True
        '    End If
        '    Return False
        'Catch ex As Exception
        '    MsgBox("Can't Connect to Meter" & vbCrLf & ex.Message)
        '    Return False
        'End Try
        Return True
    End Function

    Public Function StartMeasure() As Boolean
        'If Open() Then
        '    port.WriteLine("Measure, Start")
        '    If port.ReadLine().StartsWith("R+0000") Then
        '        Return True
        '    End If
        'End If
        'Return False
        Return True
    End Function

    Public Function StopMeasure() As Boolean
        'If Open() Then
        '    port.WriteLine("Measure, Stop")
        '    If port.ReadLine().StartsWith("R+0000") Then
        '        Return True
        '    End If
        'End If
        'Return False
        Return True
    End Function

    '14 figures
    Public Function GetMeasurementsFromBuffer(ByVal part As Measurements) As String()
        'If Open() Then
        '    port.WriteLine("DOD?")
        '    If port.ReadLine().StartsWith("R+0000") Then
        '        Dim s = port.ReadLine()
        '        Return s
        '    End If
        'End If
        'Return False
        Dim r = New Random()
        Dim temp(5) As String
        Dim result(5) As String
        For i = 0 To 5
            'result(i) = (r.Next(1, 100) / 10) + 90
            Dim tempR = (r.Next(1, 10) / 10) + r.Next(1, 119)
            temp(i) = tempR & "," & (tempR - 1) & ",--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-"
            result(i) = temp(i).Split(",")(part)
        Next

        buffer = temp
        Return result
    End Function

    Public Function GetMeasurementsFromMeters(ByVal part As Measurements) As String()
        Return GetMeasurementsFromBuffer(part)
    End Function

    Public Function GetMeasurement(ByVal meterNum As Integer, ByVal part As Measurements) As String
        Dim s() As String = buffer(meterNum / 2 - 1).Split(",")
        If Not IsNothing(s) And Not s.Length <= 0 Then
            Return s(part)
        End If
        Return False
    End Function
End Class
