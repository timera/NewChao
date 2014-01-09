Imports System.IO.Ports
Imports System.Threading

Public Class Communication
    Enum Measurements
        Lp
        Leq
        Le
        Lmax
        Lmin
        Ln1
        Ln2
        Ln3
        Ln4
        Ln5
    End Enum

    Shared WithEvents port As SerialPort = New  _
    System.IO.Ports.SerialPort("COM4",
                            9600,
                            Parity.None,
                            8,
                            StopBits.One)

    Shared Function Open()
        Try
            If Not port.IsOpen Then
                port.Open()
            End If
            If port.IsOpen Then
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox("Can't Connect to Meter" & vbCrLf & ex.Message)
            Return False
        End Try
    End Function

    Shared Function Close()
        Try
            If Not port.IsOpen Then
                Return True
            End If
            If port.IsOpen Then
                port.Close()
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox("Can't Connect to Meter" & vbCrLf & ex.Message)
            Return False
        End Try
    End Function

    Shared Function StartMeasure()
        If Open() Then
            port.WriteLine("Measure, Start")
            If port.ReadLine().StartsWith("R+0000") Then
                Return True
            End If
        End If
        Return False
    End Function

    Shared Function StopMeasure()
        If Open() Then
            port.WriteLine("Measure, Stop")
            If port.ReadLine().StartsWith("R+0000") Then
                Return True
            End If
        End If
        Return False
    End Function

    Shared Function GetMeasurements()
        If Open() Then
            port.WriteLine("DOD?")
            If port.ReadLine().StartsWith("R+0000") Then
                Dim s = port.ReadLine()
                Return s
            End If
        End If
        Return False
    End Function

    Shared Function GetOneOutput(ByVal output As String, ByVal part As Measurements)
        Dim s() As String = output.Split(",")
        If Not IsNothing(s) And Not s.Length <= 0 Then
            If part = Measurements.Lp Then
                Return s(0)
            ElseIf part = Measurements.Leq Then
                Return s(1)
            ElseIf part = Measurements.Le Then
                Return s(2)
            ElseIf part = Measurements.Lmax Then
                Return s(3)
            ElseIf part = Measurements.Lmin Then
                Return s(4)
            End If
        End If
        Return False
    End Function
End Class
