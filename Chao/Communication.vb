﻿Imports System.IO.Ports
Imports System.Threading
Imports System.Text
Imports System.IO

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

    Private ServerMac() As Byte = BitConverter.GetBytes(CLng(&H5067A2F4A6))

    Private MeterMacs As List(Of Byte()) = New List(Of Byte())
    Private Meter2Mac() As Byte = New Byte(2) {&HA2, &HF4, &H5A} 'BitConverter.GetBytes(CLng(&HA2F45A))
    Private Meter4Mac() As Byte = New Byte(2) {&HA2, &HF4, &H91}
    Private Meter6Mac() As Byte = New Byte(2) {&HA2, &HF4, &H92}
    Private Meter8Mac() As Byte = New Byte(2) {&HA2, &HF4, &H93}
    Private Meter10Mac() As Byte = New Byte(2) {&HA2, &HF4, &H94}
    Private Meter12Mac() As Byte = New Byte(2) {&HA2, &HF4, &H95}


    Private buffer() As List(Of String)

    Public WithEvents port As SerialPort = New  _
    System.IO.Ports.SerialPort("COM5",
                            9600,
                            Parity.None,
                            8,
                            StopBits.One)

    Public Sub New()
        MeterMacs.Add(Meter2Mac)
        MeterMacs.Add(Meter4Mac)
        MeterMacs.Add(Meter6Mac)
        MeterMacs.Add(Meter8Mac)
        MeterMacs.Add(Meter10Mac)
        MeterMacs.Add(Meter12Mac)
        buffer = New List(Of String)(MeterMacs.Count - 1) {}
        For i = 0 To MeterMacs.Count - 1
            buffer(i) = New List(Of String)
        Next
    End Sub

    Public Function Open() As Boolean
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
        Return True
    End Function

    Public Function Close() As Boolean
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
        Return True
    End Function

    'process incoming messages, categorize them into different meters and strip them down to only the message
    Private Function ProcessMsgs(ByRef Input As Byte()) As List(Of String)()
        Dim msgs As List(Of Byte()) = New List(Of Byte())
        Dim start As Integer = -1

        For i = 0 To Input.Length - 1
            If Input(i) = &H81 Or i = Input.Length - 1 Then
                If Not start = -1 Then
                    Dim size = i - start
                    If i = Input.Length - 1 Then
                        size += 1
                    End If
                    Dim buffer(size - 1) As Byte
                    Array.Copy(Input, start, buffer, 0, size)
                    msgs.Add(buffer)
                End If
                start = i
            End If
        Next

        Dim msgArray(MeterMacs.Count - 1) As List(Of String)
        For i = 0 To MeterMacs.Count - 1
            msgArray(i) = New List(Of String)
        Next

        'for 6 meters, start index(incl) to end index(excl)

        For i = 0 To msgs.Count - 1
            Dim add(2) As Byte
            Array.Copy(msgs(i), 4, add, 0, 3)

            For j = 0 To MeterMacs.Count - 1
                If add.SequenceEqual(MeterMacs(j)) Then
                    Dim bytes(msgs(i).Length - 7 - 1) As Byte
                    Array.Copy(msgs(i), 7, bytes, 0, msgs(i).Length - 7)
                    Dim s As String = Encoding.Unicode.GetString(bytes)
                    msgArray(j).Add(s)
                End If
            Next
        Next
        Return msgArray
    End Function

    Private Function GetInputFromPort() As Byte()
        If Open() Then
            Dim numBytes = port.BytesToRead
            Dim input() As Byte = New Byte(numBytes) {}
            port.Read(input, 0, numBytes)
            Return input
        End If
        Return Nothing
    End Function

    Private Function CheckTrue() As Boolean()
        Dim result() As Boolean = New Boolean(MeterMacs.Count - 1) {}
        Try
            Dim input() As Byte = GetInputFromPort()
            buffer = ProcessMsgs(Input)

            For i = 0 To MeterMacs.Count - 1
                Dim s As String = buffer(i)(0)
                If s.StartsWith("R+0000") Then
                    result(i) = True
                Else
                    result(i) = False
                End If
            Next
        Catch ex As Exception
            MsgBox("In CheckTrue: " & ex.Message)
        End Try
        Return result
    End Function

    'broadcasts start measuring, returns the ones that confirm start measuring
    Public Function StartMeasure() As Boolean()
        Dim result() As Boolean = New Boolean(MeterMacs.Count - 1) {}
        If Open() Then
            port.WriteLine("Measure, Start")
            Thread.Sleep(1000)
            result = CheckTrue()
        End If
        Return result
    End Function

    'broadcasts stop measuring, returns the ones that confrim stop measuring
    Public Function StopMeasure() As Boolean()
    Dim result() As Boolean = New Boolean(MeterMacs.Count - 1) {}
        If Open() Then
            port.WriteLine("Measure, Stop")
            Thread.Sleep(1000)
            result = CheckTrue()
        End If
        Return result
    End Function

    '14 figures
    Public Function GetMeasurementsFromBuffer(ByVal part As Measurements) As String()
        Dim result() As String = New String(MeterMacs.Count - 1) {}
        If Not IsNothing(buffer) Then
            Dim temp() As List(Of String) = buffer
            For i = 0 To MeterMacs.Count - 1
                If temp(i).Count > 0 Then
                    result(i) = temp(i)(1).Split(",")(part)
                End If
            Next
        End If
        Return result
    End Function

    Public Function GetMeasurementsFromMeters(ByVal part As Measurements) As String()
        'Dim result() As String = New String(MeterMacs.Count - 1) {}
        'Try
        '    Dim input() As Byte = GetInputFromPort()
        '    buffer = ProcessMsgs(input)
        '    For i = 0 To MeterMacs.Count - 1
        '        Dim s As String = buffer(i)(0)
        '        If s.StartsWith("R+0000") Then
        '            result(i) = buffer(i)(1).Split(",")(part)
        '        End If
        '    Next
        'Catch ex As Exception
        '    MsgBox("GetMeasurementsFromMeters: " & ex.Message)
        'End Try
        'Return result
        Dim r = New Random()
        Dim temp(MeterMacs.Count - 1) As List(Of String)
        For i = 0 To MeterMacs.Count - 1
            temp(i) = New List(Of String)
        Next
        Dim result() As String = New String(MeterMacs.Count - 1) {}
        For i = 0 To MeterMacs.Count - 1
            'result(i) = (r.Next(1, 100) / 10) + 90
            Dim tempR = (r.Next(1, 10) / 10) + r.Next(1, 119)
            temp(i).Add("R+0000")
            temp(i).Add(tempR & "," & (tempR - 1) & ",--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-")
            result(i) = temp(i)(1).Split(",")(part)
        Next

        buffer = temp
        Return result
    End Function

    Public Function GetMeasurement(ByVal meterNum As Integer, ByVal part As Measurements) As String
        Dim s() As String = buffer(meterNum / 2 - 1)(1).Split(",")
        If Not IsNothing(s) And Not s.Length <= 0 Then
            Return s(part)
        End If
        Return False
    End Function
End Class
