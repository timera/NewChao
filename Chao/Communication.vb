Imports System.IO.Ports
Imports System.Threading
Imports System.Text
Imports System.IO

Public Class Communication
    Dim _Latency As Integer = 500 'miliseconds to wait for response
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

    Private ServerMac() As Byte = New Byte(5) {&H0, &H50, &H67, &HA2, &HF4, &HA6} 'BitConverter.GetBytes(CLng(&H5067A2F4A6))

    Private MeterMacs As List(Of Byte()) = New List(Of Byte())
    Private Meter2Mac() As Byte = New Byte(2) {&HA2, &HF4, &H5A}
    Private Meter4Mac() As Byte = New Byte(2) {&HA2, &HF4, &H91}
    ' the following don't have the right Mac addresses yet
    Private Meter6Mac() As Byte = New Byte(2) {&HA2, &HF4, &H92}
    Private Meter8Mac() As Byte = New Byte(2) {&HA2, &HF4, &H93}
    Private Meter10Mac() As Byte = New Byte(2) {&HA2, &HF4, &H94}
    Private Meter12Mac() As Byte = New Byte(2) {&HA2, &HF4, &H95}


    Private buffer() As List(Of String)

    Public WithEvents port As SerialPort

    Public Sub New()
        MeterMacs.Add(Meter2Mac)
        'TEMP
        MeterMacs.Add(Meter4Mac)
        'TEMP because we don't have these meters yet
        If Program.sim Then
            MeterMacs.Add(Meter6Mac)
            MeterMacs.Add(Meter8Mac)
            MeterMacs.Add(Meter10Mac)
            MeterMacs.Add(Meter12Mac)
        End If
        buffer = New List(Of String)(MeterMacs.Count - 1) {}
        For i = 0 To MeterMacs.Count - 1
            buffer(i) = New List(Of String)
        Next
    End Sub
    'TEMP
    Public Sub Sim()
        If port IsNot Nothing Then
            If port.IsOpen Then
                Close()
            End If
        End If
        MeterMacs.Clear()
        MeterMacs.Add(Meter2Mac)
        MeterMacs.Add(Meter4Mac)
        MeterMacs.Add(Meter6Mac)
        MeterMacs.Add(Meter8Mac)
        MeterMacs.Add(Meter10Mac)
        MeterMacs.Add(Meter12Mac)
    End Sub

    'TEMP
    Public Sub Real()
        MeterMacs.Clear()
        MeterMacs.Add(Meter2Mac)
        MeterMacs.Add(Meter4Mac)
    End Sub

    Public Shared Function GetComs() As String()
        Return SerialPort.GetPortNames()
    End Function

    Public Function Open() As Boolean
        Try
            If Program.sim Then
                Return True
            End If
            ' if port not initialized
            If port Is Nothing Then
                Dim com As String = "COM5"
                If Program.ComboBoxComs.SelectedItem IsNot Nothing Then
                    com = Program.ComboBoxComs.SelectedItem
                End If
                port = New  _
                System.IO.Ports.SerialPort(com,
                                            9600,
                                            Parity.None,
                                            8,
                                            StopBits.One)
                port.Open()
            ElseIf Not port.IsOpen Then
                Dim com As String = "COM5"
                If Program.ComboBoxComs.SelectedItem IsNot Nothing Then
                    com = Program.ComboBoxComs.SelectedItem
                End If
                port = New  _
                System.IO.Ports.SerialPort(com,
                                            9600,
                                            Parity.None,
                                            8,
                                            StopBits.One)
                port.Open()
            End If
            If port.IsOpen Then
                Return True
            End If
        Catch ex As Exception
            MsgBox("Can't Connect to Meter" & vbCrLf & ex.Message)
        End Try
        Return False
    End Function

    Public Function Close() As Boolean
        If Program.sim Then
            Return True
        End If
        Try
            If Not port.IsOpen Then
                Return True
            End If
            If port.IsOpen Then
                port.Close()
                Return True
            End If

        Catch ex As Exception
            MsgBox("Can't Connect to Meter" & vbCrLf & ex.Message)
        End Try
        Return False
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
                    Dim s As String = Encoding.Default.GetString(bytes)
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
            buffer = ProcessMsgs(input)

            For i = 0 To MeterMacs.Count - 1
                If Not buffer(i).Count = 0 Then
                    Dim s As String = buffer(i)(0)
                    If s.StartsWith("R+0000") Then
                        result(i) = True
                    Else
                        result(i) = False
                    End If

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
        If Program.sim Then
            For i = 0 To result.Length - 1
                result(i) = True
            Next
        ElseIf Open() Then
            port.WriteLine("Measure, Start")
            Thread.Sleep(_Latency)
            result = CheckTrue()
        End If
        Return result
    End Function

    'broadcasts stop measuring, returns the ones that confrim stop measuring
    Public Function StopMeasure() As Boolean()
        Dim result() As Boolean = New Boolean(MeterMacs.Count - 1) {}
        If Program.sim Then
            For i = 0 To result.Length - 1
                result(i) = True
            Next
        ElseIf Open() Then
            port.WriteLine("Measure, Stop")
            Thread.Sleep(_Latency)
            result = CheckTrue()
        End If
        Return result
    End Function

    '14 figures
    Public Function GetMeasurementsFromBuffer(ByVal part As Measurements) As String()
        Dim result() As String = New String(MeterMacs.Count - 1) {}
        Try
            If Not IsNothing(buffer) Then
                Dim temp() As List(Of String) = buffer
                For i = 0 To MeterMacs.Count - 1
                    If temp(i).Count > 0 Then
                        result(i) = temp(i)(0).Substring(8).Split(",")(part)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox("GetMeasurementsFromBuffer: " & ex.Message)
        End Try
        Return result
    End Function

    Public Function GetMeasurementsFromMeters(ByVal part As Measurements) As String()
        Dim result() As String = New String(MeterMacs.Count - 1) {}
        If Not Program.sim Then
            Try
                If Open() Then
                    port.WriteLine("DOD?")
                    Thread.Sleep(_Latency)
                    Dim input() As Byte = GetInputFromPort()
                    buffer = ProcessMsgs(input)
                    For i = 0 To MeterMacs.Count - 1
                        If Not buffer(i).Count = 0 Then
                            Dim s As String = buffer(i)(0)
                            If s.StartsWith("R+0000") Then
                                s = s.Substring(8)
                                result(i) = s.Split(",")(part)
                            End If
                        Else
                            result(i) = 0
                        End If
                    Next
                End If
            Catch ex As Exception
                MsgBox("GetMeasurementsFromMeters: " & ex.Message)
            End Try
            Return result
        Else
            ''TEMP
            Dim r = New Random()
            Dim temp(MeterMacs.Count - 1) As List(Of String)
            For i = 0 To MeterMacs.Count - 1
                temp(i) = New List(Of String)
            Next
            For i = 0 To MeterMacs.Count - 1
                'result(i) = (r.Next(1, 100) / 10) + 90
                Dim tempR = (r.Next(1, 10) / 10) + r.Next(40, 119)
                temp(i).Add("R+0000  " & tempR & "," & (tempR - 1) & ",--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-")
                result(i) = temp(i)(0).Substring(8).Split(",")(part)
            Next

            buffer = temp
        End If
        Return result
    End Function

    Public Function GetMeasurement(ByVal meterNum As Integer, ByVal part As Measurements) As String
        Dim s() As String = buffer(meterNum / 2 - 1)(1).Split(",")
        If Not IsNothing(s) And Not s.Length <= 0 Then
            Return s(part)
        End If
        Return False
    End Function

    Public Function SetupServer() As Boolean
        If Open() Then
            Try
                Dim bufCmd() As Byte = {&H41, &H54, &H2B, &H2B, &H2B, &HD}
                Dim sucCmd() As Byte = {&HCC, &H43, &H4F, &H4D, &H0}

                Dim bufWriteAPI() As Byte = {&HCC, &H17, &H1}
                Dim sucWriteAPI() As Byte = {&HCC, &H1, &H0}

                Dim bufExitCmd() As Byte = {&HCC, &H41, &H54, &H4F, &HD}
                Dim sucExitCmd() As Byte = {&HCC, &H44, &H41, &H54, &H0}

                port.Write(bufCmd, 0, bufCmd.Length)
                Thread.Sleep(_Latency)
                Dim buffer() As Byte = New Byte(port.BytesToRead) {}
                port.Read(buffer, 0, buffer.Length)
                If buffer.SequenceEqual(sucCmd) Then
                    port.Write(bufWriteAPI, 0, bufWriteAPI.Length)
                    Thread.Sleep(_Latency)
                    buffer = New Byte(port.BytesToRead) {}
                    port.Read(buffer, 0, buffer.Length)
                    If buffer.SequenceEqual(sucWriteAPI) Then
                        port.Write(bufExitCmd, 0, bufExitCmd.Length)
                        Thread.Sleep(_Latency)
                        buffer = New Byte(port.BytesToRead) {}
                        port.Read(buffer, 0, buffer.Length)
                        If buffer.SequenceEqual(sucExitCmd) Then
                            Return True
                        End If
                    End If
                End If
            Catch ex As Exception
                MsgBox("Setup Server: " & ex.Message)
            End Try
        End If
        Return False
    End Function
End Class
