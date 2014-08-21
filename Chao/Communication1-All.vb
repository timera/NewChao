Imports System.IO.Ports
Imports System.Threading
Imports System.Text
Imports System.IO
Imports System.Management

'This is the class dealing with Noise Meters
Public Class Communication1_All
    Dim _Latency As Integer = 700 'miliseconds to wait for response
    Dim _serverLatency As Integer = 30 'miliseconds to wait for response via serial port
    Dim _btwCmdLatency As Integer = 10
    Dim Meters() As String = {
        "P2",
        "P4",
        "P6",
        "P8",
        "P10",
        "P12"
    }

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

    '這是已從server那讀取的資訊，但還沒有顯示
    Private buffer() As String

    Public WithEvents port As SerialPort

    Public Sub New()
        'PollingThread = New Thread(AddressOf Poll)
        'PollingThread.IsBackground = True
        MeterMacs.Add(Meter2Mac)
        MeterMacs.Add(Meter4Mac)
        'TEMP，有正確的MAC之後再把if拿掉
        If Program.sim Then
            MeterMacs.Add(Meter6Mac)
            MeterMacs.Add(Meter8Mac)
            MeterMacs.Add(Meter10Mac)
            MeterMacs.Add(Meter12Mac)
        End If
        ClearBuffer()
    End Sub

    Private Sub ClearBuffer()
        buffer = New String(MeterMacs.Count - 1) {}
        For i = 0 To buffer.Length - 1
            buffer(i) = ""
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

    'TEMP,有真的剩下四個噪音計之後再把'拿掉
    Public Sub Real()
        MeterMacs.Clear()
        MeterMacs.Add(Meter2Mac)
        MeterMacs.Add(Meter4Mac)
        'MeterMacs.Add(Meter6Mac)
        'MeterMacs.Add(Meter8Mac)
        'MeterMacs.Add(Meter10Mac)
        'MeterMacs.Add(Meter12Mac)
    End Sub

    'return all the serial port names
    Public Shared Function GetComs() As String()
        Dim ports() As String = SerialPort.GetPortNames()
        Dim ManObjReturn As ManagementObjectCollection
        Dim ManObjSearch As ManagementObjectSearcher
        ManObjSearch = New ManagementObjectSearcher("Select * from Win32_SerialPort")
        ManObjReturn = ManObjSearch.Get()
        For Each ManObj As ManagementObject In ManObjReturn
            For i = 0 To ports.Length - 1
                If ManObj("DeviceID").ToString() = ports(i) Then
                    ports(i) = ManObj("DeviceID") & ": " & ManObj("Name").ToString()
                End If
            Next

        Next
        Return ports
    End Function

    'Open Port so we can talk to server
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

    'we need to close port before leaving or else the port remains open and no one else can talk to this port
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

    'Filters out the other messages than the given filter
    Private Sub FilterMsgsFromBuffer(ByVal filter As String)
        For i = 0 To buffer.Length - 1
            If Not buffer(i).Contains(filter) Then
                buffer(i) = ""
            End If
        Next
    End Sub

    'Should only be called by ReadAndSortMsgs because of ease of control of race conditions
    Private Function GetInputFromPort() As Byte()
        If Open() Then
            Dim numBytes = port.BytesToRead
            Dim inputTemp() As Byte = New Byte(numBytes) {}
            If numBytes > 1 Then
                port.Read(inputTemp, 0, numBytes)
                Return inputTemp
            End If
        End If
        Return Nothing
    End Function

    Function DetermineMeter(ByVal name As String) As Integer
        Dim i As Integer = 0
        For Each meter In Meters
            If name.Contains(meter) Then
                Return i
            End If
            i += 1
        Next
        Return -1
    End Function

    'Thread to poll data from server buffer
    Private PollingThread As Thread

    'Multi Threading Lock for buffer
    Dim bufferLock As New Object
    Dim portLock As New Object

    Public Sub Poll(ByVal meterIndex As Integer)
        SyncLock portLock
            Dim i As Integer = 0
            For Each mac As Byte() In MeterMacs
                If meterIndex = -1 Then
                    'destination address
                    Dim command(2 + 3) As Byte
                    command(0) = &HCC
                    command(1) = &H10
                    Array.Copy(mac, 0, command, 2, 3)
                    'success code
                    Dim sucCode(1 + 3) As Byte
                    sucCode(0) = &HCC
                    Array.Copy(mac, 0, sucCode, 1, 3)

                    If SetupAndSendCommand(command, sucCode) Then
                        port.WriteLine("DOD?")
                    End If
                ElseIf meterIndex = i Then 'specific meter
                    'destination address
                    Dim command(2 + 3) As Byte
                    command(0) = &HCC
                    command(1) = &H10
                    Array.Copy(mac, 0, command, 2, 3)
                    'success code
                    Dim sucCode(1 + 3) As Byte
                    sucCode(0) = &HCC
                    Array.Copy(mac, 0, sucCode, 1, 3)

                    If SetupAndSendCommand(command, sucCode) Then
                        port.WriteLine("DOD?")
                    End If
                    Exit For
                End If
                i += 1
            Next
        End SyncLock
        Thread.Sleep(_Latency)
        Dim input() As Byte
        SyncLock portLock
            input = GetInputFromPort()
        End SyncLock
        SyncLock bufferLock
            If input IsNot Nothing Then
                buffer = Encoding.Default.GetString(input).Split("$")
            End If
        End SyncLock
    End Sub

    Function EnableBroadCast()
        Return SetBroadCast(1)
    End Function

    Function DisableBroadCast()
        Return SetBroadCast(0)
    End Function

    Function SetBroadCast(ByVal code As Integer) As Boolean
        Dim bufCmd() As Byte = {&HCC, &H8, code}
        Dim sucCmd() As Byte = {&HCC, code, &H0}
        Return SetupAndSendCommand(bufCmd, sucCmd)
    End Function

    Function SendCommandinAT(ByRef command As Byte(), ByRef returnCode As Byte()) As Boolean
        Dim bufWrite() As Byte = command
        Dim sucWrite() As Byte = returnCode
        Dim buffer() As Byte

        port.Write(bufWrite, 0, bufWrite.Length)
        Thread.Sleep(_serverLatency)
        buffer = New Byte(port.BytesToRead) {}
        port.Read(buffer, 0, buffer.Length)
        'If Successfully write command
        If buffer.SequenceEqual(sucWrite) Then
            Return True
        End If
        Return False
    End Function

    Function SetupAndSendCommand(ByRef command As Byte(), ByRef returnCode As Byte())
        Dim bufCmd() As Byte = {&H41, &H54, &H2B, &H2B, &H2B, &HD}
        Dim sucCmd() As Byte = {&HCC, &H43, &H4F, &H4D}

        Dim bufWrite() As Byte = command
        Dim sucWrite() As Byte = returnCode

        Dim bufExitCmd() As Byte = {&HCC, &H41, &H54, &H4F, &HD}
        Dim sucExitCmd() As Byte = {&HCC, &H44, &H41, &H54}

        Dim sucTotal(sucCmd.Length + sucWrite.Length + sucExitCmd.Length - 1) As Byte
        Array.Copy(sucCmd, sucTotal, sucCmd.Length)
        Array.Copy(sucWrite, 0, sucTotal, sucCmd.Length, sucWrite.Length)
        Array.Copy(sucExitCmd, 0, sucTotal, sucCmd.Length + sucWrite.Length - 1, sucExitCmd.Length)

        port.Write(bufCmd, 0, bufCmd.Length)
        Thread.Sleep(_btwCmdLatency)
        port.Write(bufWrite, 0, bufWrite.Length)
        Thread.Sleep(_btwCmdLatency)
        port.Write(bufExitCmd, 0, bufExitCmd.Length)

        Thread.Sleep(_serverLatency)
        If port.BytesToRead > 0 Then
            Dim buffer() As Byte = New Byte(port.BytesToRead) {}
            port.Read(buffer, 0, buffer.Length)
            Dim str As String = Encoding.Default.GetString(buffer)
            'If Successfully write command
            If buffer.SequenceEqual(sucTotal) Then
                Return True
            End If
        End If
        Return False

    End Function

    Function EnterATMode() As Boolean
        Dim bufCmd() As Byte = {&H41, &H54, &H2B, &H2B, &H2B, &HD}
        Dim sucCmd() As Byte = {&HCC, &H43, &H4F, &H4D, &H0}
        'Enter AT Command Mode
        port.Write(bufCmd, 0, bufCmd.Length)
        Thread.Sleep(_serverLatency)
        Dim buffer() As Byte = New Byte(port.BytesToRead) {}
        If port.BytesToRead > 0 Then
            port.Read(buffer, 0, buffer.Length)
            If buffer.SequenceEqual(sucCmd) Then
                Return True
            End If
        End If
        Return False
    End Function

    Function ExitATMode() As Boolean
        Dim bufExitCmd() As Byte = {&HCC, &H41, &H54, &H4F, &HD}
        Dim sucExitCmd() As Byte = {&HCC, &H44, &H41, &H54, &H0}
        Dim buffer() As Byte
        'write exit command
        port.Write(bufExitCmd, 0, bufExitCmd.Length)
        Thread.Sleep(_serverLatency)
        buffer = New Byte(port.BytesToRead) {}
        port.Read(buffer, 0, buffer.Length)
        If buffer.SequenceEqual(sucExitCmd) Then
            Return True
        End If
        Return False
    End Function

    Public Function EnsureActionOnAll()

    End Function

    'broadcasts start measuring, returns the ones that confirm start measuring
    Public Function StartMeasure(ByVal name As String) As Boolean
        If Program.sim Then
            Return True
        ElseIf Open() Then
            ClearBuffer()
            If MakeSureReady() Then
                SyncLock portLock
                    If name.Contains("Cal") Then

                        Dim meterIndex As Integer = DetermineMeter(name)
                        Dim mac() As Byte = MeterMacs(meterIndex)
                        Dim command(2 + 3) As Byte
                        command(0) = &HCC
                        command(1) = &H10
                        Array.Copy(mac, 0, command, 2, 3)
                        'success code
                        Dim sucCode(1 + 3) As Byte
                        sucCode(0) = &HCC
                        Array.Copy(mac, 0, sucCode, 1, 3)

                        If SetupAndSendCommand(command, sucCode) Then
                            port.WriteLine("Measure, Start")
                        End If
                    Else
                        If EnableBroadCast() Then
                            port.WriteLine("Measure, Start")
                        Else
                            MsgBox("Broadcast cannot be enabled!")
                            Return False
                        End If
                    End If
                End SyncLock

                'PollingThread = New Thread(AddressOf Poll)
                'PollingThread.Start(DetermineMeter(name))
                Poll(DetermineMeter(name))
                Return True
            Else
                MsgBox("Meters Not Ready")
            End If
        End If
        Return False
    End Function

    'broadcasts pause measuring, returns the ones that confirm pause measuring
    Public Function PauseMeasure(ByVal pause As Boolean) As Boolean
        If Program.sim Then
            Return True
        ElseIf Open() Then
            ClearBuffer()
            If MakeSureReady() Then
                SyncLock portLock
                    If EnableBroadCast() Then
                        If pause Then
                            port.WriteLine("Pause, Pause")
                        Else
                            port.WriteLine("Pause, Clear")
                        End If
                    Else
                        MsgBox("Broadcast cannot be enabled!")
                        Return False
                    End If
                End SyncLock
                If pause Then
                    'PollingThread = New Thread(AddressOf Poll)
                    'PollingThread.Start()
                    'Poll()
                End If
                Return True
            Else
                MsgBox("Meters Not Ready")
            End If
        End If
        Return False
    End Function

    Public Function MakeSureReady() As Boolean
        Dim result() As Boolean = New Boolean(MeterMacs.Count - 1) {}
        'This will prevent other threads from accessing buffer at the same time
        SyncLock portLock
            Dim i As Integer = 0

            For Each mac As Byte() In MeterMacs
                    'destination address
                    Dim command(2 + 3) As Byte
                    command(0) = &HCC
                    command(1) = &H10
                    Array.Copy(mac, 0, command, 2, 3)
                    'success code
                    Dim sucCode(1 + 3) As Byte
                    sucCode(0) = &HCC
                    Array.Copy(mac, 0, sucCode, 1, 3)
                    If SetupAndSendCommand(command, sucCode) Then
                        port.WriteLine("")
                        Thread.Sleep(_Latency)
                        Dim input() As Byte = GetInputFromPort()
                        If input Is Nothing Then
                            result(i) = False
                            i += 1
                            Continue For
                        ElseIf input.Length = 0 Then
                            result(i) = False
                            i += 1
                            Continue For
                        ElseIf input.Length > 0 Then
                            result(i) = True
                        End If
                    End If
                    i += 1
            Next
        End SyncLock

        For i = 0 To result.Length - 1
            If Not result(i) Then
                Return False
            End If
        Next
        Return True
    End Function

    'broadcasts stop measuring, returns the ones that confrim stop measuring
    Public Sub StopMeasure()
        Dim result() As Boolean = New Boolean(MeterMacs.Count - 1) {}
        If Program.sim Then
            For i = 0 To result.Length - 1
                result(i) = True
            Next
        ElseIf Open() Then
            SyncLock portLock
                If EnableBroadCast() Then
                    port.WriteLine("Measure, Stop")

                End If
            End SyncLock
        End If
    End Sub

    '14 figures
    Public Function ConsumeMeasurementsFromBuffer(ByVal part As Measurements, ByVal name As String) As String()
        Dim result() As String = New String(MeterMacs.Count - 1) {}
        Dim meterIndex As Integer = DetermineMeter(name)
        Try
            If Not IsNothing(buffer) Then
                SyncLock bufferLock
                    FilterMsgsFromBuffer("R+0000")
                    For i = 0 To MeterMacs.Count - 1
                        If Not meterIndex = -1 Then
                            If Not i = meterIndex Then
                                result(i) = "0"
                                Continue For
                            End If
                        End If
                        If buffer(i).Length > 8 Then
                            Dim temp() As String = buffer(i).Substring(8).Split(",")
                            If part < temp.Length Then
                                result(i) = temp(part)
                                Continue For
                            End If
                            buffer(i) = ""
                            result(i) = "0"
                        Else
                            buffer(i) = ""
                            result(i) = "0"
                        End If

                    Next
                End SyncLock
            End If
        Catch ex As Exception
            MsgBox("ConsumeMeasurementsFromBuffer: " & ex.Message)
        End Try
        Return result
    End Function

    'if meterIndex = -1 -> we want all meters, else it's a specific meter
    Public Function GetMeasurementsFromMeters(ByVal part As Measurements, ByVal name As String) As String()
        Dim meterIndex As Integer = DetermineMeter(name)
        Dim result() As String = New String(MeterMacs.Count - 1) {}
        If Not Program.sim Then
            Try
                If Open() Then
                    'PollingThread = New Thread(AddressOf Poll)
                    'PollingThread.Start()
                    Poll(DetermineMeter(name))
                    result = ConsumeMeasurementsFromBuffer(part, name)
                End If
            Catch ex As Exception
                MsgBox("GetMeasurementsFromMeters: " & ex.Message)
            End Try
            Return result
        Else
            ''TEMP
            Dim r = New Random()
            Dim temp(MeterMacs.Count - 1) As String

            For i = 0 To MeterMacs.Count - 1
                Dim tempR = (r.Next(1, 10) / 10) + r.Next(40, 119)
                temp(i) = ("R+0000  " & tempR & "," & (tempR - 1) & ",--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-")
                result(i) = temp(i).Substring(8).Split(",")(part)
                If Not meterIndex = -1 Then
                    If Not i = meterIndex Then
                        result(i) = 0
                    End If
                End If
            Next

            buffer = temp
        End If
        Return result
    End Function

End Class