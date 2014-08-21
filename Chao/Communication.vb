Imports System.IO.Ports
Imports System.Threading
Imports System.Text
Imports System.IO
Imports System.Management

'This is the class dealing with Noise Meters
Public Class Communication
    Dim _Latency As Integer = 500 'miliseconds to wait for response
    Dim _PollingFrequency As Integer = 400 'miliseconds for every poll
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
    Private buffer(6 - 1) As StringBuilder

    Public WithEvents port As SerialPort

    Public Sub New()
        PollingThread = New Thread(AddressOf Poll)
        PollingThread.IsBackground = True
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

    Function EnableBroadCast() As Boolean
        Return SetBroadCast(1)
    End Function

    Function DisableBroadCast() As Boolean
        Return SetBroadCast(0)
    End Function

    Function SetBroadCast(ByVal code As Integer) As Boolean
        Dim bufCmd() As Byte = {&HCC, &H8, code}
        Dim sucCmd() As Byte = {&HCC, code, &H0}
        Return SetupAndSendCommand(bufCmd, sucCmd)
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
        Thread.Sleep(10)
        port.Write(bufWrite, 0, bufWrite.Length)
        Thread.Sleep(10)
        port.Write(bufExitCmd, 0, bufExitCmd.Length)

        Thread.Sleep(30)
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

    Public Function DetermineMeter(ByVal name As String) As Integer
        Dim i As Integer = 0
        For Each meter In Meters
            If name.Contains(meter) Then
                Return i
            End If
            i += 1
        Next
        Return -1
    End Function

    Private Sub ClearBuffer()
        For i = 0 To MeterMacs.Count - 1
            buffer(i) = New StringBuilder
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
    Private Function FilterMsgsFromBuffer(ByVal filter As String)
        Dim tempBuffer = New List(Of String)(MeterMacs.Count - 1) {}
        For i = 0 To MeterMacs.Count - 1
            tempBuffer(i) = New List(Of String)
        Next
        For i = 0 To buffer.Length - 1
            Dim str As String = buffer(i).ToString().Replace(filter, "%")
            If str IsNot Nothing Then
                If str.Length > 0 Then
                    tempBuffer(i).AddRange(str.Split("%"))
                End If
            End If
        Next
        Return tempBuffer
    End Function

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

    Dim LeftFromLastRead() As Byte

    'process incoming messages, categorize them into different meters and strip them down to only the message
    Private Sub ReadAndSortMsgs(ByVal skipLeftFromLast As Boolean, ByVal text As String)
        Dim input() As Byte
        SyncLock portLock
            If text IsNot Nothing Then
                port.WriteLine(text)
                Thread.Sleep(_Latency)
            End If

            input = GetInputFromPort()
        End SyncLock
        If input Is Nothing Then
            Return
        End If
        If input.Length = 0 Then
            Return
        End If

        Dim msgs As List(Of Byte()) = New List(Of Byte())
        'start是這個封包的開頭index
        Dim start As Integer = -1

        '每個BYTE看內容,然後將一個個封包分開放入msgs中
        For i = 0 To Input.Length - 1
            '&H81是毎個不同receiver回的封包的開頭
            If Input(i) = &H81 Then
                '如果不是第一個封包
                If Not start = -1 Then
                    Dim size = i - start
                    Dim bufferTemp(size - 1) As Byte
                    Array.Copy(Input, start, bufferTemp, 0, size)
                    msgs.Add(bufferTemp)
                End If
                '如果有切斷的資訊
                If start = -1 Then
                    If i > 0 Then
                        If LeftFromLastRead IsNot Nothing Then
                            Dim bufferTemp(LeftFromLastRead.Length + i) As Byte
                            Array.Copy(LeftFromLastRead, bufferTemp, LeftFromLastRead.Length)
                            Array.Copy(Input, 0, bufferTemp, LeftFromLastRead.Length, i + 1)
                            msgs.Add(bufferTemp)
                            LeftFromLastRead = Nothing
                        End If
                    ElseIf i = 0 Then '雖然沒被切斷但還是要處理LeftFromLastRead
                        If LeftFromLastRead IsNot Nothing Then
                            msgs.Add(LeftFromLastRead)
                            LeftFromLastRead = Nothing
                        End If
                    End If
                End If
                start = i
            End If
            If i = Input.Length - 1 Then 'assume it's always chopped off at the end
                If start = -1 Then
                    start = 0
                End If
                Dim size = i - start
                Dim bufferTemp(size - 1) As Byte
                Array.Copy(Input, start, bufferTemp, 0, size)
                If skipLeftFromLast Then
                    msgs.Add(bufferTemp)
                Else
                    LeftFromLastRead = bufferTemp
                End If
            End If
        Next

        'for 6 meters, start index(incl) to end index(excl)

        '一個個封包分到不同的MAC address下
        For i = 0 To msgs.Count - 1
            If msgs(i).Length >= 8 Then
                Dim add(3 - 1) As Byte
                '封包第五到第七是MAC Address
                Array.Copy(msgs(i), 4, add, 0, 3)

                For j = 0 To MeterMacs.Count - 1
                    If add.SequenceEqual(MeterMacs(j)) Then
                        Dim bytes(msgs(i).Length - 7 - 1) As Byte
                        '封包第八開始是資料
                        Array.Copy(msgs(i), 7, bytes, 0, msgs(i).Length - 7)
                        Dim s As String = Encoding.Default.GetString(bytes)

                        buffer(j).Append(s)
                        Exit For
                    End If
                Next
            Else '不完整的資料留下來跟下次讀入資料結合
                LeftFromLastRead = msgs(i)
            End If
        Next
    End Sub

    'Thread to poll data from server buffer
    Private PollingThread As Thread

    'Multi Threading Lock for buffer
    Dim bufferLock As New Object
    Dim portLock As New Object

    Private Sub Poll()
        Do
            'This will prevent other threads from accessing buffer at the same time
            'If EnterAPIMode(2) Then
            '    For i = 0 To MeterMacs.Count - 1
            '        Dim dodLength As Integer = Encoding.Default.GetByteCount("DOD?" & vbCrLf)
            '        Dim tempBytes(dodLength + 7) As Byte
            '        tempBytes(0) = &H81
            '        tempBytes(1) = dodLength
            '        tempBytes(3) = 2
            '        Array.Copy(MeterMacs(i), 0, tempBytes, 4, 3)
            '        Array.Copy(Encoding.Default.GetBytes("DOD?" & vbCrLf), 0, tempBytes, 7, dodLength)
            '        SyncLock portLock
            '            port.Write(tempBytes, 0, tempBytes.Length)
            '        End SyncLock
            '    Next
            '    If EnterAPIMode(1) Then
            SyncLock portLock
                port.WriteLine("DOD?")
                port.WriteLine("DOD?")
                port.WriteLine("DOD?")
            End SyncLock
            Thread.Sleep(_PollingFrequency)
            SyncLock bufferLock
                ReadAndSortMsgs(False, Nothing)
            End SyncLock
            '    End If
            'End If
        Loop
    End Sub

    'broadcasts start measuring, returns the ones that confirm start measuring
    Public Function StartMeasure() As Boolean
        If Program.sim Then
            Return True
        ElseIf Open() Then
            ClearBuffer()
            If MakeSureReady() Then
                SyncLock portLock
                    port.WriteLine("Measure, Start")
                End SyncLock
                If IsNothing(PollingThread) Then
                    PollingThread = New Thread(AddressOf Poll)
                End If
                PollingThread.Start()
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
                    If pause Then
                        port.WriteLine("Pause, Pause")
                    Else
                        port.WriteLine("Pause, Clear")
                    End If
                End SyncLock
                If pause Then
                    PollingThread = New Thread(AddressOf Poll)
                    PollingThread.Start()
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
        SyncLock bufferLock
            ReadAndSortMsgs(True, "")
        End SyncLock
        SyncLock bufferLock

            For i = 0 To MeterMacs.Count - 1
                If buffer(i).Length > 0 Then
                    result(i) = True
                    buffer(i).Clear()
                End If
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
            port.WriteLine("Measure, Stop")
            PollingThread.Abort()
            PollingThread = Nothing
        End If
    End Sub

    '14 figures
    Public Function ConsumeMeasurementsFromBuffer(ByVal part As Measurements, ByVal name As String) As String()
        Dim result() As String = New String(MeterMacs.Count - 1) {}
        Dim meterIndex As Integer = DetermineMeter(name)
        If Not Program.sim Then
            Try
                If Not IsNothing(buffer) Then
                    SyncLock bufferLock
                        Dim arrayOfLists() As List(Of String) = FilterMsgsFromBuffer("R+0000" + vbCrLf)
                        For i = 0 To MeterMacs.Count - 1
                            If Not meterIndex = -1 Then
                                If Not i = meterIndex Then
                                    result(i) = "0"
                                    Continue For
                                End If
                            End If
                            Dim count As Integer = arrayOfLists(i).Count
                            If count >= 2 Then
                                If arrayOfLists(i)(1).EndsWith("$") Then
                                    Dim temp() As String = arrayOfLists(i)(1).Split(",")
                                    result(i) = temp(part)
                                End If
                            End If
                            If result(i) Is Nothing And count = 1 Then
                                If arrayOfLists(i)(0).EndsWith("$") Then
                                    Dim temp() As String = arrayOfLists(i)(count - 1).Split(",")
                                    result(i) = temp(part)
                                End If

                            End If
                            'always keeping the last one
                            Dim tempItem As String = arrayOfLists(i)(arrayOfLists(i).Count - 1)
                            buffer(i).Clear()
                            buffer(i).Append(tempItem)

                        Next
                    End SyncLock
                End If
            Catch ex As Exception
                MsgBox("ConsumeMeasurementsFromBuffer: " & ex.Message)
            End Try
        Else
            ''TEMP
            Dim r = New Random()
            For i = 0 To MeterMacs.Count - 1
                Dim tempR = (r.Next(1, 10) / 10) + r.Next(40, 119)
                Dim str As String = (tempR & "," & (tempR - 1) & ",--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-")
                result(i) = str.Split(",")(part)
                If Not meterIndex = -1 Then
                    If Not i = meterIndex Then
                        result(i) = 0
                    End If
                End If
            Next
        End If
        
        Return result
    End Function

    Public Function GetMeasurementsFromMeters(ByVal part As Measurements, ByVal name As String) As String()
        Dim result() As String = New String(MeterMacs.Count - 1) {}
        Dim meterIndex As Integer = DetermineMeter(name)
        If Not Program.sim Then
            Try
                If Open() Then
                    result = ConsumeMeasurementsFromBuffer(part, name)
                End If
            Catch ex As Exception
                MsgBox("GetMeasurementsFromMeters: " & ex.Message)
            End Try
            Return result
        Else
            ''TEMP
            Dim r = New Random()
            For i = 0 To MeterMacs.Count - 1
                Dim tempR = (r.Next(1, 10) / 10) + r.Next(40, 119)
                Dim str As String = (tempR & "," & (tempR - 1) & ",--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-,--.-")
                result(i) = str.Split(",")(part)
                If Not meterIndex = -1 Then
                    If Not i = meterIndex Then
                        result(i) = 0
                    End If
                End If
            Next



        End If
        Return result
    End Function

    'API Mode
    'API mode is an alternative to the default transparent operation of the RM024 and provides dynamic packet routing and packet accounting abilities to the OEM host without requiring extensive programming by the OEM host. API mode utilizes specific frame-based packet formats, specifying various vital parameters used to control radio settings and packet routing on a packet-by-packet basis. The API features can be used in any combination that suits the OEM’s application specific needs.
    'The RM024 has three API functions:
    ' Send Data Complete
    ' Receive API
    ' Transmit API

    'mode = 1 > Receive packet
    'mode = 2 > Transmit packet
    'mode = 4 > Send Data complete
    Private Function EnterAPIMode(ByVal mode) As Boolean
        If Open() Then
            Try
                Dim bufCmd() As Byte = {&H41, &H54, &H2B, &H2B, &H2B, &HD}
                Dim sucCmd() As Byte = {&HCC, &H43, &H4F, &H4D, &H0}

                Dim bufWriteAPI() As Byte = {&HCC, &H17, mode}
                Dim sucWriteAPI() As Byte = {&HCC, mode, &H0}

                Dim bufExitCmd() As Byte = {&HCC, &H41, &H54, &H4F, &HD}
                Dim sucExitCmd() As Byte = {&HCC, &H44, &H41, &H54, &H0}

                'Enter AT Command Mode
                SyncLock portLock
                    port.Write(bufCmd, 0, bufCmd.Length)
                    Thread.Sleep(_Latency)
                    Dim buffer() As Byte = New Byte(port.BytesToRead) {}
                    port.Read(buffer, 0, buffer.Length)

                    'if Enter AT Command Mode successful
                    If buffer.SequenceEqual(sucCmd) Then

                        'Write API Control-Transmit API
                        port.Write(bufWriteAPI, 0, bufWriteAPI.Length)
                        Thread.Sleep(_Latency)
                        buffer = New Byte(port.BytesToRead) {}
                        port.Read(buffer, 0, buffer.Length)

                        'If Successfully write API Control-Transmit API
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
                End SyncLock
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        Return False
    End Function

    'This Will Set up the server to enter API mode
    Public Function SetupServer() As Boolean
        EnableBroadCast()
        Return EnterAPIMode(1)
    End Function
End Class
