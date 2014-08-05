Imports System
Imports System.Text
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports Microsoft.VisualBasic
Imports System.Net.NetworkInformation

'此class是跟ANDROID用TCP的方式溝通
Public Class CommunicationWithAndroid
    '與ANDROID的IP address 和 port連線
    Public Shared Function ConnectSocket(ByVal serverIP As String, ByVal port As Integer) As Socket
        Dim hostEntry As IPHostEntry = Nothing

        Dim ip As IPAddress = System.Net.Dns.GetHostAddresses(serverIP)(0)
        'Dim endPoint As New IPEndPoint(System.Net.Dns.GetHostEntry(serverIP).AddressList(0), port)
        Dim endpoint As New IPEndPoint(ip, port)
        Dim tempSocket As New Socket(endpoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp)
        Try
            tempSocket.Connect(endpoint)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        If tempSocket.Connected Then
            Return tempSocket
        End If
        Return Nothing

    End Function

    Public Shared Function CloseSocket(ByRef socket As Socket) As Boolean
        Try
            socket.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        If socket IsNot Nothing Then
            If Not socket.Connected Then
                Return True
            End If
        End If
        Return False
    End Function

    '一個好用的轉換string 到 long integer 的function
    Public Shared Function LongIntegerFromIP(ByVal p_strIP As String) As Long
        Dim arrTemp As Object
        Dim i As Integer
        Dim lngTemp As Long

        arrTemp = Split(p_strIP, ".")

        For i = 0 To UBound(arrTemp)
            lngTemp = lngTemp + CLng(arrTemp(i)) * (256 ^ (3 - i))
        Next
        Return lngTemp
    End Function
End Class
