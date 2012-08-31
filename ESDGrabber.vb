Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Text.RegularExpressions

'***********************************************************************
'
' ESD Grabber Class         5 March 2012
' Author: Capt Brandon Magnuson
'
' Description: Use this class to connect to an ESD Server, and get parsed messages
'
'
' How to use:
'
' Create a new instance of this class
' Initialize the class by calling the connectToESD() method
' Then, create a background thread that runs UpdateESD()
'     this thread will run as long as you have a goodConnection
'
' Access all properties by the "get" prefix then the name of the property
'
' When you are done with the ESD, call the closeESD() method
'
'**************************************************************************

Public Class ESDGrabber

    'Constants
    Private Const kMinValue = -2
    Private Const kMaxValue = 999
    Private Const kNullString = "None"

    'Values that cannot be negative
    Private _Sl As Integer = kMinValue
    Private _Da As Integer = kMinValue
    Private _Te As Integer = kMinValue
    Private _Wd As Integer = kMinValue
    Private _Ws As Integer = kMinValue
    Private _Tas As Integer = kMinValue
    Private _Ai As Integer = kMinValue
    Private _Sn As Integer = kMinValue
    Private _Ih As Double = kMinValue
    Private _Sp As Double = kMinValue
    Private _Sr As Double = kMinValue
    Private _Tw As Double = kMinValue
    Private _Ic As Integer = kMinValue
    Private _Fv As Double = kMinValue
    Private _Az As Double = kMinValue
    Private _Fr As Integer = kMinValue
    Private _Gr As Double = kMinValue
    Private _Gv As Integer = kMinValue
    Private _Ha As Integer = kMinValue
    Private _Id As Integer = kMinValue
    Private _La As Integer = kMinValue
    Private _Lc As Integer = kMinValue
    Private _Lf As Integer = kMinValue
    Private _Pt As Integer = kMinValue
    Private _Vn As Integer = kMinValue

    'Values that can be negative
    Private _Ir As Double = kMaxValue
    Private _Ip As Double = kMaxValue
    Private _Se As Double = kMaxValue
    Private _At As Integer = kMaxValue
    Private _De As Double = kMaxValue
    Private _Sz As Double = kMaxValue

    'Strings
    Private _ALat As String = kNullString
    Private _ALon As String = kNullString
    Private _TLat As String = kNullString
    Private _TLon As String = kNullString
    Private _Ct As String = kNullString
    Private _Cd As String = kNullString
    Private _Dump As String = kNullString

    Private _lastGoodUpdate As Date = Now()

    'TCP Variables
    Private tcpConnection As System.Net.Sockets.TcpClient
    Private esdStream As NetworkStream
    Public goodConnection As Boolean = False
    Private failedStreamCount As Integer = 0
    Private failedStreamAllowance As Integer = 1

    Public Sub connectToESD(ByVal myipAddress As String, ByVal myport As Integer)
        'goodConnection = False
        'Exit Sub
        Try
            If Not goodConnection Then
                tcpConnection = New System.Net.Sockets.TcpClient
                If Not tcpConnection.Connected Then
                    tcpConnection.Connect(myipAddress, myport)
                    esdStream = tcpConnection.GetStream()
                    If esdStream.CanRead Then
                        goodConnection = True
                    End If
                End If
            End If
        Catch ex As Exception
            goodConnection = False
        End Try
    End Sub
    Public Sub updateESD()
        While goodConnection
            Try
                'Grab some data
                Dim bytes(tcpConnection.ReceiveBufferSize) As Byte
                esdStream.ReadTimeout = 250
                esdStream.Read(bytes, 0, CInt(tcpConnection.ReceiveBufferSize))
                Dim returndata As String = Encoding.ASCII.GetString(bytes)
                For i = 0 To 20
                    If bytes(i) > 0 Then Exit For
                    Throw New Exception
                Next
                'Parse Data and send to public variables
                Dim m As Match

                m = New Regex("Sl([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Sl = m.Groups(1).Value
                m = New Regex("Da([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Da = m.Groups(1).Value
                m = New Regex("Te([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Te = m.Groups(1).Value
                m = New Regex("Wd([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Wd = m.Groups(1).Value
                m = New Regex("Ws([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Ws = m.Groups(1).Value
                m = New Regex("As([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Tas = m.Groups(1).Value
                m = New Regex("Sn([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Sn = m.Groups(1).Value
                m = New Regex("Ai([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Ai = m.Groups(1).Value
                m = New Regex("Ih([0-9]*.[0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Ih = m.Groups(1).Value
                m = New Regex("Ip([+-][0-9]*.[0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Ip = m.Groups(1).Value
                m = New Regex("Ir([+-][0-9]*.[0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Ir = m.Groups(1).Value
                m = New Regex("Sp([0-9]*.[0-9]*)[A-Z]").Match(returndata)
                If m.Success Then
                    _Sp = m.Groups(1).Value
                    _lastGoodUpdate = Now()
                End If
                m = New Regex("Se([+-][0-9]*.[0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Se = m.Groups(1).Value
                m = New Regex("Sr([0-9]*.[0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Sr = m.Groups(1).Value
                m = New Regex("Tw([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Tw = m.Groups(1).Value
                m = New Regex("Ic([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Ic = m.Groups(1).Value
                m = New Regex("Fv([+-][0-9]*.[0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Fv = m.Groups(1).Value
                m = New Regex("At([+-][0-9]*.[0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _At = m.Groups(1).Value
                m = New Regex("Ta([+|-][0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _TLat = m.Groups(1).Value
                m = New Regex("To([+|-][0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _TLon = m.Groups(1).Value
                m = New Regex("Sa([+|-][0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _ALat = m.Groups(1).Value
                m = New Regex("So([+|-][0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _ALon = m.Groups(1).Value
                m = New Regex("Ct([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Ct = m.Groups(1).Value
                m = New Regex("Cd([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Cd = m.Groups(1).Value
                m = New Regex("Az([0-9]*.[0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Az = m.Groups(1).Value
                m = New Regex("Fr([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Fr = m.Groups(1).Value
                m = New Regex("Gr([0-9]*.[0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Gr = m.Groups(1).Value
                m = New Regex("Gv([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Gv = m.Groups(1).Value
                m = New Regex("Ha([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Ha = m.Groups(1).Value
                m = New Regex("Id([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Id = m.Groups(1).Value
                m = New Regex("La([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _La = m.Groups(1).Value
                m = New Regex("Lc([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Lc = m.Groups(1).Value
                m = New Regex("Lf([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Lf = m.Groups(1).Value
                m = New Regex("Pt([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Pt = m.Groups(1).Value
                m = New Regex("Lc([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Lc = m.Groups(1).Value
                m = New Regex("Vn([0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Vn = m.Groups(1).Value
                m = New Regex("De([0-9]*.[0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _De = m.Groups(1).Value
                m = New Regex("Sz([0-9]*.[0-9]*)[A-Z]").Match(returndata)
                If m.Success Then _Sz = m.Groups(1).Value
                _Dump = returndata

                'reset Failed Stream
                failedStreamCount = 0
            Catch ex As Exception
                failedStreamCount += 1
                If failedStreamCount > failedStreamAllowance Then
                    closeConnection()
                    failedStreamCount = 0
                End If
            End Try
        End While
    End Sub
    Public Sub closeConnection()
        Try


            esdStream.Close()
            esdStream.Dispose()
            tcpConnection.Close()
            goodConnection = False
        Catch ex As Exception
            goodConnection = False
        End Try
    End Sub
    ''' <summary>
    ''' MSL Altitude (0-128,070)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Sl() As Integer
        Get
            Return _Sl
        End Get
    End Property
    ''' <summary>
    ''' Density Alitude (-3000 - 128,070)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Da() As Integer
        Get
            Return _Da
        End Get
    End Property
    ''' <summary>
    ''' Target Elevation (-32,767 - +32,767)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Te() As Integer
        Get
            Return _Te
        End Get
    End Property
    ''' <summary>
    ''' Wind Direction (0-360)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Wd() As Integer
        Get
            Return _Wd
        End Get
    End Property
    ''' <summary>
    ''' Wind Speed (0-255)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Ws() As Integer
        Get
            Return _Ws
        End Get
    End Property
    ''' <summary>
    ''' True Airspeed (0-999)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Tas() As Integer
        Get
            Return _Tas
        End Get
    End Property
    ''' <summary>
    ''' Indicated Airspeed (0-999)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Ai() As Integer
        Get
            Return _Ai
        End Get
    End Property
    ''' <summary>
    ''' Source of the video Image (0:EO nose, 1:EO Zoom, 2:EO Spotter, 3:IR)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Sn() As Integer
        Get
            Return _Sn
        End Get
    End Property
    ''' <summary>
    ''' Heading (0-359.99)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Ih() As Double
        Get
            Return _Ih
        End Get
    End Property
    ''' <summary>
    ''' Laser Target Line (0-359.99)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Sp() As Double
        Get
            Return _Sp
        End Get
    End Property
    ''' <summary>
    ''' Slant Range (0-655.35)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Sr() As Double
        Get
            Return _Sr
        End Get
    End Property
    ''' <summary>
    ''' Target Width (0-608027)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Tw() As Integer
        Get
            Return _Tw
        End Get
    End Property
    ''' <summary>
    ''' Image Coordinate Format (0-2)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Ic() As Integer
        Get
            Return _Ic
        End Get
    End Property
    ''' <summary>
    ''' Angle of the Field of view (0-179.99)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Fv() As Double
        Get
            Return _Fv
        End Get
    End Property
    ''' <summary>
    ''' Azimuth (0-359.99)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Az() As Double
        Get
            Return _Az
        End Get
    End Property
    ''' <summary>
    ''' Fuel Remaining (0-1020?)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Fr() As Integer
        Get
            Return _Fr
        End Get
    End Property
    ''' <summary>
    ''' Ground Range (0-655.35)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Gr() As Double
        Get
            Return _Gr
        End Get
    End Property
    ''' <summary>
    ''' Ground Speed (0-999)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Gv() As Integer
        Get
            Return _Gv
        End Get
    End Property
    ''' <summary>
    ''' HAT (-1 if no target) (-32,767- +32,767)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Ha() As Integer
        Get
            Return _Ha
        End Get
    End Property
    ''' <summary>
    ''' Icing Detected (0=off,1=No,2=Yes)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Id() As Integer
        Get
            Return _Id
        End Get
    End Property
    ''' <summary>
    ''' Laser Armed
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property La() As Boolean
        Get
            If _La = 1 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
    ''' <summary>
    ''' Laser PRF
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Lc() As Integer
        Get
            Return _Lc
        End Get
    End Property
    ''' <summary>
    ''' Laser Firing
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Lf() As Boolean
        Get
            If _Lf = 1 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
    ''' <summary>
    ''' Tail Number (0-99,999)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Pt() As Integer
        Get
            Return _Pt
        End Get
    End Property
    ''' <summary>
    ''' Field of view name (0:UN,1:UW,2:Med,3:Nar,4:UN,5:Wide,6:2x,7:4x)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Vn() As Integer
        Get
            Return _Vn
        End Get
    End Property
    ''' <summary>
    ''' Roll (-60 - +60)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Ir() As Double
        Get
            Return _Ir
        End Get
    End Property
    ''' <summary>
    ''' Pitch (-60 - +60)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Ip() As Double
        Get
            Return _Ip
        End Get
    End Property
    ''' <summary>
    ''' Wings Level Depression (-90 - +90) (0=Horizon,-90=Straight Down)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Se() As Double
        Get
            Return _Se
        End Get
    End Property
    ''' <summary>
    ''' Outside Air Temperature (-127-+127)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property At() As Integer
        Get
            Return _At
        End Get
    End Property
    ''' <summary>
    ''' Depression Angle (-180-+180)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property De() As Double
        Get
            Return _De
        End Get
    End Property
    ''' <summary>
    ''' Payload FOV Roll angle (-90 - +90)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Sz() As Double
        Get
            Return _Sz
        End Get
    End Property
    ''' <summary>
    ''' Aircraft Latitude (PDDMMSST)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ALat() As String
        Get
            Return _ALat
        End Get
    End Property
    ''' <summary>
    ''' Aircraft Longitude (PDDDMMSST)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ALon() As String
        Get
            Return _ALon
        End Get
    End Property
    ''' <summary>
    ''' Target Latitude (PDDMMSST)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property TLat() As String
        Get
            Return _TLat
        End Get
    End Property
    ''' <summary>
    ''' Target Longitude (PDDDMMSST)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property TLon() As String
        Get
            Return _TLon
        End Get
    End Property
    ''' <summary>
    ''' Collection Time (HHMMSS)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Ct() As String
        Get
            Return _Ct
        End Get
    End Property
    ''' <summary>
    ''' Collection Date (CCYYMMDD)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Cd() As String
        Get
            Return _Cd
        End Get
    End Property
    ''' <summary>
    ''' get timeStamp of GCS date/time
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property timeStamp() As Date
        Get
            Try
                Dim year As Integer = _Cd.Substring(0, 4)
                Dim month As Integer = _Cd.Substring(4, 2)
                Dim day As Integer = _Cd.Substring(6, 2)
                Dim hour As Integer = _Ct.Substring(0, 2)
                Dim minute As Integer = _Ct.Substring(2, 2)
                Dim second As Integer = _Ct.Substring(4, 2)

                Dim myDate As DateTime = New DateTime(year, month, day, hour, minute, second, DateTimeKind.Utc)
                Return myDate
            Catch ex As Exception
                Return New DateTime(1, 1, 1, 1, 1, 1, DateTimeKind.Utc)
            End Try
        End Get
    End Property
    ''' <summary>
    ''' ESD Dump
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Dump() As String
        Get
            Return _Dump
        End Get
    End Property
    ''' <summary>
    ''' gets the status of the esdconnection (None, Ground,Degraded,Transit,Full)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getStatus() As String
        Dim returnString As String = Nothing
        If Not (goodConnection And Not (_Sl = -2)) Then
            returnString = "None"
        ElseIf _lastGoodUpdate < Now.AddSeconds(-30) Then
            returnString = "Stale"
        ElseIf (_ALat = "+0000000" And _ALon = "+00000000") Then
            returnString = "Idle"
        ElseIf _Ws = -2 Or _Wd = -2 Or _Tas = -2 Then
            returnString = "Degraded"
        ElseIf _TLat = "+0000000" And _TLon = "+00000000" Then
            returnString = "Transit"
        Else
            returnString = "Full"
        End If
        Return returnString
    End Function

End Class
