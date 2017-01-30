Attribute VB_Name = "NmeaDeclarations"
Public Type ggaInfo_t
' $GPGGA,141449.00,2948.24652,N,09533.52845,W,2,9,0.8,34.29,M,-25.92,M,8,0108*6F
    Utc As Double
    lat As Double
    LatHemi As String
    lon As Double
    LonHemi As String
    Quality As Integer
    SatsUsed As Integer
    hdop As Single
    Altitude As Single
    AltUnit As String
    GeoSep As Single
    GeoSepUnit As String
    hae As Single
    diffage As Single
    StationID As Integer
    LastUpdate As Long
End Type

Public Type gllInfo_t
' $GPGLL,2527.2572,N,05520.5600,E,095837,A*xx
' $GPGLL,2948.245862,N,09533.528832,W,000010.00,A,D*7D
    lat As Double
    LatHemi As String
    lon As Double
    LonHemi As String
    Utc As Double
    Status As String
    Mode As String
    LastUpdate As Long
End Type

Public Type gsaInfo_t
' $GPGSA,M,3,,8,1,2,7,,3,13,11,31,28,,2.2,0.8,2.0*01
    Mode As String
    ModeStat As Integer
    SatID(42) As Integer
    pdop As Single
    hdop As Single
    vdop As Single
    LastUpdate As Long
End Type

Public Type gstInfo_t
' $GPGST,141448.0,0.0432,0.0768,0.0613,-53.2385,0.0673,0.0716,0.2887*43
    Utc As Double
    Rms As Single
    SDSemiMajor As Single
    SDSemiMinor As Single
    OrientSemiMajor As Single
    SDLat As Single
    SDLon As Single
    SDAlt As Single
    LastUpdate As Long
End Type

Public Type satsInfo_t
' $PNCTR,SATS,4,1,01,15,169,47,39,02,08,071,43,38,03,13,043,40,38*70
' $PNCTR,SATS,4,2,07,11,217,45,39,08,44,325,51,45,11,28,131,53,42*70
' $PNCTR,SATS,4,3,13,42,195,51,42,28,31,279,49,43,29,02,326,00,00*72
' $PNCTR,SATS,4,4,31,51,048,52,46*71
    MessageMax As Integer
    MessageNum As Integer
    SatID(36) As Integer
    SatElev(36) As Integer
    SatAzimuth(36) As Integer
    SatL1SNR(36) As Integer
    SatL2SNR(36) As Integer
    SatsInView As Integer
    LastUpdate As Long
End Type

Public Type gsvInfo_t
' $GPGSV,3,1,09,05,14,050,08,11,22,296,10,14,51,030,18,15,11,174,07*79
' $GPGSV,3,2,09,16,03,184,,18,20,126,07,23,46,118,,25,68,237,18*7C
' $GPGSV,3,3,09,30,30,077,13,,,,,,,,,,,,*42
    MessageMax As Integer
    MessageNum As Integer
    SatsInView As Integer
    SatID(36) As Integer
    SatElev(36) As Integer
    SatAzimuth(36) As Integer
    SatSNR(36) As Integer
    LastUpdate As Long
End Type

Public Type hdtInfo_t
' $HEHDT,348.0,T*20
    HeadingT As Single
    HeadingTID As String
    LastUpdate As Long
End Type

Public Type hdmInfo_t
' $HEHDT,348.0,M*xx
    HeadingM As Single
    HeadingMID As String
    LastUpdate As Long
End Type

Public Type sgbInfo_t
    HeadingT As Single
    LastUpdate As Long
End Type

Public Type vtgInfo_t
' $GPVTG,0,T,,,0.00,N,0.00,K*33
' $GPVTG,0.0,T,,M,0.04,N,0.07,K,D*0B
    cogt As Single
    CogTID As String
    CogM As Single
    CogMID As String
    SpeedKt As Single
    SpeedKtID As String
    SpeedKmh As Single
    SpeedKmhID As String
    Mode As String
    LastUpdate As Long
End Type

Public Type zdaInfo_t
' $GPZDA,141449.00,3,2,2004,+0,+0*6C
    Utc As Double
    Day As Integer
    Month As Integer
    Year As Integer
    LocalTZhr As Integer
    LocalTZmn As Integer
    LastUpdate As Long
End Type

Public Type rd1Info_t
    WeekSec As Double
    Week As Integer
    Frequency As Double
    Lock As Integer
    BitErrorRate1 As Integer
    BitErrorRate2 As Integer
    Agc As Integer
    Dds As Integer
    Doppler As Integer
    DspStatus As String
    ArmStatus As String
    DiffStatus As String
    NavCondition As String
    LastUpdate As Long
End Type

Public Type rxqInfo_t
' $PNCTR,RXQ,141450,Y,12.3,2,0*40
    Utc As Double
    SFLock As String
    SFSNR As Double
    PerIdlePacket As Integer
    PerBadPacket As Integer
    LastUpdate As Long
End Type

Public Type navqInfo_t
' either - $PNCTR,NAVQ,123519,3D,RTG,DUAL*55
'     or - $PNCTR,NAVQ,202759,NN*74
    Utc As Double
    NavMode As String
    CorrType As String
    SignalType As String
    LastUpdate As Long
End Type

Public Type ohprInfo_t
'$OHPR,78.8,-0.5,-1.4,-0.009,-0.025,0.976*34
    Heading As Double
    Pitch As Double
    Roll As Double
    Depth As Double
    LastUpdate As Long
End Type

Public Type encoderInfo_t
'Encoder String
    Value As Double
    LastUpdate As Long
End Type

Public Type flowInfo_t
'Encoder String
    Value As Double
    LastUpdate As Long
End Type

Public Type NmeaInfo_t
    gga As ggaInfo_t
    gll As gllInfo_t
    gsa As gsaInfo_t
    gst As gstInfo_t
    gsv As gsvInfo_t
    sats As satsInfo_t
    hdt As hdtInfo_t
    hdm As hdmInfo_t
    sgb As sgbInfo_t
    vtg As vtgInfo_t
    zda As zdaInfo_t
    rd1 As rd1Info_t
    rxq As rxqInfo_t
    navq As navqInfo_t
    ohpr As ohprInfo_t
    encoder As encoderInfo_t
    flow As flowInfo_t
End Type

Public Declare Function GetTickCount& Lib "kernel32" () 'added 000222

'$GPTRF
'Transit Fix Data
'Time, date, position, and information related to a TRANSIT Fix.
'$--TRF,hhmmss.ss,xxxxxx,llll.ll,a,yyyyy.yy,a,x.x,x.x,x.x,x.x,xxx
'hhmmss.ss = UTC of position fix
'xxxxxx = Date: dd/mm/yy
'llll.ll,a = Latitude of position fix, N/S
'yyyyy.yy,a = Longitude of position fix, E/W
'x.x = Elevation angle
'x.x = Number of iterations
'x.x = Number of Doppler intervals
'x.x = Update distance, nautical miles
'x.x = Satellite ID
