VERSION 5.00
Object = "{E0110BE7-EBF0-4612-B2F8-817194557140}#1.0#0"; "icimagingcontrol.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmVideo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Video"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19350
   Icon            =   "frmVideo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   19350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin ComctlLib.Slider Slider2 
      Height          =   495
      Left            =   9960
      TabIndex        =   18
      Top             =   7920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   327682
   End
   Begin ComctlLib.Slider SliderLight1 
      Height          =   495
      Left            =   4800
      TabIndex        =   17
      Top             =   7920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   327682
      Max             =   10000
   End
   Begin ComctlLib.Slider Slider3 
      Height          =   495
      Left            =   2520
      TabIndex        =   16
      Top             =   7920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   327682
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   7920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      _Version        =   327682
   End
   Begin VB.CommandButton VideoInit 
      Caption         =   "Init Video"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   14
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer OverlayTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9360
      Top             =   8160
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   9720
      TabIndex        =   6
      Top             =   7560
      Width           =   9495
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Timer Timer2 
         Interval        =   30000
         Left            =   9000
         Top             =   120
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Record Video 2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin ComctlLib.Slider Slider4 
         Height          =   495
         Left            =   2520
         TabIndex        =   19
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   327682
      End
      Begin ComctlLib.Slider SliderLight2 
         Height          =   495
         Left            =   4680
         TabIndex        =   20
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   327682
         Max             =   10000
      End
      Begin VB.Label Label6 
         Caption         =   "Video 2 LED Light"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   11
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Video 2 Brightness"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Video 2 Contrast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   7560
      Width           =   9495
      Begin VB.Timer Timer1 
         Interval        =   30000
         Left            =   9000
         Top             =   120
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Record Video 1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Video 1 LED Light"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   10
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Video 1 Brightness"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Video 1 Contrast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
   End
   Begin ICImagingControl3Ctl.ICImagingControl ICImagingControl2 
      Height          =   7455
      Left            =   9720
      OleObjectBlob   =   "frmVideo.frx":014A
      TabIndex        =   1
      Top             =   0
      Width           =   9615
   End
   Begin ICImagingControl3Ctl.ICImagingControl ICImagingControl1 
      Height          =   7455
      Left            =   0
      OleObjectBlob   =   "frmVideo.frx":01BC
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private List() As Control
Private curr_obj As Object
Private iHeight As Integer
Private iWidth As Integer
Private x_size As Double
Private y_size As Double

Private Type Control
    Index As Integer
    Name As String
    Left As Integer
    Top As Integer
    width As Integer
    height As Integer
End Type

Dim current1File As String
Dim current2File As String
Dim file1Time As Double
Dim file2Time As Double
Public Logo1 As Integer
Public Logo2 As Integer
Public Date1 As Integer
Public Date2 As Integer
Public Attitude1 As Integer
Public Attitude2 As Integer
Public Circle1 As Integer
Public Circle2 As Integer
Public Info1 As Integer
Public Info2 As Integer
Private VCDProp1 As VCDSimpleProperty
Private VCDProp2 As VCDSimpleProperty
Dim Avi1Comp As AviCompressor
Dim Avi2Comp As AviCompressor

Private Sub ResizeControls(frm As Form)
Dim i As Integer
'   Get ratio of initial form size to current form size
x_size = frm.height / iHeight
y_size = frm.width / iWidth

'Loop though all the objects on the form
'Based on the upper bound of the # of controls
For i = 0 To UBound(List)
    'Grad each control individually
    For Each curr_obj In frm
        'Check to make sure its the right control
        If curr_obj.TabIndex = List(i).Index Then
            'Then resize the control
             With curr_obj
                .Left = List(i).Left * y_size
                .width = List(i).width * y_size
                .height = List(i).height * x_size
                .Top = List(i).Top * x_size
             End With
        End If
    'Get the next control
    Next curr_obj
Next i
End Sub

Private Function SetFontSize() As Integer
    'Make sure x_size is greater than 0
    If Int(x_size) > 0 Then
    'Set the font size
        SetFontSize = Int(x_size * 8)
    End If
End Function

Private Sub GetLocation(frm As Form)
Dim i As Integer
'   Load the current positions of each object into a user defined type array.
'   This information will be used to rescale them in the Resize function.

'Loop through each control
For Each curr_obj In frm
'Resize the Array by 1, and preserve
'the original objects in the array
If TypeOf curr_obj Is Timer Then
    'Wrong type of control
Else
    ReDim Preserve List(i)
    With List(i)
'        .Name = curr_obj
        .Index = curr_obj.TabIndex
        .Left = curr_obj.Left
        .Top = curr_obj.Top
        .width = curr_obj.width
        .height = curr_obj.height
    End With
    i = i + 1
End If
Next curr_obj
    
'   This is what the object sizes will be compared to on rescaling.
    iHeight = frm.height
    iWidth = frm.width
End Sub

Private Sub Check1_Click()
    Dim CompressorNumber As Integer
    If ICImagingControl1.DeviceValid And ICImagingControl1.SignalDetected Then
        If Check1.Value = 1 Then
            ' Get the index of the selected compressor in the combo box.
'            CompressorNumber = Combo1.ListIndex + 1
            ' Get the AviCompressor object from ICImagingControl.
'            Set Avi1Comp = ICImagingControl1.AviCompressors.Item(CompressorNumber)
            current1File = Format(Now, "yyyymmddHHmmss")
            ' create current time in decimal minutes
            file1Time = (CDbl(Format(Now, "dd")) * 3600) + (CDbl(Format(Now, "Hh")) * 60) + CDbl(Format(Now, "Nn"))
            ICImagingControl1.AviStartCapture ViewForm.LoggingDir & "\" & current1File & "V1.avi", Avi1Comp.Name
            Combo1.Enabled = False
            ViewForm.Video1Start = ViewForm.LoggingDir & current1File & "V1.avi"
        Else
            ICImagingControl1.AviStopCapture
            ICImagingControl1.LiveDisplayDefault = False
            ICImagingControl1.LiveDisplayHeight = ICImagingControl1.height / 15 ' 15 converts twips to pixels
            ICImagingControl1.LiveDisplayWidth = ICImagingControl1.width / 15
            Call sm_wait(1000)
            ICImagingControl1.LiveStart
            Combo1.Enabled = True
        End If
    Else
        Check1.Value = 0
    End If
    
End Sub

Private Sub Check2_Click()
    Dim CompressorNumber As Integer
    ViewForm.Video2Start = makeUnixTime(Now)
    If ICImagingControl2.DeviceValid And ICImagingControl2.SignalDetected Then
        If Check2.Value = 1 Then
            ' Get the index of the selected compressor in the combo box.
'            CompressorNumber = Combo2.ListIndex + 1
            ' Get the AviCompressor object from ICImagingControl.
'            Set Avi2Comp = ICImagingControl2.AviCompressors.Item(CompressorNumber)
            current2File = Format(Now, "yyyymmddHHmmss")
            ' create current time in decimal minutes
            file2Time = (CDbl(Format(Now, "dd")) * 3600) + (CDbl(Format(Now, "Hh")) * 60) + CDbl(Format(Now, "Nn"))
            ICImagingControl2.AviStartCapture ViewForm.LoggingDir & "\" & current2File & "V2.avi", Avi2Comp.Name
            Combo2.Enabled = False
            ViewForm.Video2Start = ViewForm.LoggingDir & current2File & "V2.avi"
        Else
            ICImagingControl2.AviStopCapture
            ICImagingControl2.LiveDisplayDefault = False
            ICImagingControl2.LiveDisplayHeight = ICImagingControl1.height / 15 ' 15 converts twips to pixels
            ICImagingControl2.LiveDisplayWidth = ICImagingControl1.width / 15
            Call sm_wait(1000)
            ICImagingControl2.LiveStart
            Combo2.Enabled = True
        End If
    Else
        Check2.Value = 0
    End If
    
End Sub

Private Sub Combo1_Click()
    Dim CompressorNumber As Integer
'    Dim AviComp As AviCompressor
    
    ' Get the index of the selected compressor in the combo box.
    CompressorNumber = Combo1.ListIndex
        
    ' Get the AviCompressor object from ICImagingControl.
'    Set Avi1Comp = ICImagingControl1.AviCompressors.Item(CompressorNumber)
    
    Check1.Enabled = True

End Sub

Private Sub Combo2_Click()
    Dim CompressorNumber As Integer
'    Dim AviComp As AviCompressor
    
    ' Get the index of the selected compressor in the combo box.
    CompressorNumber = Combo2.ListIndex
        
    ' Get the AviCompressor object from ICImagingControl.
'    Set Avi2Comp = ICImagingControl2.AviCompressors.Item(CompressorNumber)
    
    Check2.Enabled = True


End Sub

Private Sub Form_Load()
Dim Dev1 As Device
Dim Dev2 As Device

'Call GetLocation(frmVideo)

'On Error Resume Next
' Get the first video capture device
Set Dev1 = ICImagingControl1.Devices.Item(1)
Set Dev2 = ICImagingControl2.Devices.Item(2)
' Set the video capture device with the name property of Dev
ICImagingControl1.Device = Dev1.Name
ICImagingControl2.Device = Dev2.Name
' Make sure the video capture device is valid and initialize
'Call initVideo
If ICImagingControl1.DeviceValid And ICImagingControl2.DeviceValid Then
    Call initVideo
Else
    Call MsgBox("Video Devices not detected!", vbOKOnly, "Device Error")
End If
    
End Sub

Public Function initVideo()
On Error Resume Next
    Dim Codec As AviCompressor
    ' Fill the combobox with the available avi compressors (codecs).
    For Each Codec In ICImagingControl1.AviCompressors
        Combo1.AddItem Codec.Name
        Debug.Print Codec.Name
'        If Codec.Name = "x264vfw - H.264/MPEG-4 AVC codec" Then
        If LCase(Codec.Name) = "xvid mpeg-4 codec" Then
            VideoCompressorNum = Combo1.ListCount
            Debug.Print CStr(VideoCompressorNum)
        End If
    Next
    
    Set Avi1Comp = ICImagingControl1.AviCompressors.Item(VideoCompressorNum)
    Debug.Print ICImagingControl1.AviCompressors.Item(VideoCompressorNum).Name
    
    Combo1.ListIndex = VideoCompressorNum
    
    For Each Codec In ICImagingControl1.AviCompressors
        Combo1.AddItem Codec.Name
    Next
    Combo1.ListIndex = 0
    
    ' Fill the combobox with the available avi compressors (codecs).
    For Each Codec In ICImagingControl2.AviCompressors
        Combo2.AddItem Codec.Name
        If LCase(Codec.Name) = "xvid mpeg-4 codec" Then
            VideoCompressorNum = Combo2.ListCount
        End If
    Next
    
    Set Avi2Comp = ICImagingControl2.AviCompressors.Item(VideoCompressorNum)
    Debug.Print ICImagingControl2.AviCompressors.Item(VideoCompressorNum).Name
    
    Combo2.ListIndex = VideoCompressorNum
    
    For Each Codec In ICImagingControl2.AviCompressors
        Combo2.AddItem Codec.Name
    Next
    Combo2.ListIndex = 0
    
    If ICImagingControl1.DeviceValid Then
'    If ICImagingControl1.DeviceValid And ICImagingControl1.SignalDetected Then
        ' Initialize the VCDProp class to access the properties of our ICImagingControl
        ' object
        Set VCDProp1 = GetSimplePropertyContainer(ICImagingControl1.VCDPropertyItems)
        
        ' Setup the range of the brightness slider.
        Slider1.Min = VCDProp1.RangeMin(VCDID_Brightness)
        Slider1.Max = VCDProp1.RangeMax(VCDID_Brightness)
        ' Set the slider to the current brightness value.
        Slider1.Value = VCDProp1.RangeValue(VCDID_Brightness)
        Slider3.Min = VCDProp1.RangeMin(VCDID_Contrast)
        Slider3.Max = VCDProp1.RangeMax(VCDID_Contrast)
        Slider3.Value = VCDProp1.RangeValue(VCDID_Contrast)
        ICImagingControl1.LiveDisplayDefault = False
        ICImagingControl1.LiveDisplayHeight = ICImagingControl1.height / 15 ' 15 converts twips to pixels
        ICImagingControl1.LiveDisplayWidth = ICImagingControl1.width / 15
        ICImagingControl1.LiveStart
    End If
    If ICImagingControl2.DeviceValid Then
'    If ICImagingControl2.DeviceValid And ICImagingControl2.SignalDetected Then
        ' Initialize the VCDProp class to access the properties of our ICImagingControl
        ' object
        Set VCDProp2 = GetSimplePropertyContainer(ICImagingControl2.VCDPropertyItems)
        
        ' Setup the range of the brightness slider.
        Slider2.Min = VCDProp2.RangeMin(VCDID_Brightness)
        Slider2.Max = VCDProp2.RangeMax(VCDID_Brightness)
        ' Set the slider to the current brightness value.
        Slider2.Value = VCDProp2.RangeValue(VCDID_Brightness)
        Slider4.Min = VCDProp2.RangeMin(VCDID_Contrast)
        Slider4.Max = VCDProp2.RangeMax(VCDID_Contrast)
        Slider4.Value = VCDProp2.RangeValue(VCDID_Contrast)
        ICImagingControl2.LiveDisplayDefault = False
        ICImagingControl2.LiveDisplayHeight = ICImagingControl2.height / 15 ' 15 converts twips to pixels
        ICImagingControl2.LiveDisplayWidth = ICImagingControl2.width / 15
        ICImagingControl2.LiveStart
    End If

Logo1 = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Logo1", Default:="0"))
Date1 = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Date1", Default:="0"))
Attitude1 = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Attitude1", Default:="0"))
Circle1 = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Circle1", Default:="0"))
Info1 = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Info1", Default:="0"))
Logo2 = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Logo2", Default:="0"))
Date2 = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Date2", Default:="0"))
Attitude2 = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Attitude2", Default:="0"))
Circle2 = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Circle2", Default:="0"))
Info2 = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Info2", Default:="0"))

OverlayTimer.Enabled = True

End Function

Private Sub Form_Resize()
'Call ResizeControls(frmVideo)
End Sub

Private Sub Form_Unload(Cancel As Integer)
OverlayTimer.Enabled = False
Call sm_wait(100)
If Check1.Value = 1 Then
    ICImagingControl1.AviStopCapture
End If
If ICImagingControl1.LiveVideoRunning Then
    ICImagingControl1.LiveStop
End If
If Check2.Value = 1 Then
    ICImagingControl2.AviStopCapture
End If
If ICImagingControl2.LiveVideoRunning Then
    ICImagingControl2.LiveStop
End If
End Sub

Private Sub OverlayTimer_Timer()
Call DrawOverlay
End Sub

Private Sub Slider1_Scroll()
    If ICImagingControl1.LiveVideoRunning Then
        VCDProp1.RangeValue(VCDID_Brightness) = Slider1.Value
    End If
End Sub
Private Sub Slider2_Scroll()
    If ICImagingControl2.LiveVideoRunning Then
        VCDProp2.RangeValue(VCDID_Brightness) = Slider2.Value
    End If
End Sub

Public Function DrawOverlay%()
Dim ob1 As OverlayBitmap
Dim ob2 As OverlayBitmap
Dim Font As New StdFont
Set ob1 = frmVideo.ICImagingControl1.OverlayBitmap
Set ob2 = frmVideo.ICImagingControl2.OverlayBitmap
' Enable the overlay bitmap for drawing.
ob1.Enable = True
ob2.Enable = True
' Set magenta as dropout color.
ob1.DropOutColor = RGB(255, 0, 255)
ob2.DropOutColor = RGB(255, 0, 255)
' Fill the overlay bitmap with the dropout color.
ob1.Fill ob1.DropOutColor
ob2.Fill ob2.DropOutColor
Font.Name = "Arial"
Font.Size = 14
Font.Bold = False
' Set the created font.
ob1.Font = Font
ob1.FontTransparent = True
ob1.FontBackColor = RGB(255, 255, 255)
ob2.Font = Font
ob2.FontTransparent = True
ob2.FontBackColor = RGB(255, 255, 255)

If Info1 = 1 Then
    ob1.DrawText RGB(255, 0, 0), 0, 400, ViewForm.JobClient + " " + ViewForm.JobLocation
    ob1.DrawText RGB(255, 0, 0), 0, 420, ViewForm.JobDescription
End If
If Info2 = 1 Then
    ob2.DrawText RGB(255, 0, 0), 0, 400, ViewForm.JobClient + " " + ViewForm.JobLocation
    ob2.DrawText RGB(255, 0, 0), 0, 420, ViewForm.JobDescription
End If

' Print Date/Time with red color.
If Date1 = 1 Then
    ob1.DrawText RGB(255, 0, 0), ICImagingControl1.LiveDisplayWidth - 90, 0, Str(DateValue(Now)) + " " + Format(TimeValue(Now), "HH:mm:ss")
End If
If Date2 = 1 Then
    ob2.DrawText RGB(255, 0, 0), ICImagingControl2.LiveDisplayWidth - 90, 0, Str(DateValue(Now)) + " " + Format(TimeValue(Now), "HH:mm:ss")
End If

If Attitude1 = 1 Then
'    ob1.FontTransparent = True
    ob1.DrawText RGB(255, 0, 0), 0, 0, "Roll: " + Format(ViewForm.SysRoll, "000.00") + " Pitch: " + Format(ViewForm.SysPitch, "000.00")
    ob1.DrawText RGB(255, 0, 0), 0, 25, "Depth: " + Format(ViewForm.SysDepth, "000.00") + " Pressure: " + Format(ViewForm.SysPressure, "000.00")
    ob1.DrawText RGB(255, 0, 0), ICImagingControl1.LiveDisplayWidth - 90, 400, "Encoder1: " + Format(ViewForm.AdjustedEncoder1, "000.00")
    ob1.DrawText RGB(255, 0, 0), ICImagingControl1.LiveDisplayWidth - 90, 420, "Encoder2: " + Format(ViewForm.AdjustedEncoder2, "000.00")
End If
If Attitude2 = 1 Then
'    ob2.FontTransparent = True
    ob2.DrawText RGB(255, 0, 0), 0, 0, "Roll: " + Format(ViewForm.SysRoll, "000.00") + " Pitch: " + Format(ViewForm.SysPitch, "000.00")
    ob2.DrawText RGB(255, 0, 0), 0, 25, "Depth: " + Format(ViewForm.SysDepth, "000.00") + " Pressure: " + Format(ViewForm.SysPressure, "000.00")
    ob2.DrawText RGB(255, 0, 0), ICImagingControl2.LiveDisplayWidth - 90, 400, "Encoder1: " + Format(ViewForm.AdjustedEncoder1, "000.00")
    ob2.DrawText RGB(255, 0, 0), ICImagingControl2.LiveDisplayWidth - 90, 420, "Encoder2: " + Format(ViewForm.AdjustedEncoder2, "000.00")
End If

If Logo1 = 1 Then
    ShowBitmap1 frmVideo.ICImagingControl1.OverlayBitmap
End If
If Logo2 = 1 Then
    ShowBitmap2 frmVideo.ICImagingControl2.OverlayBitmap
End If

If Circle1 = 1 Then
Dim points(0 To 3) As POINTAPI ' the points to draw to/from
Dim polypoints(0 To 4) As POINTAPI
    
    MidX% = 100 ' / Center of Plot Y
    MidY% = 350 ' / Center of Plot X
    offset% = 25
    ' draw target rings
    For Counter = 100 To 0 Step -20
        DrawCircle frmVideo.ICImagingControl1.OverlayBitmap, CLng(MidX), CLng(MidY), CLng(Counter), 1, RGB(0, 0, 0)
    Next Counter
    ' draw vertical cross line
    DrawLine frmVideo.ICImagingControl1.OverlayBitmap, CDbl(MidX), CDbl(MidY - 100), CDbl(MidX), CDbl(MidY + 100), 1, RGB(0, 0, 0)
    ' draw vertical cross line
    DrawLine frmVideo.ICImagingControl1.OverlayBitmap, CDbl(MidX - 100), CDbl(MidY), CDbl(MidX + 100), CDbl(MidY), 1, RGB(0, 0, 0)
    ' draw encoder marker
    Call Traverse(CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl((360 - ViewForm.AdjustedEncoder1 + 2)), CDbl(1))
    polypoints(0).x = TraverseX: polypoints(0).Y = TraverseY
    Call Traverse(CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl((360 - ViewForm.AdjustedEncoder1 + 2)), CDbl(100))
    polypoints(1).x = TraverseX: polypoints(1).Y = TraverseY
    Call Traverse(CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl((360 - ViewForm.AdjustedEncoder1 - 2)), CDbl(100))
    polypoints(2).x = TraverseX: polypoints(2).Y = TraverseY
    Call Traverse(CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl((360 - ViewForm.AdjustedEncoder1 - 2)), CDbl(1))
    polypoints(3).x = TraverseX: polypoints(3).Y = TraverseY
    polypoints(4).x = polypoints(0).x: polypoints(4).Y = polypoints(0).Y
    DrawPolygon frmVideo.ICImagingControl1.OverlayBitmap, polypoints(0), 5, 1, RGB(0, 255, 0)
    ' draw heading object
    DrawCircle frmVideo.ICImagingControl1.OverlayBitmap, CDbl(MidX + SysRoll), CDbl(MidY - SysPitch), CDbl(offset), 2, RGB(255, 0, 0)
    'Call Traverse(CDbl(MidX + SysRoll), CDbl(MidY - SysPitch), CDbl((360 - SysHeading) - 180), CDbl(offset * 2))
    'DrawLine ICImagingControl1.OverlayBitmap, CDbl(MidX + SysRoll), CDbl(MidY - SysPitch), CDbl(TraverseX), CDbl(TraverseY), 2, RGB(255, 0, 0)
    ' draw north pointer
    'points(0).X = TraverseX: points(0).Y = TraverseY
    'Call Traverse(CDbl(points(0).X), CDbl(points(0).Y), CDbl(360 - SysHeading - 10), 15)
    'points(1).X = TraverseX: points(1).Y = TraverseY
    'Call Traverse(CDbl(points(0).X), CDbl(points(0).Y), CDbl(360 - SysHeading + 10), 15)
    'points(2).X = TraverseX: points(2).Y = TraverseY
    'points(3).X = points(0).X: points(3).Y = points(0).Y
    'DrawTriangle ICImagingControl1.OverlayBitmap, points(0), 4, CLng(3), RGB(255, 0, 0)
    ' draw encoder line
    Call Traverse(CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl((360 - ViewForm.SysEncoder1) + ((360 - ViewForm.SysHeading) - 180)), CDbl(offset * 2))
    DrawLine frmVideo.ICImagingControl1.OverlayBitmap, CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl(TraverseX), CDbl(TraverseY), 2, RGB(0, 255, 0)
End If
If Circle2 = 1 Then
Dim points2(0 To 3) As POINTAPI ' the points to draw to/from
Dim polypoints2(0 To 4) As POINTAPI
    
    MidX% = 100 ' / Center of Plot Y
    MidY% = 350 ' / Center of Plot X
    offset% = 25
    ' draw target rings
    For Counter = 100 To 0 Step -20
        DrawCircle frmVideo.ICImagingControl2.OverlayBitmap, CLng(MidX), CLng(MidY), CLng(Counter), 1, RGB(0, 0, 0)
    Next Counter
    ' draw vertical cross line
    DrawLine frmVideo.ICImagingControl2.OverlayBitmap, CDbl(MidX), CDbl(MidY - 100), CDbl(MidX), CDbl(MidY + 100), 1, RGB(0, 0, 0)
    ' draw vertical cross line
    DrawLine frmVideo.ICImagingControl2.OverlayBitmap, CDbl(MidX - 100), CDbl(MidY), CDbl(MidX + 100), CDbl(MidY), 1, RGB(0, 0, 0)
    ' draw encoder marker
    Call Traverse(CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl((360 - ViewForm.AdjustedEncoder1 + 2)), CDbl(1))
    polypoints2(0).x = TraverseX: polypoints2(0).Y = TraverseY
    Call Traverse(CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl((360 - ViewForm.AdjustedEncoder1 + 2)), CDbl(100))
    polypoints2(1).x = TraverseX: polypoints2(1).Y = TraverseY
    Call Traverse(CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl((360 - ViewForm.AdjustedEncoder1 - 2)), CDbl(100))
    polypoints2(2).x = TraverseX: polypoints2(2).Y = TraverseY
    Call Traverse(CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl((360 - ViewForm.AdjustedEncoder1 - 2)), CDbl(1))
    polypoints2(3).x = TraverseX: polypoints2(3).Y = TraverseY
    polypoints2(4).x = polypoints2(0).x: polypoints2(4).Y = polypoints2(0).Y
    DrawPolygon frmVideo.ICImagingControl2.OverlayBitmap, polypoints2(0), 5, 1, RGB(0, 255, 0)
    ' draw heading object
    DrawCircle frmVideo.ICImagingControl2.OverlayBitmap, CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl(offset), 2, RGB(255, 0, 0)
    'Call Traverse(CDbl(MidX + SysRoll), CDbl(MidY - SysPitch), CDbl((360 - SysHeading) - 180), CDbl(offset * 2))
    'DrawLine ICImagingControl1.OverlayBitmap, CDbl(MidX + SysRoll), CDbl(MidY - SysPitch), CDbl(TraverseX), CDbl(TraverseY), 2, RGB(255, 0, 0)
    ' draw north pointer
    'points(0).X = TraverseX: points(0).Y = TraverseY
    'Call Traverse(CDbl(points(0).X), CDbl(points(0).Y), CDbl(360 - SysHeading - 10), 15)
    'points(1).X = TraverseX: points(1).Y = TraverseY
    'Call Traverse(CDbl(points(0).X), CDbl(points(0).Y), CDbl(360 - SysHeading + 10), 15)
    'points(2).X = TraverseX: points(2).Y = TraverseY
    'points(3).X = points(0).X: points(3).Y = points(0).Y
    'DrawTriangle ICImagingControl1.OverlayBitmap, points(0), 4, CLng(3), RGB(255, 0, 0)
    ' draw encoder line
    Call Traverse(CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl((360 - ViewForm.SysEncoder1) + ((360 - ViewForm.SysHeading) - 180)), CDbl(offset * 2))
    DrawLine frmVideo.ICImagingControl2.OverlayBitmap, CDbl(MidX + ViewForm.SysRoll), CDbl(MidY - ViewForm.SysPitch), CDbl(TraverseX), CDbl(TraverseY), 2, RGB(0, 255, 0)
End If
End Function


'
' ShowBitmap
'
' This sub demonstrates how to use OverlayBitmap.GetDC to blit a bitmap
' from a file on the live video.
' The bitmap will be blitted with transparency on the live video because
' it's background color is magenta (load the image "Hardware.BMP"
' with "Paint.exe" to verify this). Magenta is the currently set
' dropout color. The used GDI graphic functions are based on pixel units.
' Therefore, the scaling functions of Visual Basic are called to get the
' size of the loaded bitmap in pixels.
'
Private Sub ShowBitmap1(ob As OverlayBitmap)
    Dim Pic As New StdPicture
    Dim PicWidth As Integer
    Dim PicHeight As Integer
    Dim obDC As Long
    Dim SourceDC As Long
    Dim Col As Integer
 
    ' Load a BMP file.
    Set Pic = LoadPicture("C:\Program Files\Microsoft Visual Studio\VB98\Work\AquaView\AquaTechLogoMagenta.bmp")
 
    ' Transform the size of the picture from himetric to pixel units.
    PicWidth = ScaleX(Pic.width, vbHimetric, vbPixels)
    PicHeight = ScaleX(Pic.height, vbHimetric, vbPixels)
 
    ' Calculate a column to display the bitmap in the
    ' upper right corner of Imaging Control.
    Col = 640 - 5 - PicWidth
 
    ' Get the DC of the OverlayBitmap object.
    obDC = ob.GetDC
    If obDC <> 0 Then
        ' Create a compatible DC that is used to hold the loaded bitmap.
        SourceDC = CreateCompatibleDC(obDC)
        If SourceDC <> 0 Then
            ' Select the loaded bitmap in the source DC.
            SelectObject SourceDC, Pic.Handle
 
            ' Now blit the source DC in the overlay bitmap DC.
            ' This copies the loaded bitmap to the overlay bitmap.
            BitBlt obDC, Col, 5, PicWidth, PicHeight, SourceDC, 0, 0, SRCCOPY  ' 13369376 is SRCCOPY
 
            ' Delete the DC source that is no longer needed to avoid handle leaks.
            DeleteDC SourceDC
        End If
 
        ' Release the DC.
        ob.ReleaseDC obDC
    End If
End Sub
Private Sub ShowBitmap2(ob As OverlayBitmap)
    Dim Pic As New StdPicture
    Dim PicWidth As Integer
    Dim PicHeight As Integer
    Dim obDC As Long
    Dim SourceDC As Long
    Dim Col As Integer
 
    ' Load a BMP file.
    Set Pic = LoadPicture("C:\Program Files\Microsoft Visual Studio\VB98\Work\AquaView\AquaTechLogoMagenta.bmp")
 
    ' Transform the size of the picture from himetric to pixel units.
    PicWidth = ScaleX(Pic.width, vbHimetric, vbPixels)
    PicHeight = ScaleX(Pic.height, vbHimetric, vbPixels)
 
    ' Calculate a column to display the bitmap in the
    ' upper right corner of Imaging Control.
    Col = 640 - 5 - PicWidth
 
    ' Get the DC of the OverlayBitmap object.
    obDC = ob.GetDC
    If obDC <> 0 Then
        ' Create a compatible DC that is used to hold the loaded bitmap.
        SourceDC = CreateCompatibleDC(obDC)
        If SourceDC <> 0 Then
            ' Select the loaded bitmap in the source DC.
            SelectObject SourceDC, Pic.Handle
 
            ' Now blit the source DC in the overlay bitmap DC.
            ' This copies the loaded bitmap to the overlay bitmap.
            BitBlt obDC, Col, 5, PicWidth, PicHeight, SourceDC, 0, 0, SRCCOPY  ' 13369376 is SRCCOPY
 
            ' Delete the DC source that is no longer needed to avoid handle leaks.
            DeleteDC SourceDC
        End If
 
        ' Release the DC.
        ob.ReleaseDC obDC
    End If
End Sub

Private Sub Slider3_Click()
    If ICImagingControl1.LiveVideoRunning Then
        VCDProp1.RangeValue(VCDID_Contrast) = Slider3.Value
    End If
End Sub

Private Sub Slider4_Click()
    If ICImagingControl2.LiveVideoRunning Then
        VCDProp2.RangeValue(VCDID_Contrast) = Slider4.Value
    End If
End Sub

Private Sub Slider7_Click()
    If ICImagingControl2.LiveVideoRunning Then
        VCDProp2.RangeValue(VCDID_Focus) = Slider7.Value
    End If
End Sub

Private Sub Slider5_Click()

End Sub

Private Sub SliderLight1_Click()
    Dim AdjustedLight As Double
    
    AdjustedLight = (SliderLight1.Value * -1) + 10000
    If ViewForm.Light2Comm.PortOpen = True Then
        ' set for 5v LED
        ViewForm.Light2Comm.Output = "$" & "1" & "AO+" & Format((AdjustedLight), "00000.00") & Chr$(13)
        Call sm_wait(100)
    End If
    SliderLight1.ToolTipText = Format(AdjustedLight, "00000")
End Sub

Private Sub SliderLight2_Click()
    Dim AdjustedLight As Double
    
    AdjustedLight = (SliderLight2.Value * -1) + 10000
    If ViewForm.Light1Comm.PortOpen = True Then
        ' set for 5v LED
        ViewForm.Light1Comm.Output = "$" & "1" & "AO+" & Format((AdjustedLight), "00000.00") & Chr$(13)
        Call sm_wait(100)
    End If
    SliderLight2.ToolTipText = Format(AdjustedLight, "00000")
End Sub

Private Sub Timer1_Timer()

Dim diffTime As Double
Dim nowTime As Double

If Check1.Value = 1 Then
    nowTime = (CDbl(Format(Now, "dd")) * 3600) + (CDbl(Format(Now, "Hh")) * 60) + CDbl(Format(Now, "Nn"))
    diffTime = nowTime - file1Time

'Debug.Print Format(Now, "Nn")
'Debug.Print Format(Now, "yyyymmddHhNnSs")
'Debug.Print CStr(nowTime) & " " & CStr(fileTime)
    
    If diffTime > 50 Then ' diffTime is in minutes
        ICImagingControl1.AviStopCapture
        ICImagingControl1.LiveDisplayDefault = False
        ICImagingControl1.LiveDisplayHeight = ICImagingControl1.height / 15 ' 15 converts twips to pixels
        ICImagingControl1.LiveDisplayWidth = ICImagingControl1.width / 15
        Call sm_wait(200)
        ICImagingControl1.LiveStart
        current1File = Format(Now, "yyyymmddHHmmss")
        file1Time = (CDbl(Format(Now, "dd")) * 3600) + (CDbl(Format(Now, "Hh")) * 60) + CDbl(Format(Now, "Nn"))
        ICImagingControl1.AviStartCapture ViewForm.LoggingDir & "\" & current1File & "V1.avi", Avi1Comp.Name
    End If
End If
End Sub

Private Sub Timer2_Timer()

Dim diffTime As Double
Dim nowTime As Double

If Check2.Value = 1 Then
    nowTime = (CDbl(Format(Now, "dd")) * 3600) + (CDbl(Format(Now, "Hh")) * 60) + CDbl(Format(Now, "Nn"))
    diffTime = nowTime - file2Time

'Debug.Print Format(Now, "Nn")
'Debug.Print Format(Now, "yyyymmddHhNnSs")
'Debug.Print CStr(nowTime) & " " & CStr(fileTime)
    
    If diffTime > 50 Then ' diffTime is in minutes
        ICImagingControl2.AviStopCapture
        ICImagingControl2.LiveDisplayDefault = False
        ICImagingControl2.LiveDisplayHeight = ICImagingControl1.height / 15 ' 15 converts twips to pixels
        ICImagingControl2.LiveDisplayWidth = ICImagingControl1.width / 15
        Call sm_wait(200)
        ICImagingControl2.LiveStart
        current2File = Format(Now, "yyyymmddHHmmss")
        file2Time = (CDbl(Format(Now, "dd")) * 3600) + (CDbl(Format(Now, "Hh")) * 60) + CDbl(Format(Now, "Nn"))
        ICImagingControl2.AviStartCapture ViewForm.LoggingDir & "\" & current2File & "V2.avi", Avi2Comp.Name
    End If
End If

End Sub

Private Sub VideoInit_Click()
If ICImagingControl1.DeviceValid And ICImagingControl2.DeviceValid Then
    Call initVideo
Else
    Call MsgBox("Video Devices not detected!", vbOKOnly, "Device Error")
End If
End Sub
