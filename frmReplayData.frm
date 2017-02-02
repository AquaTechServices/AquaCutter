VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmReplayData 
   Caption         =   "Replay Project Data"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4920
   Icon            =   "frmReplayData.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Slider sldDataPosition 
      Height          =   375
      Left            =   0
      TabIndex        =   52
      Top             =   2040
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   327682
   End
   Begin VB.OptionButton optPlaybackSpeed 
      Caption         =   "200X"
      Height          =   255
      Index           =   7
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   495
   End
   Begin VB.OptionButton optPlaybackSpeed 
      Caption         =   "100X"
      Height          =   255
      Index           =   6
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   495
   End
   Begin VB.OptionButton optPlaybackSpeed 
      Caption         =   "50 X"
      Height          =   255
      Index           =   5
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.OptionButton optPlaybackSpeed 
      Caption         =   "20 X"
      Height          =   255
      Index           =   4
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.OptionButton optPlaybackSpeed 
      Caption         =   "10 X"
      Height          =   255
      Index           =   3
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.OptionButton optPlaybackSpeed 
      Caption         =   "5 X"
      Height          =   255
      Index           =   2
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.OptionButton optPlaybackSpeed 
      Caption         =   "2 X"
      Height          =   255
      Index           =   1
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
   Begin VB.OptionButton optPlaybackSpeed 
      Caption         =   "1 X"
      Height          =   255
      Index           =   0
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Value           =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Timer tmrReplay 
      Left            =   3960
      Top             =   720
   End
   Begin VB.CommandButton cmdOpenLogFile 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "log"
      DialogTitle     =   "Open Project log file"
      Filter          =   "log files|*.log"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Playback Speed"
      Height          =   255
      Left            =   1320
      TabIndex        =   51
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblSysAirPressure 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   31
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lblSysMediaWeight 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3720
      TabIndex        =   30
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lblSysWaterPressure 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lblSysNitrogenPressure 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   28
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lblSysCutterForward 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   27
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblSysCutterReverse 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3720
      TabIndex        =   26
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblEncoderOffset1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblEncoderOffset2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   24
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblAdjustedEncoder1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblAdjustedEncoder2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblSysEncoder1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblSysEncoder2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblSysFlow 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblSysSpeed 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblTime 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblRecordCount 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblSysPitch 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblSysRoll 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblSysPressure 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblLogFilePath 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Log file name"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      Caption         =   "Media Weight"
      Height          =   255
      Left            =   3720
      TabIndex        =   50
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Caption         =   "Air Pressure"
      Height          =   255
      Left            =   2520
      TabIndex        =   49
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Caption         =   "Nitrogen Pres."
      Height          =   255
      Left            =   1320
      TabIndex        =   48
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Caption         =   "Water Pressure"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Caption         =   "Cutter Rev."
      Height          =   255
      Left            =   3720
      TabIndex        =   46
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "Cutter Forw."
      Height          =   255
      Left            =   2520
      TabIndex        =   45
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Caption         =   "Enc. Offset 2"
      Height          =   255
      Left            =   1320
      TabIndex        =   44
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Caption         =   "Enc. Offset 1"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Caption         =   "Adj. Encoder 2"
      Height          =   255
      Left            =   3720
      TabIndex        =   42
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "Adj. Encoder 1"
      Height          =   255
      Left            =   2520
      TabIndex        =   41
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "Encoder 2"
      Height          =   255
      Left            =   1320
      TabIndex        =   40
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "Encoder 1"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "Speed"
      Height          =   255
      Left            =   3720
      TabIndex        =   38
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "Flow"
      Height          =   255
      Left            =   2520
      TabIndex        =   37
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "Pitch"
      Height          =   255
      Left            =   1320
      TabIndex        =   36
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "Roll"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "Pressure"
      Height          =   255
      Left            =   3720
      TabIndex        =   34
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Date and Time"
      Height          =   255
      Left            =   1320
      TabIndex        =   33
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Record Index"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmReplayData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gTextRecord() As String
Dim gPrevUnixTime As Long
Dim gRecordIndex As Long
Dim gTimerInterval As Integer

Private Sub cmdOpenLogFile_Click()
    Dim logRecord As String
    Dim paramName As String
    Dim paramValue As String
    Dim recordIndex As Long
    
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        ' Make sure array is empty
        ReDim gTextRecord(0)
        lblLogFilePath.Caption = CommonDialog1.FileName
        Open CommonDialog1.FileName For Input As #2
        ' Note: it takes about 6 seconds to load 12 hours of data into the array
        Do Until EOF(2)
            ReDim Preserve gTextRecord(recordIndex + 1)
            Line Input #2, gTextRecord(recordIndex)
            recordIndex = recordIndex + 1
        Loop
        Close #2
        
        ' Make sure timer is stopped
        cmdStop_Click
        
        ' Initialize the slider
        sldDataPosition.Enabled = True
        sldDataPosition.Max = recordIndex
        sldDataPosition.Value = 1
        sldDataPosition.LargeChange = 1710
    End If
    
End Sub

Private Sub cmdPlay_Click()
    Dim unixTime As Long
    Dim params() As Variant
    Dim paramCount As Integer
    Dim atLastRecord As Boolean
    
    On Error Resume Next
    ' Make sure gTextRecord array has data and there has been a log file selected
    If UBound(gTextRecord) > 0 And lblLogFilePath.Caption <> "" Then
        cmdPlay.Enabled = False
        cmdStop.Enabled = True
        If frmVidReplay.WMPlayer1.URL <> "" Then
            frmVidReplay.WMPlayer1.Controls.play
        End If
        If frmVidReplay.WMPlayer2.URL <> "" Then
            frmVidReplay.WMPlayer2.Controls.play
        End If
        ' Enable the timer
        tmrReplay.Interval = gTimerInterval
    Else
        MsgBox "No data to play back.  Please browse for a log file."
    End If
    ViewForm.Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    gTimerInterval = 1000
    frmVidReplay.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
ViewForm.Timer1.Enabled = True
ViewForm.EncoderTimer.Enabled = True
End Sub

Private Sub optPlaybackSpeed_Click(Index As Integer)
    Select Case Index
        Case 0
            gTimerInterval = 1000
        Case 1
            gTimerInterval = 500
        Case 2
            gTimerInterval = 200
        Case 3
            gTimerInterval = 100
        Case 4
            gTimerInterval = 50
        Case 5
            gTimerInterval = 20
        Case 6
            gTimerInterval = 10
        Case 7
            gTimerInterval = 5
        Case Else
            gTimerInterval = 1000
    End Select
    
    If tmrReplay.Interval > 0 Then
        tmrReplay.Interval = gTimerInterval
        frmVidReplay.WMPlayer1.Settings.Rate = 1000 / gTimerInterval
        frmVidReplay.WMPlayer2.Settings.Rate = 1000 / gTimerInterval
    End If
    
End Sub

Private Sub sldDataPosition_Scroll()
    gRecordIndex = sldDataPosition.Value
End Sub

Private Sub tmrReplay_Timer()
    ' Set the slider position to the global record index which will trigger the sldDataPosition_Change event
    sldDataPosition.Value = gRecordIndex + 1
    If sldDataPosition.Value = sldDataPosition.Max Then
        cmdStop_Click
    End If
End Sub

Private Sub cmdStop_Click()
        tmrReplay.Interval = 0
        cmdPlay.Enabled = True
        cmdStop.Enabled = False
        If frmVidReplay.WMPlayer1.URL <> "" Then
            frmVidReplay.WMPlayer1.Controls.Pause
        End If
        If frmVidReplay.WMPlayer2.URL <> "" Then
            frmVidReplay.WMPlayer2.Controls.Pause
        End If
End Sub

Private Function GetNextGoodRecord(params() As Variant) As Boolean
    Dim paramCount As Long
    Dim blnError As Boolean
    Dim unixTime As String
    Dim paramName As String
    Dim paramValue As String
    
    paramCount = sm_parse(gTextRecord(gRecordIndex), ",", params())
    If paramCount = 3 Then
        On Error Resume Next
        unixTime = CLng(params(1))
        If Err.Number > 0 Then
            blnError = True
            Err.Clear
        End If
        paramName = params(2)
        If Err.Number > 0 Then
            blnError = True
            Err.Clear
        End If
        paramValue = params(3)
        If Err.Number > 0 Then
            blnError = True
            Err.Clear
        End If
        On Error GoTo 0
    Else
        blnError = True
    End If
    
    If blnError = True Then
        GetNextGoodRecord = False
    Else
        GetNextGoodRecord = True
    End If
    
End Function

Private Sub sldDataPosition_Change()
    Dim params() As Variant
    Dim blnGoodRecord As Boolean
    Dim unixTime As Long
    
    ' Set the global record index to the index of the slider
    gRecordIndex = sldDataPosition.Value
    
    ' Read the record at gRecordIndex
    While GetNextGoodRecord(params) = False
        If gRecordIndex < sldDataPosition.Max Then
            gRecordIndex = gRecordIndex + 1
        Else
            Exit Sub
        End If
    Wend
    ' Set the gPrevUnixTime
    gPrevUnixTime = CLng(params(1))
    ' Set the current unix time to previous unix time so the while wend loop will run the first time
    unixTime = gPrevUnixTime
    ' Set the parameter collected from this record
    SetValues params
    DoEvents
    ' Read all the records with the same unixtime as the first one as long as the global index is less
    ' than the slider max value
    ' As soon as the time changes, the sub will exit
    While unixTime = gPrevUnixTime
        ' Get the next good record (keep incrementing the record index until a good record is returned
        While GetNextGoodRecord(params) = False
            If gRecordIndex < sldDataPosition.Max Then
                gRecordIndex = gRecordIndex + 1
            Else
                Exit Sub
            End If
        Wend
        
        ' Set the parameter collected from this record
        SetValues params
        ' Set the current unix time for use in the next while wend iteration
        unixTime = CLng(params(1))
        ' Increment the record index for the next while wend iteration
        If gRecordIndex < sldDataPosition.Max Then
            gRecordIndex = gRecordIndex + 1
        Else
            Exit Sub
        End If
    Wend
End Sub

Private Sub SetValues(params() As Variant)
    Dim unixTime As Long
    Dim paramCount As Integer
    Dim paramName As String
    Dim paramValue As String
    Dim theTime As Date
    
    On Error Resume Next
    
    ' Extract the parameters from the log record
    unixTime = CLng(params(1))
    paramName = params(2)
    paramValue = params(3)
    
    ' Set the human readable date and time from the unix time
    theTime = getUnixTime(unixTime)
    
    ' Set the record count and time in the dialog
    lblRecordCount.Caption = gRecordIndex + 1
    lblTime.Caption = theTime
    
    ' Update the dialog's labels and the global Aquacutter parameters
    Select Case paramName
        Case "SysPressure"
            ViewForm.SysPressure = CDbl(paramValue)
            lblSysPressure.Caption = paramValue
        Case "SysRoll"
            ViewForm.SysRoll = CDbl(paramValue)
            lblSysRoll.Caption = paramValue
        Case "SysPitch"
            ViewForm.SysPitch = CDbl(paramValue)
            lblSysPitch.Caption = paramValue
        Case "SysFlow"
            ViewForm.SysFlow = CDbl(paramValue)
            lblSysFlow.Caption = paramValue
        Case "SysSpeed"
            ViewForm.SysSpeed = CDbl(paramValue)
            lblSysSpeed.Caption = paramValue
            ViewForm.SM_Speed11.Value = CDbl(paramValue)
        Case "SysEncoder1"
            ViewForm.SysEncoder1 = CDbl(paramValue)
            lblSysEncoder1.Caption = paramValue
        Case "SysEncoder2"
            ViewForm.SysEncoder2 = CDbl(paramValue)
            lblSysEncoder2.Caption = paramValue
        Case "AdjustedEncoder1"
            ViewForm.AdjustedEncoder1 = CDbl(paramValue)
            lblAdjustedEncoder1.Caption = paramValue
            ViewForm.SM_Encoder11.Value1 = CDbl(paramValue)
        Case "AdjustedEncoder2"
            ViewForm.AdjustedEncoder2 = CDbl(paramValue)
            lblAdjustedEncoder2.Caption = paramValue
            ViewForm.SM_Encoder11.Value2 = CDbl(paramValue)
        Case "EncoderOffset1"
            ViewForm.EncoderOffset1 = CDbl(paramValue)
            lblEncoderOffset1.Caption = paramValue
        Case "EncoderOffset2"
            ViewForm.EncoderOffset2 = CDbl(paramValue)
            lblEncoderOffset2.Caption = paramValue
        Case "SysCutterForward"
            ViewForm.SysCutterForward = CDbl(paramValue)
            lblSysCutterForward.Caption = paramValue
        Case "SysCutterReverse"
            ViewForm.SysCutterReverse = CDbl(paramValue)
            lblSysCutterReverse.Caption = paramValue
        Case "SysWaterPressure"
            ViewForm.SysWaterPressure = CDbl(paramValue)
            lblSysWaterPressure.Caption = paramValue
        Case "SysNitrogenPressure"
            ViewForm.SysNitrogenPressure = CDbl(paramValue)
            lblSysNitrogenPressure.Caption = paramValue
        Case "SysAirPressure"
            ViewForm.SysAirPressure = CDbl(paramValue)
            lblSysAirPressure.Caption = paramValue
        Case "SysMediaWeight"
            ViewForm.SysMediaWeight = CDbl(paramValue)
            lblSysMediaWeight.Caption = paramValue
        Case "Video1Start"
            If CheckPath(paramValue) Then
                frmVidReplay.WMPlayer1.URL = paramValue
                frmVidReplay.WMPlayer1.Settings.Rate = 1000 / gTimerInterval
                frmVidReplay.WMPlayer1.Controls.play
            End If
        Case "Video2Start"
            If CheckPath(paramValue) Then
                frmVidReplay.WMPlayer2.URL = paramValue
                frmVidReplay.WMPlayer2.Settings.Rate = 1000 / gTimerInterval
                frmVidReplay.WMPlayer2.Controls.play
            End If
        Case Else
    End Select

End Sub
