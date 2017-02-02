VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{3A4A6317-A2E1-406F-A101-673EDA0B4016}#147.0#0"; "SM_Speed.ocx"
Object = "{A8CC1A9C-3022-4F2F-B71C-5CB4EE584698}#55.0#0"; "SM_Encoder.ocx"
Object = "{AE6EB25D-9E61-4F96-83C6-D51B582C8296}#7.0#0"; "SM_StripChart.ocx"
Object = "{A4168463-2E31-4A21-BCD1-C6653DD5363E}#1.0#0"; "SM_PLOT.ocx"
Begin VB.Form ViewForm 
   Caption         =   "AquaCutter"
   ClientHeight    =   7965
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11085
   Icon            =   "ViewForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin SM_Plot1.SM_Plot OrientPlot 
      Height          =   3615
      Left            =   7440
      TabIndex        =   51
      Top             =   360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6376
      DataColor       =   16777215
      GridBackColor   =   0
      GridColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      GridSpacing     =   0
      plotScale       =   0
      LabelText       =   ""
      GraphPrecision  =   0
   End
   Begin VB.Timer ReplayTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   1440
   End
   Begin MSCommLib.MSComm ScaleComm 
      Left            =   2280
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   12
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin MSCommLib.MSComm Analog2Comm 
      Left            =   1680
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   11
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   300
   End
   Begin VB.Timer AnalogTimer 
      Interval        =   2000
      Left            =   0
      Top             =   1080
   End
   Begin MSCommLib.MSComm Analog1Comm 
      Left            =   480
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   6
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   300
   End
   Begin MSCommLib.MSComm StabilComm 
      Left            =   1080
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   10
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Timer EncoderTimer 
      Interval        =   1000
      Left            =   0
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SET"
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
      Left            =   6720
      TabIndex        =   5
      ToolTipText     =   "Set Water Pressure Baseline Prior to Cut"
      Top             =   360
      Width           =   615
   End
   Begin MSCommLib.MSComm Light2Comm 
      Left            =   2760
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   9
      DTREnable       =   0   'False
      RThreshold      =   1
   End
   Begin SM_StripChart.SM_StripChart1 PressureChart 
      Height          =   3255
      Left            =   3720
      TabIndex        =   3
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5741
      Value1Color     =   255
      Value2Color     =   65280
      GridBackColor   =   0
      GridColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      AutoRedraw      =   0   'False
      Step            =   2
      Value1Scale     =   2
      DataWrap        =   0
      DataWidth       =   2
      GridMove        =   0   'False
      GridSpacing     =   25
   End
   Begin SM_Encoder.SM_Encoder1 SM_Encoder11 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   4320
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6376
      DataColor       =   -2147483640
      MeterBackColor  =   0
      MeterColor      =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
   End
   Begin MSCommLib.MSComm PotComm 
      Left            =   2160
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   8
      DTREnable       =   0   'False
      RThreshold      =   1
      BaudRate        =   38400
   End
   Begin VB.Frame frameControl 
      Caption         =   "Aqua Cutter Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   3720
      TabIndex        =   1
      Top             =   4080
      Width           =   7335
      Begin VB.Frame Frame6 
         Caption         =   "Inclination Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3600
         TabIndex        =   43
         Top             =   240
         Width           =   3615
         Begin VB.CommandButton ZeroMotion 
            Caption         =   "Zero"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            ToolTipText     =   "Set Motion Offsets Once Tool is Locked In Position"
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton ClearMotionOffset 
            Caption         =   "Clear"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            ToolTipText     =   "Clear Motion Offsets"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblRoll 
            Alignment       =   2  'Center
            Caption         =   "00.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   47
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblPitch 
            Alignment       =   2  'Center
            Caption         =   "00.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   46
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Roll:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   45
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "Pitch:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   44
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Encoder Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3600
         TabIndex        =   36
         Top             =   2400
         Width           =   3615
         Begin VB.CommandButton btnZeroEncoder 
            Caption         =   "Zero"
            Height          =   255
            Left            =   1560
            TabIndex        =   40
            ToolTipText     =   "Offset Encoder Display After Preloading Tool"
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton EncoderOption 
            Caption         =   "Encoder1"
            Height          =   255
            Index           =   0
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Select Encoder 1 as Primary"
            Top             =   960
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton EncoderOption 
            Caption         =   "Encoder2"
            Height          =   255
            Index           =   1
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Select Encoder 2 as Primary"
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton ClearEncoderOffset 
            Caption         =   "Clear"
            Height          =   255
            Left            =   1560
            TabIndex        =   37
            ToolTipText     =   "Clear Encoder Offset"
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label22 
            Caption         =   "Encoder Zero"
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
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Encoder Clear"
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
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Active Encoder"
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
            Left            =   120
            TabIndex        =   41
            Top             =   960
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Stabilizer Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         TabIndex        =   28
         Top             =   1440
         Width           =   3615
         Begin VB.CheckBox StabilOff 
            Caption         =   "IN"
            Height          =   495
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Toggle Stabilizer Arms Valve IN"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox StabilOn 
            Caption         =   "OUT"
            Height          =   495
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Toggle Stabilizer Arms Valve OUT"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Stabilizer Arms Out/In"
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
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.OptionButton LocalControl 
         Caption         =   "Manual"
         Height          =   495
         Index           =   1
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton LocalControl 
         Caption         =   "Auto"
         Height          =   495
         Index           =   0
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2535
         Begin VB.HScrollBar scrPot 
            Height          =   255
            LargeChange     =   50
            Left            =   120
            Max             =   5000
            SmallChange     =   10
            TabIndex        =   25
            Top             =   840
            Value           =   1
            Width           =   2295
         End
         Begin VB.Label lblFlowPercent 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0.00%"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   5
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblFlowProgress 
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   360
            TabIndex        =   23
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Manual Flow Control"
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
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   2295
         End
         Begin VB.Label Label2 
            Caption         =   "0%"
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
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "100%"
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
            Left            =   1920
            TabIndex        =   20
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblFlowBack 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   360
            TabIndex        =   24
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   3375
         Begin VB.OptionButton OptionDir 
            Caption         =   "Reverse"
            Height          =   375
            Index           =   1
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton OptionDir 
            Caption         =   "Forward"
            Height          =   375
            Index           =   0
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Pipe Diameter"
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
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Cutting Speed"
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
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Cut Direction"
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
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Speed Control"
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
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label PipeDiam 
            Caption         =   "8 in"
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
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.Label CuttingSpeed 
            Caption         =   "2.0 in/min"
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
            Left            =   1440
            TabIndex        =   12
            Top             =   480
            Width           =   975
         End
         Begin VB.Label CutDir 
            Caption         =   "Forward"
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
            Left            =   1440
            TabIndex        =   11
            Top             =   720
            Width           =   975
         End
         Begin VB.Label SpeedCont 
            Caption         =   "Auto"
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
            Left            =   1440
            TabIndex        =   10
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   3375
         Begin VB.OptionButton StartOption 
            BackColor       =   &H0000FF00&
            Caption         =   "Start"
            Height          =   495
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton StartOption 
            BackColor       =   &H000000FF&
            Caption         =   "Stop"
            Height          =   495
            Index           =   1
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
   End
   Begin MSCommLib.MSComm Light1Comm 
      Left            =   1560
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   7
      DTREnable       =   0   'False
      RThreshold      =   1
   End
   Begin MSCommLib.MSComm EncoderComm 
      Left            =   960
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   6
      DTREnable       =   0   'False
      RThreshold      =   1
   End
   Begin MSCommLib.MSComm OHPRComm 
      Left            =   360
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   19200
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   360
   End
   Begin SM_Speed.SM_Speed1 SM_Speed11 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6376
      DataColor       =   -2147483630
      MeterBackColor  =   0
      MeterColor      =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      AutoRedraw      =   0   'False
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cutter Inclination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   35
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Encoder Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   34
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Under Water Pressure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   33
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cutter Speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label lblPressure 
      Alignment       =   2  'Center
      Caption         =   "0000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   360
      Width           =   2895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuView 
         Caption         =   "View"
         Begin VB.Menu mnuGauge 
            Caption         =   "Gauges"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuVideo 
            Caption         =   "Video"
         End
      End
      Begin VB.Menu mnuPlayback 
         Caption         =   "Playback"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuReport 
         Caption         =   "Report"
         Enabled         =   0   'False
      End
      Begin VB.Menu nothing1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSysTools 
      Caption         =   "&System Tools"
      Begin VB.Menu mnuConfig 
         Caption         =   "&Configuration"
      End
      Begin VB.Menu mnuPortSetup 
         Caption         =   "&Port Setup"
         Enabled         =   0   'False
         Begin VB.Menu mnuMotion 
            Caption         =   "MotionSensor/Depth"
         End
         Begin VB.Menu mnuEncoder 
            Caption         =   "Encoder1/2"
         End
         Begin VB.Menu mnuPressure 
            Caption         =   "PressureSensor"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFlow 
            Caption         =   "FlowMeter"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPot 
            Caption         =   "FlowControl"
         End
         Begin VB.Menu mnuAnalog1 
            Caption         =   "Analog1"
         End
      End
   End
End
Attribute VB_Name = "ViewForm"
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

Dim runArgument As String
Dim Analog1Flag As Integer
Dim Analog2Flag As Integer
Dim ColorArray() As OLE_COLOR
Dim TestArray() As Double
Dim SpeedArray(10) As Double
Dim SpeedPtArray(10) As DoubleXY
Dim SpeedCount As Integer
Public SysSpeed As Double
Public SysHeading As Double
Public SysEncoder1 As Double
Public SysEncoder2 As Double
Dim Encoder1Time As Double
Dim Encoder2Time As Double
Public SysRoll As Double
Public SysPitch As Double
Public RollOffset As Double
Public PitchOffset As Double
Public SysFlow As Double
Public SysPot As Integer
Public SysPressure As Double
Public SysDepth As Double
Public SysStabilOut As Double
Public SysCutterForward As Double
Public SysCutterReverse As Double
Public SysWaterPressure As Double
Public SysNitrogenPressure As Double
Public SysGarnetPressure As Double
Public SysAirPressure As Double
Public SysMediaWeight As Double
Public SysMediaWeightOld As Double
Public SysMediaFlow As Double
Public EncoderOffset1 As Double
Public EncoderOffset2 As Double
Public Encoder1Bit As Integer
Public Encoder2Bit As Integer
Public AdjustedEncoder1 As Double
Public AdjustedEncoder2 As Double
Public ConfigChange As Boolean
Public PipeDiameter As Double
Public CutSpeed As Double
Public CutDirection As Integer
Public SpeedControl As Integer
Dim TargetSpeed As Integer
Public LoggingDir As String
Public LoggingFile As String
Public LoggingChange As Boolean
Public LoggingOn As Boolean
Public LogP As Boolean ' Pressure
Public LogI As Boolean ' Inclination
Public LogS As Boolean ' Speed
Public LogF As Boolean ' Flow
Public LogE1 As Boolean ' Encoder1
Public LogE2 As Boolean ' Encoder2
Public JobClient As String
Public JobVessel As String
Public JobLocation As String
Public JobDescription As String
Const Encoder1Bias = 0
Const Encoder2Bias = 0
Public SwapCheck As Integer
Public Video1Start As String
Public Video2Start As String
Private Declare Function GetWindowsVersion Lib "kernel32" Alias "GetVersion" () As Long
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Public lastKey As Integer
Public dupKey As Boolean
Const HKEY_CURRENT_USER As Long = &H80000001

Private Sub Analog1Comm_OnComm()
'Command: $1RB
'Response:  *+00072.00
'           *+00123.00
'           *+78900.00
'           *-00072.00
'Analog1Comm.Output = "$" & Chr$(1) & "RB" & Chr$(13)
'DGH 5251 Needs to be setup. Min = 0.0 Max = 25.0

Static outString As String
Dim InString As String
Dim TmpString As String
Dim GenericString As String
Dim Cnt As Integer
Dim Max As Integer
Dim XVal As Double
Dim YVal As Double
Dim GenericLength As Integer
Cnt = 1
On Error Resume Next
InString = Analog1Comm.Input
Max = Len(InString)
While Cnt <= Max
    TmpString = Mid$(InString, Cnt, 1)
    If TmpString = Chr$(13) Then
        GenericString = outString
        GenericLength = Len(GenericString)
        If GenericLength > 1 Then
            outlength = Len(GenericString)
            Analog1Flag = Analog1Flag + 1
            'Debug.Print "Analog Data " & CStr(Analog1Flag) & Trim(GenericString)
            If Analog1Flag = 1 Then
                SysFlow = (right(Trim(GenericString), Len(Trim(GenericString)) - 2) / 4) - 1
            ElseIf Analog1Flag = 2 Then
                SysStabilOut = (right(Trim(GenericString), Len(Trim(GenericString)) - 2) - 4) * 375 ' 0 - 6000 psi
            ElseIf Analog1Flag = 3 Then
                SysCutterForward = (right(Trim(GenericString), Len(Trim(GenericString)) - 2) - 3.98) * 375 ' 0 - 6000 psi
            ElseIf Analog1Flag = 4 Then
                SysCutterReverse = (right(Trim(GenericString), Len(Trim(GenericString)) - 2) - 4.01) * 375 ' 0 - 6000 psi
            End If
        End If
        TmpString = ""
        outString = ""
        GenericString = ""
        GoTo NextChar
    End If
    outString = outString + TmpString
NextChar:
    Cnt = Cnt + 1
Wend

End Sub

Private Sub Analog2Comm_OnComm()
'Command: $1RB
'Response:  *+00072.00
'           *+00123.00
'           *+78900.00
'           *-00072.00
'Analog1Comm.Output = "$" & Chr$(1) & "RB" & Chr$(13)
'DGH 5251 Needs to be setup. Min = 0.0 Max = 25.0

Static outString As String
Dim InString As String
Dim TmpString As String
Dim GenericString As String
Dim Cnt As Integer
Dim Max As Integer
Dim XVal As Double
Dim YVal As Double
Dim GenericLength As Integer
Cnt = 1
On Error Resume Next
InString = Analog2Comm.Input
Max = Len(InString)
While Cnt <= Max
    TmpString = Mid$(InString, Cnt, 1)
    If TmpString = Chr$(13) Then
        GenericString = outString
        GenericLength = Len(GenericString)
Debug.Print GenericString
        If GenericLength > 1 Then
            outlength = Len(GenericString)
            Analog2Flag = Analog2Flag + 1
            If Analog2Flag = 1 Then
                SysGarnetPressure = (right(Trim(GenericString), Len(Trim(GenericString)) - 2) - 4) * 31.25 ' 0 - 500 psi
            ElseIf Analog2Flag = 2 Then
                SysNitrogenPressure = (right(Trim(GenericString), Len(Trim(GenericString)) - 2) - 4) * 31.25 ' 0 - 500 psi
            ElseIf Analog2Flag = 3 Then
                SysAirPressure = (right(Trim(GenericString), Len(Trim(GenericString)) - 2) - 4) * 31.25 ' 0 - 500 psi
            ElseIf Analog2Flag = 4 Then
                SysWaterPressure = (right(Trim(GenericString), Len(Trim(GenericString)) - 2) - 4) * 1250 ' 0 - 20000 psi
            End If
        End If
        TmpString = ""
        outString = ""
        GenericString = ""
        GoTo NextChar
    End If
    outString = outString + TmpString
NextChar:
    Cnt = Cnt + 1
Wend
End Sub

Private Sub AnalogTimer_Timer()
If Analog1Comm.PortOpen = True Then
    Analog1Comm.Output = "$1RB" & Chr$(13)
    Analog1Flag = 0
End If
If Analog2Comm.PortOpen = True Then
    Analog2Comm.Output = "$1RB" & Chr$(13)
    Analog2Flag = 0
End If
End Sub

Private Sub btnZeroEncoder_Click()
EncoderOffset1 = SysEncoder1
EncoderOffset2 = SysEncoder2
End Sub

Private Sub ClearEncoderOffset_Click()
EncoderOffset1 = 0
EncoderOffset2 = 0
End Sub

Private Sub ClearMotionOffset_Click()
RollOffset = 0
PitchOffset = 0
End Sub

Private Sub Command1_Click()
PressureChart.Value2 = SysDepth / 2.24489794 ' converts depth back to psi
End Sub

Private Sub EncoderComm_OnComm()
Static outString As String
Dim InString As String
Dim TmpString As String
Dim GenericString As String
Dim Cnt As Integer
Dim Max As Integer
Dim XVal As Double
Dim YVal As Double
Dim GenericLength As Integer
Dim FileString As String

Cnt = 1
InString = EncoderComm.Input
Max = Len(InString)
On Error Resume Next
While Cnt <= Max
    TmpString = Mid$(InString, Cnt, 1)
    If TmpString = Chr$(13) Then
        GenericString = outString
        GenericLength = Len(GenericString)
        If GenericLength > 1 Then
            If Mid(GenericString, 4, 1) = 1 Then
                Debug.Print "Encoder 1 " & GenericString
                If SwapCheck = 0 Then
                    If Encoder1Bit = 0 Then
                        SysEncoder1 = 360 - (right(GenericString, 5) / 4095) * 360 ' 12 BIT Encoder
                    ElseIf Encoder1Bit = 1 Then
                        SysEncoder1 = 360 - (right(GenericString, 5) / 8191) * 360 ' 13 BIT Encoder
                    End If
                ElseIf SwapCheck = 1 Then
                    If Encoder1Bit = 0 Then
                        SysEncoder1 = (right(GenericString, 5) / 4095) * 360 ' 12 BIT Encoder
                    ElseIf Encoder1Bit = 1 Then
                        SysEncoder1 = (right(GenericString, 5) / 8191) * 360 ' 13 BIT Encoder
                    End If
                End If
                Encoder1Time = GetTickCount
                Debug.Print "Encoder1 = " & SysEncoder1
            ElseIf Mid(GenericString, 4, 1) = 2 Then
                Debug.Print "Encoder 2 " & GenericString
                If SwapCheck = 0 Then
                    If Encoder2Bit = 0 Then
                        SysEncoder2 = (right(GenericString, 5) / 4095) * 360 ' 12 BIT Encoder
                    ElseIf Encoder2Bit = 1 Then
                        SysEncoder2 = (right(GenericString, 5) / 8191) * 360 ' 13 BIT Encoder
                    End If
                ElseIf SwapCheck = 1 Then
                    If Encoder2Bit = 0 Then
                        SysEncoder2 = 360 - (right(GenericString, 5) / 4095) * 360 ' 12 BIT Encoder
                    ElseIf Encoder2Bit = 1 Then
                        SysEncoder2 = 360 - (right(GenericString, 5) / 8191) * 360 ' 13 BIT Encoder
                    End If
                End If
                Encoder2Time = GetTickCount
                Debug.Print "Encoder2 = " & SysEncoder2
            End If
        End If
        TmpString = ""
        outString = ""
        GenericString = ""
        GoTo NextChar
    End If
    outString = outString + TmpString
NextChar:
    Cnt = Cnt + 1
Wend
End Sub

Private Sub EncoderTimer_Timer()
Dim ArcDist As Double
Static OldArcLength As Double
Static ArcLength As Double
Static encoder As Double
Static LastTime As Double
Static ThisTime As Double
Dim diffTime As Double
Dim OutCommand As String
Static count As Integer
Dim n As Integer
Dim x As Integer
Dim SortArray(10) As Variant
Dim TestSpeed As Double
Dim TestSlope As Double
Dim TestIntercept As Double
Dim FileString As String
Dim TempSlope As Double
Dim TempIntercept As Double
Dim RegTime As Double

If EncoderComm.PortOpen = True Then
    OutCommand = "$0R1"
    EncoderComm.Output = OutCommand & Chr$(13)
    Call sm_wait(100)
    'Debug.Print OutCommand
End If
Call sm_wait(100)
If EncoderComm.PortOpen = True Then
    OutCommand = "$0R2"
    EncoderComm.Output = OutCommand & Chr$(13)
    Call sm_wait(100)
    'Debug.Print OutCommand
End If

If (SysEncoder1 - EncoderOffset1) < 0 Or (SysEncoder1 - EncoderOffset1) = 0 Then
    AdjustedEncoder1 = (SysEncoder1 - EncoderOffset1) + 360
Else
    AdjustedEncoder1 = SysEncoder1 - EncoderOffset1
End If

If AdjustedEncoder1 > 360 Then
    AdjustedEncoder1 = AdjustedEncoder1 - 360
End If
    
If (SysEncoder2 - EncoderOffset2) < 0 Or (SysEncoder2 - EncoderOffset2) = 0 Then
    AdjustedEncoder2 = (SysEncoder2 - EncoderOffset2) + 360
Else
    AdjustedEncoder2 = SysEncoder2 - EncoderOffset2
End If

If AdjustedEncoder2 > 360 Then
    AdjustedEncoder2 = AdjustedEncoder2 - 360
End If
    
'If AdjustedEncoder1 > 359.99 Then AdjustedEncoder1 = 0
'If AdjustedEncoder2 > 359.99 Then AdjustedEncoder2 = 0

'Debug.Print "ThisTime = " & CStr(ThisTime) & " LastTime = " & CStr(LastTime) & " DiffTime = " & CStr(ThisTime - LastTime)
If EncoderOption(0).Value = True Then
    SM_Encoder11.Value1 = AdjustedEncoder1
    SM_Encoder11.Value2 = AdjustedEncoder2
    '**************Speed and Distance Calcs****************************
    ' calc length of cut      (radius)
    ArcLength = (2 * PI * (PipeDiameter / 2)) * (AdjustedEncoder1 / 360)
    ' calc distance cut since last second
    ArcDist = Abs(ArcLength - OldArcLength) ' distance in half second

ElseIf EncoderOption(1).Value = True Then
    SM_Encoder11.Value1 = AdjustedEncoder2
    SM_Encoder11.Value2 = AdjustedEncoder1
    '**************Speed and Distance Calcs****************************
    ' calc length of cut      (radius)
    ArcLength = (2 * PI * (PipeDiameter / 2)) * (AdjustedEncoder2 / 360)
    ' calc distance cut since last second
    ArcDist = Abs(ArcLength - OldArcLength) ' distance in half second

End If

' calc speed based on cutter position

If EncoderOption(0).Value = True Then
    ThisTime = Encoder1Time
Else
    ThisTime = Encoder2Time
End If
diffTime = ThisTime - LastTime

If ArcDist > 0 And diffTime > 0 Then

    If count < 10 Then
        SpeedArray(count) = ((ArcDist / (diffTime / 1000)) * 60)
        count = count + 1
    Else
        n = 0
        TempSpeed = 0
        Do While n < 9
            SpeedArray(n) = SpeedArray(n + 1)
            n = n + 1
        Loop
        SpeedArray(9) = ((ArcDist / (diffTime / 1000)) * 60)
        n = 0
        Do While n < 10
            TempSpeed = SpeedArray(n) + TempSpeed
            n = n + 1
        Loop
        SysSpeed = TempSpeed / 10
    End If

SM_Speed11.Value = SysSpeed
'Debug.Print CStr(ArcLength)
OldArcLength = ArcLength
LastTime = ThisTime
End If
End Sub

Private Sub Form_Load()
Dim tempIndex As Integer
Dim TempLog As String
Dim HeaderLog As String
Dim TimeCheck As Boolean
Dim OutCommand As String


Call GetLocation(ViewForm)

On Error Resume Next

frmGauge.Show

runArgument = Trim(Command) ' get command line argument without spaces
If runArgument = "-s" Then
    mnuPlayback.Enabled = False
Else
    mnuPlayback.Enabled = True
End If

runArgument = Trim(Command)
If runArgument = "-nl" Then
    'no license check
Else
    TimeCheck = checkSoftwareTime
    If TimeCheck = False Then
        MsgBox "Your software License has expired." & vbCrLf & _
        "Please Contact Aqua Tech Services for an updated Liscense Code." & vbCrLf & _
        "Please provide the unit number for the computers that need Licensing." & vbCrLf & _
        "337-837-3999", vbOKOnly, "License Check"
        End
    End If
End If

lastKey = 0

Analog1Flag = 0
Analog2Flag = 0

OHPRComm.Settings = GetSetting(appname:="AquaView", section:="Startup", Key:="OHPRCommSettings", Default:="4800,N,8,1")
OHPRComm.CommPort = GetSetting(appname:="AquaView", section:="Startup", Key:="OHPRCommPort", Default:=5)
OHPRComm.CommPort = 5
OHPRComm.Settings = "4800,N,8,1"
Analog1Comm.Settings = GetSetting(appname:="AquaView", section:="Startup", Key:="Analog1CommSettings", Default:="300,N,8,1")
Analog1Comm.CommPort = GetSetting(appname:="AquaView", section:="Startup", Key:="Analog1CommPort", Default:=6)
Analog1Comm.CommPort = 6
Analog1Comm.Settings = "300,N,8,1"
Analog2Comm.Settings = GetSetting(appname:="AquaView", section:="Startup", Key:="Analog2CommSettings", Default:="300,N,8,1")
Analog2Comm.CommPort = GetSetting(appname:="AquaView", section:="Startup", Key:="Analog2CommPort", Default:=11)
Analog2Comm.CommPort = 11
Analog2Comm.Settings = "300,N,8,1"
Light1Comm.Settings = GetSetting(appname:="AquaView", section:="Startup", Key:="Light1CommSettings", Default:="300,N,8,1")
Light1Comm.CommPort = GetSetting(appname:="AquaView", section:="Startup", Key:="Light1CommPort", Default:=7)
Light1Comm.CommPort = 7
Light1Comm.Settings = "300,N,8,1"
PotComm.Settings = GetSetting(appname:="AquaView", section:="Startup", Key:="PotCommSettings", Default:="300,N,8,1")
PotComm.CommPort = GetSetting(appname:="AquaView", section:="Startup", Key:="PotCommPort", Default:=8)
PotComm.CommPort = 8
PotComm.Settings = "300,N,8,1"
Light2Comm.Settings = GetSetting(appname:="AquaView", section:="Startup", Key:="Light2CommSettings", Default:="300,N,8,1")
Light2Comm.CommPort = GetSetting(appname:="AquaView", section:="Startup", Key:="Light2CommPort", Default:=9)
Light2Comm.CommPort = 9
Light2Comm.Settings = "300,N,8,1"
StabilComm.Settings = GetSetting(appname:="AquaView", section:="Startup", Key:="StabilCommSettings", Default:="9600,N,8,1")
StabilComm.CommPort = GetSetting(appname:="AquaView", section:="Startup", Key:="StabilCommPort", Default:=10)
StabilComm.CommPort = 10
StabilComm.Settings = "9600,N,8,1"
EncoderComm.Settings = GetSetting(appname:="AquaView", section:="Startup", Key:="EncoderCommSettings", Default:="115200,N,8,1")
EncoderComm.CommPort = GetSetting(appname:="AquaView", section:="Startup", Key:="EncoderCommPort", Default:=15)
EncoderComm.CommPort = 15
EncoderComm.Settings = "115200,N,8,1"
ScaleComm.Settings = GetSetting(appname:="AquaView", section:="Startup", Key:="ScaleCommSettings", Default:="9600,N,8,1")
ScaleComm.CommPort = GetSetting(appname:="AquaView", section:="Startup", Key:="ScaleCommPort", Default:=12)
ScaleComm.CommPort = 12
ScaleComm.Settings = "9600,N,8,1"

LoadConfig ' load cutter config settings from registry

TempLog = LoggingDir & "\" & LoggingFile
Open TempLog For Append Shared As #1
HeaderLog = LoggingDir & "\" & "JobInfo.log"
Open HeaderLog For Append Shared As #9

AutoSpeed = 0 ' set cut speed
Timer1.Interval = 500
EncoderTimer.Interval = 1000

' open ports
OHPRComm.PortOpen = True
PotComm.PortOpen = True
Analog1Comm.PortOpen = True
Analog2Comm.PortOpen = True
EncoderComm.PortOpen = True
Light1Comm.PortOpen = True
Light2Comm.PortOpen = True
StabilComm.PortOpen = True
ScaleComm.PortOpen = True
    
lblFlowProgress.width = (scrPot.Value / scrPot.Max) * lblFlowBack.width
lblFlowPercent.Caption = Format(((scrPot.Value / scrPot.Max)), "##0.00%")

OrientPlot.GridBackColor = RGB(59, 49, 48)
OrientPlot.GridColor = RGB(173, 186, 116)
OrientPlot.DataColor = RGB(240, 26, 38)
OrientPlot.DataWidth = 5
OrientPlot.GridSpacing = 4
OrientPlot.GridType = gridTarget
OrientPlot.PlotType = plotShape
OrientPlot.plotScale = 90

'frmGauge.FlowPlot.GridOn = False
'frmGauge.FlowPlot.GridBackColor = RGB(59, 49, 48)
'frmGauge.FlowPlot.DataColor = RGB(172, 184, 138)
'frmGauge.FlowPlot.DataWidth = 3
'frmGauge.FlowPlot.PlotType = plotGauge
'frmGauge.FlowPlot.LabelText = "GPM"

'frmGauge.WeightPlot.GridOn = False
'frmGauge.WeightPlot.GridBackColor = RGB(59, 49, 48)
'frmGauge.WeightPlot.DataColor = RGB(172, 184, 138)
'frmGauge.WeightPlot.DataWidth = 3
'frmGauge.WeightPlot.PlotType = plotGauge
'frmGauge.WeightPlot.LabelText = "LBS"

'frmGauge.ForwardPressure.Min = 0
'frmGauge.ForwardPressure.Max = 6000
'frmGauge.ForwardPressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="ForwardLowLimit", Default:="1500")
'frmGauge.ForwardPressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="ForwardUpLimit", Default:="3500")
'frmGauge.ForwardPressure.Inc = 6
'frmGauge.ForwardPressure.Value = 0

'frmGauge.ReversePressure.Min = 0
'frmGauge.ReversePressure.Max = 6000
'frmGauge.ReversePressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="ReverseLowLimit", Default:="1500")
'frmGauge.ReversePressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="ReverseUpLimit", Default:="3500")
'frmGauge.ReversePressure.Inc = 6
'frmGauge.ReversePressure.Value = 0

frmGauge.StabilPressure.Min = 0
frmGauge.StabilPressure.Max = 6000
frmGauge.StabilPressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="StabilLowLimit", Default:="1800")
frmGauge.StabilPressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="StabilUpLimit", Default:="3100")
frmGauge.StabilPressure.Inc = 6
frmGauge.StabilPressure.Value = 0

frmGauge.NitrogenPressure.Min = 0
frmGauge.NitrogenPressure.Max = 500
frmGauge.NitrogenPressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="NitrogenLowLimit", Default:="150")
frmGauge.NitrogenPressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="NitrogenUpLimit", Default:="220")
frmGauge.NitrogenPressure.Inc = 5
frmGauge.NitrogenPressure.Value = 0

frmGauge.CuttingPressure.Min = 0
frmGauge.CuttingPressure.Max = 20000
frmGauge.CuttingPressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="CuttingLowLimit", Default:="17000")
frmGauge.CuttingPressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="CuttingUpLimit", Default:="20000")
frmGauge.CuttingPressure.Inc = 10
frmGauge.CuttingPressure.Value = 0

frmGauge.AirPressure.Min = 0
frmGauge.AirPressure.Max = 500
frmGauge.AirPressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="AirLowLimit", Default:="150")
frmGauge.AirPressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="AirUpLimit", Default:="220")
frmGauge.AirPressure.Inc = 5
frmGauge.AirPressure.Value = 0

frmGauge.GarnetPressure.Min = 0
frmGauge.GarnetPressure.Max = 500
frmGauge.GarnetPressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="GarnetLowLimit", Default:="150")
frmGauge.GarnetPressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="GarnetUpLimit", Default:="250")
frmGauge.GarnetPressure.Inc = 5
frmGauge.GarnetPressure.Value = 0


' set encoders to 13 bit
'Call sm_wait(100)
'OutCommand = "$0L1130"
'EncoderComm.Output = OutCommand & Chr$(13)
'Call sm_wait(100)
'OutCommand = "$0L2130"
'EncoderComm.Output = OutCommand & Chr$(13)
'Call sm_wait(100)

End Sub

Private Sub Form_Resize()
Call ResizeControls(ViewForm)
SM_Speed11.UpdateGraph
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
'If runArgument = "-s" Then
'Else
'ICImagingControl1.LiveStop
'ICImagingControl2.LiveStop
'End If
'SaveSetting appname:="AquaView", section:="Startup", Key:="OHPRCommSettings", Setting:=OHPRComm.Settings
'SaveSetting appname:="AquaView", section:="Startup", Key:="OHPRCommPort", Setting:=OHPRComm.CommPort
'SaveSetting appname:="AquaView", section:="Startup", Key:="EncoderCommSettings", Setting:=EncoderComm.Settings
'SaveSetting appname:="AquaView", section:="Startup", Key:="EncoderCommPort", Setting:=EncoderComm.CommPort
'SaveSetting appname:="AquaView", section:="Startup", Key:="Analog1CommSettings", Setting:=Analog1Comm.Settings
'SaveSetting appname:="AquaView", section:="Startup", Key:="Analog1CommPort", Setting:=Analog1Comm.CommPort
'SaveSetting appname:="AquaView", section:="Startup", Key:="Analog2CommSettings", Setting:=Analog2Comm.Settings
'SaveSetting appname:="AquaView", section:="Startup", Key:="Analog2CommPort", Setting:=Analog2Comm.CommPort
'SaveSetting appname:="AquaView", section:="Startup", Key:="Light1CommSettings", Setting:=Light1Comm.Settings
'SaveSetting appname:="AquaView", section:="Startup", Key:="Light1CommPort", Setting:=Light1Comm.CommPort
'SaveSetting appname:="AquaView", section:="Startup", Key:="PotCommSettings", Setting:=PotComm.Settings
'SaveSetting appname:="AquaView", section:="Startup", Key:="PotCommPort", Setting:=PotComm.CommPort
'SaveSetting appname:="AquaView", section:="Startup", Key:="StabilCommSettings", Setting:=StabilComm.Settings
'SaveSetting appname:="AquaView", section:="Startup", Key:="StabilCommPort", Setting:=StabilComm.CommPort
'SaveSetting appname:="AquaView", section:="Startup", Key:="Light2CommSettings", Setting:=Light2Comm.Settings
'SaveSetting appname:="AquaView", section:="Startup", Key:="Light2CommPort", Setting:=Light2Comm.CommPort
'SaveSetting appname:="AquaView", section:="Startup", Key:="ScaleCommSettings", Setting:=ScaleComm.Settings
'SaveSetting appname:="AquaView", section:="Startup", Key:="ScaleCommPort", Setting:=ScaleComm.CommPort
'SaveSetting appname:="AquaView", section:="EncoderConfig", Key:="Encoder1BitOption", Setting:=Encoder1Bit
'SaveSetting appname:="AquaView", section:="EncoderConfig", Key:="Encoder2BitOption", Setting:=Encoder2Bit
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="ForwardUpLimit", Setting:=(CStr(frmGauge.ForwardPressure.LimitHigh))
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="ForwardLowLimit", Setting:=(CStr(frmGauge.ForwardPressure.LimitLow))
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="ReverseUpLimit", Setting:=(CStr(frmGauge.ReversePressure.LimitHigh))
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="ReverseLowLimit", Setting:=(CStr(frmGauge.ReversePressure.LimitLow))
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="StabilUpLimit", Setting:=(CStr(frmGauge.StabilPressure.LimitHigh))
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="StabilLowLimit", Setting:=(CStr(frmGauge.StabilPressure.LimitLow))
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="NitrogenUpLimit", Setting:=(CStr(frmGauge.NitrogenPressure.LimitHigh))
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="NitrogenLowLimit", Setting:=(CStr(frmGauge.NitrogenPressure.LimitLow))
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="CuttingUpLimit", Setting:=(CStr(frmGauge.CuttingPressure.LimitHigh))
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="CuttingLowLimit", Setting:=(CStr(frmGauge.CuttingPressure.LimitLow))
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="AirUpLimit", Setting:=(CStr(frmGauge.AirPressure.LimitHigh))
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="AirLowLimit", Setting:=(CStr(frmGauge.AirPressure.LimitLow))

'Call SetPot(0)
'Close #1
'If frmGauge.Visible Then
'    frmGauge.Hide
'End If
'If frmVideo.Visible Then
'    frmVideo.Hide
'End If
End
End Sub

Private Sub Light1Comm_OnComm()
Dim InString As String
On Error Resume Next
InString = Light1Comm.Input
Debug.Print InString
End Sub

Private Sub Light2Comm_OnComm()
Dim InString As String
On Error Resume Next
InString = Light2Comm.Input
Debug.Print InString
End Sub

Private Sub LocalControl_Click(Index As Integer)
If Index = 0 Then
    SpeedCont.Caption = "Auto"
    SpeedControl = 0
ElseIf Index = 1 Then
    SpeedCont.Caption = "Manual"
    SpeedControl = 1
End If
SaveSetting appname:="AquaView", section:="CutConfig", Key:="SpeedControl", Setting:=CStr(Index)
End Sub

Private Sub mnuAnalog1_Click()
    Dim iRet As VbMsgBoxResult
    iRet = frmComm.ShowComm(Analog1Comm)      ' open the dialog, which will configure the serial port
    Select Case iRet
    Case vbOK
        Analog1Comm.PortOpen = True             ' actually open the serial port
    Case Else
        ' don't open the port, because the parameters weren't set
    End Select
End Sub

Private Sub mnuConfig_Click()
    Dim iRet As VbMsgBoxResult
    Dim TempLog As String
    
    If StartOption(0).Value = False Then
        Close #1 'close file for configuration
        iRet = frmConfig.ShowConfig      ' open the config dialog
        TempLog = LoggingDir & "\" & LoggingFile
        Open TempLog For Append Shared As #1 ' open logging file after configuration change
    Else
        Call MsgBox("Cutter Must Be Stopped to Change Configuration Settings", vbOKOnly, "Warning")
        Exit Sub
    End If
    LoadConfig             ' read and impliment the settings
End Sub

Private Sub LoadConfig()
' load settings from registry
tempIndex = CInt(GetSetting(appname:="AquaView", section:="CutConfig", Key:="PipeDiam", Default:="6"))
    Select Case tempIndex
    Case 0
        PipeDiameter = 8
    Case 1
        PipeDiameter = 12
    Case 2
        PipeDiameter = 16
    Case 3
        PipeDiameter = 24
    Case 4
        PipeDiameter = 36
    Case 5
        PipeDiameter = 48
    Case 6
        PipeDiameter = 72
    Case 7
        PipeDiameter = CDbl(GetSetting(appname:="AquaView", section:="CutConfig", Key:="CustomDiam", Default:="0.0"))
    Case Else
    End Select
tempIndex = CInt(GetSetting(appname:="AquaView", section:="CutConfig", Key:="CutSpeed", Default:="3"))
    Select Case tempIndex
    Case 0
        CutSpeed = 0.5
    Case 1
        CutSpeed = 1
    Case 2
        CutSpeed = 1.5
    Case 3
        CutSpeed = 2
    Case 4
        CutSpeed = 2.5
    Case 5
        CutSpeed = 99
    Case Else
    End Select
CutDirection = CInt(GetSetting(appname:="AquaView", section:="CutConfig", Key:="CutDirection", Default:="0"))
SpeedControl = CInt(GetSetting(appname:="AquaView", section:="CutConfig", Key:="SpeedControl", Default:="0"))
LoggingDir = CStr(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogDir", Default:="C:"))
LoggingFile = CStr(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogFile", Default:="default.log"))
JobClient = CStr(GetSetting(appname:="AquaView", section:="JobConfig", Key:="Client", Default:=""))
JobVessel = CStr(GetSetting(appname:="AquaView", section:="JobConfig", Key:="Vessel", Default:=""))
JobLocation = CStr(GetSetting(appname:="AquaView", section:="JobConfig", Key:="Location", Default:=""))
JobDescription = CStr(GetSetting(appname:="AquaView", section:="JobConfig", Key:="Description", Default:=""))

SwapCheck = CInt(GetSetting(appname:="AquaView", section:="EncoderConfig", Key:="EncoderSwap", Default:="0"))


' after reading these values set encoder box to 12 or 13 bit... see frmConfig Set1Bits function
Encoder1Bit = CInt(GetSetting(appname:="AquaView", section:="EncoderConfig", Key:="Encoder1BitOption", Default:="0"))
Encoder2Bit = CInt(GetSetting(appname:="AquaView", section:="EncoderConfig", Key:="Encoder2BitOption", Default:="0"))


frmGauge.GarnetPressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="GarnetUpLimit", Default:="250")
frmGauge.GarnetPressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="GarnetLowLimit", Default:="100")
'frmGauge.ReversePressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="ReverseUpLimit", Default:="3500")
'frmGauge.ReversePressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="ReverseLowLimit", Default:="1500")
frmGauge.StabilPressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="StabilUpLimit", Default:="3100")
frmGauge.StabilPressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="StabilLowLimit", Default:="1800")
frmGauge.NitrogenPressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="NitrogenUpLimit", Default:="220")
frmGauge.NitrogenPressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="NitrogenLowLimit", Default:="150")
frmGauge.CuttingPressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="CuttingUpLimit", Default:="20000")
frmGauge.CuttingPressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="CuttingLowLimit", Default:="17000")
frmGauge.AirPressure.LimitHigh = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="AirUpLimit", Default:="220")
frmGauge.AirPressure.LimitLow = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="AirLowLimit", Default:="150")


PipeDiam.Caption = PipeDiameter & " in"
CuttingSpeed.Caption = CutSpeed & " in/min"
If CutDirection = 0 Then
    CutDir.Caption = "Forward"
    OptionDir(0).Value = True
ElseIf CutDirection = 1 Then
    CutDir.Caption = "Reverse"
    OptionDir(1).Value = True
End If


If SpeedControl = 0 Then
    SpeedCont.Caption = "Auto"
    LocalControl(0).Value = True
ElseIf SpeedControl = 1 Then
    SpeedCont.Caption = "Manual"
    LocalControl(1).Value = True
End If
tempIndex = CInt(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogOn", Default:="0"))
    Select Case tempIndex
    Case 0
        LoggingOn = False
    Case 1
        LoggingOn = True
    Case Else
    End Select
tempIndex = CInt(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogPressure", Default:="0"))
    Select Case tempIndex
    Case 0
        LogP = False
    Case 1
        LogP = True
    Case Else
    End Select
tempIndex = CInt(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogEncoder", Default:="0"))
    Select Case tempIndex
    Case 0
        LogE1 = False
    Case 1
        LogE1 = True
    Case Else
    End Select
tempIndex = CInt(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogFlow", Default:="0"))
    Select Case tempIndex
    Case 0
        LogF = False
    Case 1
        LogF = True
    Case Else
    End Select
tempIndex = CInt(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogInclination", Default:="0"))
    Select Case tempIndex
    Case 0
        LogI = False
    Case 1
        LogI = True
    Case Else
    End Select
tempIndex = CInt(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogSpeed", Default:="0"))
    Select Case tempIndex
    Case 0
        LogS = False
    Case 1
        LogS = True
    Case Else
    End Select

End Sub

Private Sub mnuEncoder_Click()
    Dim iRet As VbMsgBoxResult
    iRet = frmComm.ShowComm(EncoderComm)      ' open the dialog, which will configure the serial port
    Select Case iRet
    Case vbOK
        EncoderComm.PortOpen = True             ' actually open the serial port
    Case Else
        ' don't open the port, because the parameters weren't set
    End Select
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuGauge_Click()
frmGauge.Show
End Sub

Private Sub mnuMotion_Click()
    Dim iRet As VbMsgBoxResult
    iRet = frmComm.ShowComm(OHPRComm)      ' open the dialog, which will configure the serial port
    Select Case iRet
    Case vbOK
        OHPRComm.PortOpen = True             ' actually open the serial port
    Case Else
        ' don't open the port, because the parameters weren't set
    End Select
End Sub

Private Sub mnuPlayback_Click()
Timer1.Enabled = False
EncoderTimer.Enabled = False
frmReplayData.Show
End Sub

Private Sub mnuPot_Click()
    Dim iRet As VbMsgBoxResult
    iRet = frmComm.ShowComm(PotComm)      ' open the dialog, which will configure the serial port
    Select Case iRet
    Case vbOK
        PotComm.PortOpen = True             ' actually open the serial port
    Case Else
        ' don't open the port, because the parameters weren't set
    End Select
End Sub

Private Sub mnuVideo_Click()
If frmVideo.ICImagingControl1.DeviceValid And frmVideo.ICImagingControl2.DeviceValid Then
'    Call frmVideo.InitVideo
    frmVideo.Show
Else
    Call MsgBox("Video Devices not detected!", vbOKOnly, "Device Error")
End If
End Sub

Private Sub OHPRComm_OnComm()
Static outString As String
Dim InString As String
Dim TmpString As String
Dim GenericString As String
Dim Cnt As Integer
Dim Max As Integer
Dim XVal As Double
Dim YVal As Double
Dim GenericLength As Integer

Cnt = 1
InString = OHPRComm.Input
Max = Len(InString)
While Cnt <= Max
    TmpString = Mid$(InString, Cnt, 1)
    If TmpString = Chr$(13) Then GoTo NextChar ' skip <cr>
    If TmpString = Chr$(10) Then
        GenericLength = Len(GenericString)
        GenericString = outString ' & Chr$(13) & Chr$(10)
        GenericLength = Len(GenericString)
        If GenericLength > 1 Then
            'Debug.Print "Fucked Up String....................................."
            'Debug.Print GenericString
            Call parseNmea(GenericString)
            CurrentTime = GetTickCount
            ' all things that need to update with motion updates
            If (CurrentTime - NmeaInfo.ohpr.LastUpdate) < 1000 Then
                SysRoll = NmeaInfo.ohpr.Roll
                SysPitch = NmeaInfo.ohpr.Pitch
                SysHeading = NmeaInfo.ohpr.Heading
                SysDepth = (NmeaInfo.ohpr.Depth / 2.31) * 2.24489794
            End If
            outlength = Len(GenericString)
            'DataShift (GenericString)
        End If
        TmpString = ""
        outString = ""
        GenericString = ""
        GoTo NextChar
    End If
    outString = outString + TmpString
NextChar:
    Cnt = Cnt + 1
Wend

End Sub

Private Sub OptionDir_Click(Index As Integer)
Dim SetCount As Integer

For SetCount = 0 To 1
    If OptionDir(SetCount).Value = True Then
        SaveSetting appname:="AquaView", section:="CutConfig", Key:="CutDirection", Setting:=CStr(SetCount)
        CutDirection = SetCount
        Exit For
    End If
Next SetCount
If SetCount = 0 Then
    CutDir.Caption = "Forward"
ElseIf SetCount = 1 Then
    CutDir.Caption = "Reverse"
End If

End Sub

Private Sub PotComm_OnComm()
Static outString As String
Dim InString As String
Dim TmpString As String
Dim GenericString As String
Dim Cnt As Integer
Dim Max As Integer
Dim XVal As Double
Dim YVal As Double
Dim GenericLength As Integer
Cnt = 1
On Error Resume Next
InString = PotComm.Input
Max = Len(InString)
While Cnt <= Max
    TmpString = Mid$(InString, Cnt, 1)
    If TmpString = Chr$(13) Then GoTo NextChar ' skip <cr>
    If TmpString = Chr$(10) Then
        GenericLength = Len(GenericString)
        GenericString = outString
        GenericLength = Len(GenericString)
        If GenericLength > 1 Then
            'Debug.Print GenericString
            outlength = Len(GenericString)
            'SysPressure = CDbl(GenericString) ' use for depth calculation (CDbl(GenericString) * 1.019977334)
'            Debug.Print str(SysPressure)
        End If
        TmpString = ""
        outString = ""
        GenericString = ""
        GoTo NextChar
    End If
    outString = outString + TmpString
NextChar:
    Cnt = Cnt + 1
Wend

End Sub

Private Sub PressureComm_OnComm()
' dBar to Meter multiplier 1.019977334
Static outString As String
Dim InString As String
Dim TmpString As String
Dim GenericString As String
Dim Cnt As Integer
Dim Max As Integer
Dim XVal As Double
Dim YVal As Double
Dim GenericLength As Integer
Cnt = 1
On Error Resume Next
InString = PressureComm.Input
Max = Len(InString)
While Cnt <= Max
    TmpString = Mid$(InString, Cnt, 1)
    If TmpString = Chr$(13) Then GoTo NextChar ' skip <cr>
    If TmpString = Chr$(10) Then
        GenericLength = Len(GenericString)
        GenericString = outString
        GenericLength = Len(GenericString)
        If GenericLength > 1 Then
            'Debug.Print GenericString
            outlength = Len(GenericString)
            SysPressure = CDbl(GenericString) ' use for depth calculation (CDbl(GenericString) * 1.019977334)
            'Debug.Print "SysPressure = " & Str(SysPressure)
        End If
        TmpString = ""
        outString = ""
        GenericString = ""
        GoTo NextChar
    End If
    outString = outString + TmpString
NextChar:
    Cnt = Cnt + 1
Wend
End Sub

Private Sub ReplayTimer_Timer()
' need to read replay file and load data based on timestamp

End Sub

Private Sub ScaleComm_OnComm()
' calculate pounds per minute
' SysMediaFlow = SysMediaWeight / Time
Static outString As String
Dim InString As String
Dim TmpString As String
Dim GenericString As String
Dim Cnt As Integer
Dim Max As Integer
Dim XVal As Double
Dim YVal As Double
Dim GenericLength As Integer
Static filterCnt As Integer
Cnt = 1
On Error Resume Next
InString = ScaleComm.Input
'Debug.Print InString
Max = Len(InString)
While Cnt <= Max
    TmpString = Mid$(InString, Cnt, 1)
    If TmpString = Chr$(13) Then GoTo NextChar ' skip <cr>
    If TmpString = Chr$(10) Then
        GenericLength = Len(GenericString)
        GenericString = outString
        GenericLength = Len(GenericString)
        If GenericLength > 1 Then
            'Debug.Print GenericString
            outlength = Len(GenericString)
            If Abs(CDbl(GenericString) - SysMediaWeight) > 20 And filterCnt < 10 Then
                ' filter out spikes larger than 20 pounds
                filterCnt = filterCnt + 1
            Else
                SysMediaWeight = CDbl(GenericString)
                filterCnt = 0
            End If
            Debug.Print GenericString
        End If
        TmpString = ""
        outString = ""
        GenericString = ""
        GoTo NextChar
    End If
    outString = outString + TmpString
NextChar:
    Cnt = Cnt + 1
Wend
End Sub

Private Sub scrPot_Change()
Dim PotValue As Double
    
    If PotComm.PortOpen = True Then
        If CutDirection = 0 Then
            PotValue = scrPot.Value
            If PotValue = 0 Then
                PotComm.Output = "$" & Chr$(1) & "AO" & Format(PotValue, "-00000.00") & Chr$(13)
            Else
                PotComm.Output = "$" & Chr$(1) & "AO+" & Format(PotValue, "00000.00") & Chr$(13)
            End If
'            Debug.Print "$" & Chr$(1) & "AO+" & Format(PotValue, "00000.00") & Chr$(13)
            Call sm_wait(100)
        ElseIf CutDirection = 1 Then
            PotValue = scrPot.Value * -1
            If PotValue = 0 Then
                PotComm.Output = "$" & Chr$(1) & "AO" & Format(PotValue, "-00000.00") & Chr$(13)
            Else
                PotComm.Output = "$" & Chr$(1) & "AO" & Format(PotValue, "00000.00") & Chr$(13)
            End If
'            Debug.Print "$" & Chr$(1) & "AO" & Format(PotValue, "00000.00") & Chr$(13)
            Call sm_wait(100)
        End If
'        Call SetPot(PotValue)
    Else
        
        'Beep
    End If
    lblFlowProgress.width = (scrPot.Value / scrPot.Max) * lblFlowBack.width
    lblFlowPercent.Caption = Format(((scrPot.Value / scrPot.Max)), "##0.00%")
End Sub

Private Sub SetEncodersZero_Click()
' Set Encoders to Zer0
If EncoderComm.PortOpen = True Then
    OutCommand = "$0S1000"
    EncoderComm.Output = OutCommand & Chr$(13)
    Call sm_wait(100)
End If
Call sm_wait(100)
If EncoderComm.PortOpen = True Then
    OutCommand = "$0S2000"
    EncoderComm.Output = OutCommand & Chr$(13)
    Call sm_wait(100)
End If

End Sub

Private Sub StabilComm_OnComm()
Debug.Print "got reply from relay"
End Sub

Private Sub StabilOff_Click()
If StabilComm.PortOpen Then
    If StabilOn.Value = vbChecked Then ' make sure not working in opposite direction
        StabilOn.Value = vbUnchecked
        StabilComm.Output = Chr$(254) & Chr$(47) & Chr$(1) 'Enter Command Mode, Relay Off, Relay 0
        Call sm_wait(100)
    End If
    If StabilOff.Value = vbChecked Then
        StabilComm.Output = Chr$(254) & Chr$(47) & Chr$(1) 'Enter Command Mode, Relay Off, Relay 1
        Call sm_wait(100)
        StabilComm.Output = Chr$(254) & Chr$(48) & Chr$(0) 'Enter Command Mode, Relay On, Relay 0
        Call sm_wait(100)
    Else
        StabilComm.Output = Chr$(254) & Chr$(47) & Chr$(0) 'Enter Command Mode, Relay Off, Relay 1
        Call sm_wait(100)
    '    StabilComm.Output = Chr$(254) & Chr$(48) & Chr$(0) 'Enter Command Mode, Relay On, Relay 0
    End If
End If
End Sub

Private Sub StabilOn_Click()
If StabilComm.PortOpen Then
    If StabilOff.Value = vbChecked Then ' make sure not working in opposite direction
        StabilOff.Value = vbUnchecked
        StabilComm.Output = Chr$(254) & Chr$(47) & Chr$(0) 'Enter Command Mode, Relay Off, Relay 0
        Call sm_wait(100)
    End If
    If StabilOn.Value = vbChecked Then
        StabilComm.Output = Chr$(254) & Chr$(47) & Chr$(0) 'Enter Command Mode, Relay Off, Relay 0
        Call sm_wait(100)
        StabilComm.Output = Chr$(254) & Chr$(48) & Chr$(1) 'Enter Command Mode, Relay On, Relay 1
        Call sm_wait(100)
    Else
        StabilComm.Output = Chr$(254) & Chr$(47) & Chr$(1) 'Enter Command Mode, Relay Off, Relay 1
        Call sm_wait(100)
    '    StabilComm.Output = Chr$(254) & Chr$(48) & Chr$(0) 'Enter Command Mode, Relay On, Relay 0
    End If
End If
End Sub

Private Sub StartOption_Click(Index As Integer)
If Index = 0 Then
    LocalControl(0).Enabled = False
    LocalControl(1).Enabled = False
    OptionDir(0).Enabled = False
    OptionDir(1).Enabled = False
    If LocalControl(0).Value = True Then
        'scrPot.Enabled = False
    Else
        'scrPot.Enabled = True
    End If
    
ElseIf Index = 1 Then
    LocalControl(0).Enabled = True
    LocalControl(1).Enabled = True
    OptionDir(0).Enabled = True
    OptionDir(1).Enabled = True
    scrPot.Enabled = True
End If
SysSpeed = 0
End Sub

Private Sub Timer1_Timer()
Static ArcLength As Double
Static OldArcLength As Double
Dim count%
Dim ArcDist As Double
Static encoder As Double
Static flow As Double
Static pFlag As Integer
Dim CKSum As String
Dim OutCommand As String
Dim FileString As String
Static en1cnt As Integer
Dim en1val As Integer
Static MediaTime As Integer
Static SumMediaWeight As Double

On Error Resume Next

'*************** SIMULATOR *********************
If runArgument = "-s" Then
EncoderTimer.Enabled = False
If en1cnt < 8190 Then
    en1cnt = en1cnt + 1
Else
    en1cnt = 0
End If

SM_Encoder11.Value1 = (en1cnt / 8191) * 360
SM_Encoder11.Value2 = (en1cnt / 8191) * 360

SysFlow = 2.5

frmGauge.FwdPresChart.Value1 = 2500
frmGauge.RevPresChart.Value1 = 200
frmGauge.StabilPressure.Value = 2950
frmGauge.NitrogenPressure.Value = 200
frmGauge.CuttingPressure.Value = 18000
frmGauge.AirPressure.Value = 150

SM_Speed11.Value = 1.8
 
'SysFlow = (4 / scrPot.Max) * (scrPot.Value)
                
'If (encoder > 0) And (encoder < 359) Then
'    encoder = encoder + ((SysFlow * 7.2) / 60)
'Else
'    encoder = 0.001
'End If
'SysEncoder1 = encoder

''Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
SysRoll = Int((20 - 5 + 1) * Rnd - 10)
SysPitch = Int((25 - 5 + 1) * Rnd - 10)
lblRoll.Caption = SysRoll
lblPitch.Caption = SysPitch

'SysPressure = Int((500 - 400 + 1) * Rnd + 5)
'frmGauge.ForwardPressure.Value = SysFlow * 8008
'frmGauge.ReversePressure.Value = SysFlow * 2999
'frmGauge.StabilPressure.Value = SysFlow * 5000

'ReDim TestArray(2, 2)
'ReDim ColorArray(1)
'TestArray(0, 0) = SysFlow
'Call frmGauge.FlowPlot.LoadData(TestArray, ColorArray)
'''''''''******************************************
'ReDim TestArray(2, 2)
'ReDim ColorArray(1)
'TestArray(0, 0) = 90
'Call frmGauge.WeightPlot.LoadData(TestArray, ColorArray)

PressureChart.Value1 = 50
SysDepth = 50
lblPressure.Caption = Format(SysDepth / 2.24489794, "00.00 psi") & " " & Format(SysDepth, "00.0 ft")


' update orientation plot
ReDim TestArray(4, 2)
ReDim ColorArray(2)
If Abs(SysRoll - RollOffset) + Abs(SysPitch - PitchOffset) > 15 Then
    OrientPlot.GridColor = RGB(255, 0, 0)
Else
    OrientPlot.GridColor = RGB(173, 186, 116)
End If
TestArray(0, 0) = SysRoll - RollOffset
TestArray(1, 0) = SysPitch - PitchOffset
TestArray(2, 0) = 0 '(360 - SysHeading) - 180 ' disable heading reference
TestArray(3, 0) = 360 - AdjustedEncoder1
ColorArray(0) = RGB(240, 26, 38) ' red
ColorArray(1) = "&H00FF00" ' green
Call OrientPlot.LoadData(TestArray, ColorArray)

Exit Sub
End If
'*************** END SIMULATOR ******************

' set roll and pitch displays
lblRoll.Caption = Format(SysRoll - RollOffset, "#0.00")
lblPitch.Caption = Format(SysPitch - PitchOffset, "#0.00")
' update orientation plot
ReDim TestArray(4, 2)
ReDim ColorArray(2)
If Abs(SysRoll - RollOffset) + Abs(SysPitch - PitchOffset) > 15 Then
    OrientPlot.GridColor = RGB(255, 0, 0)
Else
    OrientPlot.GridColor = RGB(173, 186, 116)
End If
TestArray(0, 0) = SysRoll - RollOffset
TestArray(1, 0) = SysPitch - PitchOffset
TestArray(2, 0) = 0 '(360 - SysHeading) - 180 ' disable heading reference
TestArray(3, 0) = 360 - AdjustedEncoder1
ColorArray(0) = RGB(240, 26, 38) ' red
ColorArray(1) = "&H00FF00" ' green
Call OrientPlot.LoadData(TestArray, ColorArray)

'ReDim TestArray(2, 2)
'ReDim ColorArray(1)
'TestArray(0, 0) = SysFlow
'Call frmGauge.FlowPlot.LoadData(TestArray, ColorArray)
frmGauge.lblFlowRate = Format(SysFlow, "0.00 gpm")
'ReDim TestArray(2, 2)
'ReDim ColorArray(1)
'TestArray(0, 0) = SysMediaWeight
'Call frmGauge.WeightPlot.LoadData(TestArray, ColorArray)

' stuff to calculate garnet flow lbs/min
If MediaTime > 120 Then
    frmGauge.lblMediaFlow = Format((SysMediaWeightOld - (SumMediaWeight / 120)), "0.00 lbs/min")
    SysMediaWeightOld = SumMediaWeight / 120
    MediaTime = 0
    SumMediaWeight = 0
Else
    SumMediaWeight = SumMediaWeight + SysMediaWeight
    MediaTime = MediaTime + 1
End If

frmGauge.lblWeight = Format(SysMediaWeight, "000 lbs")
frmGauge.WeightChart.Value1 = SysMediaWeight

frmGauge.GarnetPressure.Value = SysGarnetPressure
'frmGauge.ReversePressure.Value = SysCutterReverse
frmGauge.FwdPresChart.Value1 = SysCutterForward
frmGauge.FwdPresChart.Value2 = SysCutterForward
frmGauge.lblFwdPres.Caption = "Pressure:  " & Format(SysCutterForward, "00.00 psi")
frmGauge.RevPresChart.Value1 = SysCutterReverse
frmGauge.RevPresChart.Value2 = SysCutterReverse
frmGauge.lblRevPres.Caption = "Pressure:  " & Format(SysCutterReverse, "00.00 psi")
frmGauge.StabilPressure.Value = SysStabilOut
frmGauge.NitrogenPressure.Value = SysNitrogenPressure
frmGauge.CuttingPressure.Value = SysWaterPressure
frmGauge.AirPressure.Value = SysAirPressure

' update Pressure
'PressureChart.Value1 = SysPressure * 0.69 * 3.280833333
'lblPressure.Caption = Format(SysPressure * 0.69 * 3.280833333, "0000.00ft")
PressureChart.Value1 = SysDepth / 2.24489794 ' converts depth back to psi
lblPressure.Caption = Format(SysDepth / 2.24489794, "00.00 psi") & " - " & Format(SysDepth, "00.0 ft") ' converts depth back to psi

' do auto speed stuff
If (StartOption(0).Value = True) And (LocalControl(0).Value = True) Then
    If Abs(SM_Speed11.Value - CutSpeed) > 0.05 Then
        If SM_Speed11.Value < CutSpeed Then
                If scrPot.Value < 4999 Then
                    scrPot.Value = scrPot.Value + 5
                End If
        ElseIf SM_Speed11.Value > CutSpeed Then
            If scrPot.Value > 0 Then
                scrPot.Value = scrPot.Value - 5
            End If
        End If
    End If
ElseIf StartOption(1).Value = True Then
    scrPot.Value = 0
End If

If LoggingOn = True Then
    If LogP = True Then
        FileString = CStr(makeUnixTime(Now)) & ",SysPressure," & Format(SysPressure, "##0.00")
        Print #1, FileString
    End If
    If LogI = True Then
        FileString = CStr(makeUnixTime(Now)) & ",SysRoll," & Format(SysRoll, "##0.00")
        Print #1, FileString
        FileString = CStr(makeUnixTime(Now)) & ",SysPitch," & Format(SysPitch, "##0.00")
        Print #1, FileString
    End If
    If LogF = True Then
        FileString = CStr(makeUnixTime(Now)) & ",SysFlow," & Format(SysFlow, "##0.00")
        Print #1, FileString
    End If
    If LogS = True Then
        FileString = CStr(makeUnixTime(Now)) & ",SysSpeed," & Format(SM_Speed11.Value, "##0.00")
        Print #1, FileString
    End If
    If LogE1 = True Then
        FileString = CStr(makeUnixTime(Now)) & ",SysEncoder1," & Format(SysEncoder1, "##0.00")
        Print #1, FileString
        FileString = CStr(makeUnixTime(Now)) & ",SysEncoder2," & Format(SysEncoder2, "##0.00")
        Print #1, FileString
        FileString = CStr(makeUnixTime(Now)) & ",AdjustedEncoder1," & Format(AdjustedEncoder1, "##0.00")
        Print #1, FileString
        FileString = CStr(makeUnixTime(Now)) & ",AdjustedEncoder2," & Format(AdjustedEncoder2, "##0.00")
        Print #1, FileString
        FileString = CStr(makeUnixTime(Now)) & ",EncoderOffset1," & Format(EncoderOffset1, "##0.00")
        Print #1, FileString
        FileString = CStr(makeUnixTime(Now)) & ",EncoderOffset2," & Format(EncoderOffset2, "##0.00")
        Print #1, FileString
    End If
    
    FileString = CStr(makeUnixTime(Now)) & ",SysCutterForward," & Format(SysCutterForward, "##0.00")
    Print #1, FileString
    FileString = CStr(makeUnixTime(Now)) & ",SysCutterReverse," & Format(SysCutterReverse, "##0.00")
    Print #1, FileString
    FileString = CStr(makeUnixTime(Now)) & ",SysWaterPressure," & Format(SysWaterPressure, "##0.00")
    Print #1, FileString
    FileString = CStr(makeUnixTime(Now)) & ",SysNitrogenPressure," & Format(SysNitrogenPressure, "##0.00")
    Print #1, FileString
    FileString = CStr(makeUnixTime(Now)) & ",SysAirPressure," & Format(SysAirPressure, "##0.00")
    Print #1, FileString
    FileString = CStr(makeUnixTime(Now)) & ",SysMediaWeight," & Format(SysMediaWeight, "##0.00")
    Print #1, FileString
    If Video1Start <> "" Then
        FileString = CStr(makeUnixTime(Now)) & ",Video1Start," & Video1Start
        Print #1, FileString
        Video1Start = ""
    End If
    If Video2Start <> "" Then
        FileString = CStr(makeUnixTime(Now)) & ",Video2Start," & Video2Start
        Print #1, FileString
        Video2Start = ""
    End If

End If

End Sub

Public Function SetPot(PotVal As Integer) As Boolean
Dim ShiftReg As Integer

SetPot = False ' assume function failure
ShiftReg = PotVal Mod 4

Select Case ShiftReg
Case 0
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(0) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4)) 'Set Potentiometer Wiper
    Call sm_wait(100)
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(1) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4)) 'Set Potentiometer Wiper
    Call sm_wait(100)
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(2) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4)) 'Set Potentiometer Wiper
    Call sm_wait(100)
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(3) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4)) 'Set Potentiometer Wiper
    Call sm_wait(100)
    'Debug.Print CStr(Int(PotVal / 4)) + " " + CStr(Int(PotVal / 4)) + " " + CStr(Int(PotVal / 4)) + " " + CStr(Int(PotVal / 4))
Case 1
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(0) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4) + 1) 'Set Potentiometer Wiper
    Call sm_wait(100)
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(1) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4)) 'Set Potentiometer Wiper
    Call sm_wait(100)
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(2) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4)) 'Set Potentiometer Wiper
    Call sm_wait(100)
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(3) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4)) 'Set Potentiometer Wiper
    Call sm_wait(100)
    'Debug.Print CStr(Int(PotVal / 4) + 1) + " " + CStr(Int(PotVal / 4)) + " " + CStr(Int(PotVal / 4)) + " " + CStr(Int(PotVal / 4))
Case 2
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(0) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4) + 1) 'Set Potentiometer Wiper
    Call sm_wait(100)
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(1) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4) + 1) 'Set Potentiometer Wiper
    Call sm_wait(100)
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(2) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4)) 'Set Potentiometer Wiper
    Call sm_wait(100)
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(3) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4)) 'Set Potentiometer Wiper
    Call sm_wait(100)
    'Debug.Print CStr(Int(PotVal / 4) + 1) + " " + CStr(Int(PotVal / 4) + 1) + " " + CStr(Int(PotVal / 4)) + " " + CStr(Int(PotVal / 4))
Case 3
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(0) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4) + 1) 'Set Potentiometer Wiper
    Call sm_wait(100)
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(1) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4) + 1) 'Set Potentiometer Wiper
    Call sm_wait(100)
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(2) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4) + 1) 'Set Potentiometer Wiper
    Call sm_wait(100)
    PotComm.Output = Chr$(254) 'Enter Command Mode
    PotComm.Output = Chr$(170) 'Change Single Potentiometer
    PotComm.Output = Chr$(3) 'Select Potentiometer
    PotComm.Output = Chr$(Int(PotVal / 4)) 'Set Potentiometer Wiper
    Call sm_wait(100)
    'Debug.Print CStr(Int(PotVal / 4) + 1) + " " + CStr(Int(PotVal / 4) + 1) + " " + CStr(Int(PotVal / 4) + 1) + " " + CStr(Int(PotVal / 4))
Case Else
    Debug.Print "Didn't get the case statement"
End Select
    
End Function

Private Sub ZeroMotion_Click()
RollOffset = SysRoll
PitchOffset = SysPitch
End Sub

Private Sub ResizeControls(frm As Form)
Dim i As Integer
'   Get ratio of initial form size to current form size
x_size = frm.height / iHeight
y_size = frm.width / iWidth
On Error Resume Next
'Loop though all the objects on the form
For i = 0 To UBound(List)
    'Resize each control individually
    For Each curr_obj In frm
        If TypeOf curr_obj Is Menu Or TypeOf curr_obj Is MSComm Or TypeOf curr_obj Is Timer Then
            'Wrong type of control
        Else
            'Check to make sure its the right control
            If curr_obj.TabIndex = List(i).Index Then
                'Debug.Print CStr(i)
                'Then resize the control
                 With curr_obj
                    .Left = List(i).Left * y_size
                    .width = List(i).width * y_size
                    .height = List(i).height * x_size
                    .Top = List(i).Top * x_size
                    .FontSize = Int(x_size * 8)
                 End With
            End If
        End If
    'Get the next control
    Next curr_obj
Next i
End Sub

Private Function SetFontSize(frm As Form) As Integer
x_size = frm.height / iHeight
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
If TypeOf curr_obj Is Menu Or TypeOf curr_obj Is MSComm Or TypeOf curr_obj Is Timer Then
    'Wrong type of control
Else
    
    ReDim Preserve List(i)
    With List(i)
'            .Name = curr_obj
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

Public Function GetVersion() As Long
    On Error GoTo ErrTrap
    GetVersion& = GetWindowsVersion&
Exit Function
ErrTrap:
    GetVersion& = 0&
End Function

Public Function GetHardDiskSerial(Optional sDrive As String) As Long
    On Error GoTo ErrTrap
    Dim lNumber As Long, sBuffer As String * 255
    If sDrive$ = "" Then sDrive$ = "C"
    Call GetVolumeInformation(sDrive$ & ":\", sBuffer$, 255, lNumber&, 0&, 0&, sBuffer$, 255)
    GetHardDiskSerial& = lNumber&
Exit Function
ErrTrap:
    Call MsgBox("Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error")
    HardDiskSerial& = 0&
End Function

Public Function CheckRegCode(regCode As String) As Boolean
    Dim chkSum As String
    Dim serNum As Long
    Dim serCheck As String
    Dim oldCode As String
    
    serCheck = Mid$(regCode, 1, Len(regCode) - 10)
    serNum = Abs(GetHardDiskSerial("C"))
    chkSum = ChkSumVal(regCode)
    Call CheckKeyCodes(regCode)
    
    Debug.Print serCheck & " " & CStr(serNum) & " " & regCode & " " & chkSum
    If chkSum = "1F" And serCheck = CStr(serNum) Then
        oldCode = GetSetting("MyApp", "ou812", "aRegCode", "")
        Call SaveSetting("MyApp", "ou812", "oldRegCode" & CStr(lastKey), oldCode)
        Call MsgBox("Good Registration Code!", vbOKOnly, "Registration Code Check")
        CheckRegCode = True
    Else
        Call MsgBox("Invalid Registration Code!", vbOKOnly, "Registration Code Check")
        CheckRegCode = False
    End If
End Function

Private Function CheckKeyCodes(regCode As String) As Boolean
Dim tempCode As String
Dim count As Integer

count = 1
tempCode = GetSetting("MyApp", "ou812", "oldRegCode" & CStr(count))
While tempCode <> ""
    If tempCode = regCode Then
        dupKey = True
    End If
    count = count + 1
    tempCode = GetSetting("MyApp", "ou812", "oldRegCode" & CStr(count))
    Debug.Print tempCode & " " & CStr(count)
Wend
lastKey = count
End Function

