VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4605
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ApplyButton 
      Caption         =   "Apply"
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
   End
   Begin TabDlg.SSTab ConfigTab 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Cutter"
      TabPicture(0)   =   "frmConfig.frx":014A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Job Info"
      TabPicture(1)   =   "frmConfig.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Logging"
      TabPicture(2)   =   "frmConfig.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Drive1"
      Tab(2).Control(1)=   "LoggingOn"
      Tab(2).Control(2)=   "LogFile"
      Tab(2).Control(3)=   "Dir1"
      Tab(2).Control(4)=   "LoggingFrame"
      Tab(2).Control(5)=   "LogText"
      Tab(2).Control(6)=   "Label6"
      Tab(2).Control(7)=   "Label5"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Video"
      TabPicture(3)   =   "frmConfig.frx":019E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblDevice1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblDevice2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame6"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cboDevice1"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cboDevice2"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Frame7"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Encoder"
      TabPicture(4)   =   "frmConfig.frx":01BA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "SetEncoder"
      Tab(4).Control(1)=   "CurrentEncoder"
      Tab(4).Control(2)=   "SwapCheck"
      Tab(4).Control(3)=   "Frame9"
      Tab(4).Control(4)=   "Frame8"
      Tab(4).Control(5)=   "Label7"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Gauge"
      TabPicture(5)   =   "frmConfig.frx":01D6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label8"
      Tab(5).Control(1)=   "Label10"
      Tab(5).Control(2)=   "Label11"
      Tab(5).Control(3)=   "Label12"
      Tab(5).Control(4)=   "Label13"
      Tab(5).Control(5)=   "gpLower"
      Tab(5).Control(6)=   "gpUpper"
      Tab(5).Control(7)=   "spLower"
      Tab(5).Control(8)=   "spUpper"
      Tab(5).Control(9)=   "npLower"
      Tab(5).Control(10)=   "npUpper"
      Tab(5).Control(11)=   "cpLower"
      Tab(5).Control(12)=   "cpUpper"
      Tab(5).Control(13)=   "apLower"
      Tab(5).Control(14)=   "apUpper"
      Tab(5).ControlCount=   15
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   -72600
         TabIndex        =   89
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox apUpper 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71640
         TabIndex        =   88
         Text            =   "Text1"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox apLower 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -72600
         TabIndex        =   87
         Text            =   "Text1"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox cpUpper 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         TabIndex        =   84
         Text            =   "Text2"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox cpLower 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -74760
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox npUpper 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71640
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox npLower 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -72600
         TabIndex        =   81
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox spUpper 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         TabIndex        =   78
         Text            =   "Text2"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox spLower 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -74760
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox gpUpper 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         TabIndex        =   75
         Text            =   "Text2"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox gpLower 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -74760
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton SetEncoder 
         Caption         =   "Set Encoder Value"
         Height          =   255
         Left            =   -72480
         TabIndex        =   73
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox CurrentEncoder 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -74640
         TabIndex        =   71
         Text            =   "360"
         Top             =   2400
         Width           =   735
      End
      Begin VB.CheckBox SwapCheck 
         Caption         =   "Swap Encoder Directions"
         Height          =   375
         Left            =   -74640
         TabIndex        =   70
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Frame Frame9 
         Caption         =   "Encoder 2 Config"
         Height          =   1815
         Left            =   -72600
         TabIndex        =   65
         Top             =   480
         Width           =   1935
         Begin VB.CommandButton Set2Bits 
            Caption         =   "Set Bits"
            Height          =   375
            Left            =   360
            TabIndex        =   69
            Top             =   1200
            Width           =   1215
         End
         Begin VB.OptionButton Encoder2BitOption 
            Caption         =   "13 Bit Encoder"
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   67
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton Encoder2BitOption 
            Caption         =   "12 Bit Encoder"
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   66
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Encoder 1 Config"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   62
         Top             =   480
         Width           =   1935
         Begin VB.CommandButton Set1Bits 
            Caption         =   "Set Bits"
            Height          =   375
            Left            =   360
            TabIndex        =   68
            Top             =   1200
            Width           =   1215
         End
         Begin VB.OptionButton Encoder1BitOption 
            Caption         =   "13 Bit Encoder"
            Height          =   615
            Index           =   1
            Left            =   240
            TabIndex        =   64
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Encoder1BitOption 
            Caption         =   "12 Bit Encoder"
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   63
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Video 2 Overlay Graphics"
         Height          =   735
         Left            =   240
         TabIndex        =   56
         Top             =   2400
         Width           =   4095
         Begin VB.CheckBox InfoCheck2 
            Caption         =   "Job Info"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox LogoCheck2 
            Caption         =   "Logo"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox AttitudeCheck2 
            Caption         =   "Attitude Data"
            Height          =   255
            Left            =   1320
            TabIndex        =   58
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox DateCheck2 
            Caption         =   "Date / Time"
            Height          =   255
            Left            =   2640
            TabIndex        =   57
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.ComboBox cboDevice2 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   1980
         Width           =   2775
      End
      Begin VB.ComboBox cboDevice1 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   540
         Width           =   2775
      End
      Begin VB.Frame Frame6 
         Caption         =   "Video 1 Overlay Graphics"
         Height          =   735
         Left            =   240
         TabIndex        =   48
         Top             =   960
         Width           =   4095
         Begin VB.CheckBox DateCheck1 
            Caption         =   "Date / Time"
            Height          =   255
            Left            =   2760
            TabIndex        =   52
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox AttitudeCheck1 
            Caption         =   "Attitude Data"
            Height          =   255
            Left            =   1320
            TabIndex        =   51
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox InfoCheck1 
            Caption         =   "Job Info"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox LogoCheck1 
            Caption         =   "Logo"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.CheckBox LoggingOn 
         Caption         =   "Enable Logging"
         Height          =   255
         Left            =   -74760
         TabIndex        =   46
         Top             =   540
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox LogFile 
         Height          =   285
         Left            =   -73080
         TabIndex        =   44
         Text            =   "default.log"
         Top             =   2460
         Width           =   2415
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   -72600
         TabIndex        =   42
         Top             =   1140
         Width           =   1935
      End
      Begin VB.Frame LoggingFrame 
         Caption         =   "Logging Parameters"
         Enabled         =   0   'False
         Height          =   1575
         Left            =   -74880
         TabIndex        =   36
         Top             =   780
         Width           =   2055
         Begin VB.CheckBox SpeedLog 
            Caption         =   "Calclulated Speed"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox InclinationLog 
            Caption         =   "Inclination Data"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox PressureLog 
            Caption         =   "Pressure Data"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox FlowLog 
            Caption         =   "Hydraulic Flow"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   480
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox EncoderLog 
            Caption         =   "Encoder Positions"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Value           =   1  'Checked
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Job Information"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   4335
         Begin VB.TextBox DescriptionText 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   1800
            Width           =   3975
         End
         Begin VB.TextBox LocationText 
            Height          =   285
            Left            =   1080
            TabIndex        =   33
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox VesselText 
            Height          =   285
            Left            =   1080
            TabIndex        =   31
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox ClientText 
            Height          =   285
            Left            =   1080
            TabIndex        =   28
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label4 
            Caption         =   "Job Description"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Location"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Vessel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Client"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Speed Control"
         Height          =   1575
         Left            =   -71760
         TabIndex        =   24
         Top             =   1740
         Width           =   1215
         Begin VB.OptionButton SpeedControl 
            Caption         =   "Manual"
            Height          =   375
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton SpeedControl 
            Caption         =   "Auto"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   480
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Direction"
         Height          =   1215
         Left            =   -71760
         TabIndex        =   21
         Top             =   420
         Width           =   1215
         Begin VB.OptionButton CutDirection 
            Caption         =   "Reverse"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Counter Clockwise"
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton CutDirection 
            Caption         =   "Forward"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "Clockwise"
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cutting Speed"
         Height          =   2895
         Left            =   -73320
         TabIndex        =   14
         Top             =   420
         Width           =   1335
         Begin VB.OptionButton CutSpeed 
            Caption         =   "0.5 in/min"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton CutSpeed 
            Caption         =   "Wash"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   19
            Top             =   2160
            Width           =   1095
         End
         Begin VB.OptionButton CutSpeed 
            Caption         =   "2.5 in/min"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Top             =   1800
            Width           =   1095
         End
         Begin VB.OptionButton CutSpeed 
            Caption         =   "2.0 in/min"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   17
            Top             =   1440
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton CutSpeed 
            Caption         =   "1.5 in/min"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton CutSpeed 
            Caption         =   "1.0 in/min"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Pipe Diameter"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   1
         Top             =   420
         Width           =   1335
         Begin VB.TextBox CustomDiam 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   10
            Top             =   2400
            Width           =   495
         End
         Begin VB.OptionButton PipeDiam 
            Caption         =   "Custom"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   9
            Top             =   2040
            Width           =   975
         End
         Begin VB.OptionButton PipeDiam 
            Caption         =   "72"""
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   8
            Top             =   1800
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton PipeDiam 
            Caption         =   "48"""
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   7
            Top             =   1560
            Width           =   975
         End
         Begin VB.OptionButton PipeDiam 
            Caption         =   "36"""
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   6
            Top             =   1320
            Width           =   975
         End
         Begin VB.OptionButton PipeDiam 
            Caption         =   "24"""
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   5
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton PipeDiam 
            Caption         =   "16"""
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   4
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton PipeDiam 
            Caption         =   "12"""
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   3
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton PipeDiam 
            Caption         =   "8"""
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Air Pressure Limits"
         Height          =   255
         Left            =   -72600
         TabIndex        =   86
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Cutting Pressure Limits"
         Height          =   255
         Left            =   -74760
         TabIndex        =   85
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Nitrogen Pressure Limits"
         Height          =   255
         Left            =   -72600
         TabIndex        =   80
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Stabil Pressure Limits"
         Height          =   255
         Left            =   -74760
         TabIndex        =   79
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Garnet Pressure Limits"
         Height          =   255
         Left            =   -74760
         TabIndex        =   76
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Encoder Value"
         Height          =   255
         Left            =   -73800
         TabIndex        =   72
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblDevice2 
         Caption         =   "Device"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label lblDevice1 
         Caption         =   "Device"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   540
         Width           =   975
      End
      Begin VB.Label LogText 
         Height          =   615
         Left            =   -74760
         TabIndex        =   47
         Top             =   2700
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "Logging File Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   45
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Logging Directory"
         Height          =   255
         Left            =   -72600
         TabIndex        =   43
         Top             =   840
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Settings
    DeviceAvail As Boolean
    Device As String
    
    VideoNormAvail As Boolean
    VideoNorm As String
    
    VideoFormat As String
    
    FrameRateAvail As Boolean
    FrameRate As Double
    
    InputChannelAvail As Boolean
    InputChannel As String
    
    FlipVAvail As Boolean
    FlipV As Boolean
        
    FlipHAvail As Boolean
    FlipH As Boolean
End Type

Public ImagingControl As ICImagingControl
Private DeviceState1 As Settings
Private DeviceState2 As Settings
Const NOT_AVAILABLE = "n/a"
Private piRetCode As VbMsgBoxResult  'return code identical to message box return code

Private Sub ApplyButton_Click()
Dim SetCount As Integer
' save settings to registry
For SetCount = 0 To 7
    If PipeDiam(SetCount).Value = True Then
        SaveSetting appname:="AquaView", section:="CutConfig", Key:="PipeDiam", Setting:=CStr(SetCount)
        SaveSetting appname:="AquaView", section:="CutConfig", Key:="CustomDiam", Setting:=CustomDiam.Text
        Exit For
    End If
Next SetCount
For SetCount = 0 To 5
    If CutSpeed(SetCount).Value = True Then
        SaveSetting appname:="AquaView", section:="CutConfig", Key:="CutSpeed", Setting:=CStr(SetCount)
        Exit For
    End If
Next SetCount
For SetCount = 0 To 1
    If CutDirection(SetCount).Value = True Then
        SaveSetting appname:="AquaView", section:="CutConfig", Key:="CutDirection", Setting:=CStr(SetCount)
        ViewForm.CutDirection = SetCount
        Exit For
    End If
Next SetCount
For SetCount = 0 To 1
    If SpeedControl(SetCount).Value = True Then
        SaveSetting appname:="AquaView", section:="CutConfig", Key:="SpeedControl", Setting:=CStr(SetCount)
        Exit For
    End If
Next SetCount
                                                                        ' these .Value are 0 or 1 (1=checked)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogEncoder", Setting:=CStr(EncoderLog.Value)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogFlow", Setting:=CStr(FlowLog.Value)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogPressure", Setting:=CStr(PressureLog.Value)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogInclination", Setting:=CStr(InclinationLog.Value)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogSpeed", Setting:=CStr(SpeedLog.Value)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogOn", Setting:=CStr(LoggingOn.Value)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogDir", Setting:=CStr(Dir1.Path)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogFile", Setting:=CStr(LogFile.Text)

ViewForm.LoggingFile = LogFile.Text
If LoggingOn.Value = 1 Then
    ViewForm.LoggingOn = True
ElseIf LoggingOn.Value = 0 Then
    ViewForm.LoggingOn = False
End If
SaveSetting appname:="AquaView", section:="JobConfig", Key:="Client", Setting:=ClientText.Text
SaveSetting appname:="AquaView", section:="JobConfig", Key:="Vessel", Setting:=VesselText.Text
SaveSetting appname:="AquaView", section:="JobConfig", Key:="Location", Setting:=LocationText.Text
SaveSetting appname:="AquaView", section:="JobConfig", Key:="Description", Setting:=DescriptionText.Text
ViewForm.JobClient = ClientText.Text
ViewForm.JobVessel = VesselText.Text
ViewForm.JobLocation = LocationText.Text
ViewForm.JobDescription = DescriptionText.Text

Print #9, ClientText.Text
Print #9, VesselText.Text
Print #9, LocationText.Text
Print #9, DescriptionText.Text

SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Logo1", Setting:=CStr(LogoCheck1.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Date1", Setting:=CStr(DateCheck1.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Attitude1", Setting:=CStr(AttitudeCheck1.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Info1", Setting:=CStr(InfoCheck1.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Logo2", Setting:=CStr(LogoCheck2.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Date2", Setting:=CStr(DateCheck2.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Attitude2", Setting:=CStr(AttitudeCheck2.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Info2", Setting:=CStr(InfoCheck2.Value)

ViewForm.SwapCheck = SwapCheck.Value
SaveSetting appname:="AquaView", section:="EncoderConfig", Key:="EncoderSwap", Setting:=CStr(SwapCheck.Value)
SaveSetting appname:="AquaView", section:="EncoderConfig", Key:="Encoder1BitOption", Setting:=CStr(ViewForm.Encoder1Bit)
SaveSetting appname:="AquaView", section:="EncoderConfig", Key:="Encoder2BitOption", Setting:=CStr(ViewForm.Encoder2Bit)

frmVideo.Logo1 = LogoCheck1.Value
frmVideo.Date1 = DateCheck1.Value
frmVideo.Attitude1 = AttitudeCheck1.Value
frmVideo.Info1 = InfoCheck1.Value
frmVideo.Logo2 = LogoCheck2.Value
frmVideo.Date2 = DateCheck2.Value
frmVideo.Attitude2 = AttitudeCheck2.Value
frmVideo.Info2 = InfoCheck2.Value

SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="GarnetUpLimit", Setting:=(gpUpper.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="GarnetLowLimit", Setting:=(gpLower.Text)
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="ReverseUpLimit", Setting:=(rpUpper.Text)
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="ReverseLowLimit", Setting:=(rpLower.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="StabilUpLimit", Setting:=(spUpper.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="StabilLowLimit", Setting:=(spLower.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="NitrogenUpLimit", Setting:=(npUpper.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="NitrogenLowLimit", Setting:=(npLower.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="CuttingUpLimit", Setting:=(cpUpper.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="CuttingLowLimit", Setting:=(cpLower.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="AirUpLimit", Setting:=(apUpper.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="AirLowLimit", Setting:=(apLower.Text)

frmGauge.GarnetPressure.LimitHigh = CDbl(gpUpper.Text)
frmGauge.GarnetPressure.LimitLow = CDbl(gpLower.Text)
'frmGauge.ReversePressure.LimitHigh = CDbl(rpUpper.Text)
'frmGauge.ReversePressure.LimitLow = CDbl(rpLower.Text)
frmGauge.StabilPressure.LimitHigh = CDbl(spUpper.Text)
frmGauge.StabilPressure.LimitLow = CDbl(spLower.Text)
frmGauge.NitrogenPressure.LimitHigh = CDbl(npUpper.Text)
frmGauge.NitrogenPressure.LimitLow = CDbl(npLower.Text)
frmGauge.CuttingPressure.LimitHigh = CDbl(cpUpper.Text)
frmGauge.CuttingPressure.LimitLow = CDbl(cpLower.Text)
frmGauge.AirPressure.LimitHigh = CDbl(apUpper.Text)
frmGauge.AirPressure.LimitLow = CDbl(apLower.Text)

End Sub

Private Sub CancelButton_Click()
    piRetCode = vbCancel
'RestoreDeviceSettings
Unload Me
End Sub

Private Sub Dir1_Change()
ViewForm.LoggingDir = Dir1.Path
LogText.Caption = Dir1.Path & "\" & LogFile.Text
End Sub
Public Function ShowConfig() As VbMsgBoxResult
Dim pDiam As Integer
Dim cSpeed As Integer
Dim cDir As Integer
Dim sControl As Integer
Dim TempLog As String
On Error Resume Next
' load settings from registry
pDiam = CInt(GetSetting(appname:="AquaView", section:="CutConfig", Key:="PipeDiam", Default:="6"))
PipeDiam(pDiam).Value = True
CustomDiam.Text = GetSetting(appname:="AquaView", section:="CutConfig", Key:="CustomDiam", Default:="0.0")
cSpeed = CInt(GetSetting(appname:="AquaView", section:="CutConfig", Key:="CutSpeed", Default:="3"))
CutSpeed(cSpeed).Value = True
cDir = CInt(GetSetting(appname:="AquaView", section:="CutConfig", Key:="CutDirection", Default:="0"))
CutDirection(cDir).Value = True
sControl = CInt(GetSetting(appname:="AquaView", section:="CutConfig", Key:="SpeedControl", Default:="0"))
SpeedControl(sControl).Value = True

Dir1.Path = CStr(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogDir", Default:="C:"))
LogFile.Text = CStr(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogFile", Default:="default.log"))
LogText.Caption = Dir1.Path & "\" & LogFile.Text

EncoderLog.Value = CInt(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogEncoder", Default:="0"))
FlowLog.Value = CInt(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogFlow", Default:="0"))
PressureLog.Value = CInt(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogPressure", Default:="0"))
InclinationLog.Value = CInt(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogInclination", Default:="0"))
SpeedLog.Value = CInt(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogSpeed", Default:="0"))
LoggingOn.Value = CInt(GetSetting(appname:="AquaView", section:="LogConfig", Key:="LogOn", Default:="0"))
ClientText.Text = ViewForm.JobClient
VesselText.Text = ViewForm.JobVessel
LocationText.Text = ViewForm.JobLocation
DescriptionText.Text = ViewForm.JobDescription

LogoCheck1.Value = 0 'CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Logo1", Default:="0"))
DateCheck1.Value = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Date1", Default:="0"))
AttitudeCheck1.Value = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Attitude1", Default:="0"))
'CircleCheck1.Value = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Circle1", Default:="0"))
InfoCheck1.Value = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Info1", Default:="0"))
LogoCheck2.Value = 0 'CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Logo2", Default:="0"))
DateCheck2.Value = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Date2", Default:="0"))
AttitudeCheck2.Value = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Attitude2", Default:="0"))
'CircleCheck2.Value = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Circle2", Default:="0"))
InfoCheck2.Value = CInt(GetSetting(appname:="AquaView", section:="VideoConfig", Key:="Info2", Default:="0"))

SwapCheck.Value = CInt(GetSetting(appname:="AquaView", section:="EncoderConfig", Key:="EncoderSwap", Default:="0"))
If ViewForm.Encoder1Bit = 0 Then
    Encoder1BitOption(0).Value = True
ElseIf ViewForm.Encoder1Bit = 1 Then
    Encoder1BitOption(1).Value = True
End If
If ViewForm.Encoder2Bit = 0 Then
    Encoder2BitOption(0).Value = True
ElseIf ViewForm.Encoder2Bit = 1 Then
    Encoder2BitOption(1).Value = True
End If

CurrentEncoder.Text = ViewForm.AdjustedEncoder1

gpUpper.Text = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="GarnetUpLimit", Default:="250")
gpLower.Text = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="GarnetLowLimit", Default:="100")
'rpUpper.Text = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="ReverseUpLimit", Default:="3500")
'rpLower.Text = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="ReverseLowLimit", Default:="1500")
spUpper.Text = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="StabilUpLimit", Default:="3100")
spLower.Text = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="StabilLowLimit", Default:="1800")
npUpper.Text = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="NitrogenUpLimit", Default:="220")
npLower.Text = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="NitrogenLowLimit", Default:="150")
cpUpper.Text = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="CuttingUpLimit", Default:="20000")
cpLower.Text = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="CuttingLowLimit", Default:="17000")
apUpper.Text = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="AirUpLimit", Default:="220")
apLower.Text = GetSetting(appname:="AquaView", section:="GaugeConfig", Key:="AirLowLimit", Default:="220")

ViewForm.LoggingOn = False ' turn off while file closed
'Close #1 'close file for configuration

Me.Show vbModal ' show the form and allow for configuration

'TempLog = Dir1.Path & "\" & LogFile.Text
'Open TempLog For Append Shared As #1
If LoggingOn.Value = 1 Then
    ViewForm.LoggingOn = True
End If

ShowConfig = piRetCode  ' return value

End Function

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    piRetCode = vbCancel
    'ImagingControl = frmVideo.ICImagingControl1
    
    If frmVideo.ICImagingControl1.DeviceValid Then
        If frmVideo.ICImagingControl1.LiveVideoRunning Then
            lblDevice1.Visible = False
            cboDevice1.Visible = False
        End If
    End If
   
    If frmVideo.ICImagingControl2.DeviceValid Then
        If frmVideo.ICImagingControl2.LiveVideoRunning Then
            lblDevice2.Visible = False
            cboDevice2.Visible = False
            Exit Sub
        End If
    End If
    
    'SaveDeviceSettings
    
    'UpdateDevices

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' generate a return value even if user manually closes window
    Select Case UnloadMode
    Case vbFormControlMenu
        piRetCode = vbCancel
    Case Else
        ' other cases are taken care of with the Ok and Cancel buttons
    End Select
End Sub

Private Sub LogFile_LostFocus()
    If right(LogFile.Text, 4) = ".log" Then
        LogText.Caption = Dir1.Path & "\" & LogFile.Text
    Else
        LogFile.Text = LogFile.Text & ".log"
        LogText.Caption = Dir1.Path & "\" & LogFile.Text
    End If
End Sub

Private Sub LoggingOn_Click()
If LoggingOn.Value = 1 Then
    LoggingFrame.Enabled = True
    EncoderLog.Enabled = True
    FlowLog.Enabled = True
    PressureLog.Enabled = True
    InclinationLog.Enabled = True
    SpeedLog.Enabled = True
Else
    LoggingFrame.Enabled = False
    EncoderLog.Enabled = False
    FlowLog.Enabled = False
    PressureLog.Enabled = False
    InclinationLog.Enabled = False
    SpeedLog.Enabled = False
End If
End Sub

Private Sub OkButton_Click()
Dim SetCount As Integer
' save settings to registry
For SetCount = 0 To 7
    If PipeDiam(SetCount).Value = True Then
        SaveSetting appname:="AquaView", section:="CutConfig", Key:="PipeDiam", Setting:=CStr(SetCount)
        SaveSetting appname:="AquaView", section:="CutConfig", Key:="CustomDiam", Setting:=CustomDiam.Text
        Exit For
    End If
Next SetCount
For SetCount = 0 To 5
    If CutSpeed(SetCount).Value = True Then
        SaveSetting appname:="AquaView", section:="CutConfig", Key:="CutSpeed", Setting:=CStr(SetCount)
        Exit For
    End If
Next SetCount
For SetCount = 0 To 1
    If CutDirection(SetCount).Value = True Then
        SaveSetting appname:="AquaView", section:="CutConfig", Key:="CutDirection", Setting:=CStr(SetCount)
        ViewForm.CutDirection = SetCount
        Exit For
    End If
Next SetCount
For SetCount = 0 To 1
    If SpeedControl(SetCount).Value = True Then
        SaveSetting appname:="AquaView", section:="CutConfig", Key:="SpeedControl", Setting:=CStr(SetCount)
        Exit For
    End If
Next SetCount
                                                                        ' these .Value are 0 or 1 (1=checked)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogEncoder", Setting:=CStr(EncoderLog.Value)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogFlow", Setting:=CStr(FlowLog.Value)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogPressure", Setting:=CStr(PressureLog.Value)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogInclination", Setting:=CStr(InclinationLog.Value)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogSpeed", Setting:=CStr(SpeedLog.Value)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogOn", Setting:=CStr(LoggingOn.Value)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogDir", Setting:=CStr(Dir1.Path)
SaveSetting appname:="AquaView", section:="LogConfig", Key:="LogFile", Setting:=CStr(LogFile.Text)

ViewForm.LoggingFile = LogFile.Text
If LoggingOn.Value = 1 Then
    ViewForm.LoggingOn = True
ElseIf LoggingOn.Value = 0 Then
    ViewForm.LoggingOn = False
End If

SaveSetting appname:="AquaView", section:="JobConfig", Key:="Client", Setting:=ClientText.Text
SaveSetting appname:="AquaView", section:="JobConfig", Key:="Vessel", Setting:=VesselText.Text
SaveSetting appname:="AquaView", section:="JobConfig", Key:="Location", Setting:=LocationText.Text
SaveSetting appname:="AquaView", section:="JobConfig", Key:="Description", Setting:=DescriptionText.Text
ViewForm.JobClient = ClientText.Text
ViewForm.JobVessel = VesselText.Text
ViewForm.JobLocation = LocationText.Text
ViewForm.JobDescription = DescriptionText.Text

Print #9, ClientText.Text
Print #9, VesselText.Text
Print #9, LocationText.Text
Print #9, DescriptionText.Text

SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Logo1", Setting:=CStr(LogoCheck1.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Date1", Setting:=CStr(DateCheck1.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Attitude1", Setting:=CStr(AttitudeCheck1.Value)
'SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Circle1", Setting:=CStr(CircleCheck1.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Info1", Setting:=CStr(InfoCheck1.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Logo2", Setting:=CStr(LogoCheck2.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Date2", Setting:=CStr(DateCheck2.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Attitude2", Setting:=CStr(AttitudeCheck2.Value)
'SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Circle2", Setting:=CStr(CircleCheck2.Value)
SaveSetting appname:="AquaView", section:="VideoConfig", Key:="Info2", Setting:=CStr(InfoCheck2.Value)

ViewForm.SwapCheck = SwapCheck.Value
SaveSetting appname:="AquaView", section:="EncoderConfig", Key:="EncoderSwap", Setting:=CStr(SwapCheck.Value)
SaveSetting appname:="AquaView", section:="EncoderConfig", Key:="Encoder1BitOption", Setting:=CStr(ViewForm.Encoder1Bit)
SaveSetting appname:="AquaView", section:="EncoderConfig", Key:="Encoder2BitOption", Setting:=CStr(ViewForm.Encoder2Bit)

frmVideo.Logo1 = LogoCheck1.Value
frmVideo.Date1 = DateCheck1.Value
frmVideo.Attitude1 = AttitudeCheck1.Value
'frmVideo.Circle1 = CircleCheck1.Value
frmVideo.Info1 = InfoCheck1.Value
frmVideo.Logo2 = LogoCheck2.Value
frmVideo.Date2 = DateCheck2.Value
frmVideo.Attitude2 = AttitudeCheck2.Value
'frmVideo.Circle2 = CircleCheck2.Value
frmVideo.Info2 = InfoCheck2.Value

SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="GarnetUpLimit", Setting:=(gpUpper.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="GarnetLowLimit", Setting:=(gpLower.Text)
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="ReverseUpLimit", Setting:=(rpUpper.Text)
'SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="ReverseLowLimit", Setting:=(rpLower.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="StabilUpLimit", Setting:=(spUpper.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="StabilLowLimit", Setting:=(spLower.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="NitrogenUpLimit", Setting:=(npUpper.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="NitrogenLowLimit", Setting:=(npLower.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="CuttingUpLimit", Setting:=(cpUpper.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="CuttingLowLimit", Setting:=(cpLower.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="AirUpLimit", Setting:=(apUpper.Text)
SaveSetting appname:="AquaView", section:="GaugeConfig", Key:="AirLowLimit", Setting:=(apLower.Text)

frmGauge.GarnetPressure.LimitHigh = CDbl(gpUpper.Text)
frmGauge.GarnetPressure.LimitLow = CDbl(gpLower.Text)
'frmGauge.ReversePressure.LimitHigh = CDbl(rpUpper.Text)
'frmGauge.ReversePressure.LimitLow = CDbl(rpLower.Text)
frmGauge.StabilPressure.LimitHigh = CDbl(spUpper.Text)
frmGauge.StabilPressure.LimitLow = CDbl(spLower.Text)
frmGauge.NitrogenPressure.LimitHigh = CDbl(npUpper.Text)
frmGauge.NitrogenPressure.LimitLow = CDbl(npLower.Text)
frmGauge.CuttingPressure.LimitHigh = CDbl(cpUpper.Text)
frmGauge.CuttingPressure.LimitLow = CDbl(cpLower.Text)
frmGauge.AirPressure.LimitHigh = CDbl(apUpper.Text)
frmGauge.AirPressure.LimitLow = CDbl(apLower.Text)

piRetCode = vbOK
Unload Me
End Sub

Private Sub PipeDiam_Click(Index As Integer)
If PipeDiam(7).Value = True Then
    CustomDiam.Enabled = True
Else
    CustomDiam.Enabled = False
End If
End Sub

Private Sub Set1Bits_Click()
Dim OutCommand As String
If ViewForm.EncoderComm.PortOpen Then
    If Encoder1BitOption(0).Value = True Then ' 12 bit encoder
        OutCommand = "$0L1120"
        ViewForm.EncoderComm.Output = OutCommand & Chr$(13)
        Call sm_wait(100)
        ViewForm.Encoder1Bit = 0
    ElseIf Encoder1BitOption(1).Value = True Then ' 13 bit encoder
        OutCommand = "$0L1130"
        ViewForm.EncoderComm.Output = OutCommand & Chr$(13)
        Call sm_wait(100)
        ViewForm.Encoder1Bit = 1
    End If
    'Debug.Print OutCommand
End If
End Sub

Private Sub Set2Bits_Click()
Dim OutCommand As String
If ViewForm.EncoderComm.PortOpen Then
    If Encoder2BitOption(0).Value = True Then ' 12 bit encoder
        OutCommand = "$0L2120"
        ViewForm.EncoderComm.Output = OutCommand & Chr$(13)
        Call sm_wait(100)
        ViewForm.Encoder2Bit = 0
    ElseIf Encoder2BitOption(1).Value = True Then ' 13 bit encoder
        OutCommand = "$0L2130"
        ViewForm.EncoderComm.Output = OutCommand & Chr$(13)
        Call sm_wait(100)
        ViewForm.Encoder2Bit = 1
    End If
    'Debug.Print OutCommand
End If
End Sub

Private Sub SetEncoder_Click()
'ViewForm.EncoderOffset1 = ViewForm.SysEncoder1
ViewForm.EncoderOffset1 = ViewForm.SysEncoder1 - Val(CurrentEncoder.Text)
ViewForm.EncoderOffset2 = ViewForm.SysEncoder2 - Val(CurrentEncoder.Text)
End Sub

Private Sub SpeedControl_Click(Index As Integer)
If Index = 0 Then
    ViewForm.LocalControl(0).Value = True
ElseIf Index = 1 Then
    ViewForm.LocalControl(1).Value = True
End If
End Sub

Private Sub cboDevice1_Click()
    Dim Serial As String

    On Error GoTo err_cboDevice1_Click
        'Open the device
        If cboDevice1.Enabled Then
            Dim Item As Variant
            Dim Index As Integer
            Index = 1

            frmVideo.ICImagingControl1.Device = cboDevice1.Text

            For Each Item In frmVideo.ICImagingControl1.Devices
                If Item.Name = cboDevice1.Text Then
                End If
                Index = Index + 1
            Next
        End If

    Exit Sub
err_cboDevice1_Click:
    MsgBox Err.Description
End Sub
Private Sub cboDevice2_Click()
    Dim Serial As String

    On Error GoTo err_cboDevice2_Click
        'Open the device
        If cboDevice2.Enabled Then
            Dim Item As Variant
            Dim Index As Integer
            Index = 1

            frmVideo.ICImagingControl2.Device = cboDevice2.Text

            For Each Item In frmVideo.ICImagingControl2.Devices
                If Item.Name = cboDevice2.Text Then
                End If
                Index = Index + 1
            Next
        End If

    Exit Sub
err_cboDevice2_Click:
    MsgBox Err.Description
End Sub
'
' UpdateDevices
'
' Fills cboDevice combo list box
'
Private Sub UpdateDevices()
    cboDevice1.Clear
    If frmVideo.ICImagingControl1.Devices.count > 0 Then
        Dim Item As Variant
        Dim Index As Long
       
        For Each Item In frmVideo.ICImagingControl1.Devices
            cboDevice1.AddItem Item.Name
        Next

        If frmVideo.ICImagingControl1.DeviceValid Then
            Index = frmVideo.ICImagingControl1.Devices.FindIndex(frmVideo.ICImagingControl1.Device)
            cboDevice1.ListIndex = Index - 1
        Else
            cboDevice1.ListIndex = 0
        End If
        cboDevice1.Enabled = True
    Else
        cboDevice1.AddItem NOT_AVAILABLE
        cboDevice1.Enabled = False
        cboDevice1.ListIndex = 0
    End If
    cboDevice2.Clear
    If frmVideo.ICImagingControl2.Devices.count > 0 Then
        Dim Item2 As Variant
        Dim index2 As Long
       
        For Each Item2 In frmVideo.ICImagingControl2.Devices
            cboDevice2.AddItem Item2.Name
        Next

        If frmVideo.ICImagingControl2.DeviceValid Then
            index2 = frmVideo.ICImagingControl2.Devices.FindIndex(frmVideo.ICImagingControl2.Device)
            cboDevice2.ListIndex = index2 - 1
        Else
            cboDevice2.ListIndex = 0
        End If
        cboDevice2.Enabled = True
    Else
        cboDevice2.AddItem NOT_AVAILABLE
        cboDevice2.Enabled = False
        cboDevice2.ListIndex = 0
    End If
    
End Sub

Private Sub RestoreDeviceSettings()
    If DeviceState1.DeviceAvail Then
        frmVideo.ICImagingControl1.Device = DeviceState1.Device
    End If
    If DeviceState2.DeviceAvail Then
        frmVideo.ICImagingControl2.Device = DeviceState2.Device
    End If
End Sub

Private Sub SaveDeviceSettings()
    If frmVideo.ICImagingControl1.DeviceValid Then
        DeviceState1.DeviceAvail = True
        DeviceState1.Device = frmVideo.ICImagingControl1.Device
    Else
        DeviceState1.DeviceAvail = False
    End If
    If frmVideo.ICImagingControl2.DeviceValid Then
        DeviceState2.DeviceAvail = True
        DeviceState2.Device = frmVideo.ICImagingControl2.Device
    Else
        DeviceState2.DeviceAvail = False
    End If
End Sub
