VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSetup 
   Caption         =   "Device Setup"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   58
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   57
      Top             =   4200
      Width           =   1095
   End
   Begin TabDlg.SSTab SetupTab 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Motion Sensor"
      TabPicture(0)   =   "frmSetup.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboParity(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboFlowControl(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboStopBits(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboDataBits(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboBPS(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cboPort(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboType(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Encoder"
      TabPicture(1)   =   "frmSetup.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboType(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cboPort(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cboBPS(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cboDataBits(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cboStopBits(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cboFlowControl(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cboParity(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label7(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label2(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label3(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label4(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label5(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label6(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Flow Meter"
      TabPicture(2)   =   "frmSetup.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboType(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cboPort(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cboBPS(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cboDataBits(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cboStopBits(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cboFlowControl(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cboParity(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label7(2)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label1(2)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label2(2)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label3(2)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label4(2)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label5(2)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label6(2)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Potentiometer"
      TabPicture(3)   =   "frmSetup.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cboType(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cboPort(3)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cboBPS(3)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cboDataBits(3)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cboStopBits(3)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cboFlowControl(3)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cboParity(3)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label7(3)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label1(3)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label2(3)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label3(3)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label4(3)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label5(3)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Label6(3)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).ControlCount=   14
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   3
         Left            =   -72240
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   2
         Left            =   -72240
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   1
         Left            =   -72240
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   0
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cboPort 
         Height          =   315
         Index           =   3
         Left            =   -72240
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cboBPS 
         Height          =   315
         Index           =   3
         ItemData        =   "frmSetup.frx":0070
         Left            =   -72240
         List            =   "frmSetup.frx":008C
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cboDataBits 
         Height          =   315
         Index           =   3
         ItemData        =   "frmSetup.frx":00C5
         Left            =   -72240
         List            =   "frmSetup.frx":00D5
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cboStopBits 
         Height          =   315
         Index           =   3
         ItemData        =   "frmSetup.frx":00E5
         Left            =   -72240
         List            =   "frmSetup.frx":00F2
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox cboFlowControl 
         Height          =   315
         Index           =   3
         ItemData        =   "frmSetup.frx":0101
         Left            =   -72240
         List            =   "frmSetup.frx":0108
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox cboParity 
         Height          =   315
         Index           =   3
         ItemData        =   "frmSetup.frx":0112
         Left            =   -72240
         List            =   "frmSetup.frx":012F
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ComboBox cboPort 
         Height          =   315
         Index           =   2
         Left            =   -72240
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cboBPS 
         Height          =   315
         Index           =   2
         ItemData        =   "frmSetup.frx":0151
         Left            =   -72240
         List            =   "frmSetup.frx":016D
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cboDataBits 
         Height          =   315
         Index           =   2
         ItemData        =   "frmSetup.frx":01A6
         Left            =   -72240
         List            =   "frmSetup.frx":01B6
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cboStopBits 
         Height          =   315
         Index           =   2
         ItemData        =   "frmSetup.frx":01C6
         Left            =   -72240
         List            =   "frmSetup.frx":01D3
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox cboFlowControl 
         Height          =   315
         Index           =   2
         ItemData        =   "frmSetup.frx":01E2
         Left            =   -72240
         List            =   "frmSetup.frx":01E9
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox cboParity 
         Height          =   315
         Index           =   2
         ItemData        =   "frmSetup.frx":01F3
         Left            =   -72240
         List            =   "frmSetup.frx":0210
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ComboBox cboPort 
         Height          =   315
         Index           =   1
         Left            =   -72240
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cboBPS 
         Height          =   315
         Index           =   1
         ItemData        =   "frmSetup.frx":0232
         Left            =   -72240
         List            =   "frmSetup.frx":024E
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cboDataBits 
         Height          =   315
         Index           =   1
         ItemData        =   "frmSetup.frx":0287
         Left            =   -72240
         List            =   "frmSetup.frx":0297
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cboStopBits 
         Height          =   315
         Index           =   1
         ItemData        =   "frmSetup.frx":02A7
         Left            =   -72240
         List            =   "frmSetup.frx":02B4
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox cboFlowControl 
         Height          =   315
         Index           =   1
         ItemData        =   "frmSetup.frx":02C3
         Left            =   -72240
         List            =   "frmSetup.frx":02CA
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox cboParity 
         Height          =   315
         Index           =   1
         ItemData        =   "frmSetup.frx":02D4
         Left            =   -72240
         List            =   "frmSetup.frx":02F1
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ComboBox cboPort 
         Height          =   315
         Index           =   0
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cboBPS 
         Height          =   315
         Index           =   0
         ItemData        =   "frmSetup.frx":0313
         Left            =   2760
         List            =   "frmSetup.frx":032F
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cboDataBits 
         Height          =   315
         Index           =   0
         ItemData        =   "frmSetup.frx":0368
         Left            =   2760
         List            =   "frmSetup.frx":0378
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cboStopBits 
         Height          =   315
         Index           =   0
         ItemData        =   "frmSetup.frx":0388
         Left            =   2760
         List            =   "frmSetup.frx":0395
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox cboFlowControl 
         Height          =   315
         Index           =   0
         ItemData        =   "frmSetup.frx":03A4
         Left            =   2760
         List            =   "frmSetup.frx":03AB
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox cboParity 
         Height          =   315
         Index           =   0
         ItemData        =   "frmSetup.frx":03B5
         Left            =   2760
         List            =   "frmSetup.frx":03D2
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Type:"
         Height          =   255
         Index           =   3
         Left            =   -74520
         TabIndex        =   55
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Type:"
         Height          =   255
         Index           =   2
         Left            =   -74520
         TabIndex        =   53
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Type:"
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   51
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Type:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   49
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Port:"
         Height          =   255
         Index           =   3
         Left            =   -74520
         TabIndex        =   48
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "&Bits per second (baud rate):"
         Height          =   255
         Index           =   3
         Left            =   -74520
         TabIndex        =   47
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "&Data bits:"
         Height          =   255
         Index           =   3
         Left            =   -74520
         TabIndex        =   46
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "&Stop bits:"
         Height          =   255
         Index           =   3
         Left            =   -74520
         TabIndex        =   45
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Flow control:"
         Height          =   255
         Index           =   3
         Left            =   -74520
         TabIndex        =   44
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "&Parity:"
         Height          =   255
         Index           =   3
         Left            =   -74520
         TabIndex        =   43
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Port:"
         Height          =   255
         Index           =   2
         Left            =   -74520
         TabIndex        =   36
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "&Bits per second (baud rate):"
         Height          =   255
         Index           =   2
         Left            =   -74520
         TabIndex        =   35
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "&Data bits:"
         Height          =   255
         Index           =   2
         Left            =   -74520
         TabIndex        =   34
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "&Stop bits:"
         Height          =   255
         Index           =   2
         Left            =   -74520
         TabIndex        =   33
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Flow control:"
         Height          =   255
         Index           =   2
         Left            =   -74520
         TabIndex        =   32
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "&Parity:"
         Height          =   255
         Index           =   2
         Left            =   -74520
         TabIndex        =   31
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Port:"
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   24
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "&Bits per second (baud rate):"
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   23
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "&Data bits:"
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   22
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "&Stop bits:"
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   21
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Flow control:"
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   20
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "&Parity:"
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   19
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Port:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "&Bits per second (baud rate):"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "&Data bits:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "&Stop bits:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Flow control:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "&Parity:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   2520
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
