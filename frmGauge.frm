VERSION 5.00
Object = "{AE6EB25D-9E61-4F96-83C6-D51B582C8296}#7.0#0"; "SM_StripChart.ocx"
Object = "{B1CFBB97-7427-409F-A664-DE2DDAAA155E}#11.0#0"; "SM_Pressure.ocx"
Begin VB.Form frmGauge 
   Caption         =   "Hydraulic Gauges"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14790
   Icon            =   "frmGauge.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin SM_StripChart.SM_StripChart1 WeightChart 
      Height          =   1935
      Left            =   11160
      TabIndex        =   22
      Top             =   2040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3413
      Value1Color     =   255
      Value2Color     =   0
      GridBackColor   =   0
      GridColor       =   12632256
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
      Value1Scale     =   2
      DataWrap        =   0
      DataWidth       =   2
      GridMove        =   0   'False
      GridSpacing     =   20
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   11160
      TabIndex        =   21
      Top             =   360
      Width           =   3615
      Begin VB.Label lblMediaFlow 
         Alignment       =   2  'Center
         Caption         =   "5.00 lbs/min"
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
         Left            =   1800
         TabIndex        =   28
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Garnet Flow Rate:"
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
         TabIndex        =   27
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblFlowRate 
         Alignment       =   2  'Center
         Caption         =   "1.00 gpm"
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
         Left            =   2160
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Hydraulic Flow Rate:"
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
         TabIndex        =   25
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblWeight 
         Alignment       =   2  'Center
         Caption         =   "181 lbs"
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
         Left            =   2160
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Garnet Weight:"
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
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
   End
   Begin SM_Pressure.SM_Pressure1 GarnetPressure 
      Height          =   3615
      Left            =   11160
      TabIndex        =   20
      Top             =   4440
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6376
      DataColor       =   -2147483640
      MeterBackColor  =   -2147483643
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
      DataWidth       =   0
      Max             =   0
      Inc             =   0
   End
   Begin SM_StripChart.SM_StripChart1 RevPresChart 
      Height          =   3135
      Left            =   3720
      TabIndex        =   13
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5530
      Value1Color     =   255
      Value2Color     =   65280
      GridBackColor   =   0
      GridColor       =   12632256
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
      Value1Scale     =   2
      Value2Scale     =   4
      Value2Max       =   1000
      DataWrap        =   0
      DataWidth       =   2
      GridMove        =   0   'False
      GridSpacing     =   20
   End
   Begin SM_StripChart.SM_StripChart1 FwdPresChart 
      Height          =   3135
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5530
      Value1Color     =   255
      Value2Color     =   65280
      GridBackColor   =   0
      GridColor       =   12632256
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
      Value1Scale     =   2
      Value2Scale     =   4
      Value2Max       =   1000
      DataWrap        =   0
      DataWidth       =   2
      GridMove        =   0   'False
      GridSpacing     =   20
   End
   Begin SM_Pressure.SM_Pressure1 AirPressure 
      Height          =   3615
      Left            =   3720
      TabIndex        =   8
      Top             =   4440
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6376
      DataColor       =   -2147483640
      MeterBackColor  =   -2147483643
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
      DataWidth       =   0
      Max             =   0
      Inc             =   0
   End
   Begin SM_Pressure.SM_Pressure1 CuttingPressure 
      Height          =   3615
      Left            =   7440
      TabIndex        =   9
      Top             =   4440
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6376
      DataColor       =   -2147483640
      MeterBackColor  =   -2147483643
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
      DataWidth       =   0
      Max             =   0
      Inc             =   0
   End
   Begin SM_Pressure.SM_Pressure1 NitrogenPressure 
      Height          =   3615
      Left            =   0
      TabIndex        =   11
      Top             =   4440
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6376
      DataColor       =   -2147483640
      MeterBackColor  =   -2147483643
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
      DataWidth       =   0
      Max             =   0
      Inc             =   0
   End
   Begin SM_Pressure.SM_Pressure1 StabilPressure 
      Height          =   3615
      Left            =   7440
      TabIndex        =   10
      Top             =   360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6376
      DataColor       =   -2147483640
      MeterBackColor  =   -2147483643
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
      DataWidth       =   0
      Max             =   0
      Inc             =   0
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Garnet Weight"
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
      Left            =   11160
      TabIndex        =   29
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "0-100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6360
      TabIndex        =   19
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "0-1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6360
      TabIndex        =   18
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblRevPres 
      Caption         =   "Pressure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   17
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "0-100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "0-1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblFwdPres 
      Caption         =   "Pressure"
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
      Left            =   0
      TabIndex        =   14
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Garnet Air Pressure"
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
      Left            =   11160
      TabIndex        =   7
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "De-watering Air Pressure"
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
      TabIndex        =   6
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cutting Water Pressure"
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
      TabIndex        =   5
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nitrogen Pressure"
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
      TabIndex        =   4
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Media / Hydraulic"
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
      Left            =   11160
      TabIndex        =   3
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Forward Motor Pressure"
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
      TabIndex        =   2
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stabilizer Arms Pressure"
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
      TabIndex        =   1
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reverse Motor Pressure"
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
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmGauge"
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

Private Sub ResizeControls(frm As Form)
Dim i As Integer
'   Get ratio of initial form size to current form size
x_size = frm.height / iHeight
y_size = frm.width / iWidth
On Error Resume Next
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
                .FontSize = Int(x_size * 8)
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
    ReDim Preserve List(i)
    With List(i)
'        .Name = curr_obj
        .Index = curr_obj.TabIndex
        .Left = curr_obj.Left
        .Top = curr_obj.Top
        .width = curr_obj.width
        .height = curr_obj.height
        Debug.Print CStr(curr_obj.Name)
    End With
    i = i + 1
Next curr_obj
    
'   This is what the object sizes will be compared to on rescaling.
    iHeight = frm.height
    iWidth = frm.width
End Sub

Private Sub Form_Load()
Call GetLocation(frmGauge)
End Sub

'Public Sub CenterForm(frm As Form)
'    frm.Move (Screen.width - frm.width) \ 2, (Screen.height - frm.height) \ 2
'End Sub
'Public Sub ResizeForm(frm As Form)
'    'Set the forms height
'    frm.height = Screen.height / 2
    'Set the forms width
'    frm.width = Screen.width / 2
    'Resize all of the controls
    'based on the forms new size
'    ResizeControls frm
'End Sub


Private Sub Form_Resize()
Call ResizeControls(frmGauge)
End Sub
