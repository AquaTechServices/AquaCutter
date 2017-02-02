VERSION 5.00
Begin VB.Form frmComm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Serial Port Settings"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4155
   Icon            =   "frmComm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox InParity 
      Height          =   315
      ItemData        =   "frmComm.frx":014A
      Left            =   2520
      List            =   "frmComm.frx":014C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ComboBox InStopBit 
      Height          =   315
      ItemData        =   "frmComm.frx":014E
      Left            =   2520
      List            =   "frmComm.frx":0150
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox InDataBit 
      Height          =   315
      ItemData        =   "frmComm.frx":0152
      Left            =   2520
      List            =   "frmComm.frx":0154
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox InBaud 
      Height          =   315
      ItemData        =   "frmComm.frx":0156
      Left            =   2520
      List            =   "frmComm.frx":0158
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox InPort 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "&Parity:"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "&Stop bits:"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "&Data bits:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "&Bits per second (baud rate):"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Port:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pComm As MSComm     ' reference to the comm port being configured
Private piRetCode As VbMsgBoxResult  'return code identical to message box return code

Public Function ShowComm(ByRef InComm As MSComm) As VbMsgBoxResult
Dim PortFlag As Integer
Dim OpenFlag As Integer
Dim PortNumber As Integer
Dim PortCount As Integer

On Error Resume Next
If InComm.PortOpen = True Then ' see if port is already open
    InComm.PortOpen = False ' close port for configuration
End If
PortNumber = InComm.CommPort ' save current comm port
' enumerate ports
For PortCount = 1 To 16
    InComm.CommPort = PortCount ' change port number
    InComm.PortOpen = True ' open port
    If InComm.CommEvent = 68 Or Err.Number = 8002 Then ' invalid device Or invalid port
        InPort.AddItem "N/A"
    ElseIf Err.Number = 8005 Then ' port in use
        InPort.AddItem "Com " & CStr(PortCount) & " (In Use)"
    Else
        InComm.PortOpen = False ' close port
        InPort.AddItem "Com " & CStr(PortCount)
    End If
    Err.Clear ' clear errors
Next PortCount

InComm.CommPort = PortNumber ' reset comm to original port
LoadListBoxes ' load comm settings list boxes
Call GetCommParams(InComm) ' setup list boxes with current comm settings
Set pComm = InComm ' set local comm reference for configuration from OkButton

Me.Show vbModal ' show the form and allow for configuration

ShowComm = piRetCode  ' return value

End Function

Private Function GetCommParams(ByRef InComm As MSComm) ' setup list boxes with current comm settings
Dim TempArr() As Variant
    
    InPort.ListIndex = Val(InComm.CommPort) - 1 ' set port list box to current port
    Call sm_parse(InComm.Settings, ",", TempArr)
    Select Case TempArr(1) ' set baud list box to current baud
        Case "300"
            InBaud.ListIndex = 0
        Case "1200"
            InBaud.ListIndex = 1
        Case "2400"
            InBaud.ListIndex = 2
        Case "4800"
            InBaud.ListIndex = 3
        Case "9600"
            InBaud.ListIndex = 4
        Case "14400"
            InBaud.ListIndex = 5
        Case "19200"
            InBaud.ListIndex = 6
        Case "38400"
            InBaud.ListIndex = 7
        Case "57600"
            InBaud.ListIndex = 8
        Case "115200"
            InBaud.ListIndex = 9
        Case Else
    End Select
    Select Case TempArr(3) ' set data bit list box to current data bit
        Case "4"
            InDataBit.ListIndex = 0
        Case "5"
            InDataBit.ListIndex = 1
        Case "6"
            InDataBit.ListIndex = 2
        Case "7"
            InDataBit.ListIndex = 3
        Case "8"
            InDataBit.ListIndex = 4
        Case Else
    End Select
    TempArr(2) = UCase(TempArr(2))
    Select Case TempArr(2) ' set parity list box to current parity
        Case "E"
            InParity.ListIndex = 0
        Case "M"
            InParity.ListIndex = 1
        Case "N"
            InParity.ListIndex = 2
        Case "O"
            InParity.ListIndex = 3
        Case "S"
            InParity.ListIndex = 4
        Case Else
    End Select
    Select Case TempArr(4) ' set stop bit list box to current stop bit
        Case "1"
            InStopBit.ListIndex = 0
        Case "1.5"
            InStopBit.ListIndex = 1
        Case "2"
            InStopBit.ListIndex = 2
        Case Else
    End Select

End Function

Private Sub LoadListBoxes() ' load comm settings list boxes
    InBaud.AddItem "300"
    InBaud.AddItem "1200"
    InBaud.AddItem "2400"
    InBaud.AddItem "4800"
    InBaud.AddItem "9600"
    InBaud.AddItem "14400"
    InBaud.AddItem "19200"
    InBaud.AddItem "38400"
    InBaud.AddItem "57600"
    InBaud.AddItem "115200"
    InDataBit.AddItem "4"
    InDataBit.AddItem "5"
    InDataBit.AddItem "6"
    InDataBit.AddItem "7"
    InDataBit.AddItem "8"
    InParity.AddItem "E"
    InParity.AddItem "M"
    InParity.AddItem "N"
    InParity.AddItem "O"
    InParity.AddItem "S"
    InStopBit.AddItem "1"
    InStopBit.AddItem "1.5"
    InStopBit.AddItem "2"
End Sub

Private Sub CancelButton_Click()
    piRetCode = vbCancel
    Unload Me
End Sub

Private Sub Form_Load()
    piRetCode = vbCancel
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

Private Sub Form_Unload(Cancel As Integer)
' clean up when the form closes
'    pComm = Nothing ' release the local reference object
End Sub

Private Function ConfigurePort(ByRef InComm As MSComm)
' validate the selections
    If InPort.ListIndex = -1 Then
        MsgBox "Port has to be selected", vbExclamation, "COM port configuration"
        InPort.SetFocus
        Exit Function
    End If
    
    If InBaud.ListIndex = -1 Then
        MsgBox "Baud rate has to be selected", vbExclamation, "COM port configuration"
        InBaud.SetFocus
        Exit Function
    End If
    
    If InDataBit.ListIndex = -1 Then
        MsgBox "Number of data bits has to be selected", vbExclamation, "COM port configuration"
        InDataBit.SetFocus
        Exit Function
    End If
        
    If InParity.ListIndex = -1 Then
        MsgBox "Parity has to be selected", vbExclamation, "COM port configuration"
        InParity.SetFocus
        Exit Function
    End If
    
    If InStopBit.ListIndex = -1 Then
        MsgBox "Number of stop bits has to be selected", vbExclamation, "COM port configuration"
        InStopBit.SetFocus
        Exit Function
    End If

    'actually configure the port
    InComm.CommPort = InPort.ListIndex + 1
    InComm.Settings = InBaud & "," & InParity & "," & InDataBit & "," & InStopBit
    

End Function

Private Sub OkButton_Click()
' set the parameters of the MSComm from the GUI
    piRetCode = vbOK
    Call ConfigurePort(pComm) ' configure port settings using local reference
    'pComm = Nothing ' release the local reference object
    Unload Me
End Sub


