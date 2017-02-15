VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Key Generator"
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7740
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtChkSum 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "1F"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtSerNum 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3240
      Top             =   2400
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox lstDetails 
      Height          =   2595
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "Check Sum"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Serial Number"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowsVersion Lib "kernel32" Alias "GetVersion" () As Long
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
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

Public Function GetHDSerialNumber() As String()
Dim arrComputers() As String
Dim strComputer As Variant
Dim objWMIService As Object
Dim colItems As Object
Dim objItem As Object
Const wbemFlagReturnImmediately = &H10
Const wbemFlagForwardOnly = &H20
Dim arHD() As String
Dim i As Integer


    On Error Resume Next
    
    arrComputers = Array(".")
    For Each strComputer In arrComputers
       Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
       Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PhysicalMedia", "WQL", _
                                              wbemFlagReturnImmediately + wbemFlagForwardOnly)
                                              
       
       For Each objItem In colItems
          Debug.Print "SerialNumber: " & objItem.SerialNumber
          Debug.Print "Tag: " & objItem.Tag
          Debug.Print
            ReDim Preserve arHD(i)
            arHD(i) = Trim(objItem.SerialNumber)
            i = i + 1
       Next
    Next

    GetHDSerialNumber = arHD
    
    Set colItems = Nothing
    Set objWMIService = Nothing
    
End Function
Private Sub Form_Load()

ComputerName = "."
Set wmiServices = GetObject( _
    "winmgmts:{impersonationLevel=Impersonate}!//" _
    & ComputerName)
' Get physical disk drive
Set wmiDiskDrives = wmiServices.ExecQuery( _
    "SELECT * FROM Win32_DiskDrive")

For Each wmiDiskDrive In wmiDiskDrives
    txtSerNum.Text = NoSpace(wmiDiskDrive.SerialNumber)
    Exit For
Next

'txtSerNum.Text = Abs(GetHardDiskSerial("C"))
''CheckHDDSerial
''GetHDSerialNumber
'MsgBox GetHardDiskSerial("C")
'MsgBox GetVersion
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #1
End Sub

Private Sub Option1_Click()
Timer1.Enabled = True
Open App.Path & "\" & "1Fcodes.txt" For Append As #1
End Sub

Private Sub Option2_Click()
Timer1.Enabled = False
Close #1
End Sub

Private Sub Timer1_Timer()
Dim TryString As String
Dim cSum As String

TryString = RandomString(10)
cSum = ChkSumVal(txtSerNum.Text & TryString)
If cSum = txtChkSum.Text Then
    Call DisplayData(txtSerNum.Text & TryString & " " & cSum)
    Print #1, txtSerNum.Text & TryString & " " & cSum
End If

End Sub

Function RandomString(cb As Integer) As String

    Randomize
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789"

    Dim i As Long
    For i = 1 To cb
        RandomString = RandomString & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next

End Function

Public Sub DisplayData(RawData)
Dim x
' Add the contents of variable "RawData" to listbox lstInput1.
lstDetails.AddItem (RawData)
' Limit the amount of data in the list box
lstDetails.ListIndex = (lstDetails.ListCount - 1)
If lstDetails.ListCount > 25 Then
    x = 5
    Do While x > 0
        lstDetails.RemoveItem 0
        x = x - 1
    Loop
End If
End Sub

