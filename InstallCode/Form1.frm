VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Install Code"
   ClientHeight    =   1290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5985
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Install Code"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Code"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtRegCode 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Registration Code"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowsVersion Lib "kernel32" Alias "GetVersion" () As Long
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Public lastKey As Integer
Public dupKey As Boolean
Const HKEY_CURRENT_USER As Long = &H80000001

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
    Dim serNum As String
    Dim serCheck As String
    Dim oldCode As String
    
    serCheck = Mid$(regCode, 1, Len(regCode) - 10)
    serNum = GetHDSerial
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

Private Sub Command1_Click()
Call CheckRegCode(txtRegCode.Text)
End Sub

Private Sub Command2_Click()
Dim goodCode As Boolean
    goodCode = CheckRegCode(txtRegCode.Text)
    If goodCode And dupKey = False Then
        Call SaveSetting("MyApp", "ou812", "aRegCode", txtRegCode.Text)
        Call checkSoftwareTime(True)
    Else
        If dupKey = True Then
            Call MsgBox("Registration Code Has Already Been Used!", vbOKOnly, "Registration Code Install")
        Else
            Call MsgBox("Registration Code Failed!", vbOKOnly, "Registration Code Install")
        End If
    End If
End Sub

Private Sub Form_Load()
lastKey = 0
'CheckHDDSerial
'hdSerial.Text = SysInfo1.GetDiskSerialNum(0)
End Sub

Private Function GetHDSerial() As String
ComputerName = "."
Set wmiServices = GetObject( _
    "winmgmts:{impersonationLevel=Impersonate}!//" _
    & ComputerName)
' Get physical disk drive
Set wmiDiskDrives = wmiServices.ExecQuery( _
    "SELECT * FROM Win32_DiskDrive")

For Each wmiDiskDrive In wmiDiskDrives
    GetHDSerial = NoSpace(wmiDiskDrive.SerialNumber)
    Exit For
Next

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

    ' Recursively get information on the keys.
'    For subkey_num = 1 To subkeys.count
'        Debug.Print subkeys(subkey_num)
'        If subkeys(subkey_num) = "aRegCode" Then
'            tempCode = subkey_values(subkey_num)
            'Debug.Print tempCode
'        ElseIf Mid$(subkeys(subkey_num), 1, Len(subkeys(subkey_num)) - 1) = "oldRegCode" Then
            'Debug.Print subkeys(subkey_num)
'            If tempCode = subkey_values(subkey_num) Then
'                dupKey = True
'            End If
'            lastKey = Mid$(subkeys(subkey_num), Len(subkeys(subkey_num)) - 1, 1)
'        End If
        
Public Function checkSoftwareTime(Optional reset As Boolean) As Boolean
'Output is 1/27/2009 11:14:58 PM
Dim NinetyDays As Long
Dim curTime As String
Dim curSeconds As Long
Dim startTime As Long
Dim endTime As Long
Dim encTime As Long

NinetyDays = 7776000 '324000 seconds
curTime = Now 'get the time
curSeconds = makeUnixTime(curTime) 'convert to unix seconds
'encrypt seconds
'encTime = encryptUnixTime(curSeconds)
'get registry entries
startTime = CLng(GetSetting("MyApp", "ou812", "StartTime", "0"))
endTime = CLng(GetSetting("MyApp", "ou812", "EndTime", "0"))

If reset Then
    Call SaveSetting("MyApp", "ou812", "StartTime", CStr(curSeconds))
    startTime = curSeconds
    If endTime > 0 Then
        Call SaveSetting("MyApp", "ou812", "EndTime", CStr(endTime + NinetyDays))
        endTime = endTime + NinetyDays
    Else
        Call SaveSetting("MyApp", "ou812", "EndTime", CStr(curSeconds + NinetyDays))
        endTime = curSeconds + NinetyDays
    End If
    Call MsgBox(CStr((endTime - startTime) / 86400) & " Days left on current registration code.", vbOKOnly, "Registration Check")
    checkSoftwareTime = True
Else
    'check if this time is between registry times
    If curSeconds < endTime And curSeconds > startTime Then
        'all ok
        If curSeconds > endTime Then
            Call SaveSetting("MyApp", "ou812", "StartTime", curSeconds)
            Call SaveSetting("MyApp", "ou812", "EndTime", CStr(curSeconds + NinetyDays))
        Else
            Call SaveSetting("MyApp", "ou812", "StartTime", curSeconds)
            Call SaveSetting("MyApp", "ou812", "EndTime", CStr(endTime + NinetyDays))
        End If
        Call MsgBox(CStr((endTime - curSeconds) / 86400) & " Days left on current registration code.", vbOKOnly, "Registration Check")
        checkSoftwareTime = True
    Else
        'shut down software
        checkSoftwareTime = False
    End If
End If
End Function

