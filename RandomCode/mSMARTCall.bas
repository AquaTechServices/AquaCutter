Attribute VB_Name = "mSMARTCall"
Option Explicit

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long

Private Type ATTR_DATA
    AttrID As Byte
    AttrName As String
    AttrValue As Byte
    ThresholdValue As Byte
    WorstValue As Byte
    StatusFlags As STATUS_FLAGS
End Type

Public Type DRIVE_INFO
    bDriveType As Byte
    SerialNumber As String
    Model As String
    FirmWare As String
    Cilinders As Long
    Heads As Long
    SecPerTrack As Long
    BytesPerSector As Long
    BytesperTrack As Long
    NumAttributes As Byte
    Attributes() As ATTR_DATA
End Type

Public Enum IDE_DRIVE_NUMBER
    PRIMARY_MASTER
    PRIMARY_SLAVE
    SECONDARY_MASTER
    SECONDARY_SLAVE
End Enum

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const CREATE_NEW = 1

Private Const INVALID_HANDLE_VALUE = -1
Dim di As DRIVE_INFO
Dim colAttrNames As Collection
'***************************************************************************
' Open SMART to allow DeviceIoControl communications. Return SMART handle
'***************************************************************************
Private Function OpenSmart(drv_num As IDE_DRIVE_NUMBER) As Long
   If IsWindowsNT Then
      OpenSmart = CreateFile("\\.\PhysicalDrive" & CStr(drv_num), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
   Else
      OpenSmart = CreateFile("\\.\SMARTVSD", 0, 0, ByVal 0&, CREATE_NEW, 0, 0)
   End If
End Function

'****************************************************************************
' CheckSMARTEnable - Check if SMART enable
' FUNCTION: Send a SMART_ENABLE_SMART_OPERATIONS command to the drive
' bDriveNum = 0-3
'***************************************************************************}
Private Function CheckSMARTEnable(ByVal hDrive As Long, DriveNum As IDE_DRIVE_NUMBER) As Boolean
   'Set up data structures for Enable SMART Command.
   Dim SCIP As SENDCMDINPARAMS
   Dim SCOP As SENDCMDOUTPARAMS
   Dim lpcbBytesReturned As Long
   With SCIP
       .cBufferSize = 0
       With .irDriveRegs
            .bFeaturesReg = SMART_ENABLE_SMART_OPERATIONS
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = SMART_CYL_LOW
            .bCylHighReg = SMART_CYL_HI
   'Compute the drive number.
            .bDriveHeadReg = &HA0 ' Or (DriveNum And 1) * 16
            .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
        End With
        .bDriveNumber = DriveNum
   End With
   CheckSMARTEnable = DeviceIoControl(hDrive, DFP_SEND_DRIVE_COMMAND, SCIP, Len(SCIP) - 4, SCOP, Len(SCOP) - 4, lpcbBytesReturned, ByVal 0&)
End Function

'****************************************************************************
' DoIdentify
' Function: Send an IDENTIFY command to the drive
' DriveNum = 0-3
' IDCmd = IDE_ID_FUNCTION or IDE_ATAPI_ID
'*****************************************************************************

Private Function IdentifyDrive(ByVal hDrive As Long, ByVal IDCmd As Byte, ByVal DriveNum As IDE_DRIVE_NUMBER) As Boolean
    Dim SCIP As SENDCMDINPARAMS
    Dim IDSEC As IDSECTOR
    Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
    Dim sMsg As String
    Dim lpcbBytesReturned As Long
    Dim barrfound(100) As Long
    Dim i As Long
    Dim lng As Long
'   Set up data structures for IDENTIFY command.
    With SCIP
        .cBufferSize = IDENTIFY_BUFFER_SIZE
        .bDriveNumber = CByte(DriveNum)
        With .irDriveRegs
             .bFeaturesReg = 0
             .bSectorCountReg = 1
             .bSectorNumberReg = 1
             .bCylLowReg = 0
             .bCylHighReg = 0
' Compute the drive number.
             .bDriveHeadReg = &HA0
             If Not IsWindowsNT Then .bDriveHeadReg = .bDriveHeadReg Or (DriveNum And 1) * 16
' The command can either be IDE identify or ATAPI identify.
             .bCommandReg = CByte(IDCmd)
        End With
    End With
    If DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), OUTPUT_DATA_SIZE, lpcbBytesReturned, ByVal 0&) Then
       IdentifyDrive = True
       CopyMemory IDSEC, bArrOut(16), Len(IDSEC)
       di.Model = SwapStringBytes(StrConv(IDSEC.sModelNumber, vbUnicode))
       di.FirmWare = SwapStringBytes(StrConv(IDSEC.sFirmwareRev, vbUnicode))
       di.SerialNumber = SwapStringBytes(StrConv(IDSEC.sSerialNumber, vbUnicode))
       di.Cilinders = IDSEC.wNumCyls
       di.Heads = IDSEC.wNumHeads
       di.SecPerTrack = IDSEC.wSectorsPerTrack
    End If
End Function

'****************************************************************************
' ReadAttributesCmd
' FUNCTION: Send a SMART_READ_ATTRIBUTE_VALUES command to the drive
' bDriveNum = 0-3
'***************************************************************************}
Private Function ReadAttributesCmd(ByVal hDrive As Long, DriveNum As IDE_DRIVE_NUMBER) As Boolean
   Dim cbBytesReturned As Long
   Dim SCIP As SENDCMDINPARAMS
   Dim drv_attr As DRIVEATTRIBUTE
   Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
   Dim sMsg As String
   Dim i As Long
   With SCIP
 ' Set up data structures for Read Attributes SMART Command.
       .cBufferSize = READ_ATTRIBUTE_BUFFER_SIZE
       .bDriveNumber = DriveNum
       With .irDriveRegs
            .bFeaturesReg = SMART_READ_ATTRIBUTE_VALUES
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = SMART_CYL_LOW
            .bCylHighReg = SMART_CYL_HI
'  Compute the drive number.
            .bDriveHeadReg = &HA0
            If Not IsWindowsNT Then .bDriveHeadReg = .bDriveHeadReg Or (DriveNum And 1) * 16
            .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
       End With
  End With
  ReadAttributesCmd = DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), OUTPUT_DATA_SIZE, cbBytesReturned, ByVal 0&)
  On Error Resume Next
  For i = 0 To NUM_ATTRIBUTE_STRUCTS - 1
      If bArrOut(18 + i * 12) > 0 Then
         di.Attributes(di.NumAttributes).AttrID = bArrOut(18 + i * 12)
         di.Attributes(di.NumAttributes).AttrName = "Unknown value (" & bArrOut(18 + i * 12) & ")"
         di.Attributes(di.NumAttributes).AttrName = colAttrNames(CStr(bArrOut(18 + i * 12)))
         di.NumAttributes = di.NumAttributes + 1
         ReDim Preserve di.Attributes(di.NumAttributes)
         CopyMemory di.Attributes(di.NumAttributes).StatusFlags, bArrOut(19 + i * 12), 2
         di.Attributes(di.NumAttributes).AttrValue = bArrOut(21 + i * 12)
         di.Attributes(di.NumAttributes).WorstValue = bArrOut(22 + i * 12)
      End If
  Next i
End Function

Private Function ReadThresholdsCmd(ByVal hDrive As Long, DriveNum As IDE_DRIVE_NUMBER) As Boolean
   Dim cbBytesReturned As Long
   Dim SCIP As SENDCMDINPARAMS
   Dim IDSEC As IDSECTOR
   Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
   Dim sMsg As String
   Dim thr_attr As ATTRTHRESHOLD
   Dim i As Long, J As Long
   With SCIP
 ' Set up data structures for Read Attributes SMART Command.
       .cBufferSize = READ_THRESHOLD_BUFFER_SIZE
       .bDriveNumber = DriveNum
       With .irDriveRegs
            .bFeaturesReg = SMART_READ_ATTRIBUTE_THRESHOLDS
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = SMART_CYL_LOW
            .bCylHighReg = SMART_CYL_HI
'  Compute the drive number.
            .bDriveHeadReg = &HA0
            If Not IsWindowsNT Then .bDriveHeadReg = .bDriveHeadReg Or (DriveNum And 1) * 16
            .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
       End With
  End With
  ReadThresholdsCmd = DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), OUTPUT_DATA_SIZE, cbBytesReturned, ByVal 0&)
  For i = 0 To NUM_ATTRIBUTE_STRUCTS - 1
      CopyMemory thr_attr, bArrOut(18 + i * Len(thr_attr)), Len(thr_attr)
      If thr_attr.bAttrID > 0 Then
         For J = 0 To UBound(di.Attributes)
             If thr_attr.bAttrID = di.Attributes(J).AttrID Then
                di.Attributes(J).ThresholdValue = thr_attr.bWarrantyThreshold
                Exit For
             End If
         Next J
      End If
  Next i
End Function

Private Function GetSmartVersion(ByVal hDrive As Long, VersionParams As GETVERSIONOUTPARAMS) As Boolean
   Dim cbBytesReturned As Long
   GetSmartVersion = DeviceIoControl(hDrive, DFP_GET_VERSION, ByVal 0&, 0, VersionParams, Len(VersionParams), cbBytesReturned, ByVal 0&)
End Function

Public Function GetDriveInfo(DriveNum As IDE_DRIVE_NUMBER) As DRIVE_INFO
    Dim hDrive As Long
    Dim VerParam As GETVERSIONOUTPARAMS
    Dim cb As Long
    di.bDriveType = 0
    di.NumAttributes = 0
    ReDim di.Attributes(0)
    hDrive = OpenSmart(DriveNum)
    If hDrive = INVALID_HANDLE_VALUE Then Exit Function
    If Not GetSmartVersion(hDrive, VerParam) Then Exit Function
    'for Others Computer Now salman also
'    If Not IsBitSet(VerParam.bIDEDeviceMap, DriveNum) Then Exit Function
    'For Zeshan Office Computer
'    If IsBitSet(VerParam.bIDEDeviceMap, DriveNum) Then Exit Function
    di.bDriveType = 1 + Abs(IsBitSet(VerParam.bIDEDeviceMap, DriveNum + 4))
    If Not CheckSMARTEnable(hDrive, DriveNum) Then Exit Function
    FillAttrNameCollection
    Call IdentifyDrive(hDrive, IDE_ID_FUNCTION, DriveNum)
    Call ReadAttributesCmd(hDrive, DriveNum)
    Call ReadThresholdsCmd(hDrive, DriveNum)
    GetDriveInfo = di
    CloseHandle hDrive
    Set colAttrNames = Nothing
End Function

Private Function IsWindowsNT() As Boolean
   Dim verinfo As OSVERSIONINFO
   verinfo.dwOSVersionInfoSize = Len(verinfo)
   If (GetVersionEx(verinfo)) = 0 Then Exit Function
   If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
End Function

Private Function IsBitSet(iBitString As Byte, ByVal lBitNo As Integer) As Boolean
    If lBitNo = 7 Then
        IsBitSet = iBitString < 0
    Else
        IsBitSet = iBitString And (2 ^ lBitNo)
    End If
End Function

Private Function SwapStringBytes(ByVal sIn As String) As String
   Dim sTemp As String
   Dim i As Integer
   sTemp = Space(Len(sIn))
   For i = 1 To Len(sIn) - 1 Step 2
       Mid(sTemp, i, 1) = Mid(sIn, i + 1, 1)
       Mid(sTemp, i + 1, 1) = Mid(sIn, i, 1)
   Next i
   SwapStringBytes = sTemp
End Function

Public Sub FillAttrNameCollection()
   Set colAttrNames = New Collection
   With colAttrNames
       .Add "ATTR_INVALID", "0"
       .Add "READ_ERROR_RATE", "1"
       .Add "THROUGHPUT_PERF", "2"
       .Add "SPIN_UP_TIME", "3"
       .Add "START_STOP_COUNT", "4"
       .Add "REALLOC_SECTOR_COUNT", "5"
       .Add "READ_CHANNEL_MARGIN", "6"
       .Add "SEEK_ERROR_RATE", "7"
       .Add "SEEK_TIME_PERF", "8"
       .Add "POWER_ON_HRS_COUNT", "9"
       .Add "SPIN_RETRY_COUNT", "10"
       .Add "CALIBRATION_RETRY_COUNT", "11"
       .Add "POWER_CYCLE_COUNT", "12"
       .Add "SOFT_READ_ERROR_RATE", "13"
       .Add "G_SENSE_ERROR_RATE", "191"
       .Add "POWER_OFF_RETRACT_CYCLE", "192"
       .Add "LOAD_UNLOAD_CYCLE_COUNT", "193"
       .Add "TEMPERATURE", "194"
       .Add "REALLOCATION_EVENTS_COUNT", "196"
       .Add "CURRENT_PENDING_SECTOR_COUNT", "197"
       .Add "UNCORRECTABLE_SECTOR_COUNT", "198"
       .Add "ULTRADMA_CRC_ERROR_RATE", "199"
       .Add "WRITE_ERROR_RATE", "200"
       .Add "DISK_SHIFT", "220"
       .Add "G_SENSE_ERROR_RATEII", "221"
       .Add "LOADED_HOURS", "222"
       .Add "LOAD_UNLOAD_RETRY_COUNT", "223"
       .Add "LOAD_FRICTION", "224"
       .Add "LOAD_UNLOAD_CYCLE_COUNTII", "225"
       .Add "LOAD_IN_TIME", "226"
       .Add "TORQUE_AMPLIFICATION_COUNT", "227"
       .Add "POWER_OFF_RETRACT_COUNT", "228"
       .Add "GMR_HEAD_AMPLITUDE", "230"
       .Add "TEMPERATUREII", "231"
       .Add "READ_ERROR_RETRY_RATE", "250"
   End With
End Sub
