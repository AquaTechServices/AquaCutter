Attribute VB_Name = "mSMARTDef"
Public Const MAX_IDE_DRIVES = 4         ' // Max number of drives assuming primary/secondary, master/slave topology
Public Const READ_ATTRIBUTE_BUFFER_SIZE = 512
Public Const IDENTIFY_BUFFER_SIZE = 512
Public Const READ_THRESHOLD_BUFFER_SIZE = 512
Public Const OUTPUT_DATA_SIZE = IDENTIFY_BUFFER_SIZE + 16

'IOCTL commands
Public Const DFP_GET_VERSION = &H74080
Public Const DFP_SEND_DRIVE_COMMAND = &H7C084
Public Const DFP_RECEIVE_DRIVE_DATA = &H7C088

'---------------------------------------------------------------------
' GETVERSIONOUTPARAMS contains the data returned from the
' Get Driver Version function.
'---------------------------------------------------------------------
Public Type GETVERSIONOUTPARAMS
       bVersion       As Byte ' Binary driver version.
       bRevision      As Byte ' Binary driver revision.
       bReserved      As Byte ' Not used.
       bIDEDeviceMap  As Byte ' Bit map of IDE devices.
       fCapabilities  As Long ' Bit mask of driver capabilities.
       dwReserved(3)  As Long ' For future use.
End Type

'Bits returned in the fCapabilities member of GETVERSIONOUTPARAMS
Public Const CAP_IDE_ID_FUNCTION = 1             ' ATA ID command supported
Public Const CAP_IDE_ATAPI_ID = 2                ' ATAPI ID command supported
Public Const CAP_IDE_EXECUTE_SMART_FUNCTION = 4  ' SMART commannds supported

'---------------------------------------------------------------------
' IDE registers
'---------------------------------------------------------------------
Public Type IDEREGS
   bFeaturesReg     As Byte ' // Used for specifying SMART "commands".
   bSectorCountReg  As Byte ' // IDE sector count register
   bSectorNumberReg As Byte ' // IDE sector number register
   bCylLowReg       As Byte ' // IDE low order cylinder value
   bCylHighReg      As Byte ' // IDE high order cylinder value
   bDriveHeadReg    As Byte ' // IDE drive/head register
   bCommandReg      As Byte ' // Actual IDE command.
   bReserved        As Byte ' // reserved for future use.  Must be zero.
End Type

'---------------------------------------------------------------------
' SENDCMDINPARAMS contains the input parameters for the
' Send Command to Drive function.
'---------------------------------------------------------------------
Public Type SENDCMDINPARAMS
   cBufferSize     As Long     ' Buffer size in bytes
   irDriveRegs     As IDEREGS  ' Structure with drive register values.
   bDriveNumber    As Byte     ' Physical drive number to send command to (0,1,2,3).
   bReserved(2)    As Byte     ' Bytes reserved
   dwReserved(3)   As Long     ' DWORDS reserved
   bBuffer()      As Byte      ' Input buffer.
End Type

' Valid values for the bCommandReg member of IDEREGS.
Public Const IDE_ATAPI_ID = &HA1               ' Returns ID sector for ATAPI.
Public Const IDE_ID_FUNCTION = &HEC            ' Returns ID sector for ATA.
Public Const IDE_EXECUTE_SMART_FUNCTION = &HB0 ' Performs SMART cmd.
                                               ' Requires valid bFeaturesReg,
                                               ' bCylLowReg, and bCylHighReg

' Cylinder register values required when issuing SMART command
Public Const SMART_CYL_LOW = &H4F
Public Const SMART_CYL_HI = &HC2

'---------------------------------------------------------------------
' Status returned from driver
'---------------------------------------------------------------------
Public Type DRIVERSTATUS
   bDriverError  As Byte          ' Error code from driver, or 0 if no error.
   bIDEStatus    As Byte          ' Contents of IDE Error register.
                                  ' Only valid when bDriverError is SMART_IDE_ERROR.
   bReserved(1)  As Byte
   dwReserved(1) As Long
 End Type

' bDriverError values
Public Enum DRIVER_ERRORS
       SMART_NO_ERROR = 0         ' No error
       SMART_IDE_ERROR = 1        ' Error from IDE controller
       SMART_INVALID_FLAG = 2     ' Invalid command flag
       SMART_INVALID_COMMAND = 3  ' Invalid command byte
       SMART_INVALID_BUFFER = 4   ' Bad buffer (null, invalid addr..)
       SMART_INVALID_DRIVE = 5    ' Drive number not valid
       SMART_INVALID_IOCTL = 6    ' Invalid IOCTL
       SMART_ERROR_NO_MEM = 7     ' Could not lock user's buffer
       SMART_INVALID_REGISTER = 8 ' Some IDE Register not valid
       SMART_NOT_SUPPORTED = 9    ' Invalid cmd flag set
       SMART_NO_IDE_DEVICE = 10   ' Cmd issued to device not present
                                  ' although drive number is valid
       ' 11-255 reserved
End Enum
'---------------------------------------------------------------------
' The following struct defines the interesting part of the IDENTIFY
' buffer:
'---------------------------------------------------------------------
Public Type IDSECTOR
   wGenConfig                 As Integer
   wNumCyls                   As Integer
   wReserved                  As Integer
   wNumHeads                  As Integer
   wBytesPerTrack             As Integer
   wBytesPerSector            As Integer
   wSectorsPerTrack           As Integer
   wVendorUnique(2)           As Integer
   sSerialNumber(19)          As Byte
   wBufferType                As Integer
   wBufferSize                As Integer
   wECCSize                   As Integer
   sFirmwareRev(7)            As Byte
   sModelNumber(39)           As Byte
   wMoreVendorUnique          As Integer
   wDoubleWordIO              As Integer
   wCapabilities              As Integer
   wReserved1                 As Integer
   wPIOTiming                 As Integer
   wDMATiming                 As Integer
   wBS                        As Integer
   wNumCurrentCyls            As Integer
   wNumCurrentHeads           As Integer
   wNumCurrentSectorsPerTrack As Integer
   ulCurrentSectorCapacity    As Long
   wMultSectorStuff           As Integer
   ulTotalAddressableSectors  As Long
   wSingleWordDMA             As Integer
   wMultiWordDMA              As Integer
   bReserved(127)             As Byte
End Type

'---------------------------------------------------------------------
' Structure returned by SMART IOCTL for several commands
'---------------------------------------------------------------------
Public Type SENDCMDOUTPARAMS
  cBufferSize   As Long         ' Size of bBuffer in bytes (IDENTIFY_BUFFER_SIZE in our case)
  DRIVERSTATUS  As DRIVERSTATUS ' Driver status structure.
  bBuffer()    As Byte          ' Buffer of arbitrary length in which to store the data read from the drive.
End Type

'---------------------------------------------------------------------
' Feature register defines for SMART "sub commands"
'---------------------------------------------------------------------

Public Const SMART_READ_ATTRIBUTE_VALUES = &HD0
Public Const SMART_READ_ATTRIBUTE_THRESHOLDS = &HD1
Public Const SMART_ENABLE_DISABLE_ATTRIBUTE_AUTOSAVE = &HD2
Public Const SMART_SAVE_ATTRIBUTE_VALUES = &HD3
Public Const SMART_EXECUTE_OFFLINE_IMMEDIATE = &HD4
' Vendor specific commands:
Public Const SMART_ENABLE_SMART_OPERATIONS = &HD8
Public Const SMART_DISABLE_SMART_OPERATIONS = &HD9
Public Const SMART_RETURN_SMART_STATUS = &HDA

'---------------------------------------------------------------------
' The following structure defines the structure of a Drive Attribute
'---------------------------------------------------------------------

Public Const NUM_ATTRIBUTE_STRUCTS = 30

Public Type DRIVEATTRIBUTE
       bAttrID As Byte         ' Identifies which attribute
       wStatusFlags As Integer 'Integer ' see bit definitions below
       bAttrValue As Byte      ' Current normalized value
       bWorstValue As Byte     ' How bad has it ever been?
       bRawValue(5) As Byte    ' Un-normalized value
       bReserved As Byte       ' ...
End Type
'---------------------------------------------------------------------
' Status Flags Values
'---------------------------------------------------------------------
Public Enum STATUS_FLAGS
       PRE_FAILURE_WARRANTY = &H1
       ON_LINE_COLLECTION = &H2
       PERFORMANCE_ATTRIBUTE = &H4
       ERROR_RATE_ATTRIBUTE = &H8
       EVENT_COUNT_ATTRIBUTE = &H10
       SELF_PRESERVING_ATTRIBUTE = &H20
End Enum

'---------------------------------------------------------------------
' The following structure defines the structure of a Warranty Threshold
' Obsoleted in ATA4!
'---------------------------------------------------------------------

Public Type ATTRTHRESHOLD
       bAttrID As Byte            ' Identifies which attribute
       bWarrantyThreshold As Byte ' Triggering value
       bReserved(9) As Byte       ' ...
End Type

'---------------------------------------------------------------------
' Valid Attribute IDs
'---------------------------------------------------------------------
Public Enum ATTRIBUTE_ID
       ATTR_INVALID = 0
       ATTR_READ_ERROR_RATE = 1
       ATTR_THROUGHPUT_PERF = 2
       ATTR_SPIN_UP_TIME = 3
       ATTR_START_STOP_COUNT = 4
       ATTR_REALLOC_SECTOR_COUNT = 5
       ATTR_READ_CHANNEL_MARGIN = 6
       ATTR_SEEK_ERROR_RATE = 7
       ATTR_SEEK_TIME_PERF = 8
       ATTR_POWER_ON_HRS_COUNT = 9
       ATTR_SPIN_RETRY_COUNT = 10
       ATTR_CALIBRATION_RETRY_COUNT = 11
       ATTR_POWER_CYCLE_COUNT = 12
       ATTR_SOFT_READ_ERROR_RATE = 13
       ATTR_G_SENSE_ERROR_RATE = 191
       ATTR_POWER_OFF_RETRACT_CYCLE = 192
       ATTR_LOAD_UNLOAD_CYCLE_COUNT = 193
       ATTR_TEMPERATURE = 194
       ATTR_REALLOCATION_EVENTS_COUNT = 196
       ATTR_CURRENT_PENDING_SECTOR_COUNT = 197
       ATTR_UNCORRECTABLE_SECTOR_COUNT = 198
       ATTR_ULTRADMA_CRC_ERROR_RATE = 199
       ATTR_WRITE_ERROR_RATE = 200
       ATTR_DISK_SHIFT = 220
       ATTR_G_SENSE_ERROR_RATEII = 221
       ATTR_LOADED_HOURS = 222
       ATTR_LOAD_UNLOAD_RETRY_COUNT = 223
       ATTR_LOAD_FRICTION = 224
       ATTR_LOAD_UNLOAD_CYCLE_COUNTII = 225
       ATTR_LOAD_IN_TIME = 226
       ATTR_TORQUE_AMPLIFICATION_COUNT = 227
       ATTR_POWER_OFF_RETRACT_COUNT = 228
       ATTR_GMR_HEAD_AMPLITUDE = 230
       ATTR_TEMPERATUREII = 231
       ATTR_READ_ERROR_RETRY_RATE = 250
End Enum
