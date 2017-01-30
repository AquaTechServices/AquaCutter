Attribute VB_Name = "SM_Moldule"
Option Explicit

'    pixXValue = twipXValue \ Screen.TwipsPerPixelX
'    pixYValue = twipYValue \ Screen.TwipsPerPixelY

'    twipXValue = pixXValue * Screen.TwipsPerPixelX
'    twipYValue = pixYValue * Screen.TwipsPerPixelY


#If Win32 Then

Public Const LF_FACESIZE = 32
   
Public Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type Bounds
    UpperX As Long
    UpperY As Long
    LowerX As Long
    LowerY As Long
End Type

'Public Type DoubleXY
'    X As Double
'    Y As Double
'End Type

Public Type DynamicData
    Sets() As DoubleXY
End Type

Public Type PointXY
    X As Long
    Y As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum SymbolTypes
    squareSymbol = 0
    circleSymbol = 1
    triangleSymbol = 2
    pointSymbol = 3
End Enum

Public Enum LineTypes    ' Pen Type Styles:
    solidLine = 0
    dashedLine = 1          '  -------
    dottedLine = 2          '  .......
End Enum

Public Type LogFont
     lfHeight As Long
     lfWidth As Long
     lfEscapement As Long
     lfOrientation As Long
     lfWeight As Long
     lfItalic As Byte
     lfUnderline As Byte
     lfStrikeOut As Byte
     lfCharSet As Byte
     lfOutPrecision As Byte
     lfClipPrecision As Byte
     lfQuality As Byte
     lfPitchAndFamily As Byte
     lfFaceName As String * LF_FACESIZE
End Type

    'Dynamic Integer Array
Public Type Dyn_Array
    D_Array() As POINTAPI
    D_Index As Integer
End Type
    
Public Type ShapeArray
    S_Array() As DoubleXY
    S_Rotation As Double
    S_Color As OLE_COLOR
End Type
    
Public Type PointArray
    P_Data As DoubleXY
    P_Color As OLE_COLOR
    P_Symbol As SymbolTypes
End Type

Public Type RangeRings
    RingX As Double
    RingY As Double
    RingCount As Integer
    RingSpacing As Double
    RingColor As OLE_COLOR
End Type

Public Type LineArray
    SOL As DoubleXY
    EOL As DoubleXY
    L_Color As OLE_COLOR
    L_Type As LineTypes
End Type

Public Type XYLine
    XYArray() As DoubleXY
End Type

#End If 'WIN32 Types
Public dl&, savedDC&, gcurrent_index%, glast_index%, DataFlag%
Public Gamma!()
'Public PolyArray() As POINTAPI
'**********************************
'**  Type Definitions:

Public DataPoints%
Public GridStep%
Public PlotWidth&

#If Win32 Then
Public Const DC_ACTIVE = &H1
Public Const DC_ICON = &H4
Public Const DC_TEXT = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const DFC_BUTTON = 4
Public Const DFC_POPUPMENU = 5            'Only Win98/2000 !!
Public Const DFCS_BUTTON3STATE = &H10
Public Const DC_GRADIENT = &H20          'Only Win98/2000 !!


Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000
Public Const DT_EDITCONTROL = &H2000
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_RTLREADING = &H20000
Public Const DT_WORD_ELLIPSIS = &H40000

Public Const SRCPAINT& = &HEE0086
Public Const SRCCOPY& = &HCC0020
Public Const SRCAND& = &H8800C6
Public Const SRCERASE& = &H440328
Public Const SRCINVERT& = &H660046
Public Const TRANSPARENT& = 1
Public Const RGN_COPY& = 5
Public Const RGN_AND& = 1
Public Const RGN_DIFF& = 4
Public Const RGN_XOR& = 3
Public Const RGN_OR& = 2
Public Const BLACK_BRUSH = 4
Public Const BLACK_PEN = 7
Public Const BLACKONWHITE = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3
'  Pen Styles
Public Const PS_SOLID = 0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DOT = 2                     '  .......
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_NULL = 5
Public Const PS_INSIDEFRAME = 6
Public Const PS_USERSTYLE = 7
Public Const PS_ALTERNATE = 8
Public Const PS_STYLE_MASK = &HF
'  Brush Styles
Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Const BS_HOLLOW = BS_NULL
Public Const BS_HATCHED = 2
Public Const BS_PATTERN = 3
Public Const BS_INDEXED = 4
Public Const BS_DIBPATTERN = 5
Public Const BS_DIBPATTERNPT = 6
Public Const BS_PATTERN8X8 = 7
Public Const BS_DIBPATTERN8X8 = 8
Public Const BS_MONOPATTERN = 9
'  Hatch Styles
Public Const HS_HORIZONTAL = 0     '/* ----- */
Public Const HS_VERTICAL = 1       '/* ||||| */
Public Const HS_FDIAGONAL = 2      '/* \\\\\ */
Public Const HS_BDIAGONAL = 3      '/* ///// */
Public Const HS_CROSS = 4          '/* +++++ */
Public Const HS_DIAGCROSS = 5      '/* xxxxx */
' Text Modes
Public Const TEXT_OPAQUE = 2
Public Const TEXT_TRANSPARENT = 1
Public Const GridWidth& = 0 ' pixels
'Public Const GridSpacing = 8 ' pixels

' font crap
'used with fnWeight
Public Const FW_DONTCARE = 0
Public Const FW_THIN = 100
Public Const FW_EXTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_NORMAL = 400
Public Const FW_MEDIUM = 500
Public Const FW_SEMIBOLD = 600
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_HEAVY = 900
Public Const FW_BLACK = FW_HEAVY
Public Const FW_DEMIBOLD = FW_SEMIBOLD
Public Const FW_REGULAR = FW_NORMAL
Public Const FW_ULTRABOLD = FW_EXTRABOLD
Public Const FW_ULTRALIGHT = FW_EXTRALIGHT
'used with fdwCharSet
Public Const ANSI_CHARSET = 0
Public Const DEFAULT_CHARSET = 1
Public Const SYMBOL_CHARSET = 2
Public Const SHIFTJIS_CHARSET = 128
Public Const HANGEUL_CHARSET = 129
Public Const CHINESEBIG5_CHARSET = 136
Public Const OEM_CHARSET = 255
'used with fdwOutputPrecision
Public Const OUT_CHARACTER_PRECIS = 2
Public Const OUT_DEFAULT_PRECIS = 0
Public Const OUT_DEVICE_PRECIS = 5
'used with fdwClipPrecision
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_CHARACTER_PRECIS = 1
Public Const CLIP_STROKE_PRECIS = 2
'used with fdwQuality
Public Const DEFAULT_QUALITY = 0
Public Const DRAFT_QUALITY = 1
Public Const PROOF_QUALITY = 2
'used with fdwPitchAndFamily
Public Const DEFAULT_PITCH = 0
Public Const FIXED_PITCH = 1
Public Const VARIABLE_PITCH = 2
'used with SetBkMode
'Const OPAQUE = 2
'Const TRANSPARENT = 1

Const LOGPIXELSY = 90
Const COLOR_WINDOW = 5
Const Message = "Hello !"

Const DRIVE_REMOVABLE = 2


#End If 'WIN32


'**********************************
'**  Function Declarations:

#If Win32 Then
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function Pie Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, ByVal lpRect As Any, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long)
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long)
Public Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long, pcRect As RECT, ByVal un As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long
Public Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function AngleArc Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Radius As Long, ByVal StartAngle As Single, ByVal SweepAngle As Single) As Boolean
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Public Declare Function CreateCompatibleBitmap& Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long)
Public Declare Function SelectObject& Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long)
Public Declare Function CreateCompatibleDC& Lib "gdi32" (ByVal hDC As Long)
Public Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Public Declare Function CreatePen& Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long)
Public Declare Function GetStockObject& Lib "gdi32" (ByVal nIndex As Long)
Public Declare Function MoveToEx& Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long)
Public Declare Function LineTo& Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long)
Public Declare Function Polyline& Lib "gdi32" (ByVal hDC&, lpPoints As POINTAPI, ByVal nCount&)
Public Declare Function PolylineTo Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function PolyBezier Lib "gdi32.dll" (ByVal hDC As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function SetMapMode& Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long)
Public Declare Function GetClipRgn& Lib "gdi32" (ByVal hDC As Long, ByVal hRegion As Long)
Public Declare Function SaveDC& Lib "gdi32" (ByVal hDC As Long)
Public Declare Function RestoreDC& Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long)
Public Declare Function DeleteDC& Lib "gdi32" (ByVal hDC As Long)
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateSolidBrush& Lib "gdi32" (ByVal crColor As Long)
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long)
Public Declare Function StretchBlt& Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long)
Public Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC%, ByVal i%, ByVal i%, ByVal w&, ByVal i%) As Integer
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As LargeInt, lpTotalNumberOfBytes As LargeInt, lpTotalNumberofFreeBytes As LargeInt) As Long

#End If 'WIN32


