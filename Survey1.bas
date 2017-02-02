Attribute VB_Name = "survey"
' updated: 20 June 2000

'**************************************************
'*   Written by S. McBay on 28 July 1999          *
'*   Includes the following commonly used         *
'*   Survey functions.                            *
'*   ToJulian - Calculate Julian Date             *
'*   TruncTime - Calculates Winfrog Datfile time  *
'*   XY Traverse - Simple X Y Traverse, Returns   *
'*         Traverse X Y                           *
'*   XY Inverse - Simple X Y Inverse, Returns     *
'*         Inverse Distance and Bearing (Azmuth)  *
'**************************************************
'*   Added by J. Richard on 05 Aug 1999           *
'*   Four Point Intersection Calculation          *
'*   FourPtIntersection - Calculates the          *
'*         Intersection Point of two lines        *
'**************************************************
'*   Added By S. McBay on 24 October 1999         *
'*   Real Time Standard Deviation Calculation     *
'*   StdDevRealTime - Calculates standard         *
'*         deviation of current set each time     *
'*         a new value is submitted to function   *
'**************************************************
'*   Added by S. McBay on 22 Feb 2000            *
'*   sm_wait% (WaitTime%)                         *
'*   waits specified time in milliseconds         *
'**************************************************
'*   Added by S. McBay on 13 Feb 2000             *
'*   DegreesToRadians()                           *
'*   converts degrees to radians, returns radians *
'**************************************************
'*   Added by S. McBay on 13 Feb 2000             *
'*   RadiansToDegrees()                           *
'*   converts redians to degrees, returns degrees *
'**************************************************
'*   updated functions traverse and inverse       *
'*   to use the above two functions               *
'**************************************************
'*   Added by S. McBay on 19 April 2000           *
'*   FileUnixToDos(FileName, FileNumber)          *
'*   converts an open file from unix to dos format*
'**************************************************
'*   Added by S. McBay on 29 May 2000             *
'*   LinearRegression(PointArray(),nPoints,Slope, *
'*                        Intercept)              *
'*   calculates slope and Y intercept of the best *
'*   fit line through nPoints number of points    *
'**************************************************
'*   Added by S. McBay on 19 June 2000            *
'*   WindowsDirectory()                           *
'*   returns windows directory                    *
'**************************************************
'*   Added by S. McBay on 19 June 2000            *
'*   SystemDirectory()                            *
'*   returns system directory                     *
'**************************************************
'*   Fixed FileUnixToDos()                        *
'*   it was not working properly                  *
'*   Scott McBay, June 20, 2000                   *
'**************************************************
'*   Added by S. McBay on 20 June 2000            *
'*   sm_parse(InString As String, DynamicArray()) *
'*   takes a string and a reference to a dynamic  *
'*   array and parses the comma delimited values  *
'*   it places them in the "1" indexed array and  *
'*   returns the number of variables              *
'**************************************************
'*   Added by S. McBay on 18 July 2001            *
'*   Check_Sum(CheckString As String) As Boolean  *
'*   usage: Result = Check_Sum(CheckString)       *
'*   returns: True or False                       *
'*   what does it do? Exclusive OR each character *
'*   in string and convert to hex                 *
'**************************************************
'*   Bool2Int(ByVal InBool As Boolean) As Integer *
'*   usage: Result = Bool2Int(InBool)             *
'*   returns: 1 if True, 0 if False               *
'**************************************************
'*   Int2Bool(ByVal InInt As Integer) As Boolean  *
'*   usage: Result = Int2Bool(InInt)              *
'*   returns: True if 1, False if 0               *
'**************************************************
'*   Added by S. McBay on July 30 2001            *
'*   ChkSumVal(InString As String) As String      *
'*   usage: CheckSum = ChkSumVal(InString)        *
'*   returns: string containing hex checksum      *
'**************************************************
'*   Updated Inverse() on October 23 2001         *
'*   InverseBearing was returning values relative *
'*   to the quadrant. Added code to compute       *
'*   quadrant and adjust InverseBearing           *
'**************************************************
'*   GreatCircleDistance(ByVal Lat1 As Double,    *
'*                       ByVal Lon1 As Double,    *
'*                       ByVal Lat2 As Double,    *
'*                       ByVal Lon2 As Double)    *
'*   returns great circle distance                *
'**************************************************
'*   InitialHeading(ByVal Lat1 As Double,         *
'*                  ByVal Lon1 As Double,         *
'*                  ByVal Lat2 As Double,         *
'*                  ByVal Lon2 As Double)         *
'*   compute the initial bearing (in degrees)     *
'**************************************************
'*   BoxCarFilter(ByRef ValueArray As Double)     *
'*   return the average of the number of values   *
'**************************************************

Public Type PointXY
    x As Long
    Y As Long
End Type

Public Type DoubleXY
    x As Double
    Y As Double
End Type

Public Type LargeInt
  lngLower As Long
  lngUpper As Long
End Type

Public NmeaArray()

Public Const earth_radius = 6367000 'meters
Public TraverseX As Double
Public TraverseY As Double
Public InverseBearing As Double
Public InverseDistance As Double
Public FourPtIntersect_Y As Double
Public FourPtIntersect_X As Double
'Public StandardDeviationRT As Double
Private StdDevSetArrayRT() As Double
Public Const PI = 3.14159265358979
'Public Const PI = 4 * Atn(1 / 1)
Public Declare Function GetTickCount& Lib "kernel32" () 'added 000222
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' convert from degrees to radians

Public Function DegreesToRadians(degrees As Double) As Double
    DegreesToRadians = degrees / (180 / PI)
End Function

' convert from radians to degrees

Public Function RadiansToDegrees(radians As Double) As Double
    RadiansToDegrees = radians * (180 / PI)
End Function

' sm_wait (WaitTime As Long) As Long
' ex: RetVal = sm_wait(500)
' returns: non zero upon completion of delay

Public Function sm_wait&(WaitTime&)
Dim ClockTicks&
Dim ClockStart&
ClockStart = GetTickCount
Do While ClockTicks < (ClockStart + WaitTime)
ClockTicks = GetTickCount
DoEvents
Loop
sm_wait = ClockTicks

End Function

' ToJulian()
' arguments: WfgDate - date string from winfrog dat file
' returns: julian day
' ex: JulianDay = ToJulian(07-27-99) = 208

Public Function ToJulian(WfgDate)

WfgDay = Mid$(WfgDate, 4, 2)
WfgMon = Mid$(WfgDate, 1, 2)
WfgYer = Mid$(WfgDate, 7, 2)

'Debug.Print WfgDay
'Debug.Print WfgMon
'Debug.Print WfgYer

If (Val(WfgYer) > 50) Then
    WfgYer = "19" + WfgYer
    WfgYear = WfgYer
Else
    WfgYer = "20" + WfgYer
    WfgYear = WfgYer
End If

' Leap Year Calculation
If (WfgYer Mod 4) < 0 Or (WfgYer Mod 4) > 0 Then
    Leap = 0
ElseIf (WfgYer Mod 400) = 0 Then
    Leap = 1
ElseIf (WfgYer Mod 100) = 0 Then
    Leap = 0
Else
    Leap = 1
End If

If WfgMon = "01" Then
    ToJulian = Val(WfgDay)
ElseIf WfgMon = "02" Then
    ToJulian = Val(Leap) + 31 + Val(WfgDay)
ElseIf WfgMon = "03" Then
    ToJulian = Val(Leap) + 59 + Val(WfgDay)
ElseIf WfgMon = "04" Then
    ToJulian = Val(Leap) + 90 + Val(WfgDay)
ElseIf WfgMon = "05" Then
    ToJulian = Val(Leap) + 120 + Val(WfgDay)
ElseIf WfgMon = "06" Then
    ToJulian = Val(Leap) + 151 + Val(WfgDay)
ElseIf WfgMon = "07" Then
    ToJulian = Val(Leap) + 181 + Val(WfgDay)
ElseIf WfgMon = "08" Then
    ToJulian = Val(Leap) + 212 + Val(WfgDay)
ElseIf WfgMon = "09" Then
    ToJulian = Val(Leap) + 243 + Val(WfgDay)
ElseIf WfgMon = "10" Then
    ToJulian = Val(Leap) + 273 + Val(WfgDay)
ElseIf WfgMon = "11" Then
    ToJulian = Val(Leap) + 304 + Val(WfgDay)
ElseIf WfgMon = "12" Then
    ToJulian = Val(Leap) + 334 + Val(WfgDay)
End If

End Function

' TruncTime()
' arguments: WfgTime - time string from winfrog dat file
' returns: time without delimiters (hrmnscd)
' ex: FixTime = TruncTime(01:00:31.4) = 0100314

Public Function TruncTime(WfgTime)
Max = Len(WfgTime)
Cnt = 1
NewTime = ""

While Cnt <= Max
    If Mid$(WfgTime, Cnt, 1) = ":" Then
        Cnt = Cnt + 1
    ElseIf Mid$(WfgTime, Cnt, 1) = "." Then
        Cnt = Cnt + 1
    Else
        NewTime = NewTime + Mid$(WfgTime, Cnt, 1)
        Cnt = Cnt + 1
    End If
Wend
TruncTime = NewTime

End Function

' Traverse()
' arguments: StartX - starting X coordinate for traverse
'            StartY - starting Y coordinate for traverse
'            Bearing - bearing of traverse
'            Distance - distance of traverse
' returns: assigns values to Public variables TraverseX and TraverseY
' ex: Call Traverse(StartX, StartY, Bearing, Distance)

Public Sub Traverse(StartX As Double, StartY As Double, Bearing As Double, distance As Double)
Dim dy As Double
Dim dx As Double

dy = (Cos(DegreesToRadians(Bearing))) * distance
dx = (Sin(DegreesToRadians(Bearing))) * distance
TraverseX = StartX + dx
TraverseY = StartY + dy

End Sub

' Inverse()
' arguments: StartX - starting X coordinate of inverse
'            StartY - starting Y coordinate of inverse
'            EndX - ending X of inverse
'            EndY - ending Y of inverse
' returns: assigns values to Public variables InverseBearing and InverseDistance
' ex: Call Inverse(StartX, StartY, EndX, EndY)
Public Sub Inverse(StartX As Double, StartY As Double, EndX As Double, EndY As Double)
Dim DiffX As Double
Dim DiffY As Double
If StartX = 0 Then
    StartX = StartX + 0.000000001
End If
If StartY = 0 Then
    StartY = StartY + 0.000000001
End If
If EndX = 0 Then
    EndX = EndX + 0.000000001
End If
If EndY = 0 Then
    EndY = EndY + 0.000000001
End If

DiffX = EndX - StartX
DiffY = EndY - StartY
If DiffX = 0 Then
    DiffX = DiffX + 0.000000001
End If
If DiffY = 0 Then
    DiffY = DiffY + 0.000000001
End If


InverseDistance = Sqr(((DiffY) * (DiffY)) + ((DiffX) * (DiffX)))
InverseBearing = RadiansToDegrees((Atn(CDbl((DiffY) / (DiffX)))))

' put InverseBearing between 0 and 360
If InverseBearing < 0 Then
    InverseBearing = InverseBearing + 360
ElseIf InverseBearing > 360 Then
    InverseBearing = InverseBearing - 360
End If

' find quadrant and adjust bearing
If DiffX > 0 And DiffY > 0 Then
    InverseBearing = 90 - InverseBearing
ElseIf DiffX > 0 And DiffY < 0 Then
    InverseBearing = 180 - (InverseBearing - 270)
ElseIf DiffX < 0 And DiffY < 0 Then
    InverseBearing = 270 - InverseBearing
ElseIf DiffX < 0 And DiffY > 0 Then
    InverseBearing = 360 - (InverseBearing - 270)
End If

End Sub

' FourPointIntersection()
' arguments: Start1_Y - Starting Y on 1st Line
'            Start1_X - Starting X on 1st Line
'            End1_Y - Ending Y on 1st Line
'            End1_X - Ending X on 1st Line
'            Start2_Y - Starting Y on 2nd Line
'            Start2_X - Starting X on 2nd Line
'            End2_Y - Ending Y on 2nd Line
'            End2_X - Ending X on 2nd Line
' Returns:  Assigns values to Public Variables FourPtIntersect_Y and FourPtIntersect_X
' Ex:  Call FourPointIntersection(Start1_Y,Start1_X,End1_Y,End1_X,Start2_Y,Start2_X,End2_Y,End2_X)

Public Sub FourPtIntersection(YA As Double, XA As Double, YB As Double, XB As Double, YC As Double, XC As Double, YD As Double, XD As Double)
Static RValue As Double
a = (YA - YC) * (XD - XC)
b = (XA - XC) * (YD - YC)
C = (XB - XA) * (YD - YC)
d = (YB - YA) * (XD - XC)
RValue = (a - b) / (C - d)
'Debug.Print RValue
FourPtIntersect_Y = YA + RValue * (YB - YA)
'Debug.Print Intersect_Y
FourPtIntersect_X = XA + RValue * (XB - XA)
'Debug.Print Intersect_X
End Sub

' StdDevRealTime()
' arguments: NewValue - New Value in set
'            Reset = 1 (optional) - Resets NumValues to 0 to restart calculation
' Returns: Calculates standard deviation of current set and returns standard deviation
' Ex: Call StdDevRealTime(NewValue [,1])

Public Function StdDevRealTime(NewValue As Double, Optional reset As Integer) As Double

' variable declaration
Static NumValues As Integer
Dim ArithmeticMean As Double
Dim MeanOfSquaresOfResiduals

' variable initialization
Cnt = 0
TotalValue = 0

' dynamically increase array size
ReDim Preserve StdDevSetArrayRT(3, NumValues)

' check for calculation reset
If reset = 1 Then
    NumValues = 0
End If

' store NewValue in array
StdDevSetArrayRT(0, NumValues) = NewValue

' calculate standard deviation
NumValues = NumValues + 1
While Cnt < NumValues
    TotalValue = TotalValue + StdDevSetArrayRT(0, Cnt)
    Cnt = Cnt + 1
Wend
ArithmeticMean = TotalValue / NumValues
Cnt = 0
While Cnt < NumValues
    StdDevSetArrayRT(1, Cnt) = StdDevSetArrayRT(0, Cnt) - ArithmeticMean
    StdDevSetArrayRT(2, Cnt) = StdDevSetArrayRT(1, Cnt) * StdDevSetArrayRT(1, Cnt)
    Cnt = Cnt + 1
Wend
Cnt = 0
While Cnt < NumValues
    TotalSquaresOfResiduals = TotalSquaresOfResiduals + StdDevSetArrayRT(2, Cnt)
    Cnt = Cnt + 1
Wend
MeanOfSquaresOfResiduals = TotalSquaresOfResiduals / NumValues

' return current standard deviation
StdDevRealTime = Sqr(MeanOfSquaresOfResiduals)

End Function

Public Function StandardDeviation(ByRef arr() As Double) As Double

'standard deviation
Dim Sum As Double
Dim sumSquare As Double
Dim Value As Double
Dim count As Long
Dim Index As Long
Sum = 0
sumSquare = 0
Value = 0
count = 0
Index = 0

' evaluate sum of values
For Index = LBound(arr) To UBound(arr)
    Value = arr(Index)
    count = count + 1
    Sum = Sum + Value
    sumSquare = sumSquare + Value * Value
Next

If ((Sum * Sum / count)) > sumSquare Then
    StandardDeviation = 0
Else
    StandardDeviation = Sqr((sumSquare - (Sum * Sum / count)) / count)
End If

'StandardDeviation = Sqr((sumSquare - (sum * sum / count)) / count)

End Function

' VarianceRealTime()
' arguments: NewValue - New Value in set
'            Reset = 1 (optional) - Resets NumValues to 0 to restart calculation
' Returns: Calculates standard deviation of current set and returns standard deviation
' Ex: Call VarianceRealTime(NewValue [,1])

Public Function VarianceRealTime(NewValue As Double, Optional reset As Integer) As Double

' variable declaration
Static NumValues As Integer
Dim ArithmeticMean As Double
Dim MeanOfSquaresOfResiduals

' variable initialization
Cnt = 0
TotalValue = 0

' dynamically increase array size
ReDim Preserve StdDevSetArrayRT(3, NumValues)

' check for calculation reset
If reset = 1 Then
    NumValues = 0
End If

' store NewValue in array
StdDevSetArrayRT(0, NumValues) = NewValue

' calculate standard deviation
NumValues = NumValues + 1
While Cnt < NumValues
    DoEvents
    TotalValue = TotalValue + StdDevSetArrayRT(0, Cnt)
    Cnt = Cnt + 1
Wend
ArithmeticMean = TotalValue / NumValues
Cnt = 0
While Cnt < NumValues
    DoEvents
    StdDevSetArrayRT(1, Cnt) = StdDevSetArrayRT(0, Cnt) - ArithmeticMean
    StdDevSetArrayRT(2, Cnt) = StdDevSetArrayRT(1, Cnt) * StdDevSetArrayRT(1, Cnt)
    Cnt = Cnt + 1
Wend
Cnt = 0
While Cnt < NumValues
    DoEvents
    TotalSquaresOfResiduals = TotalSquaresOfResiduals + StdDevSetArrayRT(2, Cnt)
    Cnt = Cnt + 1
Wend
MeanOfSquaresOfResiduals = TotalSquaresOfResiduals / NumValues

' return current standard deviation
VarianceRealTime = MeanOfSquaresOfResiduals

End Function

' Variance()
' arguments: NewValues() - Array of new values
'            NumValues   - Long containing the number of values
' Returns: Calculates variance of current set and returns
' Ex: Call VarianceRealTime(NewValues, NumValues)

Public Function Variance(NewValues() As Double, NumValues As Long) As Double

' variable declaration
'Static NumValues As Integer
Dim ArithmeticMean As Double
Dim MeanOfSquaresOfResiduals

' variable initialization
Cnt = 0
TotalValue = 0

' dynamically increase array size
ReDim Preserve StdDevSetArrayRT(3, NumValues)

While Cnt < NumValues
    DoEvents
    StdDevSetArrayRT(0, Cnt) = NewValues(Cnt)
    Cnt = Cnt + 1
Wend

Cnt = 0
While Cnt < NumValues
    DoEvents
    TotalValue = TotalValue + StdDevSetArrayRT(0, Cnt)
    Cnt = Cnt + 1
Wend
ArithmeticMean = TotalValue / NumValues

Cnt = 0
While Cnt < NumValues
    DoEvents
    StdDevSetArrayRT(1, Cnt) = StdDevSetArrayRT(0, Cnt) - ArithmeticMean
    StdDevSetArrayRT(2, Cnt) = StdDevSetArrayRT(1, Cnt) * StdDevSetArrayRT(1, Cnt)
    Cnt = Cnt + 1
Wend

Cnt = 0
While Cnt < NumValues
    DoEvents
    TotalSquaresOfResiduals = TotalSquaresOfResiduals + StdDevSetArrayRT(2, Cnt)
    Cnt = Cnt + 1
Wend
MeanOfSquaresOfResiduals = TotalSquaresOfResiduals / NumValues

' return current variance
Variance = MeanOfSquaresOfResiduals

End Function

Public Function HarmonicMean(ByRef DataArray() As Double, ByVal nPoints As Integer) As Double
    Dim datacount As Integer
    Dim reciprocals As Double
    For datacount = 0 To (nPoints - 1)
        reciprocals = reciprocals + (1 / DataArray(datacount))
    Next datacount
    HarmonicMean = nPoints / reciprocals
End Function

Public Function Median(ByRef DataArray() As Double) As Double
Dim TempArray() As Double
ReDim TempArray(UBound(DataArray))
TempArray = DataArray
'QuickSort (TempArray)

If (UBound(TempArray) Mod 2) = 0 Then
    ' odd number of items
    Median = TempArray(UBound(TempArray) \ 2)
Else
    ' even number of items
    Median = (TempArray(UBound(TempArray) \ 2) + TempArray(1 + UBound(TempArray) \ 2)) / 2
End If
End Function

Public Function FileUnixToDos(ByVal FileName As String, ByVal FileNumber As Integer) As Integer
Cnt = 0
Max = FileLen(FileName)
TempString = ""
FileUnixToDos = -1 ' assume failure
If Max = 0 Then
    Exit Function
End If
TempNumber = FreeFile(0)
Open "dosfile.tmp" For Output As #TempNumber
While EOF(FileNumber) = False
    LineChar = Input(1, #FileNumber)
    If LineChar = Chr$(10) Then
        Print #TempNumber, TempString
        TempString = ""
    ElseIf LineChar = Chr$(13) Then
        'do nothing
    Else
        TempString = TempString + LineChar
    End If
Wend
Close #FileNumber
Close #TempNumber
FileCopy "dosfile.tmp", FileName
Kill "dosfile.tmp"
Open FileName For Input As #FileNumber
FileUnixToDos = FileNumber ' return success
End Function

Public Function LinearRegression(ByRef PointArray() As DoubleXY, ByVal nPoints As Integer, ByRef Slope As Double, ByRef Intercept As Double)

' linear regression to determine the best fit line for a given set of
' points. solves for the slope and the Y intercept. given the equation
' of a line y = m(slope)x + b(y-intercept), any point along the line can
' be calculated.

Dim sum_X As Double
Dim avg_X As Double ' (x1 + x2 + ...xN) / N
Dim sum_Y As Double
Dim avg_Y As Double ' (y1 + y2 + ...yN) / N
Dim sum_X_squares As Double
Dim avg_X_squares As Double ' ((x1)ª + (x2)ª + ...(xN)ª) / N
Dim sum_XY As Double
Dim avg_XY As Double ' ((x1 * y1) + (x2 * y2) + ...(xN * yN)) / N
Dim Cnt As Integer

LinearRegression = -1 ' assume failure

sum_X = 0
sum_Y = 0
sum_X_squares = 0
sum_XY = 0

For Cnt = 0 To (nPoints - 1)
    sum_X = sum_X + PointArray(Cnt).x
    sum_Y = sum_Y + PointArray(Cnt).Y
    sum_X_squares = sum_X_squares + (PointArray(Cnt).x * PointArray(Cnt).x)
    sum_XY = sum_XY + (PointArray(Cnt).x * PointArray(Cnt).Y)
Next Cnt

avg_X = sum_X / nPoints
avg_Y = sum_Y / nPoints
avg_X_squares = sum_X_squares / nPoints
avg_XY = sum_XY / nPoints

Slope = (avg_XY - (avg_X * avg_Y)) / (avg_X_squares - (avg_X * avg_X))
Intercept = avg_Y - (Slope * avg_X)

LinearRegression = Cnt + 1 ' return should equal nPoints

End Function
        
Public Function WindowsDirectory() As String
    Dim WinPath As String
    WinPath = String(255, Chr(0))
    WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))
End Function

Public Function SystemDirectory() As String
    Dim SysPath As String
    SysPath = String(255, Chr(0))
    SystemDirectory = Left(SysPath, GetSystemDirectory(SysPath, Len(SysPath)))
End Function

Public Function sm_parse(ByVal InString As String, Delimiter As String, ByRef DynArray()) As Integer
' returns: elements in the array (tokens?)
    Dim Max As Integer
    Dim LoopCnt As Integer
    Dim ArrayCnt As Integer
    Dim DCnt As Integer
    
    Max = Len(InString)
    ArrayCnt = 1
    DCnt = 0
    ReDim DynArray(ArrayCnt)
    
    If Delimiter = " " Then
        For LoopCnt = 1 To Max
            InChar = Mid$(InString, LoopCnt, 1)
            If InChar = Delimiter Then
                If DCnt = 0 Then
                    ArrayCnt = ArrayCnt + 1
                    ReDim Preserve DynArray(ArrayCnt)
                    DCnt = DCnt + 1
                End If
            Else
                DynArray(ArrayCnt) = DynArray(ArrayCnt) + InChar
                DCnt = 0
            End If
        Next LoopCnt
    Else
        For LoopCnt = 1 To Max
            InChar = Mid$(InString, LoopCnt, 1)
            If InChar = Delimiter Then
                ArrayCnt = ArrayCnt + 1
                ReDim Preserve DynArray(ArrayCnt)
            Else
                DynArray(ArrayCnt) = DynArray(ArrayCnt) + InChar
    '            Debug.Print DynArray(ArrayCnt)
            End If
        Next LoopCnt
    End If
    sm_parse = ArrayCnt
End Function

Public Function Check_Sum(CheckString As String) As Boolean
' usage: Result = Check_Sum(CheckString)
' returns: True or False
' what does it do? Exclusive OR each character in string, convert to hex
Dim Cnt As Integer
Dim Sum As Integer
Dim SumCheck As String
Dim CheckChar As String
Cnt = 2 ' skip over ($) and start with GPGGA,XXX...
Sum = 0
Max = Len(CheckString) - 3 ' leave off (*HH) checksum
While Cnt <= Max
    CheckChar = Mid$(CheckString, Cnt, 1)
    Sum = Sum Xor Asc(Mid$(CheckString, Cnt, 1))
    Cnt = Cnt + 1
Wend
SumCheck = Hex(Sum)
' verify calculated checksum against observed checksum
If Val(SumCheck) = Val(Mid$(CheckString, Max + 2, 2)) Then
    Check_Sum = True
Else
    Check_Sum = False
End If
End Function

Public Function ChkSumVal(InString As String) As String
' Added by S. McBay on July 30 2001
' ChkSumVal(InString As String) As String
' usage: CheckSum = ChkSumVal(InString)
' returns: string containing hex checksum
' what does it do? Exclusive OR each character in string, convert to hex
Dim Cnt As Integer
Dim Sum As Integer
Dim SumCheck As String
Dim CheckChar As String
Cnt = 1
Sum = 0
Max = Len(InString)
While Cnt <= Max
    CheckChar = Mid$(InString, Cnt, 1)
    Sum = Sum Xor Asc(Mid$(InString, Cnt, 1))
    Cnt = Cnt + 1
Wend
ChkSumVal = CStr(Hex(Sum))
End Function

Public Function WebChkSum(InString As String) As String
' Added by S. McBay on February 19 2013
' WebChkSum(InString As String) As String
' usage: CheckSum = WebChkSum(InString)
' returns: string containing hex checksum
' Adds the ASCII value of each character, convert to hex, return 2 least significant digits
Dim Cnt As Integer
Dim Sum As Integer
Dim Max As Integer
Dim CheckChar As String
Dim HexSum As String
Cnt = 1
Sum = 0
Max = Len(InString)
While Cnt <= Max
    CheckChar = Mid$(InString, Cnt, 1)
    Sum = Sum + Asc(CheckChar)
    Cnt = Cnt + 1
Wend
Max = Len(Str(Sum))
If Max > 2 Then
    WebChkSum = Mid$(Hex(Sum), Max - 2, 2)
Else
    WebChkSum = Hex(Sum)
End If

End Function

Public Function Bool2Int(ByVal InBool As Boolean) As Integer
    If InBool = True Then
        Bool2Int = 1
    ElseIf InBool = False Then
        Bool2Int = 0
    Else
        Bool2Int = -1
    End If
End Function

Public Function Int2Bool(ByVal InInt As Integer) As Boolean
    If InInt = 0 Then
        Int2Bool = False
    ElseIf InInt > 0 Then
        Int2Bool = True
    Else
        Int2Bool = False
    End If
End Function

Public Function BaseConvert(NumIn As String, BaseIn As Byte, BaseOut As Byte) As String
   'Binary       = Base 2
   'Octal       = Base 8
   'Decimal     = Base 10
   'Hexadecimal = Base 16

   Dim i As Integer, CurrentCharacter As String, CharacterValue As Integer
   Dim PlaceValue As Integer, RunningTotal As Double, Remainder As Double
   Dim BaseOutDouble As Double, NumInCaps As String

   If NumIn = "" Or BaseIn < 2 Or BaseIn > 36 Or BaseOut < 1 Or BaseOut > 36 Then
      BaseConvert = "Error"
      Exit Function
   End If

   NumInCaps = UCase(NumIn)

   PlaceValue = Len(NumInCaps)

   For i = 1 To Len(NumInCaps)
      PlaceValue = PlaceValue - 1
      CurrentCharacter = Mid$(NumInCaps, i, 1)
      CharacterValue = 0
      If Asc(CurrentCharacter) > 64 And Asc(CurrentCharacter) < 91 Then
         CharacterValue = Asc(CurrentCharacter) - 55
      End If

      If CharacterValue = 0 Then
         If Asc(CurrentCharacter) < 48 Or Asc(CurrentCharacter) > 57 Then
            BaseConvert = "Error"
            Exit Function
         Else
            CharacterValue = Val(CurrentCharacter)
         End If
      End If

      If CharacterValue < 0 Or CharacterValue > BaseIn - 1 Then
         BaseConvert = "Error"
         Exit Function
      End If
      RunningTotal = RunningTotal + CharacterValue * (BaseIn ^ PlaceValue)
   Next i

   Do
      BaseOutDouble = CDbl(BaseOut)
      Remainder = RunningTotal - (Int(RunningTotal / BaseOutDouble) * BaseOutDouble)
      RunningTotal = (RunningTotal - Remainder) / BaseOut

      If Remainder >= 10 Then
         CurrentCharacter = Chr$(Remainder + 55)
      Else
         CurrentCharacter = right$(Str$(Remainder), Len(Str$(Remainder)) - 1)
      End If
      BaseConvert = CurrentCharacter & BaseConvert
   Loop While RunningTotal > 0
End Function

Public Function FileSize(File As String) As String
Dim LSize As String
If File = "" Then
    FileSize = ""
    Exit Function
End If
LSize = FileLen(File)
FileSize = LSize 'Size in bytes
End Function

Public Function GreatCircleDistance(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As Double
    ' approx radius of Earth in meters.  True radius varies from
    ' 6357km (polar) to 6378km (equatorial).
    Dim dlon As Double
    Dim dlat As Double
    Dim d As Double
    Dim a As Double
    
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = (Sin(dlat / 2)) ^ 2 + Cos(lat1) * Cos(lat2) * (Sin(dlon / 2)) ^ 2
    d = 2 * Atn(Sqr(a) / Sqr(1 - a))

    ' This is a simpler formula, but it's subject to rounding errors
    ' for small distances.  See http://www.census.gov/cgi-bin/geo/gisfaq?Q5.1
    ' d = acos(sin(Lat1) * sin(Lat2) + cos(Lat1) * cos(Lat2) * cos(Lon1-Lon2))
    
    GreatCircleDistance = earth_radius * d ' earth_radius defined above as earth_radius = 6367000 meters

End Function

' compute the initial bearing (in degrees) to get from lat1/long1 to lat2/long2
Public Function InitialHeading(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As Double
    ' note that this is the same d calculation as above.
    ' duplicated for clarity.
    Dim dlon As Double
    Dim dlat As Double
    Dim d As Double
    Dim a As Double
    Dim Heading As Double
    
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = (Sin(dlat / 2)) ^ 2 + Cos(lat1) * Cos(lat2) * (Sin(dlon / 2)) ^ 2
    d = 2 * Atn(Sqr(a) / Sqr(1 - a))
    
    Heading = Arcos((Sin(lat2) - Sin(lat1) * Cos(d)) / (Sin(d) * Cos(lat1)))
    If (Sin(lon2 - lon1) < 0) Then
        Heading = 2 * PI - Heading ' PI defined above as PI = 4 * Atn(1 / 1)
    End If
    
    InitialHeading = RadiansToDegrees(Heading)

End Function

Public Function Arcos(ByVal x As Double) As Double
    Arcos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

Public Function BoxCarFilter(ByRef ValueArray() As Double)
    Dim Min As Integer
    Dim Max As Integer
    Dim Sum As Double
    Min = LBound(ValueArray)
    Max = UBound(ValueArray)
    For count = Min To Max
        Sum = Sum + ValueArray(count)
    Next
    BoxCarFilter = Sum / (Max - Min)
End Function

Public Function makeUnixTime(ByVal myTime As Date) As Long
' to make a current stamp--> TimeStamp = makeUnixTime(Now)
makeUnixTime = (myTime - DateSerial(1970, 1, 1)) * 86400
End Function

Public Function getUnixTime(ByVal unixTime As Long) As Date
getUnixTime = CDate(unixTime / 86400 + DateSerial(1970, 1, 1))
End Function

Public Function encryptUnixTime(ByVal unixTime As Long) As Long
Dim timeString As String
Dim outString As String
Dim count As Integer
count = 1
timeString = CStr(unixTime)
While count < Len(timeString)
    Select Case Mid$(timeString, count, 1)
    Case "1"
        outString = outString + "9"
    Case "2"
        outString = outString + "8"
    Case "3"
        outString = outString + "7"
    Case "4"
        outString = outString + "6"
    Case "5"
        outString = outString + "5"
    Case "6"
        outString = outString + "4"
    Case "7"
        outString = outString + "3"
    Case "8"
        outString = outString + "2"
    Case "9"
        outString = outString + "1"
    Case Else
        outString = outString + Mid$(timeString, count, 1)
    End Select
        
    count = count + 1
Wend
'If outString <> "" Then
'    encryptUnixTime = CLng(outString)
'End If

End Function

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
        Call SaveSetting("MyApp", "ou812", "StartTime", curSeconds)
        Call MsgBox(Format((endTime - curSeconds) / 86400, 0) & " Days left on current registration code.", vbOKOnly, "Registration Check")
        checkSoftwareTime = True
    Else
        'shut down software
        checkSoftwareTime = False
    End If
End If
End Function

Public Function parseNmea(ByVal InString As String) As Boolean
Dim Max As Integer
Dim temp As Integer
Dim i As Integer
Dim NmeaID As String
Dim TempTime As Long
Dim TempDiff As Double
Dim SatCount As Integer
Dim MessageCount As Integer
Dim BeginMessage As Integer
Dim EndMessage As Integer
Dim TempArray()
Dim BitErrArray()
Static GsvSats As Integer
Dim errmsg As String
'On Error GoTo ErrorHandler

If InStr(1, InString, "$", vbTextCompare) <= 0 Then ' probably an invalid BAUD rate selected
    parseNmea = False
    Exit Function
End If

InString = Mid$(InString, InStr(1, InString, "$", vbTextCompare)) ' Strip-off any data before the dollar character

If Check_Sum(InString) Or DeviceType = Asci Then ' device EITHER with OR without CheckSum validation
    Call sm_parse(InString, "*", TempArray) ' parse out checksum if present (ie; strip off * etc...)
    Call sm_parse(TempArray(1), ",", NmeaArray)
    Max = UBound(NmeaArray)
    NmeaID = right$(NmeaArray(1), Len(NmeaArray(1)) - 3) ' skip over $xx characters
    parseNmea = False
    Select Case NmeaID
    Case "GGA"
        If Max = 15 Then
            If IsNumeric(NmeaArray(2)) Then
                NmeaInfo.gga.Utc = NmeaArray(2)
            End If
            If IsNumeric(NmeaArray(3)) Then
                NmeaInfo.gga.lat = NmeaArray(3)
            End If
            NmeaInfo.gga.LatHemi = NmeaArray(4)
            If IsNumeric(NmeaArray(5)) Then
                NmeaInfo.gga.lon = NmeaArray(5)
            End If
            NmeaInfo.gga.LonHemi = NmeaArray(6)
            If IsNumeric(NmeaArray(7)) Then
                NmeaInfo.gga.Quality = NmeaArray(7)
            End If
            If IsNumeric(NmeaArray(8)) Then
                NmeaInfo.gga.SatsUsed = NmeaArray(8)
            End If
            If IsNumeric(NmeaArray(9)) Then
                NmeaInfo.gga.hdop = NmeaArray(9)
            End If
            If IsNumeric(NmeaArray(10)) Then
                NmeaInfo.gga.Altitude = NmeaArray(10)
            End If
            NmeaInfo.gga.AltUnit = NmeaArray(11)
            If IsNumeric(NmeaArray(12)) Then
                NmeaInfo.gga.GeoSep = NmeaArray(12)
            End If
            NmeaInfo.gga.GeoSepUnit = NmeaArray(13)
            If IsNumeric(NmeaArray(14)) Then
                NmeaInfo.gga.diffage = NmeaArray(14)
            Else
                NmeaInfo.gga.diffage = -1
            End If
            If IsNumeric(NmeaArray(15)) Then
                NmeaInfo.gga.StationID = NmeaArray(15)
            Else
                NmeaInfo.gga.StationID = -1
            End If
            NmeaInfo.gga.LastUpdate = GetTickCount
            parseNmea = True
        Else
            Debug.Print "ERROR Decoding GGA " + InString + " Expected 15 but got " + Str(Max)
        End If
        Exit Function
    Case "GLL"
        If Max = 7 Or Max = 8 Then
            If IsNumeric(NmeaArray(2)) Then
                NmeaInfo.gll.lat = NmeaArray(2)
            End If
            NmeaInfo.gll.LatHemi = NmeaArray(3)
            If IsNumeric(NmeaArray(4)) Then
                NmeaInfo.gll.lon = NmeaArray(4)
            End If
            NmeaInfo.gll.LonHemi = NmeaArray(5)
            If IsNumeric(NmeaArray(6)) Then
                NmeaInfo.gll.Utc = NmeaArray(6)
            End If
            NmeaInfo.gll.Status = NmeaArray(7)
            ' for NMEA-0183 version 3.0
            If Max = 8 Then
                NmeaInfo.gll.Mode = NmeaArray(8)
            End If
            NmeaInfo.gll.LastUpdate = GetTickCount
            parseNmea = True
        Else
            Debug.Print "ERROR Decoding GLL " + InString + " Expected 7 or 8 but got " + Str(Max)
        End If
        Exit Function
    Case "GSA"
        If Max = 18 Then
            NmeaInfo.gsa.Mode = NmeaArray(2)
            NmeaInfo.gsa.ModeStat = NmeaArray(3)
            For SatCount = 4 To Max - 3
                NmeaInfo.gsa.SatID(SatCount - 3) = NmeaArray(SatCount)
            Next SatCount
            If IsNumeric(NmeaArray(16)) Then
                NmeaInfo.gsa.pdop = NmeaArray(16)
            End If
            If IsNumeric(NmeaArray(17)) Then
                NmeaInfo.gsa.hdop = NmeaArray(17)
            End If
            If IsNumeric(NmeaArray(18)) Then
                NmeaInfo.gsa.vdop = NmeaArray(18)
            End If
            NmeaInfo.gsa.LastUpdate = GetTickCount
            parseNmea = True
        Else
            Debug.Print "ERROR Decoding GSA " + InString + " Expected 18 but got " + Str(Max)
        End If
        Exit Function
    Case "GST"
        If Max = 9 Then
            If IsNumeric(NmeaArray(2)) Then
                NmeaInfo.gst.Utc = NmeaArray(2)
            Else
                Exit Function
            End If
            If IsNumeric(NmeaArray(3)) Then
                NmeaInfo.gst.Rms = NmeaArray(3)
            Else
                Exit Function
            End If
            If IsNumeric(NmeaArray(4)) Then
                NmeaInfo.gst.SDSemiMajor = NmeaArray(4)
            End If
            If IsNumeric(NmeaArray(5)) Then
                NmeaInfo.gst.SDSemiMinor = NmeaArray(5)
            End If
            If IsNumeric(NmeaArray(6)) Then
                NmeaInfo.gst.OrientSemiMajor = NmeaArray(6)
            End If
            If IsNumeric(NmeaArray(7)) Then
                NmeaInfo.gst.SDLat = NmeaArray(7)
            End If
            If IsNumeric(NmeaArray(8)) Then
                NmeaInfo.gst.SDLon = NmeaArray(8)
            End If
            If IsNumeric(NmeaArray(9)) Then
                NmeaInfo.gst.SDAlt = NmeaArray(9)
            End If
            NmeaInfo.gst.LastUpdate = GetTickCount
            parseNmea = True
        Else
            Debug.Print "ERROR Decoding GST " + InString + " Expected 9 but got " + Str(Max)
        End If
        Exit Function
    Case "GSV"
        If IsNumeric(NmeaArray(2)) Then
            NmeaInfo.gsv.MessageMax = NmeaArray(2)
        End If
        If IsNumeric(NmeaArray(3)) Then
            NmeaInfo.gsv.MessageNum = NmeaArray(3)
        End If
        If IsNumeric(NmeaArray(4)) Then
            NmeaInfo.gsv.SatsInView = NmeaArray(4)
        End If
        If NmeaInfo.gsv.MessageNum = 1 Then
            GsvSats = 0
        End If
        SatCount = 0
        For MessageCount = 1 To (Max - 5) / 4
            If IsNumeric(NmeaArray((MessageCount * 4) + 1)) Then
                NmeaInfo.gsv.SatID(GsvSats) = NmeaArray((MessageCount * 4) + 1)
            End If
            If IsNumeric(NmeaArray(1 + ((MessageCount * 4) + 1))) Then
                NmeaInfo.gsv.SatElev(GsvSats) = NmeaArray(1 + ((MessageCount * 4) + 1))
            End If
            If IsNumeric(NmeaArray(2 + ((MessageCount * 4) + 1))) Then
                NmeaInfo.gsv.SatAzimuth(GsvSats) = NmeaArray(2 + ((MessageCount * 4) + 1))
            End If
            If IsNumeric(NmeaArray(3 + ((MessageCount * 4) + 1))) Then
                NmeaInfo.gsv.SatSNR(GsvSats) = NmeaArray(3 + ((MessageCount * 4) + 1))
            End If
            GsvSats = GsvSats + 1
        Next MessageCount
        If NmeaInfo.gsv.MessageMax = NmeaInfo.gsv.MessageNum Then
            NmeaInfo.gsv.LastUpdate = GetTickCount
        End If
        parseNmea = True
        Exit Function
    Case "CTR"  ' these are the C-Nav2000 messages following the $PN prefix
        If NmeaArray(2) Like "SATS" Then
            If IsNumeric(NmeaArray(3)) Then
                NmeaInfo.sats.MessageMax = NmeaArray(3)
            End If
            If IsNumeric(NmeaArray(4)) Then
                NmeaInfo.sats.MessageNum = NmeaArray(4)
            End If
            If NmeaInfo.sats.MessageNum = 1 Then
                NmeaInfo.sats.SatsInView = 0
            End If
            For MessageCount = 1 To (Max - 5) / 5
                If IsNumeric(NmeaArray(MessageCount * 5)) Then
                    NmeaInfo.sats.SatID(NmeaInfo.sats.SatsInView) = NmeaArray(MessageCount * 5)
                End If
                If IsNumeric(NmeaArray(1 + (MessageCount * 5))) Then
                    NmeaInfo.sats.SatElev(NmeaInfo.sats.SatsInView) = NmeaArray(1 + (MessageCount * 5))
                End If
                If IsNumeric(NmeaArray(2 + (MessageCount * 5))) Then
                    NmeaInfo.sats.SatAzimuth(NmeaInfo.sats.SatsInView) = NmeaArray(2 + (MessageCount * 5))
                End If
                If IsNumeric(NmeaArray(3 + (MessageCount * 5))) Then
                    NmeaInfo.sats.SatL1SNR(NmeaInfo.sats.SatsInView) = NmeaArray(3 + (MessageCount * 5))
                End If
                If IsNumeric(NmeaArray(4 + (MessageCount * 5))) Then
                    NmeaInfo.sats.SatL2SNR(NmeaInfo.sats.SatsInView) = NmeaArray(4 + (MessageCount * 5))
                End If
                NmeaInfo.sats.SatsInView = NmeaInfo.sats.SatsInView + 1
            Next MessageCount
            If NmeaInfo.sats.MessageMax = NmeaInfo.sats.MessageNum Then
                NmeaInfo.sats.LastUpdate = GetTickCount
            End If
            parseNmea = True
            Exit Function
        End If
        If NmeaArray(2) Like "RXQ" Then
            '$PNCTR,RXQ,123519,Y,9.6,54,0* 78
            If Max > 3 Then
                If IsNumeric(NmeaArray(3)) Then
                    NmeaInfo.rxq.Utc = NmeaArray(3)
                End If
                NmeaInfo.rxq.SFLock = NmeaArray(4)
                If NmeaInfo.rxq.SFLock = "Y" Then
                    NmeaInfo.rxq.SFSNR = NmeaArray(5)
                    NmeaInfo.rxq.PerIdlePacket = NmeaArray(6)
                    NmeaInfo.rxq.PerBadPacket = NmeaArray(7)
                End If
                NmeaInfo.rxq.LastUpdate = GetTickCount
                parseNmea = True
            Else
                Debug.Print "ERROR Decoding RXQ " + InString + " Expected 4 or 7 but got " + Str(Max)
            End If
            Exit Function
        End If
        If NmeaArray(2) Like "NAVQ" Then
            ' either - $PNCTR,NAVQ,123519,3D,RTG,DUAL*55
            '     or - $PNCTR,NAVQ,202759,NN*74
            If Max > 3 Then
                If IsNumeric(NmeaArray(3)) Then
                    NmeaInfo.navq.Utc = NmeaArray(3)
                End If
                If IsNumeric(NmeaArray(4)) Then
                    NmeaInfo.navq.NavMode = NmeaArray(4)
                End If
                If NmeaArray(4) = "NN" Then
                    NmeaInfo.navq.NavMode = "None"
                Else
                    NmeaInfo.navq.CorrType = NmeaArray(5)
                    NmeaInfo.navq.SignalType = NmeaArray(6)
                End If
                NmeaInfo.navq.LastUpdate = GetTickCount
                parseNmea = True
            Else
                Debug.Print "ERROR Decoding NAVQ " + InString + " Expected 4 or 6 but got " + Str(Max)
            End If
            Exit Function
        End If
    Case "VTG"
        If Max = 9 Or Max = 10 Then
            If IsNumeric(NmeaArray(2)) Then
                NmeaInfo.vtg.cogt = NmeaArray(2)
            Else
                Exit Function
            End If
            NmeaInfo.vtg.CogTID = NmeaArray(3)
            If IsNumeric(NmeaArray(4)) Then
                NmeaInfo.vtg.CogM = NmeaArray(4)
            End If
            NmeaInfo.vtg.CogMID = NmeaArray(5)
            If IsNumeric(NmeaArray(6)) Then
                NmeaInfo.vtg.SpeedKt = NmeaArray(6)
            End If
            NmeaInfo.vtg.SpeedKtID = NmeaArray(7)
            If IsNumeric(NmeaArray(8)) Then
                NmeaInfo.vtg.SpeedKmh = NmeaArray(8)
            End If
            NmeaInfo.vtg.SpeedKmhID = NmeaArray(9)
            NmeaInfo.vtg.LastUpdate = GetTickCount
            parseNmea = True
            ' for NMEA-0183 version 3.0
            If Max = 10 Then
                NmeaInfo.vtg.Mode = NmeaArray(10)
            End If
            NmeaInfo.vtg.LastUpdate = GetTickCount
            parseNmea = True
        Else
            Debug.Print "ERROR Decoding VTG " + InString + " Expected 9 or 10 but got " + Str(Max)
        End If
        Exit Function
    Case "ZDA"
        If Max = 7 Then
            If IsNumeric(NmeaArray(2)) Then
                NmeaInfo.zda.Utc = NmeaArray(2)
            End If
            If IsNumeric(NmeaArray(3)) Then
                NmeaInfo.zda.Day = NmeaArray(3)
            End If
            If IsNumeric(NmeaArray(4)) Then
                NmeaInfo.zda.Month = NmeaArray(4)
            End If
            If IsNumeric(NmeaArray(5)) Then
                NmeaInfo.zda.Year = NmeaArray(5)
            End If
            If IsNumeric(NmeaArray(6)) Then
                NmeaInfo.zda.LocalTZhr = NmeaArray(6)
            End If
            If IsNumeric(NmeaArray(7)) Then
                NmeaInfo.zda.LocalTZmn = NmeaArray(7)
            End If
            NmeaInfo.zda.LastUpdate = GetTickCount
            parseNmea = True
        Else
            Debug.Print "ERROR Decoding ZDA " + InString + " Expected 7 but got " + Str(Max)
        End If
        Exit Function
    Case "HDT"
        If Max = 3 Then
            If IsNumeric(NmeaArray(2)) Then
                NmeaInfo.hdt.HeadingT = NmeaArray(2)
            End If
            NmeaInfo.hdt.HeadingTID = NmeaArray(3)
            NmeaInfo.hdt.LastUpdate = GetTickCount
            parseNmea = True
        Else
            Debug.Print "ERROR Decoding " + DeviceType + " HDT " + InString + " Expected 3 but got " + Str(Max)
        End If
        Exit Function
'    Case "HDM"
'        If Max = 3 Then
'            If IsNumeric(NmeaArray(2)) Then
'                NmeaInfo.hdm.HeadingM = NmeaArray(2)
'            End If
'            NmeaInfo.hdm.HeadingMID = NmeaArray(3)
'            NmeaInfo.hdm.LastUpdate = GetTickCount
'            parseNmea = True
'        Else
'            Debug.Print "ERROR Decoding Gyro HDM " + InString + " Expected 3 but got " + Str(Max)
'        End If
'        Exit Function
    Case "PR"
        If Max = 5 Then
            If IsNumeric(NmeaArray(2)) Then
                NmeaInfo.ohpr.Heading = NmeaArray(2)
            End If
            If IsNumeric(NmeaArray(3)) Then
                NmeaInfo.ohpr.Pitch = NmeaArray(3)
            End If
            If IsNumeric(NmeaArray(4)) Then
                NmeaInfo.ohpr.Roll = NmeaArray(4)
            End If
            If IsNumeric(NmeaArray(5)) Then
                NmeaInfo.ohpr.Depth = NmeaArray(5)
            End If
            NmeaInfo.ohpr.LastUpdate = GetTickCount
            parseNmea = True
        Else
            Debug.Print "ERROR Decoding OHPR " + InString + " Expected 7 but got " + Str(Max)
        End If
        Exit Function
    Case Else
        parseNmea = False
    End Select
Else
    Call sm_parse(InString, ",", NmeaArray)
    Max = UBound(NmeaArray)
    NmeaID = right$(NmeaArray(1), 3)
    Select Case NmeaID
    Case "HDT"
        If Max = 3 Then
            If IsNumeric(NmeaArray(2)) Then
                NmeaInfo.hdt.HeadingT = NmeaArray(2)
            End If
            NmeaInfo.hdt.HeadingTID = NmeaArray(3)
            NmeaInfo.hdt.LastUpdate = GetTickCount
            parseNmea = True
        Else
            Debug.Print "ERROR Decoding HDT " + InString + " Expected 3 but got " + Str(Max)
        End If
        Exit Function
    Case "HDM"
        If Max = 3 Then
            If IsNumeric(NmeaArray(2)) Then
                NmeaInfo.hdm.HeadingM = NmeaArray(2)
            End If
            NmeaInfo.hdm.HeadingMID = NmeaArray(3)
            NmeaInfo.hdm.LastUpdate = GetTickCount
            parseNmea = True
        Else
            Debug.Print "ERROR Decoding HDM " + InString + " Expected 3 but got " + Str(Max)
        End If
        Exit Function
    Case "RD1"
        If Max >= 6 Then
            If IsNumeric(NmeaArray(6)) Then
                NmeaInfo.rd1.BitErrorRate1 = NmeaArray(6)
                NmeaInfo.rd1.BitErrorRate2 = -1
            Else
                Call sm_parse(NmeaArray(6), "-", BitErrArray)
                If IsNumeric(BitErrArray(1)) Then
                    NmeaInfo.rd1.BitErrorRate1 = BitErrArray(1)
                End If
                If IsNumeric(BitErrArray(2)) Then
                    NmeaInfo.rd1.BitErrorRate2 = BitErrArray(2)
                End If
            End If
            NmeaInfo.rd1.LastUpdate = GetTickCount
            parseNmea = True
        End If
        Exit Function
    Case Else
        Debug.Print "Decoding Failed " + InString + " - detected fields = " + Str(Max)
    End Select
End If
Exit Function

ErrorHandler:
    If Err.Number <> 5 Then ' ignore Invalid procedure call or argument errors
        Debug.Print "Error decoding : " + InString
        'ErrorTrap Err.Number, "System(parseNmea)"
    End If
    Resume Next
End Function

Public Function CheckPath(strPath As String) As Boolean
    If Dir$(strPath) <> "" Then
        CheckPath = True
    Else
        CheckPath = False
    End If
End Function

'Public Function FreeDiskSpace(ByVal sDriveLetter As String) As Double
'Dim udtFreeBytesAvail As LargeInt, udtTtlBytes As LargeInt
'Dim udtTTlFree As LargeInt
'Dim dblFreeSpace As Double

'If GetDiskFreeSpaceEx(sDriveLetter, udtFreeBytesAvail, udtTtlBytes, udtTTlFree) Then
'    If udtFreeBytesAvail.lngLower < 0 Then
'        dblFreeSpace = udtFreeBytesAvail.lngUpper * 2 ^ 32 + udtFreeBytesAvail.lngLower + 4294967296#
'    Else
'        dblFreeSpace = udtFreeBytesAvail.lngUpper * 2 ^ 32 + udtFreeBytesAvail.lngLower
'    End If
'End If
'FreeDiskSpace = dblFreeSpace
'End Function

Public Function IsADirectory(ByVal TheName As String) As Boolean
If GetAttr(TheName) And vbDirectory Then
    IsADirectory = True
End If
End Function

Public Function GetDirectories(myPath As String, dirArray() As String)
Dim myName As String
Dim count As Integer
' Display the names in C:\ that represent directories.
myName = Dir(myPath, vbDirectory)   ' Retrieve the first entry.
Do While myName <> ""   ' Start the loop.
   ' Ignore the current directory and the encompassing directory.
    If myName <> "." And myName <> ".." Then
      ' Use bitwise comparison to make sure MyName is a directory.
        If (GetAttr(myPath & myName) And vbDirectory) = vbDirectory Then
            count = count + 1
            ReDim Preserve dirArray(count)
            dirArray(count - 1) = myName
        End If
    End If
    myName = Dir   ' Get next entry.
Loop
End Function

'    Public Type STAT_ANALYSIS
'       N As Long
'       Sum As Double
'       Mean As Double
'       Min As Double
'       Max As Double
'       Range As Double
'       SumSquared As Double
'       SumOfXiSquared As Double
'       Variance As Double
'       StandardDeviation As Double
'       SourceData As Variant
'       Population As Boolean
'    End Type
     
'    Public Function Analyze(varData As Variant, Optional ByVal blnPopulation As Boolean = True) As STAT_ANALYSIS
'       Dim blnFirstPass As Boolean 'flag to tell if it's the first pass
'       Dim i As Long
'       'passing anything other than an array is useless, so terminate if varData doesn't have an array
'       If Not CBool(VarType(varData) And vbArray) Then Exit Function
'       With Analyze
'          'total number of items
'          .N = UBound(varData) - LBound(varData) + 1
'          'statistical analysis needs at the very very least 2 items
'          If .N < 2 Then Exit Function
'          .Sum = 0
'          .SumOfXiSquared = 0
'          blnFirstPass = True
'          'go through each item in the data set
'          For i = LBound(varData) To UBound(varData)
'             'get the running sum of the items
'             .Sum = .Sum + varData(i)
'             'get the running sum of the square of each item
'             .SumOfXiSquared = .SumOfXiSquared + (varData(i) ^ 2)
'             'if this isn't the first pass, give the max and min a value to start off
'             If blnFirstPass Then
'                .Min = varData(i)
'                .Max = varData(i)
'                blnFirstPass = False
'             Else
'                'these two statements will eventually find the max and min of the data set
'                If varData(i) > .Max Then .Max = varData(i)
'                If varData(i) < .Min Then .Min = varData(i)
'             End If
'          Next i
'          'get the range of the data
'          .Range = .Max - .Min
'          'get the mean (average) of the data
'          .Mean = .Sum / .N
'          'get the sum of all items squared
'          .SumSquared = .Sum ^ 2
'          'get the variance, based on whether the data set is a population or a sample
'          .Variance = (.SumOfXiSquared - (.SumSquared / .N)) / IIf(blnPopulation, .N, .N - 1)
'          'get the standard deviation
'          .StandardDeviation = Sqr(.Variance)
'          'stick the source data into the struct in case you need it later
'          .SourceData = varData
'          'tell whether or not the source data was treated as a population
'          .Population = blnPopulation
'       End With
'    End Function

Sub QuickSort(ByRef arr() As Double, Optional numEls As Variant, Optional descending As Boolean)

Dim Value As Variant, temp As Variant
Dim sp As Integer
Dim leftStk(32) As Long, rightStk(32) As Long
Dim leftNdx As Long, rightNdx As Long
Dim i As Long, j As Long

' account for optional arguments
If IsMissing(numEls) Then numEls = UBound(arr)
' init pointers
leftNdx = LBound(arr)
rightNdx = numEls
' init stack
sp = 1
leftStk(sp) = leftNdx
rightStk(sp) = rightNdx

Do
    If rightNdx > leftNdx Then
        Value = arr(rightNdx)
        i = leftNdx - 1
        j = rightNdx
        ' find the pivot item
        If descending Then
            Do
                Do: i = i + 1: Loop Until arr(i) <= Value
                Do: j = j - 1: Loop Until j = leftNdx Or arr(j) >= Value
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            Loop Until j <= i
        Else
            Do
                Do: i = i + 1: Loop Until arr(i) >= Value
                Do: j = j - 1: Loop Until j = leftNdx Or arr(j) <= Value
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            Loop Until j <= i
        End If
        ' swap found items
        temp = arr(j)
        arr(j) = arr(i)
        arr(i) = arr(rightNdx)
        arr(rightNdx) = temp
        ' push on the stack the pair of pointers that differ most
        sp = sp + 1
        If (i - leftNdx) > (rightNdx - i) Then
            leftStk(sp) = leftNdx
            rightStk(sp) = i - 1
            leftNdx = i + 1
        Else
            leftStk(sp) = i + 1
            rightStk(sp) = rightNdx
            rightNdx = i - 1
        End If
    Else
        ' pop a new pair of pointers off the stacks
        leftNdx = leftStk(sp)
        rightNdx = rightStk(sp)
        sp = sp - 1
        If sp = 0 Then Exit Do
    End If
Loop
End Sub

