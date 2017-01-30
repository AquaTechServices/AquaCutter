Attribute VB_Name = "addtrig"

Private Const PI = 3.14159265358979

' cosecant of an angle
Function Csc(radians As Double) As Double
    Csc = 1 / Sin(radians)
End Function

' arc sine
' error if value is outside the range [-1,1]

Function ASin(value As Double) As Double
If Abs(value) <> 1 Then
    ASin = Atn(value / Sqr(1 - value * value))
Else
    ASin = 1.5707963267949 * Sgn(value)
End If
End Function

' arc cosine
' error if NUMBER is outside the range [-1,1]

Function ACos(value As Double) As Double
If Abs(value) <> 1 Then
    ACos = 1.5707963267949 - Atn(value / Sqr(1 - value * value))
Else
    ACos = 3.14159265358979 * Sgn(value)
End If
End Function

' arc cotangent
' error if NUMBER is zero

Function ACot(value As Double) As Double
    ACot = Atn(1 / value)
End Function

' arc secant
' error if value is inside the range [-1,1]

Function ASec(value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ASec = ACos(1 / value)
If Abs(value) <> 1 Then
    ASec = 1.5707963267949 - Atn((1 / value) / Sqr(1 - 1 / (value * value)))
Else
    ASec = 3.14159265358979 * Sgn(value)
End If
End Function

' arc cosecant
' error if value is inside the range [-1,1]

Function ACsc(value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ACsc = ASin(1 / value)
If Abs(value) <> 1 Then
    ACsc = Atn((1 / value) / Sqr(1 - 1 / (value * value)))
Else
    ACsc = 1.5707963267949 * Sgn(value)
End If
End Function


