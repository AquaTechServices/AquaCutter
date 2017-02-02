Attribute VB_Name = "KalmanFilter"
Option Explicit

' ********************************************************************
' *   Kalman Filter                                                  *
' *   Based on Kalman Filter info found at:                                                      *
' *   http://bilgin.esme.org/BitsBytes/KalmanFilterforDummies.aspx   *
' *   By: IGutYa                                                     *
' ********************************************************************



Private Const Noise = 0.1                     ' Noise Preset



Public Function Kalman(mVal() As Double) As Double
    Dim C As Double, X As Double, P As Double
    Dim i As Integer
    
    C = 0
    X = 0
    P = 1
        
    For i = 1 To UBound(mVal)
        
        ' Basic Kalman Filter, Were the Magic Happens:
        
        C = P / (P + Noise)
        X = X + C * (mVal(i) - X)
        P = (1 - C) * P
        
    Next i

Kalman = X

End Function

