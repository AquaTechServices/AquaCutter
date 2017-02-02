VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VCDSimpleProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private internalTree As VCDPropertyItems

Public Sub init(Tree As VCDPropertyItems)
    Set internalTree = Tree
End Sub

Private Function fetchItf(id As String) As VCDPropertyItem
End Function

Public Function AutoAvailable(id As String) As Boolean
On Error GoTo na
    Dim x As VCDSwitchProperty
    Set x = fetchCheckedSwitchItf(id, internalTree)
    AutoAvailable = x.Available
    Exit Function
na:
    AutoAvailable = False
End Function

Public Function Available(id As String) As Boolean
On Error GoTo na
    Dim fetchItf As VCDPropertyItem
    
    If id = VCDElement_WhiteBalanceBlue Then
        id = VCDID_WhiteBalance + ":" + VCDElement_WhiteBalanceBlue
    ElseIf id = VCDElement_WhiteBalanceRed Then
        id = VCDID_WhiteBalance + ":" + VCDElement_WhiteBalanceRed
    End If

    Set fetchItf = internalTree.FindItem(id)
    If fetchItf Is Nothing Then
        Available = False
    Else
        Available = True
    End If
    Exit Function
na:
    Available = False
End Function

Property Get RangeValue(id As String) As Long
    Dim RangeProp As VCDRangeProperty
    Set RangeProp = fetchCheckedRangeItf(id, internalTree)
    
    If RangeProp.Delta <> 0 Then
        RangeValue = RangeProp.value / RangeProp.Delta
    Else
        RangeValue = RangeProp.value
    End If
End Property

Property Let RangeValue(id As String, l As Long)
    Dim RangeProp As VCDRangeProperty
    Set RangeProp = fetchCheckedRangeItf(id, internalTree)
    RangeProp.value = RangeProp.Delta * l
End Property

Property Get Automation(id As String) As Boolean
    Automation = fetchCheckedSwitchItf(id, internalTree).Switch
End Property

Property Let Automation(id As String, b As Boolean)
    fetchCheckedSwitchItf(id, internalTree).Switch = b
End Property

Property Get Switch(id As String) As Boolean
    Switch = fetchCheckedSwitchItf(id, internalTree).Switch
End Property

Property Let Switch(id As String, b As Boolean)
    fetchCheckedSwitchItf(id, internalTree).Switch = b
End Property


'  neu eingefügt!
Public Sub OnePush(id As String)
    fetchCheckedOnePushItf(id, internalTree).Push
End Sub

' ende eingefügt!
Property Get RangeMin(id As String) As Long
    Dim RangeProp As VCDRangeProperty
    Set RangeProp = fetchCheckedRangeItf(id, internalTree)
    RangeMin = RangeProp.RangeMin / RangeProp.Delta
End Property

Property Get RangeMax(id As String) As Long
    Dim RangeProp As VCDRangeProperty
    Set RangeProp = fetchCheckedRangeItf(id, internalTree)
    RangeMax = RangeProp.RangeMax / RangeProp.Delta
End Property

Property Get DefaultValue(id As String) As Long
    Dim RangeProp As VCDRangeProperty
    Set RangeProp = fetchCheckedRangeItf(id, internalTree)
    DefaultValue = RangeProp.Default / RangeProp.Delta
End Property

Property Get Delta(id As String) As Long
    Delta = fetchCheckedRangeItf(id, internalTree).Delta
End Property
