Attribute VB_Name = "VCDSimpleModule"
Option Explicit

Public Const errPropertyNotFound As Long = vbObjectError + 1

Public Function fetchCheckedRangeItf(itemId As String, props As VCDPropertyItems) As VCDRangeProperty
On Error GoTo raise_error
    Dim irange As VCDRangeProperty
    Dim itf As VCDPropertyInterface
    Dim id As String
    
    If itemId = VCDElement_WhiteBalanceBlue Then
        id = VCDID_WhiteBalance + ":" + VCDElement_WhiteBalanceBlue
    ElseIf itemId = VCDElement_WhiteBalanceRed Then
        id = VCDID_WhiteBalance + ":" + VCDElement_WhiteBalanceRed
    ElseIf itemId = VCDElement_GPIOIn Then
        id = VCDID_GPIO + ":" + VCDElement_GPIOIn
    ElseIf itemId = VCDElement_GPIOOut Then
        id = VCDID_GPIO + ":" + VCDElement_GPIOOut
    ElseIf itemId = VCDElement_StrobeDelay Then
        id = VCDID_Strobe + ":" + VCDElement_StrobeDelay
    ElseIf itemId = VCDElement_StrobeDuration Then
        id = VCDID_Strobe + ":" + VCDElement_StrobeDuration
    Else
        id = itemId + ":" + VCDElement_Value
    End If
    
    id = id + ":" + VCDInterface_Range

    Set irange = props.FindInterface(id)
    If Not irange Is Nothing Then
        Set fetchCheckedRangeItf = irange
    Else
raise_error:
        Err.Raise errPropertyNotFound, "VCDSimpleProperty", "Range for Property " + id + " not found"
    End If
End Function

Public Function fetchCheckedSwitchItf(itemId As String, props As VCDPropertyItems) As VCDSwitchProperty
On Error GoTo raise_error
    Dim AutoProp As VCDSwitchProperty

    Dim id As String
    If itemId = VCDElement_WhiteBalanceBlue Or itemId = VCDElement_WhiteBalanceRed Then
        id = VCDID_WhiteBalance + ":" + VCDElement_Auto
    ElseIf itemId = VCDElement_StrobePolarity Then
        id = VCDID_Strobe + ":" + VCDElement_StrobePolarity
    ElseIf itemId = VCDID_Strobe Then
        id = VCDID_Strobe + ":" + VCDElement_Value
    ElseIf itemId = VCDID_TriggerMode Then
        id = VCDID_TriggerMode + ":" + VCDElement_Value
    Else
        id = itemId + ":" + VCDElement_Auto
    End If
    
    id = id + ":" + VCDInterface_Switch

    Set AutoProp = props.FindInterface(id)
    If Not AutoProp Is Nothing Then
        Set fetchCheckedSwitchItf = AutoProp
    Else
raise_error:
        Err.Raise errPropertyNotFound, "VCDSimpleProperty", "AutoProp for Property " + id + " not found"
    End If
End Function

Public Function fetchCheckedOnePushItf(itemId As String, props As VCDPropertyItems) As VCDButtonProperty
On Error GoTo raise_error
    Dim OnePushProp As VCDButtonProperty

    Dim id As String
    If itemId = VCDElement_WhiteBalanceBlue Or itemId = VCDElement_WhiteBalanceRed Then
        id = VCDID_WhiteBalance + ":" + VCDElement_OnePush
    ElseIf itemId = VCDElement_GPIORead Then
        id = VCDID_GPIO + ":" + VCDElement_GPIORead
    ElseIf itemId = VCDElement_GPIOWrite Then
        id = VCDID_GPIO + ":" + VCDElement_GPIOWrite
    Else
        id = itemId + ":" + VCDElement_OnePush
    End If
    
    id = id + ":" + VCDInterface_Button

    Set OnePushProp = props.FindInterface(id)
    If Not OnePushProp Is Nothing Then
        Set fetchCheckedOnePushItf = OnePushProp
    Else
raise_error:
        Err.Raise errPropertyNotFound, "VCDSimpleProperty", "OnePushProp for Property " + id + " not found"
    End If
End Function

Public Function GetSimplePropertyContainer(props As VCDPropertyItems) As VCDSimpleProperty
    Dim r As VCDSimpleProperty
    Set r = New VCDSimpleProperty
    r.init props
    Set GetSimplePropertyContainer = r
End Function


