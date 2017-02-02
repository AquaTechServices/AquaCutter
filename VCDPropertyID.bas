Attribute VB_Name = "VCDPropertyID"
Option Explicit

' Standard Interface IDs
Public Const VCDInterface_Range = "{99B44940-BFE1-4083-ADA1-BE703F4B8E03}"
Public Const VCDInterface_Switch = "{99B44940-BFE1-4083-ADA1-BE703F4B8E04}"
Public Const VCDInterface_Button = "{99B44940-BFE1-4083-ADA1-BE703F4B8E05}"
Public Const VCDInterface_MapStrings = "{99B44940-BFE1-4083-ADA1-BE703F4B8E06}"
Public Const VCDInterface_AbsoluteValue = "{99B44940-BFE1-4083-ADA1-BE703F4B8E08}"

' Standard Element IDs
Public Const VCDElement_Value = "{B57D3000-0AC6-4819-A609-272A33140ACA}"
Public Const VCDElement_Auto = "{B57D3001-0AC6-4819-A609-272A33140ACA}"
Public Const VCDElement_OnePush = "{B57D3002-0AC6-4819-A609-272A33140ACA}"

' Standard Property Item IDs
Public Const VCDID_Brightness = "{284C0E06-010B-45BF-8291-09D90A459B28}"
Public Const VCDID_Contrast = "{284C0E07-010B-45BF-8291-09D90A459B28}"
Public Const VCDID_Hue = "{284C0E08-010B-45BF-8291-09D90A459B28}"
Public Const VCDID_Saturation = "{284C0E09-010B-45BF-8291-09D90A459B28}"
Public Const VCDID_Sharpness = "{284C0E0A-010B-45BF-8291-09D90A459B28}"
Public Const VCDID_Gamma = "{284C0E0B-010B-45BF-8291-09D90A459B28}"
Public Const VCDID_ColorEnable = "{284C0E0C-010B-45BF-8291-09D90A459B28}"
Public Const VCDID_WhiteBalance = "{284C0E0D-010B-45BF-8291-09D90A459B28}"
Public Const VCDID_BacklightCompensation = "{284C0E0E-010B-45BF-8291-09D90A459B28}"
Public Const VCDID_Gain = "{284C0E0F-010B-45BF-8291-09D90A459B28}"

Public Const VCDID_Pan = "{90D5702A-E43B-4366-AAEB-7A7A10B448B4}"
Public Const VCDID_Tilt = "{90D5702B-E43B-4366-AAEB-7A7A10B448B4}"
Public Const VCDID_Roll = "{90D5702C-E43B-4366-AAEB-7A7A10B448B4}"
Public Const VCDID_Zoom = "{90D5702D-E43B-4366-AAEB-7A7A10B448B4}"
Public Const VCDID_Exposure = "{90D5702E-E43B-4366-AAEB-7A7A10B448B4}"
Public Const VCDID_Iris = "{90D5702F-E43B-4366-AAEB-7A7A10B448B4}"
Public Const VCDID_Focus = "{90D57030-E43B-4366-AAEB-7A7A10B448B4}"

Public Const VCDID_TriggerMode = "{90D57031-E43B-4366-AAEB-7A7A10B448B4}"
Public Const VCDID_VCRCompatibilityMode = "{90D57032-E43B-4366-AAEB-7A7A10B448B4}"
Public Const VCDID_SignalDetected =  "{90D57033-E43B-4366-AAEB-7A7A10B448B4}"
Public Const VCDID_ColorEnhancement= "{3A3A8F77-6440-46CC-94XA-8752B02E6C29}"

' TIS DCAM Property Item IDs
Public Const VCDID_TestPattern = "{F7EAA79E-90FA-4969-B05F-9BDAF1A4328F}"

Public Const VCDID_MultiSlope = "{630B1F3E-4A0A-4963-89B1-86BA8FDA2990}"

Public Const VCDElement_MultiSlope_SlopeValue0 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3090}"
Public Const VCDElement_MultiSlope_ResetValue0 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3091}"
Public Const VCDElement_MultiSlope_SlopeValue1 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3190}"
Public Const VCDElement_MultiSlope_ResetValue1 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3191}"
Public Const VCDElement_MultiSlope_SlopeValue2 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3290}"
Public Const VCDElement_MultiSlope_ResetValue2 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3291}"
Public Const VCDElement_MultiSlope_SlopeValue3 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3390}"
Public Const VCDElement_MultiSlope_ResetValue3 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3391}"
Public Const VCDElement_MultiSlope_SlopeValue4 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3490}"
Public Const VCDElement_MultiSlope_ResetValue4 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3491}"
Public Const VCDElement_MultiSlope_SlopeValue5 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3590}"
Public Const VCDElement_MultiSlope_ResetValue5 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3591}"
Public Const VCDElement_MultiSlope_SlopeValue6 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3690}"
Public Const VCDElement_MultiSlope_ResetValue6 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3691}"
Public Const VCDElement_MultiSlope_SlopeValue7 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3790}"
Public Const VCDElement_MultiSlope_ResetValue7 = "{630B1F3E-4A0A-4963-89B1-86BA8FDA3791}"


' TIS DCAM Element IDs
Public Const VCDElement_WhiteBalanceBlue = "{6519038A-1AD8-4E91-9021-66D64090CC85}"
Public Const VCDElement_WhiteBalanceRed =   "{6519038B-1AD8-4E91-9021-66D64090CC85}"
Public Const VCDElement_WhiteBalanceGreen = "{8407E480-175A-498C-8171-08BD987CC1AC}"

Public Const VCDElement_TriggerPolarity = "{6519038D-1AD8-4E91-9021-66D64090CC85}"
Public Const VCDElement_TriggerMode = "{6519038E-1AD8-4E91-9021-66D64090CC85}"
Public Const VCDElement_SoftwareTrigger = "{FDB4003C-552C-4FAA-B87B-42E888D54147}"
Public Const VCDElement_ResetValue =   "{B57D3003-0AC6-4819-A609-272A33140ACA}"

Public Const VCDElement_AutoMaxValue     = "{6519038F-1AD8-4E91-9021-66D64090CC85}"
Public Const VCDElement_AutoMaxValueAuto = "{65190390-1AD8-4E91-9021-66D64090CC85}"
Public Const VCDElement_AutoReference = "{6519038C-1AD8-4E91-9021-66D64090CC85}"


Public Const VCDID_GPIO = 			"{86D89D69-9880-4618-9BF6-DED5E8383449}"
Public Const VCDElement_GPIOIn =	"{7D006621-761D-4B88-9C5F-8B906857A500}"
Public Const VCDElement_GPIOOut =	"{7D006621-761D-4B88-9C5F-8B906857A501}"
Public Const VCDElement_GPIOWrite =	"{7D006621-761D-4B88-9C5F-8B906857A502}"
Public Const VCDElement_GPIORead =	"{7D006621-761D-4B88-9C5F-8B906857A503}"

Public Const VCDID_Strobe = 				"{DC320EDE-DF2E-4A90-B926-71417C71C57C}"
Public Const VCDElement_StrobePolarity = 	"{B41DB628-0975-43F8-A9D9-7E0380580ACA}"
Public Const VCDElement_StrobeDuration = 	"{B41DB628-0975-43F8-A9D9-7E0380580ACB}"
Public Const VCDElement_StrobeDelay = 		"{B41DB628-0975-43F8-A9D9-7E0380580ACC}"

' Partitial Scan
Public Const VCDID_PartialScanOffset =          "{2CED6FD6-AB4D-4C74-904C-D682E53B9CC5}"
Public Const VCDElement_PartialScanAutoCenter = "{36EAA683-3321-44BE-9D73-E1FD4C3FDB87}"
Public Const VCDElement_PartialScanOffsetX =    "{5E59F654-7B47-4458-B4C6-5D4F0D175FC1}"
Public Const VCDElement_PartialScanOffsetY  =   "{87FB6C02-98A8-46B0-B18D-6442D9775CD3}"

