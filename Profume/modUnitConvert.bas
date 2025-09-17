Attribute VB_Name = "modUnitConvert"
Option Explicit
Const MODULE_NAME = "modUnitConvert"
Public Enum ConvertMode
    modeDisplay = 0
    modeSave = 1
End Enum
Public Enum UnitType
    unitLWH = 0
    unitTEMPERATURE = 1
    unitWEIGHT = 2
    unitVOLUME = 3
    unitSHOOTINGRATE = 4
    unitTIME = 5
    unitLess = 6
    unitPressure = 7
End Enum

'**************************************************************
'   SUBROUTINE:  UnitConvert
'
'   DESCRIPTION
'       Converts a given an input value to its metric or english equivalent
'
'   PARAMETERS
'           ConvMode  - the convertmode.  0 for "Display" or 1 for "Save"
'           pvUnit      - string indicating unit of measure
'           pvOldValue  - the value to convert
'           ByRef prNewValue  - the converted value
'           ByRef prUnitName  - the unit of measure name depending on the measure system
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/18/2002       Malou Espaldon    Initial Revision
'**************************************************************

Public Sub UnitConvert(pvConvMode As ConvertMode, pvUnit As UnitType, pvOldValue As Double, ByRef prNewValue As Double, Optional ByRef prUnitName As String)

    Const FUNCTION_NAME = "UnitConvert"
    
    '--The unit conversion constants are from the
    '--CRC Handbook of Chemistry and Physics 57th Edition
    Const METERS_PER_FOOT As Double = 0.3048
    Const FEET_PER_METER As Double = 3.2808399
    Const MILES_PER_KM As Double = 0.62137119
    Const KM_PER_MILE As Double = 1.609344
    Const CUM_PER_CUF As Double = 0.0283168466
    Const CUF_PER_CUM As Double = 35.3146667
    Const POUNDS_PER_KG As Double = 2.2046226
    Const KG_PER_POUND As Double = 0.45359237
    Const TIME_MIN As Double = 1#
    Const PSIG_PER_BAR = 0.0672
  

    On Error GoTo Error_Handler

    If pvConvMode = modeSave Then
        Select Case pvUnit
            Case unitLWH
                prNewValue = IIf(gbIsMetric, pvOldValue, pvOldValue * METERS_PER_FOOT)
                prUnitName = ""
            Case unitVOLUME
                prNewValue = IIf(gbIsMetric, pvOldValue, pvOldValue * CUM_PER_CUF)
                prUnitName = ""
            Case unitTEMPERATURE
                prNewValue = IIf(gbIsMetric, pvOldValue, (5 / 9) * (pvOldValue - 32))
                prUnitName = ""
            Case unitWEIGHT
                prNewValue = IIf(gbIsMetric, pvOldValue, pvOldValue * KG_PER_POUND)
                prUnitName = ""
            Case unitSHOOTINGRATE
                prNewValue = IIf(gbIsMetric, pvOldValue, pvOldValue * KG_PER_POUND)
                prUnitName = ""
            Case unitPressure
                prNewValue = IIf(gbIsMetric, pvOldValue, pvOldValue * PSIG_PER_BAR)
                prUnitName = ""
            Case unitTIME
                prNewValue = IIf(gbIsMetric, pvOldValue, pvOldValue * TIME_MIN)
                prUnitName = ""
            Case unitLess
                prNewValue = pvOldValue
                prUnitName = ""
            Case Else
                prNewValue = Null
                prUnitName = ""
        End Select
    ElseIf pvConvMode = modeDisplay Then
    '--assumption here is that value is converted FROM METRIC
        Select Case pvUnit
            Case unitLWH
                prNewValue = IIf(gbIsMetric, pvOldValue, pvOldValue * FEET_PER_METER)
                prUnitName = IIf(gbIsMetric, LoadResString(IDS_METER), LoadResString(IDS_FEET))
            Case unitVOLUME
                prNewValue = IIf(gbIsMetric, pvOldValue, pvOldValue * CUF_PER_CUM)
                prUnitName = IIf(gbIsMetric, LoadResString(IDS_METER_CUBED), LoadResString(IDS_FEET_CUBED))
            Case unitTEMPERATURE
                prNewValue = IIf(gbIsMetric, pvOldValue, ((9 / 5) * pvOldValue) + 32)
                prUnitName = IIf(gbIsMetric, LoadResString(IDS_CELSIUS), LoadResString(IDS_FAHRENHEIT))
            Case unitWEIGHT
                prNewValue = IIf(gbIsMetric, pvOldValue, pvOldValue * POUNDS_PER_KG)
                prUnitName = IIf(gbIsMetric, LoadResString(IDS_KILOGRAMS), LoadResString(IDS_POUNDS))
            Case unitSHOOTINGRATE
                prNewValue = IIf(gbIsMetric, pvOldValue, pvOldValue * POUNDS_PER_KG)
                prUnitName = IIf(gbIsMetric, LoadResString(IDS_KG_PER_MINUTE), LoadResString(IDS_POUNDS_PER_MINUTE))
            Case unitTIME
                prNewValue = IIf(gbIsMetric, pvOldValue, pvOldValue * TIME_MIN)
                prUnitName = IIf(gbIsMetric, LoadResString(IDS_MINUTE), LoadResString(IDS_MINUTE))
            Case unitPressure
                prNewValue = IIf(gbIsMetric, pvOldValue, pvOldValue / PSIG_PER_BAR)
                prUnitName = IIf(gbIsMetric, LoadResString(IDS_PRESSURE_BAR), LoadResString(IDS_PRESSURE_PSIG))
            Case unitLess
                prNewValue = pvOldValue
                prUnitName = ""
            Case Else
                prNewValue = Null
                prUnitName = ""
        End Select
    Else
        prNewValue = Null
        prUnitName = ""
    End If
    Exit Sub
Error_Handler:
    If Err.Number <> 0 Then
    '-- just return empty values
        prNewValue = Null
        prUnitName = ""
        Err.Clear
    End If
End Sub
'**************************************************************
'   SUBROUTINE:  ConvertTempPressure
'
'   DESCRIPTION
'       Converts Temperature to pressure
'
'   PARAMETERS
'
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  3/18/2005       Pat San Miguel   Initial Revision
'**************************************************************
Public Sub ConvertTempPressure(dblTemp As Double, ByRef dblPressure As Double)
    Const FUNCTION_NAME = "ConvertTempPressure"
    
    On Error GoTo Error_Handler
       
    dblPressure = (10 ^ (4.597575 - (1026.558 / (dblTemp + 280.439)))) * 1.013
    
Error_Handler:
    If Err.Number <> 0 Then
        dblPressure = Null
        Err.Clear
    End If
End Sub

'**************************************************************
'   SUBROUTINE:  ConvertPressureTemp
'
'   DESCRIPTION
'       Converts Pressure to Temperature
'
'   PARAMETERS
'
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  3/18/2005       Pat San Miguel   Initial Revision
'**************************************************************
Public Sub ConvertPressureTemp(ByRef dblTemp As Double, dblPressure As Double)
    Const FUNCTION_NAME = "ConvertPressureTemp"
    Dim dblLog As Double
    
    On Error GoTo Error_Handler
    
    dblLog = (Log(dblPressure / 1.013) / Log(10#))
        
    dblTemp = ((4.597575 * 280.4394) - 1026.558 - (dblLog * 280.4394)) / (dblLog - 4.597575)
    
    If gbIsMetric = False Then
        dblTemp = ((9 / 5) * dblTemp) + 32
    End If
    
    
Error_Handler:
    If Err.Number <> 0 Then
        dblTemp = Null
        Err.Clear
    End If
End Sub

