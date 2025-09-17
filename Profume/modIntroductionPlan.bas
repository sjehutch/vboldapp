Attribute VB_Name = "modIntroductionPlan"
' JETT HAS THOUROUGHLY REVIEWED AND OPTIMIZED THIS MODULE
'------------------------------------------------------
'*
'* Filename: modIntroductionPlan.bas
'*
'* DESCRIPTION
'* -----------
'* This code contains the functions and subroutines
'* used by frmIntroPlan.
'*
'* Copyright (c) 2002 Accenture.  All rights reserved.
'-----------------------------------------------------
'*
'*  REVISION LOG
'*      DATE       NAME              CHANGE
'*  10/09/2003  Joseph Manyaga    Initial revision.
'*  03/31/2003  Joseph Manyaga    Fumiguide Enhancements 2003:
'*                                Modified Shooting Rate error message
'*                                Set default value of RH to blank
'*                                Set options of Inside Diameter to 3 values only
'*
'*  1/07/2005  Patrick San Miguel Added new columns (1/4 length, 1/8 length and
'*                                Time (min) to empty cylinder) and
'*                                removed Diameter field and length field
'-----------------------------------------------------

Option Explicit

'Declare variables
Public sFirstIntroAreaName As String    'stores the default area name
Public sFirstInsideDiameter As String   'stores the default inside diameter
Dim arrDiameter(3) As String            'array containing the Inside Diameter list options
Dim sGridIntroLineName As String        'stores the names of the existing Intro Lines
Public bHasIntroLine As Boolean            'flag if there are Intro Line records
Dim sFirstIntroAreaId As String         'stores the id of the default area
Dim nDecimal As Integer                 'stores the number of decimal places
Dim nRecordNumber As Integer            'stores the current record number

Dim dTempPresTemp As Double
Public blnOpenFile As Boolean
              



'Define constants
'Pat 4/8/2005 - New COnstants
Private Const VAL_PRESSURE = 9.99

Private Const COL_INDEX As Integer = 0
Private Const COL_AREA_NAME As Integer = 1
Private Const COL_INTROLINE_NAME As Integer = 2
Private Const COL_RH As Integer = 3
Private Const COL_CYLINDER_TEMP As Integer = 4
'Pat 12/28/2004 Removed Diameter field and length field. Replaced them with 1/4 and 1/8 length
Private Const COL_ONE_FOURTH As Integer = 5
Private Const COL_ONE_EIGHT As Integer = 6
'Private Const COL_INSIDE_DIAMETER As Integer = 5
'Private Const COL_LENGTH As Integer = 6
Private Const COL_FAN_CAP As Integer = 7
Private Const COL_NUM_CYLINDER As Integer = 9
Private Const COL_ORDER_RS As Integer = 8

Private Const COL_ACTUAL_RATE = 3
Private Const COL_PERMIT_RATE = 4
Private Const COL_TIME_EMPTY_CYL = 5

'JNM 03/31/03 - Fumiguide Enhancements 2003 - Set inside diameter to 3 options only
Private Const SHOOTINGLINE_DIAMETER_ONE As Double = 3.175
Private Const SHOOTINGLINE_DIAMETER_TWO As Double = 4.37
Private Const SHOOTINGLINE_DIAMETER_THREE As Double = 6.35

'JNM 03/31/03 - FE 2003 - Humidity has a blank default value which we will represent with
'   the -1 value.
Private Const DEFAULT_HUMIDITY As Double = -1
Private Const DEFAULT_DIAMETER As Double = SHOOTINGLINE_DIAMETER_ONE

Const AREANAME_COLSIZE = 2880
Const INTROLINE_COLSIZE = 3240
Const RH_COLSIZE = 960
Const CYLTEMP_COLSIZE = 1350
'Pat 12/28/2004 Removed Diameter field and length field. Replaced them with 1/4 and 1/8 length
Const ONEEIGHT_LENGTH_COLSIZE = 1000
Const ONEFOURTH_LENGTH_COLSIZE = 1000
'Const DIAMETER_COLSIZE = 1335
'Const LENGTH_COLSIZE = 1140
Const FANCAP_COLSIZE = 1500
Const NUM_CYLINDER_COLSIZE = 1140

Const AREANAME2_COLSIZE = 2880
Const LINENAME_COLSIZE = 3180
Const ACTUALRATE_COLSIZE = 1410
Const PERMITRATE_COLSIZE = 2115
Const TIMEMIN_COLSIZE = 2115

Private Const MODULE_NAME = "modIntroductionPlan"

'Create the GridWrapper instances
Dim GridWrapInput As New clsGridWrapper
Dim GridWrapOutput As New clsGridWrapper

'Create database objects
Dim rsIntroLine As New ADODB.Recordset




Public Sub IntroPlanInitializeControls()
'**************************************************************
'   DESCRIPTION
'       This subroutine initializes all the controls in the form.
'       This retrieves the data from the database and displays these
'       in the UI.
'
'   PARAMETERS
'          NONE
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************

On Error GoTo Error_Handler

    Dim sIntroLineSQL
    Dim I
    
    If rsIntroLine.State Then
        rsIntroLine.Close
    End If
    
    'blnConvert = False
    
    'Create the Introduction Line recordset
    sIntroLineSQL = "select ShootingLine.* from ShootingLine left outer join " & _
        "JobArea on ShootingLine.jobarea_fkey = JobArea.id where " & _
        "ShootingLine.deleted = false and ShootingLine.job_fkey = '" & gsJobID & "'" & _
        "order by ShootingLine.row"

    rsIntroLine.Open sIntroLineSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText

    nDecimal = IIf(gbIsMetric, 2, 0)

    With frmMain
        'Initialize the labels from the Resource file
        .lblIntroInput.Caption = LoadResString(IDS_CAP_INTRO_INPUT)
        .lblIntroRates.Caption = LoadResString(IDS_CAP_INTRO_RATES)
        .lblWarnings.Caption = LoadResString(IDS_WARNING_CAP)
        'Pat 3/16/205: addded caption for cylinder temp/pressure combobox
        .lblCylinderTempPressure.Caption = LoadResString(IDS_CONVERT_CAPTION_DESC)
        .rbtTemperature.Caption = LoadResString(IDS_CYLINDER_TEMP)
        .rbtPressure.Caption = LoadResString(IDS_CYLINDER_PRESSURE)
        .btnConvert.Caption = LoadResString(IDS_CONVERT_CAPTION)
        .cmdNextStep3.Caption = LoadResString(IDS_BTN_NEXTSTEP)
        
        'Pat 4/5:
        If Not rsIntroLine.EOF And Not rsIntroLine.BOF Then
            Do While Not rsIntroLine.EOF
                 'Pat 4/5: The initial value of Pressure0_or_temperature1 is set to Temperature
                dTempPresTemp = IIf(IsNull(rsIntroLine("Pressure0_or_temperature1")), 1, rsIntroLine("Pressure0_or_temperature1"))
                
                If blnChangeTab = True Then
                    'Pat 4/5: If Temparature is selected else Pressure
                    If blnTempPres = False Then
                        dTempPresTemp = 1
                        'Pat 4/5: If the user click on the other tabs without converting the values
                        If blnChangeTab = True Then
                            gblPresTempMode = True
                        End If
                    Else
                        dTempPresTemp = 0
                        If blnChangeTab = True Then
                            gblPresTempMode = False
                        End If
                    End If
                End If
                
                'Pat 4/5: If the value of Pressure0_or_temperature1 is 1 then Temperature then
                If dTempPresTemp = 1 Then
                    .rbtTemperature.Value = True
                    .rbtPressure.Value = False
                    If blnOpenFile = False And blnChangeTab = False Then
                        .btnConvert.Enabled = False
                    End If
                'Pat 4/5: If the value of Pressure0_or_temperature1 is Null then Temperature then
                ElseIf dTempPresTemp = Null Then
                    .rbtTemperature.Value = True
                    .rbtPressure.Value = False
                'Pat 4/5: Else Pressure
                Else
                    'If it is an exisiting file that has Pressure as its Cylinder
                    blnOpenFile = True
                        If blnChangeTab = True Then
                            blnOpenFile = False
                        End If
                    .rbtPressure.Value = True
                    .rbtTemperature.Value = False
                    If blnChangeTab = True Then
                        .btnConvert.Enabled = True
                    Else
                        .btnConvert.Enabled = False
                    End If
                End If
                
                Exit Do
            Loop
        Else
            'Pat 4/5: Initial values: Always Temperature
            If blnTempPres = False Then
                .rbtTemperature.Value = True
                .rbtPressure.Value = False
            Else
                .rbtPressure.Value = True
                .rbtTemperature.Value = False
            End If
            .btnConvert.Enabled = False
        End If

        '.btnConvert.Enabled = False
        .cmdCalculateIntroRate.Caption = LoadResString(IDS_CAP_INTRO_CALCULATE)
        
        'Mouse-over instructions
        .cmdCalculate.ToolTipText = LoadResString(IDS_TIP_INTRO_CALCULATE)
        .rtfWarnings.ToolTipText = LoadResString(IDS_TIP_INTRO_WARNINGS)
        frmWarningMsg.Caption = LoadResString(IDS_WARNINGS)

        If Not bHasIntroLine And blnTempPres = False Then
            .rbtPressure.Enabled = False
        End If
    End With
    
    Call LoadIntroGrid
    Call LoadWarnings
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "IntroPlanInitializeControls", Err.Description)
    End If

End Sub

Private Sub LoadIntroGrid()
'**************************************************************
'   DESCRIPTION
'       This initializes grdIntroInput and grdIntroRates.
'
'   PARAMETERS
'          NONE
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'   04/03/2003   J. Manyaga   Fumiguide Enhancements 2003 - Set Inside
'                             Diameter to 3 options only.
'
'**************************************************************

On Error GoTo Error_Handler
    
    Dim sJobAreaSQL As String
    Dim sAreaName As String
    Dim sInsideDiameter As String
    Dim nCount As Integer
    Dim rsJobArea As ADODB.Recordset
    
    Set rsJobArea = New ADODB.Recordset
    
    'Create the Job Area recordset
    sJobAreaSQL = "select id, area_name from JobArea where deleted = false and " & _
        "job_fkey = '" & gsJobID & "' order by row"
    rsJobArea.Open sJobAreaSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    'Initialize grdIntroInput
    With GridWrapInput
        'Initialize the grid
        .Initialize "IntroLine", frmMain.grdIntroInput
        
        'Create the columns in the grid
        .AddColumn "colIndex", " ", colTypeText, flexAlignCenterTop
        .AddColumn "colAreaName", LoadResString(IDS_AREA_NAME), colTypeCombo, _
            flexAlignLeftTop
        .AddColumn "colIntroLineName", LoadResString(IDS_SHOOTING_LINE_NAME), _
            colTypeText, flexAlignLeftTop
        .AddColumn "colRH", LoadResString(IDS_REL_HUMIDITY), colTypeText, _
            flexAlignLeftTop
        .AddColumn "colCylinderTemp", LoadResString(IDS_SHOOTINGLINE_CYLINDER_TEMPERATURE), _
            colTypeText, flexAlignLeftTop
        '.AddColumn "colInsideDiameter", LoadResString(IDS_INSIDE_DIAMETER), _
            colTypeText, flexAlignLeftTop
        '.AddColumn "colLength", LoadResString(IDS_LENGTH), colTypeText, flexAlignLeftTop
        .AddColumn "colOneFourthLength", LoadResString(IDS_ONE_FOURTH_LENGTH), _
            colTypeText, flexAlignLeftTop
        .AddColumn "colOneEightLength", LoadResString(IDS_ONE_EIGHT_LENGTH), colTypeText, flexAlignLeftTop
        .AddColumn "colFanCap", LoadResString(IDS_FAN), colTypeText, flexAlignLeftTop
        
        'This column is used to store the order/location of the record in the recordset.
        .AddColumn "colOrderRS", " ", colTypeText, flexAlignLeftTop
        
        'Pat 4/5: Number of Cylinder..When adding new column, it should placed after the last column
        .AddColumn "colNumCylinder", LoadResString(IDS_NUM_CYLINDER), colTypeText, flexAlignLeftTop
           
        ' set default column widths (will be overridden if previous settings were saved)
        With frmMain.grdIntroInput
            .ColWidth(1) = AREANAME_COLSIZE
            .ColWidth(2) = INTROLINE_COLSIZE
            .ColWidth(3) = RH_COLSIZE
            .ColWidth(4) = CYLTEMP_COLSIZE
            .ColWidth(5) = ONEFOURTH_LENGTH_COLSIZE
            .ColWidth(6) = ONEEIGHT_LENGTH_COLSIZE
            '.ColWidth(5) = DIAMETER_COLSIZE
            '.ColWidth(6) = LENGTH_COLSIZE
            .ColWidth(7) = FANCAP_COLSIZE
            .ColWidth(9) = NUM_CYLINDER_COLSIZE
        End With

        'Format the columns according to the user's settings in the registry.
        .LoadGridSettings GRID_SETTINGS_PATH
    End With
    
    'Create the combo box list options for the Area Name column
    sAreaName = ""
    rsJobArea.MoveFirst
    'Store the first Area name in a global variable. This will be used during grid updates.
    sFirstIntroAreaId = rsJobArea("id")
    sFirstIntroAreaName = rsJobArea("area_name")
    While Not rsJobArea.EOF
        sAreaName = sAreaName & "#" & rsJobArea("id") & ";" & rsJobArea("area_name") & "|"
        rsJobArea.MoveNext
    Wend
    sAreaName = VBA.Left(sAreaName, Len(sAreaName) - 1)
    frmMain.grdIntroInput.ColComboList(COL_AREA_NAME) = sAreaName
    
    'Create the combo box list options for the Inside Diameter column
    'JNM 03/31/03 - Fumiguide Enhancements 2003 - Set inside diameter dropdown to 3 options only
    'If Not gbIsMetric Then
    '    arrDiameter(0) = LoadResString(IDS_EIGHTH_INCH)
    '    arrDiameter(1) = LoadResString(IDS_ELEVEN_64TH_INCH)
    '    arrDiameter(2) = LoadResString(IDS_FOURTH_INCH)
    'Else
   '     arrDiameter(0) = FormatNumber(SHOOTINGLINE_DIAMETER_ONE, 3) & LoadResString(IDS_DIAMETER_UNIT)
   '     arrDiameter(1) = FormatNumber(SHOOTINGLINE_DIAMETER_TWO, 3) & LoadResString(IDS_DIAMETER_UNIT)
    '    arrDiameter(2) = FormatNumber(SHOOTINGLINE_DIAMETER_THREE, 3) & LoadResString(IDS_DIAMETER_UNIT)
    'End If
    
    'sInsideDiameter = ""
    'For nCount = 0 To 2
    '    sInsideDiameter = sInsideDiameter & "#" & nCount & ";" & arrDiameter(nCount) & "|"
    'Next
    'sInsideDiameter = VBA.Left(sInsideDiameter, Len(sInsideDiameter) - 1)
    'frmMain.grdIntroInput.ColComboList(COL_INSIDE_DIAMETER) = sInsideDiameter
    'sFirstInsideDiameter = arrDiameter(0)
    
    frmMain.grdIntroInput.ColHidden(COL_ORDER_RS) = True    'Hide the Order column
    frmMain.grdIntroInput.Rows = GRID_MAX_ROWS              'Initialize the number of rows
    
    'Initialize grdIntroRates
    With GridWrapOutput
        'Initialize the grid
        .Initialize "IntroRates", frmMain.grdIntroRates
        
        'Create the columns in the grid
        .AddColumn "colIndex", " ", colTypeText, flexAlignCenterTop
        .AddColumn "colAreaNameRO", LoadResString(IDS_AREA_NAME), colTypeCombo, _
            flexAlignLeftTop
        .AddColumn "colIntroLineNameRO", LoadResString(IDS_SHOOTING_LINE_NAME), _
            colTypeText, flexAlignLeftTop
        .AddColumn "colActualRate", LoadResString(IDS_SHOOTING_RATE), colTypeText, _
            flexAlignLeftTop
        .AddColumn "colPermitRate", LoadResString(IDS_MAX_SHOOTING_RATE), _
            colTypeText, flexAlignLeftTop
        .AddColumn "colTimeMin", LoadResString(IDS_TIME_MIN_EMPTY_CYLINDER), _
            colTypeText, flexAlignLeftTop
            
           
        ' default column widths (will be overridden if previous settings saveD)
        With frmMain.grdIntroRates
            .ColWidth(1) = AREANAME2_COLSIZE
            .ColWidth(2) = LINENAME_COLSIZE
            .ColWidth(3) = ACTUALRATE_COLSIZE
            .ColWidth(4) = PERMITRATE_COLSIZE
            .ColWidth(5) = TIMEMIN_COLSIZE
        End With
        
        'Format the columns according to the user's settings in the registry.
        .LoadGridSettings GRID_SETTINGS_PATH
    End With
    
    frmMain.grdIntroRates.ColComboList(COL_AREA_NAME) = sAreaName
    frmMain.grdIntroRates.Rows = GRID_MAX_ROWS              'Initialize the number of rows
    
    frmMain.grdIntroInput.FixedCols = 1
    frmMain.grdIntroRates.FixedCols = 1
    
    'Initialize the colors of the grids
    With frmMain.grdIntroInput
        .GridColor = GRID_COLOR
        '.BackColorAlternate = BACK_COLOR_ALTERNATE
        .ColWidth(COL_INDEX) = COL_INDEX_WIDTH
    End With
    With frmMain.grdIntroRates
        .GridColor = GRID_COLOR
        .BackColor = READ_ONLY_COLOR
        .ColWidth(COL_INDEX) = COL_INDEX_WIDTH
        'Set the color of the read-only cells
        'For nCount = 1 To GRID_MAX_ROWS - 1 Step 2
        '    .Cell(flexcpBackColor, nCount, 1, nCount, 4) = BACK_COLOR_ALTERNATE
        'Next
    End With
    
    'Fill the grids with the database values
    Call LoadIntroUnits
    Call LoadIntroRS
    
    'Clean-up time
    rsJobArea.Close
    Set rsJobArea = Nothing
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "LoadIntroGrid", Err.Description)
    End If

End Sub

Private Sub LoadIntroUnits()
'**************************************************************
'   DESCRIPTION
'       This stores the measurement units in the flexcpData
'       attribute of the column headings.
'
'   PARAMETERS
'          NONE
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************

On Error GoTo Error_Handler
    
    Dim dTemperature As Double
    Dim sTempUnit As String
    Dim dLength As Double
    Dim sLengthUnit As String
    Dim dActualRate As Double
    Dim sActualRateUnit As String
    Dim dPermitRate As Double
    Dim sPermitRateUnit As String
    Dim dOneFourth As Double
    Dim dOneEight As Double
    Dim dTimeEmptyCylinder As Double
    Dim sTimeEmptyUnit As String
    Dim dPressure As Double
    Dim sPressureUnit As String
    
    With frmMain.grdIntroInput
        .Cell(flexcpData, 0, COL_RH) = "%"
        dTemperature = 0#
         sTempUnit = ""
         Call UnitConvert(modeDisplay, unitTEMPERATURE, 1.5, dTemperature, sTempUnit)
         .Cell(flexcpData, 0, COL_CYLINDER_TEMP) = "°" & sTempUnit
         
         dPressure = 0#
        sPressureUnit = ""
        .Cell(flexcpData, 0, COL_CYLINDER_TEMP) = " " & sPressureUnit
        Call UnitConvert(modeDisplay, unitPressure, 1.5, dPressure, sPressureUnit)
        

            
        'Pat 12/28/2004
        'dLength = 0#
        'sLengthUnit = ""
        'Call UnitConvert(modeDisplay, unitLWH, 1.5, dLength, sLengthUnit)
        '.Cell(flexcpData, 0, COL_ONE_FOURTH) = " " & sLengthUnit
        
        dOneFourth = 0#
        sLengthUnit = ""
        Call UnitConvert(modeDisplay, unitLWH, 1.5, dOneFourth, sLengthUnit)
        .Cell(flexcpData, 0, COL_ONE_FOURTH) = " " & sLengthUnit
        
        dOneEight = 0#
        sLengthUnit = ""
        Call UnitConvert(modeDisplay, unitLWH, 1.5, dOneEight, sLengthUnit)
        .Cell(flexcpData, 0, COL_ONE_EIGHT) = " " & sLengthUnit
            
        .Cell(flexcpData, 0, COL_FAN_CAP) = " " & IIf(gbIsMetric, _
            LoadResString(IDS_CUBIC_METERS_PER_MINUTE), LoadResString(IDS_CUBIC_FEET_PER_MINUTE))
    End With
    
    With frmMain.grdIntroRates
        dActualRate = 0#
        sActualRateUnit = ""
        Call UnitConvert(modeDisplay, unitSHOOTINGRATE, 1.5, dActualRate, sActualRateUnit)
        .Cell(flexcpData, 0, COL_ACTUAL_RATE) = " " & sActualRateUnit
        
        dPermitRate = 0#
        sPermitRateUnit = ""
        Call UnitConvert(modeDisplay, unitSHOOTINGRATE, 1.5, dPermitRate, sPermitRateUnit)
        .Cell(flexcpData, 0, COL_PERMIT_RATE) = " " & sPermitRateUnit
        
        dTimeEmptyCylinder = 0#
        sTimeEmptyUnit = ""
        Call UnitConvert(modeDisplay, unitTIME, 1.5, dTimeEmptyCylinder, sTimeEmptyUnit)
        .Cell(flexcpData, 0, COL_TIME_EMPTY_CYL) = " " & sTimeEmptyUnit
    End With
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "LoadIntroUnits", Err.Description)
    End If

End Sub

Public Sub LoadIntroRS()
'**************************************************************
'   DESCRIPTION
'       This displays the database values in grdIntroInput.
'
'   PARAMETERS
'          NONE
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'   04/03/03     J. Manyaga     Fumiguide Enhancements 2003 - RH will be blank
'                               if the value in the database is -1.
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim sIntroLineSQL As String
    Dim nCurrRow As Integer
    Dim dTemperature As Double
    Dim dLength As Double
    Dim dFanCap As Double
    Dim dActualRate As Double
    Dim dPermitRate As Double
    Dim nCount As Integer
    Dim nOption As Integer
    Dim dOneFourth As Double
    Dim dOneEight As Double
    Dim dNumCylinder As Double
    Dim dTemp4 As Double
    Dim dTemp8 As Double
    Dim dTempNumCylinder As Double
    Dim dTimeEmptyCylinder As Double
    Dim dTempTimeEmptyCylinder As Double
    Dim dPressure As Double
    Dim dTempunit As String
    Dim dPressureUnit As String
    Dim I
 
    
    'Clear the grid of its contents
    'Pat 4/5: Added new (1) column so the number of columns to be shown is changed from 8 to 9
    frmMain.grdIntroInput.Cell(flexcpData, 1, 0, GRID_MAX_ROWS - 1, 9) = ""
    frmMain.grdIntroInput.Cell(flexcpText, 1, 0, GRID_MAX_ROWS - 1, 9) = ""
    frmMain.grdIntroRates.Cell(flexcpData, 1, 0, GRID_MAX_ROWS - 1, 5) = ""
    frmMain.grdIntroRates.Cell(flexcpText, 1, 0, GRID_MAX_ROWS - 1, 5) = ""
    
    'Reset the grid colors AND font
    'TSG 11/25 - Reset the foreground colors to black
    With frmMain
        .grdIntroInput.Cell(flexcpForeColor, 1, COL_AREA_NAME, 100, .grdIntroInput.Cols - 1) = &H0
        .grdIntroInput.Cell(flexcpBackColor, 1, COL_AREA_NAME, 100, .grdIntroInput.Cols - 1) = &H80000005
        .grdIntroInput.Cell(flexcpFontBold, 1, COL_AREA_NAME, 100, .grdIntroInput.Cols - 1) = False
        .grdIntroRates.Cell(flexcpForeColor, 1, COL_AREA_NAME, 100, .grdIntroRates.Cols - 1) = &H0
        .grdIntroRates.Cell(flexcpBackColor, 1, COL_AREA_NAME, 100, .grdIntroRates.Cols - 1) = &H8000000B
        .grdIntroRates.Cell(flexcpFontBold, 1, COL_AREA_NAME, 100, .grdIntroRates.Cols - 1) = False
        
        If blnChangeTab = True Then
            For I = 1 To 9
                frmMain.grdIntroInput.Cell(flexcpForeColor, 1, I, 100, frmMain.grdIntroInput.Cols - 1) = &H0
                frmMain.grdIntroInput.Cell(flexcpBackColor, 1, I, 100, frmMain.grdIntroInput.Cols - 1) = &H8000000B
            Next
        End If
    End With
    
    'Initialize the string containing the list of Introduction Line Names
    sGridIntroLineName = ""
    
    'Display the values of the index column
    For nCount = 1 To GRID_MAX_ROWS - 1
        frmMain.grdIntroInput.Cell(flexcpText, nCount, COL_INDEX) = nCount
        frmMain.grdIntroRates.Cell(flexcpText, nCount, COL_INDEX) = nCount
    Next
    
    'Check if the Introduction Line recordset is already open.
    'Close it first before re-opening
    If rsIntroLine.State Then
        rsIntroLine.Close
    End If
    
    'Create the Introduction Line recordset
    sIntroLineSQL = "select ShootingLine.* from ShootingLine left outer join " & _
        "JobArea on ShootingLine.jobarea_fkey = JobArea.id where " & _
        "ShootingLine.deleted = false and ShootingLine.job_fkey = '" & gsJobID & "'" & _
        "order by ShootingLine.row"
    
    rsIntroLine.Open sIntroLineSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    

    
    'Display the recordset in the grid if the recordset is not empty. Assign to flexcpText
    'the value and the corresponding units. Assign the value only to flexcpData.
    nCurrRow = 1            'the current row in the grid
    nRecordNumber = 1       'the current record number
    If Not rsIntroLine.BOF And Not rsIntroLine.EOF Then
        bHasIntroLine = True  'set global boolean
        With frmMain
        rsIntroLine.MoveFirst
        While Not rsIntroLine.EOF
            If bHasIntroLine = True And gblPresTempMode = False Then
                frmMain.rbtPressure.Enabled = True
            End If
                   
            'Display only if we are in the correct grid location.
            If rsIntroLine("row") = nCurrRow Then
                'Area Name
                .grdIntroInput.Cell(flexcpData, nCurrRow, COL_AREA_NAME) = CStr(rsIntroLine("jobarea_fkey"))
                .grdIntroInput.Cell(flexcpText, nCurrRow, COL_AREA_NAME) = rsIntroLine("jobarea_fkey")
                .grdIntroRates.Cell(flexcpText, nCurrRow, COL_AREA_NAME) = rsIntroLine("jobarea_fkey")
                
                'Inroduction Line Name
                .grdIntroInput.Cell(flexcpData, nCurrRow, COL_INTROLINE_NAME) = CStr(rsIntroLine("name"))
                .grdIntroInput.Cell(flexcpText, nCurrRow, COL_INTROLINE_NAME) = rsIntroLine("name")
                .grdIntroRates.Cell(flexcpText, nCurrRow, COL_INTROLINE_NAME) = rsIntroLine("name")
                sGridIntroLineName = sGridIntroLineName & "|" & rsIntroLine("name") & "|"
                
                'Relative Humidity
                'JNM 03/31/03 - FE 2003 - RH can be blank, which we will represent by -1.
                If rsIntroLine("humidity") <> -1 Then
                    .grdIntroInput.Cell(flexcpData, nCurrRow, COL_RH) = FormatNumber(rsIntroLine("humidity"), 0)
                    .grdIntroInput.Cell(flexcpText, nCurrRow, COL_RH) = rsIntroLine("humidity") & _
                        .grdIntroInput.Cell(flexcpData, 0, COL_RH)
                Else
                    .grdIntroInput.Cell(flexcpData, nCurrRow, COL_RH) = ""
                    .grdIntroInput.Cell(flexcpText, nCurrRow, COL_RH) = ""
                End If

                'Cylinder Temperature/Pressure
                If .rbtTemperature.Value = True Then
                    dTemperature = 0#
                    'Pat 4/5: If the user selected Temperature
                    If gblPresTempMode = False Then
                        Call UnitConvert(modeDisplay, unitTEMPERATURE, rsIntroLine("cylinder_temperature"), _
                            dTemperature, dTempunit)
                    'Pat 4/5: If the user tries to change tabs while not performing the conversion of data, it will capture the values
                    'of the values which the user wanted to convert
                    ElseIf blnChangeTab = True Then
                        dPressure = rsIntroLine("cylinder_temperature")
                        dPressureUnit = IIf(gbIsMetric, LoadResString(IDS_PRESSURE_BAR), LoadResString(IDS_PRESSURE_PSIG))
                        Call UnitConvert(modeDisplay, unitPressure, rsIntroLine("cylinder_temperature"), _
                        dPressure, dPressureUnit)
                    'Pat 4/5: Else if the click on the Pressure Radio button and perfoem the Conversion of Temperature to pressure
                    Else
                        dTemperature = rsIntroLine("cylinder_temperature")
                        dTempunit = IIf(gbIsMetric, LoadResString(IDS_CELSIUS), LoadResString(IDS_FAHRENHEIT))
                    End If
                    
                    'Pat 4/5: If the user tries to change tabs while not performing the conversion of data, it will capture the values
                    'of the values which the user wanted to convert
                    If blnChangeTab = True Then
                        .grdIntroInput.Cell(flexcpData, 0, COL_CYLINDER_TEMP) = " " & dPressureUnit
                        dTemperature = Round(dPressure, IIf(gbIsMetric, 1, 0))
                        .grdIntroInput.Cell(flexcpData, nCurrRow, COL_CYLINDER_TEMP) = _
                        FormatNumber(dPressure, IIf(gbIsMetric, 1, 0))
                        .grdIntroInput.Cell(flexcpText, nCurrRow, COL_CYLINDER_TEMP) = _
                        FormatNumber(dPressure, IIf(gbIsMetric, 1, 0)) & _
                        .grdIntroInput.Cell(flexcpData, 0, COL_CYLINDER_TEMP)
                    'Pat 4/5: Else normal procedure
                    Else
                        .grdIntroInput.Cell(flexcpData, 0, COL_CYLINDER_TEMP) = " " & "°" & dTempunit
                        dTemperature = Round(dTemperature, IIf(gbIsMetric, 1, 0))
                        .grdIntroInput.Cell(flexcpData, nCurrRow, COL_CYLINDER_TEMP) = _
                        FormatNumber(dTemperature, IIf(gbIsMetric, 1, 0))
                        .grdIntroInput.Cell(flexcpText, nCurrRow, COL_CYLINDER_TEMP) = _
                        FormatNumber(dTemperature, IIf(gbIsMetric, 1, 0)) & _
                        .grdIntroInput.Cell(flexcpData, 0, COL_CYLINDER_TEMP)
                    End If
                    'Pat 4/5: Indicates that Temperature is selected
                    blnTempPres = False
                    dTempPresTemp = 1#
                    'Pat 4/5: Intro Input grid are enabled/editable
                    blnEnable = True
                Else
                    dPressure = 0#
                    'Pat 4/5: If the user selected Pressure
                    If gblPresTempMode = True Then
                        Call UnitConvert(modeDisplay, unitPressure, rsIntroLine("cylinder_temperature"), _
                        dPressure, dPressureUnit)
                    'Pat 4/5: If the user tries to change tabs while not performing the conversion of data, it will capture the values
                    'of the values which the user wanted to convert
                    ElseIf blnChangeTab = True Then
                        dTemperature = rsIntroLine("cylinder_temperature")
                        dTempunit = IIf(gbIsMetric, LoadResString(IDS_CELSIUS), LoadResString(IDS_FAHRENHEIT))
                        Call UnitConvert(modeDisplay, unitTEMPERATURE, rsIntroLine("cylinder_temperature"), _
                        dTemperature, dTempunit)
                    'Pat 4/5: Else if the click on the Temperature Radio button and perfoem the Conversion of pressure to temperature
                    Else
                        dPressure = rsIntroLine("cylinder_temperature")
                        dPressureUnit = IIf(gbIsMetric, LoadResString(IDS_PRESSURE_BAR), LoadResString(IDS_PRESSURE_PSIG))
                        Call UnitConvert(modeDisplay, unitPressure, rsIntroLine("cylinder_temperature"), _
                        dPressure, dPressureUnit)
                    End If
               
                    'Pat 4/5: If the user tries to change tabs while not performing the conversion of data, it will capture the values
                    'of the values which the user wanted to convert
                    If blnChangeTab = True Then
                        .grdIntroInput.Cell(flexcpData, 0, COL_CYLINDER_TEMP) = " " & "°" & dTempunit
                        dTemperature = Round(dTemperature, IIf(gbIsMetric, 1, 0))
                        .grdIntroInput.Cell(flexcpData, nCurrRow, COL_CYLINDER_TEMP) = _
                        FormatNumber(dTemperature, IIf(gbIsMetric, 1, 0))
                        .grdIntroInput.Cell(flexcpText, nCurrRow, COL_CYLINDER_TEMP) = _
                        FormatNumber(dTemperature, IIf(gbIsMetric, 1, 0)) & _
                        .grdIntroInput.Cell(flexcpData, 0, COL_CYLINDER_TEMP)
                    'Pat 4/5: Else normal procedure
                    Else
                        .grdIntroInput.Cell(flexcpData, 0, COL_CYLINDER_TEMP) = " " & dPressureUnit
                        dTemperature = Round(dPressure, IIf(gbIsMetric, 1, 0))
                        .grdIntroInput.Cell(flexcpData, nCurrRow, COL_CYLINDER_TEMP) = _
                        FormatNumber(dPressure, IIf(gbIsMetric, 1, 0))
                        .grdIntroInput.Cell(flexcpText, nCurrRow, COL_CYLINDER_TEMP) = _
                        FormatNumber(dPressure, IIf(gbIsMetric, 1, 0)) & _
                        .grdIntroInput.Cell(flexcpData, 0, COL_CYLINDER_TEMP)
                    End If
                    'Pat 4/5: Indicates that Temperature is selected
                    blnTempPres = True
                    dTempPresTemp = 0#
                    'Pat 4/5: Intro Input grid are enabled/editable
                    blnEnable = True
                End If
          
                'Inside Diameter
                'JNM 03/31/03 - Fumiguide Enhancements 2003 - There are only 3 options for the
                '   inside diameter drop-down control.
                'nOption = 0
                'Select Case CDbl(rsIntroLine("diameter"))
                '    Case SHOOTINGLINE_DIAMETER_ONE
                '        nOption = 0
                '    Case SHOOTINGLINE_DIAMETER_TWO
                '        nOption = 1
                '    Case SHOOTINGLINE_DIAMETER_THREE
                '        nOption = 2
                '    Case Else
                '        nOption = 0
                'End Select
                '.grdIntroInput.Cell(flexcpData, nCurrRow, COL_INSIDE_DIAMETER) = nOption
                '.grdIntroInput.Cell(flexcpText, nCurrRow, COL_INSIDE_DIAMETER) = nOption
                    
                'Length
                'dLength = 0#
                'Call UnitConvert(modeDisplay, unitLWH, rsIntroLine("length"), _
                '        dLength)
                'dLength = Round(dLength, 0)
                '.grdIntroInput.Cell(flexcpData, nCurrRow, COL_LENGTH) = FormatNumber(dLength, 0)
                '.grdIntroInput.Cell(flexcpText, nCurrRow, COL_LENGTH) = _
                '    FormatNumber(dLength, 0) & .grdIntroInput.Cell(flexcpData, 0, COL_LENGTH)
                
                'One Fourth Length
                dOneFourth = 0#
                dTemp4 = IIf(IsNull(rsIntroLine("one_fourth_length")), 0, rsIntroLine("one_fourth_length"))
                Call UnitConvert(modeDisplay, unitLWH, dTemp4, _
                        dOneFourth)
                dOneFourth = Round(dOneFourth, 0)
                .grdIntroInput.Cell(flexcpData, nCurrRow, COL_ONE_FOURTH) = FormatNumber(dOneFourth, 0)
                .grdIntroInput.Cell(flexcpText, nCurrRow, COL_ONE_FOURTH) = _
                    FormatNumber(dOneFourth, 0) & .grdIntroInput.Cell(flexcpData, 0, COL_ONE_FOURTH)
                    
                'One Eight Length
                dOneEight = 0#
                dTemp8 = IIf(IsNull(rsIntroLine("one_eight_length")), 0, rsIntroLine("one_eight_length"))
                Call UnitConvert(modeDisplay, unitLWH, dTemp8, _
                        dOneEight)
                dOneEight = Round(dOneEight, 0)
                .grdIntroInput.Cell(flexcpData, nCurrRow, COL_ONE_EIGHT) = FormatNumber(dOneEight, 0)
                .grdIntroInput.Cell(flexcpText, nCurrRow, COL_ONE_EIGHT) = _
                    FormatNumber(dOneEight, 0) & .grdIntroInput.Cell(flexcpData, 0, COL_ONE_EIGHT)
              
                    
                'Fan Capacity
                dFanCap = 0#
                Call UnitConvert(modeDisplay, unitVOLUME, rsIntroLine("fan"), _
                        dFanCap)
                dFanCap = Round(dFanCap, 0)
                .grdIntroInput.Cell(flexcpData, nCurrRow, COL_FAN_CAP) = FormatNumber(dFanCap, 0, , , vbFalse)
                .grdIntroInput.Cell(flexcpText, nCurrRow, COL_FAN_CAP) = _
                    FormatNumber(dFanCap, 0, , , vbFalse) & .grdIntroInput.Cell(flexcpData, 0, COL_FAN_CAP)
                
                'Number of Cylinders
                 dNumCylinder = 0#
                 dTempNumCylinder = IIf(IsNull(rsIntroLine("numcylinder")), 0, rsIntroLine("numcylinder"))
                 Call UnitConvert(modeDisplay, unitLess, dTempNumCylinder, _
                        dNumCylinder)
                .grdIntroInput.Cell(flexcpData, nCurrRow, COL_NUM_CYLINDER) = FormatNumber(dNumCylinder, 0, , , vbFalse)
                .grdIntroInput.Cell(flexcpText, nCurrRow, COL_NUM_CYLINDER) = _
                    FormatNumber(dNumCylinder, 0, , , vbFalse) & .grdIntroInput.Cell(flexcpData, 0, COL_NUM_CYLINDER)

                               
                'Input Recordset order
                .grdIntroInput.Cell(flexcpData, nCurrRow, COL_ORDER_RS) = nRecordNumber
                .grdIntroInput.Cell(flexcpText, nCurrRow, COL_ORDER_RS) = nRecordNumber   'Temporary
                
                'Actual Introduction Rate
                dActualRate = 0#
                Call UnitConvert(modeDisplay, unitSHOOTINGRATE, rsIntroLine("rate"), dActualRate)
                dActualRate = Round(dActualRate, 1)
              
                 
                'JNM 05/14/03 - PT fix - Display a message if value is < 0, instead of
                '   the value
                If rsIntroLine("calculation_warning") = LoadResString(IDS_INTRO_RATE_ZERO) Then
                    .grdIntroRates.Cell(flexcpText, nCurrRow, COL_ACTUAL_RATE) = LoadResString(IDS_INTRO_RATE_ERROR)
                Else
                    .grdIntroRates.Cell(flexcpText, nCurrRow, COL_ACTUAL_RATE) = _
                        FormatNumber(dActualRate, 1, , , vbFalse) & .grdIntroRates.Cell(flexcpData, 0, COL_ACTUAL_RATE)
                End If
                
                'Permitted Introduction Rate
                dPermitRate = 0#
                Call UnitConvert(modeDisplay, unitSHOOTINGRATE, rsIntroLine("max_rate"), dPermitRate)
                dPermitRate = Round(dPermitRate, 1)
                .grdIntroRates.Cell(flexcpText, nCurrRow, COL_PERMIT_RATE) = _
                    FormatNumber(dPermitRate, 1, , , vbFalse) & .grdIntroRates.Cell(flexcpData, 0, COL_PERMIT_RATE)

                'Pat 3/15/2005: Time Minimum To Empty Cylinder
                dTimeEmptyCylinder = 0#
                dTempTimeEmptyCylinder = IIf(IsNull(rsIntroLine("time_min_empty_cylinder")), 0, rsIntroLine("time_min_empty_cylinder"))
                Call UnitConvert(modeDisplay, unitTIME, dTempTimeEmptyCylinder, dTimeEmptyCylinder)
                If .grdIntroRates.Cell(flexcpText, nCurrRow, COL_ACTUAL_RATE) = LoadResString(IDS_INTRO_RATE_ERROR) Then
                    dTimeEmptyCylinder = 0#
                    .grdIntroRates.Cell(flexcpText, nCurrRow, COL_TIME_EMPTY_CYL) = _
                    FormatNumber(dTimeEmptyCylinder, 1, , , vbFalse) & .grdIntroRates.Cell(flexcpData, 0, COL_TIME_EMPTY_CYL)
                Else
                    .grdIntroRates.Cell(flexcpText, nCurrRow, COL_TIME_EMPTY_CYL) = _
                    FormatNumber(dTimeEmptyCylinder, 1, , , vbFalse) & .grdIntroRates.Cell(flexcpData, 0, COL_TIME_EMPTY_CYL)
                End If
                            
                'JNM 11/15 - set the background color to red if record has a warning
                'TSG 11/26 - set foreground color to white for warnings
                If rsIntroLine("calculation_warning") <> "" Then
                    'red and bold
                    'Pat 4/5/2005 -  Removed red back color
'                    .grdIntroInput.Cell(flexcpBackColor, nCurrRow, COL_AREA_NAME, nCurrRow, _
'                        .grdIntroInput.Cols - 1) = &HFF&
                    .grdIntroInput.Cell(flexcpForeColor, nCurrRow, COL_AREA_NAME, nCurrRow, _
                        .grdIntroInput.Cols - 1) = &H0&
                    '.grdIntroInput.Cell(flexcpFontBold, nCurrRow, COL_AREA_NAME, nCurrRow, _
                        .grdIntroInput.Cols - 1) = True
'                    .grdIntroRates.Cell(flexcpBackColor, nCurrRow, COL_AREA_NAME, nCurrRow, _
'                        .grdIntroRates.Cols - 1) = &HFF&
                    .grdIntroRates.Cell(flexcpForeColor, nCurrRow, COL_AREA_NAME, nCurrRow, _
                        .grdIntroRates.Cols - 1) = &H0&
                    '.grdIntroRates.Cell(flexcpFontBold, nCurrRow, COL_AREA_NAME, nCurrRow, _
                        .grdIntroRates.Cols - 1) = True
                End If
                                          
                'Increment the record number and move to the next record
                nRecordNumber = nRecordNumber + 1
                rsIntroLine.MoveNext
                                
            End If
            'Increment nCurrRow in order to move to the next grid row
            nCurrRow = nCurrRow + 1
        Wend
        End With
    Else
        bHasIntroLine = False
    End If
    blnChangeTab = False
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "LoadIntroRS", Err.Description)
    End If

End Sub

Public Sub LoadWarnings(Optional ByRef sWarnings As String)
'**************************************************************
'   DESCRIPTION
'       This displays the Warnings in the text box.
'
'   PARAMETERS
'          NONE
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim sIntroWarnings As String
    Dim arrIntroWarnings(1) As String
    'Dim sWarnings As String
    
    'Create the Warnings string.
    If bHasIntroLine Then
        rsIntroLine.MoveFirst
        While Not rsIntroLine.EOF
            If Len(rsIntroLine("calculation_warning")) > 0 Then
                sIntroWarnings = ""
                arrIntroWarnings(1) = rsIntroLine("calculation_warning")
                arrIntroWarnings(0) = rsIntroLine("name") & ": "
                sIntroWarnings = Join(arrIntroWarnings, "")
                sIntroWarnings = Replace(sIntroWarnings, vbCrLf, vbCrLf & rsIntroLine("name") & ": ")
                If sWarnings <> "" Then
                    sWarnings = sWarnings & vbCrLf
                End If
                sWarnings = sWarnings & sIntroWarnings & vbCrLf
            End If
            rsIntroLine.MoveNext
        Wend
    End If
    
    'Display and format the Results string in the text box
    With frmMain.rtfWarnings
        'Display the value
        .TextRTF = sWarnings
        
        'Format the entire text
        .SelStart = 0
        .SelLength = Len(sWarnings)
        .SelFontSize = RTF_FONT_SIZE
        .SelFontName = RTF_FONT_NAME
        'Pat 3/23/2005 - Removed Font Color
        '.SelColor = &HFF&
            
        .SelStart = 0
        .SelLength = 0
        
    End With
    

    
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "LoadWarnings", Err.Description)
    End If

End Sub

Public Function InputGridValidate(ByVal lRow As Long, ByVal lCol As Long) As Boolean
'**************************************************************
'   DESCRIPTION
'       This function validates the data entered in the cell.
'
'   PARAMETERS
'           lRow - the row number of the cell that is being edited
'           lCol - the column number of the cell that is being edited
'
'   RETURNS - True - if cell value is valid
'             False - if cell value is invalid
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'   04/03/03      J. Manyaga    Fumiguide Enhancements 2003 - Added a condition
'                               that would accept blank humidity values.
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim lHumidity As Long
    Dim dMinTemp As Double
    Dim dMaxTemp As Double
    Dim dTemperature As Double
    Dim dMinLength As Double
    Dim dMaxLength As Double
    Dim dLength As Double
    Dim dMinFanCap As Double
    Dim dMaxFanCap As Double
    Dim dFanCap As Double
    Dim dOneFourth As Double
    Dim dOneEight As Double
    Dim dNumCylinder As Integer
    Dim dPressure As Double
    Dim dMinPres As Double
    Dim dMaxPres As Double
    
    With frmMain.grdIntroInput
    
    Select Case lCol
        Case COL_INTROLINE_NAME
            'The Introduction Line Name should not be null
            If Len(Trim(.EditText)) = 0 Then
                MsgBox LoadResString(IDS_SHOOTINGLINE_NAME_EMPTY), vbExclamation
                InputGridValidate = False
            Else
                'There should be no duplicate Introduction Line Names for the Job
                If .EditText = .Cell(flexcpData, lRow, lCol) Then
                    InputGridValidate = True
                Else
                    If InStr(sGridIntroLineName, "|" & Trim(.EditText) & "|") > 0 Then
                        MsgBox LoadResString(IDS_DUPLICATE_INTRO_NAME), vbExclamation
                        InputGridValidate = False
                    Else
                        InputGridValidate = True
                    End If
                End If
            End If
        Case COL_RH
            'JNM 03/31/03 - FE 2003 - RH can be blank.
            If Len(Trim(.EditText)) = 0 Then
                InputGridValidate = True
            Else
                'Relative Humidity should be numeric
                If Not IsNumeric(Trim(.EditText)) Then
                    MsgBox FormatSWParams(LoadResString(IDS_AREA_INFO_HUMIDITY_NOT_A_NUMBER), _
                        MIN_HUMIDITY & "|" & MAX_HUMIDITY), vbExclamation
                    InputGridValidate = False
                Else
                    lHumidity = CLng(Trim(.EditText))
                    'Test if value is an integer
                    If Not IsInteger(Trim(.EditText)) Then
                        MsgBox LoadResString(IDS_HUMIDITY_NOT_INTEGER), vbExclamation
                        InputGridValidate = False
                    Else
                        'Relative Humidity should be within the allowable range
                        If lHumidity < MIN_HUMIDITY Or lHumidity > MAX_HUMIDITY Then
                            MsgBox FormatSWParams(LoadResString(IDS_AREA_INFO_HUMIDITY_OUT_OF_RANGE), _
                                MIN_HUMIDITY & "|" & MAX_HUMIDITY), vbExclamation
                            InputGridValidate = False
                        Else
                            InputGridValidate = True
                        End If
                    End If
                End If
            End If
        Case COL_CYLINDER_TEMP
            'Temperature else Pressure
            If frmMain.rbtTemperature.Value = True Then
                dMinTemp = IIf(gbIsMetric, MIN_TEMP_METRIC, MIN_TEMP_ENGLISH)
                dMaxTemp = IIf(gbIsMetric, MAX_TEMP_METRIC, MAX_TEMP_ENGLISH)
                'Temperature should be numeric
                If Len(Trim(.EditText)) = 0 Or Not IsNumeric(Trim(.EditText)) Then
                    MsgBox FormatSWParams(LoadResString(IDS_SHOOTINGLINE_CYLINDER_TEMPERATURE_NOT_A_NUMBER), _
                        dMinTemp & "|" & dMaxTemp & "|" & .Cell(flexcpData, 0, COL_CYLINDER_TEMP)), vbExclamation
                    InputGridValidate = False
                Else
                    'If we are in English units, temperature should be a whole number
                    dTemperature = Round(CDbl(Trim(.EditText)), 0)
                    If Not gbIsMetric And (dTemperature - CDbl(Trim(.EditText))) <> 0 Then
                        MsgBox LoadResString(IDS_TEMP_WHOLE_NUM), vbExclamation
                        InputGridValidate = False
                    Else
                        'Temperature should be within the allowable range
                        dTemperature = CDbl(Trim(.EditText))
                        If dTemperature < dMinTemp Or dTemperature > dMaxTemp Then
                            MsgBox FormatSWParams(LoadResString(IDS_SHOOTINGLINE_CYLINDER_TEMPERATURE_OUT_OF_RANGE), _
                                dMinTemp & "|" & dMaxTemp & "|" & .Cell(flexcpData, 0, COL_CYLINDER_TEMP)), vbExclamation
                            InputGridValidate = False
                        Else
                            InputGridValidate = True
                        End If
                    End If
                End If
            Else
                dMinPres = IIf(gbIsMetric, MIN_PRES_METRIC, MIN_PRES_ENGLISH)
                dMaxPres = IIf(gbIsMetric, MAX_PRES_METRIC, MAX_PRES_ENGLISH)
                'Pressure should be numeric
                If Len(Trim(.EditText)) = 0 Or Not IsNumeric(Trim(.EditText)) Then
                    MsgBox FormatSWParams(LoadResString(IDS_SHOOTINGLINE_CYLINDER_PRESSURE_OUT_OF_RANGE), _
                        dMinPres & "|" & dMaxPres & "|" & .Cell(flexcpData, 0, COL_CYLINDER_TEMP)), vbExclamation
                    InputGridValidate = False
                Else
                    'If we are in English units, Pressure should be a whole number
                    dPressure = Round(CDbl(Trim(.EditText)), 0)
                    If Not gbIsMetric And (dPressure - CDbl(Trim(.EditText))) <> 0 Then
                        MsgBox LoadResString(IDS_PRES_WHOLE_NUM), vbExclamation
                        InputGridValidate = False
                    Else
                        'Pressure should be within the allowable range
                        dPressure = CDbl(Trim(.EditText))
                        If dPressure < dMinPres Or dPressure > dMaxPres Then
                            MsgBox FormatSWParams(LoadResString(IDS_SHOOTINGLINE_CYLINDER_PRESSURE_NOT_A_NUMBER), _
                                dMinPres & "|" & dMaxPres & "|" & .Cell(flexcpData, 0, COL_CYLINDER_TEMP)), vbExclamation
                            InputGridValidate = False
                        Else
                            InputGridValidate = True
                        End If
                    End If
                End If
            End If
        'Case COL_LENGTH
        '    dMinLength = IIf(gbIsMetric, MIN_LINE_LENGTH_METRIC, MIN_LINE_LENGTH_ENGLISH)
        '    dMaxLength = IIf(gbIsMetric, MAX_LINE_LENGTH_METRIC, MAX_LINE_LENGTH_ENGLISH)
        '
        '    'Length should be numeric
        '    If Len(Trim(.EditText)) = 0 Or Not IsNumeric(Trim(.EditText)) Then
        '        MsgBox FormatSWParams(LoadResString(IDS_SHOOTINGLINE_LENGTH_INVALID), _
        '            FormatNumber(dMinLength, nDecimal) & "|" & FormatNumber(dMaxLength, nDecimal) _
        '            & "|" & .Cell(flexcpData, 0, COL_LENGTH)), vbExclamation
        '        InputGridValidate = False
        '    Else
        '        'If we are in English units, length should be a whole number
        '        dLength = Round(CDbl(Trim(.EditText)), 0)
        '        If Not gbIsMetric And (dLength - CDbl(Trim(.EditText))) <> 0 Then
        '            MsgBox LoadResString(IDS_LENGTH_WHOLE_NUM), vbExclamation
        '            InputGridValidate = False
        '        Else
        '            'Length should be within the allowable range
        '            dLength = CDbl(Trim(.EditText))
        '            If dLength < dMinLength Or dLength > dMaxLength Then
        '                MsgBox FormatSWParams(LoadResString(IDS_SHOOTINGLINE_LENGTH_INVALID), _
        '                    FormatNumber(dMinLength, nDecimal) & "|" & _
        '                    FormatNumber(dMaxLength, nDecimal) & "|" & .Cell(flexcpData, 0, COL_LENGTH)), vbExclamation
        '                InputGridValidate = False
        '            Else
        '                InputGridValidate = True
        '            End If
        '        End If
        '    End If
        Case COL_ONE_FOURTH
            dMinLength = IIf(gbIsMetric, MIN_LENGTH_ONEFOURTH_METRIC, MIN_LENGTH_ONEFOURTH_ENGLISH)
            dMaxLength = IIf(gbIsMetric, MAX_LENGTH_ONEFOURTH_METRIC, MAX_LENGTH_ONEFOURTH_ENGLISH)
        
            'Length should be numeric
            If Len(Trim(.EditText)) = 0 Or Not IsNumeric(Trim(.EditText)) Then
                MsgBox FormatSWParams(LoadResString(IDS_SHOOTINGLINE_LENGTH_INVALID), _
                    FormatNumber(dMinLength, nDecimal) & "|" & FormatNumber(dMaxLength, nDecimal) _
                    & "|" & .Cell(flexcpData, 0, COL_ONE_FOURTH)), vbExclamation
                InputGridValidate = False
            Else
                'If we are in English units, length should be a whole number
                dOneFourth = Round(CDbl(Trim(.EditText)), 0)
                If Not gbIsMetric And (dOneFourth - CDbl(Trim(.EditText))) <> 0 Then
                    MsgBox LoadResString(IDS_LENGTH_WHOLE_NUM), vbExclamation
                    InputGridValidate = False
                Else
                    'Length should be within the allowable range
                    dOneFourth = CDbl(Trim(.EditText))
                    If dOneFourth < dMinLength Or dOneFourth > dMaxLength Then
                        MsgBox FormatSWParams(LoadResString(IDS_SHOOTINGLINE_LENGTH_INVALID), _
                            FormatNumber(dMinLength, nDecimal) & "|" & _
                            FormatNumber(dMaxLength, nDecimal) & "|" & .Cell(flexcpData, 0, COL_ONE_FOURTH)), vbExclamation
                        InputGridValidate = False
                    Else
                        InputGridValidate = True
                    End If
                End If
            End If
        Case COL_ONE_EIGHT
            dMinLength = IIf(gbIsMetric, MIN_LENGTH_ONEEIGHT_METRIC, MIN_LENGTH_ONEEIGHT_ENGLISH)
            dMaxLength = IIf(gbIsMetric, MAX_LENGTH_ONEEIGHT_METRIC, MAX_LENGTH_ONEEIGHT_ENGLISH)
        
            'Length should be numeric
            If Len(Trim(.EditText)) = 0 Or Not IsNumeric(Trim(.EditText)) Then
                MsgBox FormatSWParams(LoadResString(IDS_SHOOTINGLINE_LENGTH_INVALID), _
                    FormatNumber(dMinLength, nDecimal) & "|" & FormatNumber(dMaxLength, nDecimal) _
                    & "|" & .Cell(flexcpData, 0, COL_ONE_EIGHT)), vbExclamation
                InputGridValidate = False
            Else
                'If we are in English units, length should be a whole number
                dOneEight = Round(CDbl(Trim(.EditText)), 0)
                If Not gbIsMetric And (dOneEight - CDbl(Trim(.EditText))) <> 0 Then
                    MsgBox LoadResString(IDS_LENGTH_WHOLE_NUM), vbExclamation
                    InputGridValidate = False
                Else
                    'Length should be within the allowable range
                    dOneEight = CDbl(Trim(.EditText))
                    If dOneEight < dMinLength Or dOneEight > dMaxLength Then
                        MsgBox FormatSWParams(LoadResString(IDS_SHOOTINGLINE_LENGTH_INVALID), _
                            FormatNumber(dMinLength, nDecimal) & "|" & _
                            FormatNumber(dMaxLength, nDecimal) & "|" & .Cell(flexcpData, 0, COL_ONE_EIGHT)), vbExclamation
                        InputGridValidate = False
                    Else
                        InputGridValidate = True
                    End If
                End If
            End If
        Case COL_FAN_CAP
            dMinFanCap = IIf(gbIsMetric, MIN_FAN_METRIC, MIN_FAN_ENGLISH)
            dMaxFanCap = IIf(gbIsMetric, MAX_FAN_METRIC, MAX_FAN_ENGLISH)
            
            'Fan Capacity should be numeric
            If Len(Trim(.EditText)) = 0 Or Not IsNumeric(Trim(.EditText)) Then
                MsgBox FormatSWParams(LoadResString(IDS_SHOOTINGLINE_FAN_INVALID), _
                    FormatNumber(dMinFanCap, nDecimal) & "|" & FormatNumber(dMaxFanCap, nDecimal) _
                    & "|" & .Cell(flexcpData, 0, COL_FAN_CAP)), vbExclamation
                InputGridValidate = False
            Else
                'If we are in English units, Fan Capacity should be a whole number
                dFanCap = Round(CDbl(Trim(.EditText)), 0)
                If Not gbIsMetric And (dFanCap - CDbl(Trim(.EditText))) <> 0 Then
                    MsgBox LoadResString(IDS_FAN_WHOLE_NUM), vbExclamation
                    InputGridValidate = False
                Else
                    'Fan Capacity should be within the allowable range
                    dFanCap = CDbl(Trim(.EditText))
                    If dFanCap < dMinFanCap Or dFanCap > dMaxFanCap Then
                        MsgBox FormatSWParams(LoadResString(IDS_SHOOTINGLINE_FAN_INVALID), _
                            FormatNumber(dMinFanCap, nDecimal) & "|" & _
                            FormatNumber(dMaxFanCap, nDecimal) & "|" & _
                            .Cell(flexcpData, 0, COL_FAN_CAP)), vbExclamation
                        InputGridValidate = False
                    Else
                        InputGridValidate = True
                    End If
                End If
            End If
        Case Else
            InputGridValidate = True
    End Select
    End With
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "InputGridValidate", Err.Description)
    End If

End Function

Public Sub InputGridCellUpdate(ByVal lRow As Long, ByVal lCol As Long)
'**************************************************************
'   DESCRIPTION
'       This subroutine saves the cell value in the database.
'
'   PARAMETERS
'           lRow - the row number of the cell that is being edited
'           lCol - the column number of the cell that is being edited
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'   04/03/03    J. Manyaga     Added code that would save the value -1
'                              in the db if humidity is blank.
'                              Limited Inside Diameter to 3 options only.
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim nCount As Integer
    Dim nCurrentRecord As Integer
    Dim bNameOnly As Boolean
    Dim dTemperature As Double
    Dim dLength As Double
    Dim nInsideDiameterOption As Integer
    Dim dFanCap As Double
    Dim dOneFourth As Double
    Dim dOneEight As Double
    Dim dNumCylinder As Double
    Dim dPressure As Double
        
    bNameOnly = True
        
    
    With frmMain.grdIntroInput
    If .Cell(flexcpData, lRow, COL_ORDER_RS) = "" Then
        Call AddIntroLineRecord(lRow)
        nCurrentRecord = nRecordNumber - 1
    Else
        nCurrentRecord = CInt(.Cell(flexcpData, lRow, COL_ORDER_RS)) - 1
    End If
        
    'Go to the Introduction Line record which is being edited
    rsIntroLine.MoveFirst
    For nCount = 1 To nCurrentRecord
        rsIntroLine.MoveNext
    Next
    
    'Update the Introduction Line record
    Select Case lCol
        Case COL_AREA_NAME
            rsIntroLine("jobarea_fkey") = .ComboData
            If IsNull(rsIntroLine("name")) Then
                rsIntroLine("name") = GenerateIntroLineName(lRow, _
                    LoadResString(IDS_DEFAULT_INTROLINE)) & ">"
            End If
            bNameOnly = True
        Case COL_INTROLINE_NAME
            rsIntroLine("name") = CStr(.Cell(flexcpData, lRow, lCol))
            If IsNull(rsIntroLine("jobarea_fkey")) Then
                rsIntroLine("jobarea_fkey") = sFirstIntroAreaId
            End If
            bNameOnly = True
        Case COL_RH
            'JNM 03/31/03 - FE 2003 - RH can be blank.
            If Len(Trim(.Cell(flexcpData, lRow, lCol))) = 0 Then
                rsIntroLine("humidity") = DEFAULT_HUMIDITY
            Else
                rsIntroLine("humidity") = CInt(.Cell(flexcpData, lRow, lCol))
            End If
            
            rsIntroLine("max_rate") = 0#
            If IsNull(rsIntroLine("name")) Then
                rsIntroLine("name") = GenerateIntroLineName(lRow, _
                    LoadResString(IDS_DEFAULT_INTROLINE)) & ">"
            End If
            If IsNull(rsIntroLine("jobarea_fkey")) Then
                rsIntroLine("jobarea_fkey") = sFirstIntroAreaId
            End If
        Case COL_CYLINDER_TEMP
            If frmMain.rbtTemperature.Value = True Then
                dTemperature = 0#
                If gblPresTempMode = False Then
                    Call UnitConvert(modeSave, unitTEMPERATURE, CDbl(.Cell(flexcpData, lRow, lCol)), _
                        dTemperature)
                End If
                rsIntroLine("cylinder_temperature") = Round(dTemperature, 2)
                rsIntroLine("rate") = 0#
                If IsNull(rsIntroLine("name")) Then
                    rsIntroLine("name") = GenerateIntroLineName(lRow, _
                        LoadResString(IDS_DEFAULT_INTROLINE)) & ">"
                End If
                If IsNull(rsIntroLine("jobarea_fkey")) Then
                    rsIntroLine("jobarea_fkey") = sFirstIntroAreaId
                End If
            Else
                dPressure = 0#
                If gblPresTempMode = True Then
                    Call UnitConvert(modeSave, unitPressure, CDbl(.Cell(flexcpData, lRow, lCol)), _
                        dPressure)
                End If
                rsIntroLine("cylinder_temperature") = Round(dPressure, 2)
                rsIntroLine("rate") = 0#
                If IsNull(rsIntroLine("name")) Then
                    rsIntroLine("name") = GenerateIntroLineName(lRow, _
                        LoadResString(IDS_DEFAULT_INTROLINE)) & ">"
                End If
                If IsNull(rsIntroLine("jobarea_fkey")) Then
                    rsIntroLine("jobarea_fkey") = sFirstIntroAreaId
                End If
            End If
        'Case COL_INSIDE_DIAMETER
            'JNM 03/31/03 - FE 2003 - There are only 3 options for the Inside Diameter.
        '    nInsideDiameterOption = CInt(.ComboData)
        '    Select Case nInsideDiameterOption
        '        Case 0
        '            rsIntroLine("diameter") = SHOOTINGLINE_DIAMETER_ONE
        '        Case 1
        '            rsIntroLine("diameter") = SHOOTINGLINE_DIAMETER_TWO
        '        Case 2
        '            rsIntroLine("diameter") = SHOOTINGLINE_DIAMETER_THREE
        '    End Select
            
        '    rsIntroLine("rate") = 0#
        '    If IsNull(rsIntroLine("name")) Then
        '        rsIntroLine("name") = GenerateIntroLineName(lRow, _
        '            LoadResString(IDS_DEFAULT_INTROLINE)) & ">"
        '    End If
        '    If IsNull(rsIntroLine("jobarea_fkey")) Then
        '        rsIntroLine("jobarea_fkey") = sFirstIntroAreaId
        '    End If
        'Case COL_LENGTH
        '    dLength = 0#
        '    Call UnitConvert(modeSave, unitLWH, CDbl(.Cell(flexcpData, lRow, lCol)), _
        '        dLength)
        '    rsIntroLine("length") = Round(dLength, 2)
        '    rsIntroLine("rate") = 0#
        '    If IsNull(rsIntroLine("name")) Then
        '        rsIntroLine("name") = GenerateIntroLineName(lRow, _
        '            LoadResString(IDS_DEFAULT_INTROLINE)) & ">"
        '    End If
        '    If IsNull(rsIntroLine("jobarea_fkey")) Then
        '        rsIntroLine("jobarea_fkey") = sFirstIntroAreaId
        '    End If
        Case COL_ONE_FOURTH
            dOneFourth = 0#
            Call UnitConvert(modeSave, unitLWH, CDbl(.Cell(flexcpData, lRow, lCol)), _
                dOneFourth)
            rsIntroLine("one_fourth_length") = Round(dOneFourth, 2)
            rsIntroLine("rate") = 0#
            If IsNull(rsIntroLine("name")) Then
                rsIntroLine("name") = GenerateIntroLineName(lRow, _
                    LoadResString(IDS_DEFAULT_INTROLINE)) & ">"
            End If
            If IsNull(rsIntroLine("jobarea_fkey")) Then
                rsIntroLine("jobarea_fkey") = sFirstIntroAreaId
            End If
        Case COL_ONE_EIGHT
            dOneEight = 0#
            Call UnitConvert(modeSave, unitLWH, CDbl(.Cell(flexcpData, lRow, lCol)), _
                dOneEight)
            rsIntroLine("one_eight_length") = Round(dOneEight, 2)
            rsIntroLine("rate") = 0#
            If IsNull(rsIntroLine("name")) Then
                rsIntroLine("name") = GenerateIntroLineName(lRow, _
                    LoadResString(IDS_DEFAULT_INTROLINE)) & ">"
            End If
            If IsNull(rsIntroLine("jobarea_fkey")) Then
                rsIntroLine("jobarea_fkey") = sFirstIntroAreaId
            End If
        Case COL_FAN_CAP
            dFanCap = 0#
            Call UnitConvert(modeSave, unitVOLUME, CDbl(.Cell(flexcpData, lRow, lCol)), _
                dFanCap)
            rsIntroLine("fan") = Round(dFanCap, 2)
            rsIntroLine("max_rate") = 0#
            If IsNull(rsIntroLine("name")) Then
                rsIntroLine("name") = GenerateIntroLineName(lRow, _
                    LoadResString(IDS_DEFAULT_INTROLINE)) & ">"
            End If
            If IsNull(rsIntroLine("jobarea_fkey")) Then
                rsIntroLine("jobarea_fkey") = sFirstIntroAreaId
            End If
        Case COL_NUM_CYLINDER
            dNumCylinder = 0#
            Call UnitConvert(modeSave, unitLess, CDbl(.Cell(flexcpData, lRow, lCol)), _
                 dNumCylinder)
            rsIntroLine("numcylinder") = Round(dNumCylinder, 2)
            rsIntroLine("rate") = 0#
            If IsNull(rsIntroLine("name")) Then
                rsIntroLine("name") = GenerateIntroLineName(lRow, _
                    LoadResString(IDS_DEFAULT_INTROLINE)) & ">"
            End If
            If IsNull(rsIntroLine("jobarea_fkey")) Then
                rsIntroLine("jobarea_fkey") = sFirstIntroAreaId
            End If
    End Select
    
    'Reset the Warnings field
    If Not bNameOnly Then
        rsIntroLine("calculation_warning") = ""
    End If
    
    'Go forth and update
    rsIntroLine.Update
    
    End With
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "InputGridCellUpdate", Err.Description)
    End If

End Sub
Private Function GenerateIntroLineName(ByVal lRow As Long, _
ByVal sIntroLineName As String) As String
'**************************************************************
'   DESCRIPTION
'       This subroutine generates a unique name for the Introduction
'       Line record.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - a unique Introduction Line name
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    sIntroLineName = sIntroLineName & CStr(lRow)
    
    If InStr(sGridIntroLineName, "|" & sIntroLineName & ">|") Then
        sIntroLineName = GenerateIntroLineName(lRow, sIntroLineName)
    End If
    
    GenerateIntroLineName = sIntroLineName
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "GenerateIntroLineName", Err.Description)
    End If

End Function

Private Sub AddIntroLineRecord(ByVal lRow As Long)
'**************************************************************
'   DESCRIPTION
'       This subroutine adds a new record in the Introduction
'       Line table.
'
'   PARAMETERS
'           lRow - the row number of the cell that is being edited
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************

On Error GoTo Error_Handler
    
    Dim sNewGUID As String
    
    'Generate a new GUID for the new Introduction Line record
    sNewGUID = GetGUID
    
    'Create the new record
    rsIntroLine.AddNew
    rsIntroLine("id") = sNewGUID
    rsIntroLine("job_fkey") = gsJobID
    rsIntroLine("humidity") = DEFAULT_HUMIDITY
    rsIntroLine("diameter") = DEFAULT_DIAMETER
    rsIntroLine("row") = lRow
    If frmMain.rbtTemperature.Value = True Then
        rsIntroLine("Pressure0_or_temperature1") = 1
    Else
        rsIntroLine("Pressure0_or_temperature1") = 0
    End If
    If frmMain.rbtPressure.Value = True Then
        rsIntroLine("cylinder_temperature") = VAL_PRESSURE
        gblPresTempMode = True
    Else
        gblPresTempMode = False
    End If
    
    rsIntroLine.Update
    
    'Since we added a new record
    bIsCleared = False
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "AddIntroLineRecord", Err.Description)
    End If

End Sub

Public Sub CalculateIntroRate()
'**************************************************************
'   DESCRIPTION
'       This subroutine calculates the Actual and Permitted
'       Introduction rates.
'
'   PARAMETERS
'           lRow - the row number of the cell that is being edited
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'   04/03/03     J. Manyaga      Added a condition that would prevent
'                                a calculation if Humidity is blank.
'**************************************************************

On Error GoTo Error_Handler
    
    'Do not calculate if there are no Introduction Line records
    If Not bHasIntroLine Then
        MsgBox LoadResString(IDS_NO_CALC_INTRO), vbExclamation
        Exit Sub
    End If
    
    Dim bWarningsFound As Boolean
    Dim sWarning As String
    Dim sActualIntroRateWarning As String
    Dim sIntroRateExceedWarning As String
    Dim dMaxIntroRate As Double
    Dim dActualIntroRate As Double
    Dim dConActualIntroRate As Double
    Dim dConMaxIntroRate As Double
    Dim sIntroRateUnit As String
    Dim dActualTimeMin As Double
    Dim sWarnings As String
    Dim dblCylinderTemp As Double
    Dim dblCylinderPress As Double
    Dim bIntroRateWarning As Boolean
    
    
    
    'Create an instance of the Shooting Line class
    Dim IntroCalc As ShootingLine
    
    Set IntroCalc = New ShootingLine
    
    bWarningsFound = False
    bIntroRateWarning = False
    
    'Loop through all the Introduction Line records
    rsIntroLine.MoveFirst
    While Not rsIntroLine.EOF
        'Clear the current Introduction Line warning
        rsIntroLine("calculation_warning") = ""
        rsIntroLine.Update
        
        rsIntroLine("Pressure0_or_temperature1") = dTempPresTemp
        
        If rsIntroLine("Pressure0_or_temperature1") = 1 Then
            dblCylinderTemp = rsIntroLine("cylinder_temperature")
            Call ConvertTempPressure(dblCylinderTemp, dblCylinderPress)
            gblPresTempMode = False
        Else
            dblCylinderPress = rsIntroLine("cylinder_temperature")
            gblPresTempMode = True
        End If
        
        If IsNull(rsIntroLine("numcylinder")) Then
            rsIntroLine("numcylinder") = 0#
        End If
        
        If IsNull(rsIntroLine("one_eight_length")) Then
            rsIntroLine("one_eight_length") = 0#
        End If
        
        If IsNull(rsIntroLine("one_fourth_length")) Then
            rsIntroLine("one_fourth_length") = 0#
        End If
                
        'Don't calculate the shooting rates if the user has not entered
        'a Fan Capacity.
        'JNM 03/31/03 - FE 2003 - Also, do not calculate if RH is blank
        If rsIntroLine("humidity") <> -1 And CDbl(rsIntroLine("fan")) > 0# Then
            'Initialize the values before calculating
            dMaxIntroRate = 0#
            dActualIntroRate = 0#
            dActualTimeMin = 0#
            sWarning = ""
            sActualIntroRateWarning = ""
            sIntroRateExceedWarning = ""
                      
            'Calculate the Maximum Shooting Rate
            IntroCalc.GetMaxShootingRate CDbl(rsIntroLine("fan")), CLng(rsIntroLine("humidity")), _
            dMaxIntroRate, sWarning
            
            'Do not continue the calculation if there is an error. This is indicated by
            'a value in sWarning
            If sWarning = "" Then
                'Calculate the Shooting Rate
                ' 5/8/03 - Jett: Rounding down the line length, when converting the max
                ' value (500) to metric, there we are .4 over
                IntroCalc.GetShootingRate CDbl(dblCylinderPress), _
                    (Round(CDbl(rsIntroLine("one_fourth_length")), 0)), CDbl(rsIntroLine("one_eight_length")), _
                    CDbl(rsIntroLine("numcylinder")), CDbl(rsIntroLine("Pressure0_or_temperature1")), _
                    dActualIntroRate, dActualTimeMin, sActualIntroRateWarning
                
                If sActualIntroRateWarning = "" Then
                    If dActualIntroRate > dMaxIntroRate Then
                        dConActualIntroRate = 0#
                        dConMaxIntroRate = 0#
                        sIntroRateUnit = ""
                    
                        Call UnitConvert(modeDisplay, unitSHOOTINGRATE, dActualIntroRate, _
                            dConActualIntroRate, sIntroRateUnit)
                        Call UnitConvert(modeDisplay, unitSHOOTINGRATE, dMaxIntroRate, _
                            dConMaxIntroRate)
                        dConActualIntroRate = Round(dConActualIntroRate, 2)
                        dConMaxIntroRate = Round(dConMaxIntroRate, 2)
                    
                        sIntroRateExceedWarning = FormatSWParams( _
                            LoadResString(IDS_WARNING_SHOOT_RATE_GREATER_THAN_MAX_SHOOT_RATE), _
                            FormatNumber(dConActualIntroRate, 2) & "|" & sIntroRateUnit & "|" & _
                            FormatNumber(dConMaxIntroRate, 2) & "|" & sIntroRateUnit)
                    End If
                End If
                
                If sActualIntroRateWarning <> "" Then
                    If sWarning <> "" Then
                        sWarning = sWarning & vbCrLf
                    End If
                    sWarning = sWarning & sActualIntroRateWarning
                End If
            
                If sIntroRateExceedWarning <> "" Then
                    bIntroRateWarning = True
                    If sWarning <> "" Then
                        sWarning = sWarning & vbCrLf
                    End If
                    sWarning = sWarning & sIntroRateExceedWarning
                End If
    
                If sWarning <> "" Then
                    bWarningsFound = True
                End If
                
                     
                'Update the Introduction Line record
                rsIntroLine("max_rate") = dMaxIntroRate
                rsIntroLine("rate") = dActualIntroRate
                rsIntroLine("time_min_empty_cylinder") = dActualTimeMin
                rsIntroLine("calculation_warning") = sWarning
                rsIntroLine.Update
            
            Else
                'Store the warning in the database
                rsIntroLine("calculation_warning") = sWarning
                rsIntroLine.Update
                bWarningsFound = True
            End If
        End If
           
        rsIntroLine.MoveNext
        
    Wend
     
    'Pat 4/20/2005: This message box shall be shown
    If bIntroRateWarning = False And bWarningsFound = True Then
        MsgBox LoadResString(IDS_SHOOTING_RATE_ERROR_SEE_WARNINGS), vbExclamation
    End If
    
    'Refresh the grids
    Call LoadIntroRS
    
    'Refresh the Warnings text box
    Call LoadWarnings(sWarnings)
    
    If Len(sWarnings) > 0 And bIntroRateWarning = True Then
        frmWarningMsg.txtErrorIntroLine = sWarnings
        frmWarningMsg.Caption = LoadResString(IDS_WARNINGS)
        frmWarningMsg.chkEnableCmdBtn.Caption = LoadResString(IDS_WARNING_VERIFICATION)
        frmWarningMsg.Show vbModal
    End If
    
    'Clean up
    Set IntroCalc = Nothing
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "CalculateIntroRate", Err.Description)
    End If

End Sub

Public Sub DeleteIntroLine(ByVal lRow As Long)
'**************************************************************
'   DESCRIPTION
'       This subroutine deletes the selected record in the Introduction
'       Line (ShootingLine)table.
'
'   PARAMETERS
'           lRow - the row number of the cell that is being edited
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim nCurrentRecord As Integer
    Dim nDeleteOk As Integer
    Dim bDeleteRow As Boolean
    Dim nCount As Integer
    
    bDeleteRow = False
    
    'Exit this sub if the Job has no Introduction Line records
    If Not bHasIntroLine Then
        Exit Sub
    End If
    
    'Determine if the current row has an Introduction Line record
    If frmMain.grdIntroInput.Cell(flexcpData, lRow, COL_ORDER_RS) = "" Then
        'Current row has no Introduction Line record. Ergo, just remove the row.
        bDeleteRow = True
    Else
        nCurrentRecord = CInt(frmMain.grdIntroInput.Cell(flexcpData, lRow, COL_ORDER_RS)) - 1
     
        'Go to the Introduction Line record which is being deleted
        rsIntroLine.MoveFirst
        For nCount = 1 To nCurrentRecord
            rsIntroLine.MoveNext
        Next
        
        'Do not delete the Introduction Line record if it has an associated
        'introduction history record.
        If HasIntroHistory(rsIntroLine("id")) Then
            bDeleteRow = False
        Else
           'Delete the record if the user agrees
           nDeleteOk = MsgBox(FormatSWParams(LoadResString(IDS_MSG_DELETE_GENERIC), _
               "Introduction Line|Introduction Line"), vbOKCancel)
           If nDeleteOk = vbOK Then
               rsIntroLine("deleted") = True
               rsIntroLine.Update
               bDeleteRow = True
           Else
               bDeleteRow = False
           End If
        End If
    End If
    
    If bDeleteRow Then
        'Update all Introduction Line records below the row
        rsIntroLine.MoveLast
        While Not rsIntroLine.BOF
            If rsIntroLine("row") > lRow Then
                rsIntroLine("row") = rsIntroLine("row") - 1
                rsIntroLine.Update
            End If
            rsIntroLine.MovePrevious
        Wend
        'Refresh the grid
        Call LoadIntroRS
        Call LoadWarnings
    End If
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "DeleteIntroLine", Err.Description)
    End If

End Sub

Private Function HasIntroHistory(sIntroLineId As String) As Boolean
'**************************************************************
'   DESCRIPTION
'       This function determines if the Introduction Line has associated
'       Introduction History records.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - True - if there are Introduction History records
'             False - if there are no Introduction History records
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim sIntroHistorySQL As String
    Dim rsIntroHistory As ADODB.Recordset
    
    Set rsIntroHistory = New ADODB.Recordset
    
    sIntroHistorySQL = "SELECT IntroHistory.* FROM IntroHistory " & _
        "where IntroHistory.shootingline_fkey = '" & sIntroLineId & "' and " & _
        "IntroHistory.deleted = FALSE"
    rsIntroHistory.Open sIntroHistorySQL, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    If rsIntroHistory.BOF And rsIntroHistory.EOF Then
        HasIntroHistory = False
    Else
        MsgBox LoadResString(IDS_MSG_NODELETE_INTROLINE_INTROHISTORY), vbExclamation
        HasIntroHistory = True
    End If
    
    'Clean-up
    rsIntroHistory.Close
    Set rsIntroHistory = Nothing
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "HasIntroHistory", Err.Description)
    End If

End Function

Public Sub IntroPlanCleanUp()
'**************************************************************
'   DESCRIPTION
'       This subroutine closes and destroys all objects.
'       This routine is called by the Unload event of the form.
'
'   PARAMETERS
'          NONE
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    'First of all, save the current grid settings of the user
    GridWrapInput.SaveGridSettings GRID_SETTINGS_PATH
    GridWrapOutput.SaveGridSettings GRID_SETTINGS_PATH
    
    rsIntroLine.Close
    Set rsIntroLine = Nothing
    
    Set GridWrapInput = Nothing
    Set GridWrapOutput = Nothing
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "IntroPlanCleanUp", Err.Description)
    End If

End Sub


Public Sub Load_IntroductionPlan()
    IntroPlanInitializeControls
    
End Sub


Public Sub Unload_IntroductionPlan()
    If frmMain.grdIntroInput.EditWindow <> 0 Then
        ' TODO: Cancel = True
        gbCanChangeTabs = False
    Else
        Call IntroPlanCleanUp
        gbCanChangeTabs = True
        ' TODO: Cancel = False
    End If
End Sub


'========================================================================================================
' -------------------- Handlers moved from the form ---------------
'========================================================================================================

Public Sub cmdCalculateIntroRate_Click_Handler()
    CalculateIntroRate
    gbHasChanged = True
End Sub


Public Sub grdIntroInput_AfterEdit_Handler(ByVal Row As Long, ByVal Col As Long)
    'Refresh the controls
    Call LoadIntroRS
    Call LoadWarnings
End Sub


Public Sub grdIntroInput_BeforeUserResize_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Do not allow the user to change the width of the Index column
    If Col = COL_INDEX Then
        Cancel = True
    End If
End Sub

Public Sub grdIntroInput_KeyDown_Handler(KeyCode As Integer, Shift As Integer)
    'If user pressed the Delete button, delete the Introduction Line record
    If KeyCode = 46 Then
        Call DeleteIntroLine(frmMain.grdIntroInput.Row)
    End If
End Sub

Public Sub grdIntroInput_MouseDown_Handler(Button As Integer, Shift As Integer, x As Single, Y As Single)

    With frmMain
        'Handle the right-click event
        If Button = vbRightButton And .grdIntroInput.MouseRow > 0 Then
            'Select the cell
            .grdIntroInput.Select .grdIntroInput.MouseRow, .grdIntroInput.MouseCol
            
            frmMain.PopupMenu frmMain.mnuIntroPlanContext
        End If
    End With

End Sub


Public Sub grdIntroInput_MouseMove_Handler(Button As Integer, Shift As Integer, x As Single, Y As Single)

    With frmMain.grdIntroInput
        Select Case .MouseCol
            Case COL_AREA_NAME
                .ToolTipText = LoadResString(IDS_TIP_INTRO_AREA)
            Case COL_INTROLINE_NAME
                .ToolTipText = LoadResString(IDS_TIP_INTRO_NAME)
            Case COL_RH
                .ToolTipText = LoadResString(IDS_TIP_INTRO_RH)
            Case COL_CYLINDER_TEMP
                .ToolTipText = LoadResString(IDS_TIP_INTRO_TEMP)
            'Pat 12/28/2004
            'Case COL_INSIDE_DIAMETER
            '    .ToolTipText = LoadResString(IDS_TIP_INTRO_DIAMETER)
            'Case COL_LENGTH
            '    .ToolTipText = LoadResString(IDS_TIP_INTRO_LENGTH)
            Case COL_ONE_FOURTH
                .ToolTipText = LoadResString(IDS_TIP_INTRO_LENGTH)
            Case COL_ONE_EIGHT
                .ToolTipText = LoadResString(IDS_TIP_INTRO_LENGTH)
            Case COL_FAN_CAP
                .ToolTipText = LoadResString(IDS_TIP_INTRO_FAN)
            Case COL_NUM_CYLINDER
                .ToolTipText = LoadResString(IDS_NUMBER_OF_CYL)
        End Select
    End With
    
End Sub


'**************************************************************
'   DESCRIPTION
'       Start Edit event of the Intro Input Grid in the Intro Plan form
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/11/02        Cel Gamboa        Set the maxlength of the RH column to 3.
                                     
'**************************************************************
Public Sub grdIntroInput_StartEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With frmMain
    
        'The first row and first column should  be read-only
        If Row = 0 Or Col = COL_INDEX Or Col = COL_ORDER_RS Then
            Cancel = True
        Else
            'Remove the units once the user starts to edit the cell
            .grdIntroInput.Cell(flexcpText, Row, Col) = .grdIntroInput.Cell(flexcpData, Row, Col)
            
            If Col = COL_AREA_NAME And .grdIntroInput.Cell(flexcpText, Row, COL_ORDER_RS) = "" Then
                .grdIntroInput.Cell(flexcpText, Row, Col) = sFirstIntroAreaName
            End If
            
            'Pat 12/24/2004 - Removed
            'If Col = COL_INSIDE_DIAMETER And .grdIntroInput.Cell(flexcpText, Row, COL_ORDER_RS) = "" Then
            '    .grdIntroInput.Cell(flexcpText, Row, Col) = sFirstInsideDiameter
            'End If
            
            If Col = COL_INTROLINE_NAME Then
                .grdIntroInput.EditMaxLength = 50
            End If
            
            If Col = COL_RH Then
                .grdIntroInput.EditMaxLength = 3
            End If
            
            If Col <> COL_RH And Col <> COL_INTROLINE_NAME Then
                .grdIntroInput.EditMaxLength = 0
            End If
        
            
            Cancel = False
        End If
    
    End With
    
    gbHasChanged = True
    
    'The line below would ensure that the tab does not change if the user enter enters
    'an invalid value and then clicks on another tab
    ReturnCurrTab TAB_INTRO_PLAN

End Sub


Public Sub grdIntroInput_ValidateEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With frmMain
        If InputGridValidate(Row, Col) Then
            'Do not include the Area Name and Inside Diameter since these are drop-down list boxes
            If Col <> COL_AREA_NAME Then
            'And Col <> COL_INSIDE_DIAMETER
                'Store the new value in flexcpData and flexcpText
                .grdIntroInput.Cell(flexcpData, Row, Col) = Trim(.grdIntroInput.EditText)
        
                'And let there be units!!! The units are hidden in the flexcpData attribute
                'of the grid heading
                .grdIntroInput.Cell(flexcpText, Row, Col) = Trim(.grdIntroInput.EditText) & _
                    .grdIntroInput.Cell(flexcpData, 0, Col)
            End If
            
            'Save in the database
            InputGridCellUpdate Row, Col
            
            Cancel = False
        Else
            'JNM PT fix - 5/7/2003 - Set to blank as default.
            .grdIntroInput.Cell(flexcpData, Row, Col) = ""
            Cancel = True
        End If
        
        'JNM 05/02/03 - PT fix - remove this since we are already in one form.
        'JNM 10/23 - Focus should return to the grid if value is invalid
        'If Cancel Then
        '    .grdIntroInput.SetFocus
        'End If
    End With
    
End Sub

Public Sub grdIntroRates_BeforeEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'This grid should be read-only
    Cancel = True
End Sub


Public Sub grdIntroRates_MouseMove_Handler(Button As Integer, Shift As Integer, x As Single, Y As Single)

    With frmMain.grdIntroRates
        Select Case .MouseCol
            Case COL_AREA_NAME
                .ToolTipText = LoadResString(IDS_TIP_INTRO_AREA_RO)
            Case COL_INTROLINE_NAME
                .ToolTipText = LoadResString(IDS_TIP_INTRO_NAME)
            Case COL_ACTUAL_RATE
                .ToolTipText = LoadResString(IDS_TIP_INTRO_ACTUAL)
            Case COL_PERMIT_RATE
                .ToolTipText = LoadResString(IDS_TIP_INTRO_PERMIT)
            Case COL_TIME_EMPTY_CYL
                .ToolTipText = LoadResString(IDS_TIME_EMPTY_CYL)
        End Select
    End With
    
End Sub


Public Sub mnuIntroPlanDelete_Click_Handler()
    Call DeleteIntroLine(frmMain.grdIntroInput.Row)
End Sub

'**************************************************************
'   DESCRIPTION
'       Will call convert functions: Temperature to Pressure
'                                    Pressure to Temperature
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  3/28/2005       Pat San Miguel       Initial Revision
'**************************************************************
Public Sub ConvertPresTemp()
   
    Dim nCurrRow As Integer
    Dim dTemperature As Double
    Dim dPressure As Double
    Dim dTempunit As String
    Dim dPressureUnit As String
    Dim nCurrentRecord
    Dim nRecordNumber
    Dim nCount
    Dim sIntroLineSQL
    Dim rsCylinder As Double

    'Will not convert if it has no introline
    If Not bHasIntroLine Then
        MsgBox LoadResString(IDS_NO_CALC_INTRO), vbExclamation
        Exit Sub
    End If
    
    If rsIntroLine.State Then
        rsIntroLine.Close
    End If

    'Create the Introduction Line recordset
    sIntroLineSQL = "select ShootingLine.* from ShootingLine left outer join " & _
        "JobArea on ShootingLine.jobarea_fkey = JobArea.id where " & _
        "ShootingLine.deleted = false and ShootingLine.job_fkey = '" & gsJobID & "'" & _
        "order by ShootingLine.row"
        
    rsIntroLine.Open sIntroLineSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    nRecordNumber = 1
    nCurrRow = 1
    With frmMain
        If Not rsIntroLine.EOF And Not rsIntroLine.BOF Then
        'loop through the recordset
            Do While Not rsIntroLine.EOF
                If .rbtTemperature.Value = True Then
                    
                    dPressure = rsIntroLine("cylinder_temperature")
                    'Convert Pressure to Temperature
                    Call ConvertPressureTemp(dTemperature, dPressure)
                    
                    'If the converted temperature is less than to its minimum values
                    If dTemperature < MIN_TEMP_METRIC Or dTemperature < MIN_TEMP_ENGLISH Then
                        'Round Off the value
                        rsIntroLine("cylinder_temperature") = RoundOff(dTemperature, 2)
                    Else
                        'Get the computed value from the database
                        rsIntroLine("cylinder_temperature") = dTemperature
                    End If
                    
                    'get the unit
                    dTempunit = IIf(gbIsMetric, LoadResString(IDS_CELSIUS), LoadResString(IDS_FAHRENHEIT))
                    
                    'Asssign values to grids
                    .grdIntroInput.Cell(flexcpData, 0, COL_CYLINDER_TEMP) = " " & "°" & dTempunit
                    dTemperature = Round(dTemperature, IIf(gbIsMetric, 1, 0))
                    .grdIntroInput.Cell(flexcpData, nCurrRow, COL_CYLINDER_TEMP) = _
                        FormatNumber(dTemperature, IIf(gbIsMetric, 1, 0))
                    .grdIntroInput.Cell(flexcpText, nCurrRow, COL_CYLINDER_TEMP) = _
                        FormatNumber(dTemperature, IIf(gbIsMetric, 1, 0)) & _
                    .grdIntroInput.Cell(flexcpData, 0, COL_CYLINDER_TEMP)
                    
                    'This will save the metric value in the database
                    If gblPressure = True Then
                        Call UnitConvert(modeSave, unitTEMPERATURE, CDbl(.grdIntroInput.Cell(flexcpData, nCurrRow, COL_CYLINDER_TEMP)), _
                        dTemperature)
                        rsIntroLine("cylinder_temperature") = dTemperature
                    End If
                    
                    'Indicate that Temperature is selected
                    gblPresTempMode = False
                    blnTempPres = False
                    rsIntroLine("Pressure0_or_temperature1") = 1
                    
                    'Enables Pressure radio button
                    frmMain.rbtPressure.Enabled = True

                    gbForcedClick = False
                Else

                    dTemperature = rsIntroLine("cylinder_temperature")
                    
                    'Convert Temperature to Pressure
                    Call ConvertTempPressure(dTemperature, dPressure)
                    
                    rsIntroLine("cylinder_temperature") = dPressure
                    dPressureUnit = IIf(gbIsMetric, LoadResString(IDS_PRESSURE_BAR), LoadResString(IDS_PRESSURE_PSIG))
                    
                    Call UnitConvert(modeDisplay, unitPressure, rsIntroLine("cylinder_temperature"), _
                    dPressure, dPressureUnit)
                    
                    .grdIntroInput.Cell(flexcpData, 0, COL_CYLINDER_TEMP) = " " & dPressureUnit
                    dPressure = Round(dPressure, IIf(gbIsMetric, 1, 0))
                    .grdIntroInput.Cell(flexcpData, nCurrRow, COL_CYLINDER_TEMP) = _
                        FormatNumber(dPressure, IIf(gbIsMetric, 1, 0))
                    .grdIntroInput.Cell(flexcpText, nCurrRow, COL_CYLINDER_TEMP) = _
                        FormatNumber(dPressure, IIf(gbIsMetric, 1, 0)) & _
                    .grdIntroInput.Cell(flexcpData, 0, COL_CYLINDER_TEMP)
                    
                    'Indicate that Pressure is selected
                    gblPresTempMode = True
                    blnTempPres = True
                    gblPressure = True
                    rsIntroLine("Pressure0_or_temperature1") = 0
                    
                    'Enables Pressure radio button
                    frmMain.rbtTemperature.Enabled = True
                    
                    gbForcedClick = False
                End If
                nCurrRow = nCurrRow + 1
                nRecordNumber = nRecordNumber + 1
                rsIntroLine.MoveNext
            Loop
        End If
        
   End With
   'Indicates that convert button is selected
    blnConvert = True
   'Enable English-Metric Conversion
    frmMain.mnuViewOptions.Enabled = True
End Sub
