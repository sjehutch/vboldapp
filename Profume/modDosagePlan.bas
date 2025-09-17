Attribute VB_Name = "modDosagePlan"
' JETT HAS THOUROUgHLY REVIEWED AND OPTIMIZED THIS MODULE

'------------------------------------------------------
'*
'* Filename: modDosagePlan.bas
'*
'* DESCRIPTION
'* -----------
'* This code contains the functions and subroutines
'* used by frmDosagePlan.
'*
'* Copyright (c) 2002 Accenture.  All rights reserved.
'-----------------------------------------------------
'*
'* REVISION HISTORY
'* Revision:    0
'* Author:      Manyaga, Joseph
'* Date:        October 3, 2002
'* Revision:    1
'* Author:      Manyaga, Joseph
'* Date:        November 4, 2002
'* Description: Added the code that would set the tooltip text
'*              for the rtfResults text box.
'* Revision:    2
'* Author:      Tan, Yasmin
'* Date:        December 20, 2004
'* Description: Added the code to allow user defined CT on the dosage plan
'*
'* Revision:    3
'* Author:      David Smith
'* Description: Removed text from mouse over for txtLoadFactor
'----------------------------------------------------
Option Explicit

'Declare database objects
Dim rsJob As ADODB.Recordset        'the Job recordset - always open
Dim rsJobPest As ADODB.Recordset    'the Job Pests recordset - always open
Dim rsJobArea As New ADODB.Recordset    'the Job Area recordset - always open

'Declare class instances
Dim GridWrap As New clsGridWrapper      'the Grid Wrapper
Dim mclsToolTip As clsToolTip       'the class that enables multi-line tooltips

'Declare global variables
Dim nDecimal As Integer                 'the number of decimal places that would be displayed
Dim nPestTotal As Integer               'total number of available pests in the Profume Calculator
Dim sGridAreaName As String             'used to store all the exisitng area names in the database
Dim nRecordNumber As Integer            'used to store the current record number of the JobArea recordset
Dim bHasPest As Boolean                 'flag if Job has Pests in the JobPest table
Public bIsLoading As Boolean            'flag if the combo boxes are still being loaded
Public bIsCleared As Boolean            'flag if the Results text box has already been cleared
Dim bHasJobArea As Boolean              'flag if the Job has Job Area records

Dim bInvalid As Boolean                 'invalid flag from the old form

Public bEnteredUserCT As Boolean
    Public nCurrRow As Integer

'Declare constants
'These are the grid columns.
Private Const COL_AREA_NAME As Integer = 0
Private Const COL_TEMPERATURE As Integer = 1
Private Const COL_EST_HLT As Integer = 2
Private Const COL_PLAN_EXPOSURE As Integer = 3
Private Const COL_VOLUME As Integer = 4
Private Const COL_SPACE As Integer = 5
Private Const COL_FUMIGANT As Integer = 6
Private Const COL_CT As Integer = 7
Private Const COL_USER_CT As Integer = 8 'Added user-defined CT, adjusted column #s 12/21/04-YYT
Private Const COL_CO As Integer = 9
Private Const COL_ORDER_RS As Integer = 10
'These are the rich text box format settings
Private Const RTF_FONT_SIZE As Integer = 8      'the font size of the Results string
Private Const RTF_FONT_NAME As String = "arial"
Private Const SPACE_NUM As Integer = 10         'number of spaces between results fields

Private Const POUNDS_PER_KG As Double = 2.2046226
Private Const MIN_LOAD_FACTOR As Integer = 0
Private Const MIN_TEMP_METRIC As Double = 4.44
Private Const MIN_TEMP_ENGLISH As Double = 40
Private Const MAX_TEMP_METRIC As Double = 65.56
Private Const MAX_TEMP_ENGLISH As Double = 150
Private Const MIN_EST_HLT As Double = 1#
Private Const MAX_EST_HLT As Double = 1000#
Private Const MIN_PLANNED_EXPOSURE As Double = 1#
Private Const MAX_PLANNED_EXPOSURE As Double = 168
Private Const MIN_VOLUME_METRIC As Double = 1#
Private Const MIN_VOLUME_ENGLISH  As Double = 35
'Private Const MAX_VOLUME_METRIC As Double = 283168.47
'Private Const MAX_VOLUME_ENGLISH As Double = 10000000
Public Const MAX_VOLUME_METRIC As Double = 2803367.81
Public Const MAX_VOLUME_ENGLISH As Double = 99000000
Private Const MAX_LOAD_FACTOR As Integer = 100
'12/21/04- Added MAX User-DEFINED CT for fumigation
Private Const MAX_VACUUM_USER_CT = 200
Private Const MAX_NORMAL_USER_CT = 1500
Private Const MIN_NORMAL_VACUUM_USER_CT = 36


Private Const DEFAULT_TEMPERATURE As Double = 23.89
Private Const DEFAULT_HLT As Double = MIN_EST_HLT
Private Const DEFAULT_EXPOSURE As Double = MIN_PLANNED_EXPOSURE
Private Const DEFAULT_VOLUME As Double = 1#

Const AREANAME_COLSIZE = 3525
Const TEMP_COLSIZE = 1140
Const ESTHLT_COLSIZE = 870
Const PLANEXP_COLSIZE = 1290
Const AREAVOL_COLSIZE = 1140
Const SPACE_COLSIZE = 45
Const FUMIGANT_COLSIZE = 1515
Const CT_COLSIZE = 900
Const CONC_COLSIZE = 1515
Const USER_CT_COLSIZE = 1140 'added user-defined CT 12/20/04 YYT

Private Const MODULE_NAME = "modDosagePlan"

Dim bRSClosed As Boolean

Private Sub ModuleCleanUp()
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
    
    mclsToolTip.RemoveTool frmMain.rtfResults
    Set mclsToolTip = Nothing
    
    'First of all, save the current grid settings of the user
    GridWrap.SaveGridSettings GRID_SETTINGS_PATH
    
    'Annihilate!!!
    rsJob.Close
    
    rsJobPest.Close
    rsJobArea.Close
    Set rsJob = Nothing
    Set rsJobPest = Nothing
    Set rsJobArea = Nothing
    Set GridWrap = Nothing
    
    'JNM 04/24/03 - AT Fix - Inform the app that the rs have been closed
    bRSClosed = True
           
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "ModuleCleanUp", Err.Description)
    End If

End Sub


Private Sub InitializeControls()
'**************************************************************
'   DESCRIPTION
'       This subroutine initializes all the controls in the form.
'       This sets the options of the list and combo boxes.
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
    
    'Declare variables
    Dim nCount As Integer           'loop counter
    Dim nLifeStageTotal As Integer  'total number of Life Stages in the Profume Calculator
    Dim nFumTypeTotal As Integer    'total number of Fumigation Types in the calculator
    Dim nPressTypeTotal As Integer  'total number of Pressure Types in the calculator
    Dim nCommodityTotal As Integer  'total number of Commodities in the calculator
    
    'Instantiate the Profume Calculator
    Dim ProCalc As New ProFumeCalculator
    
    Set mclsToolTip = New clsToolTip
    
    'Initialize the global variables.
    nDecimal = IIf(gbIsMetric, 2, 0)
    bHasJobArea = False
    bIsLoading = True
    bIsCleared = False
    bHasPest = False
    
    'JNM 04/23/03 - AT Fix - Set our global that indicates if the recordsets are
    '   still open
    bRSClosed = False
        
    'Initialize variables
    nPestTotal = 0
    nCount = 0
    nLifeStageTotal = 0
    nFumTypeTotal = 0
    nPressTypeTotal = 0
    nCommodityTotal = 0
    
    'Create the Job recordset
    Call CreateJobRS
    
    With frmMain
        'Set the form labels using the values from the resource file
        .fraGeneralInfo.Caption = LoadResString(IDS_GENINFO_CAP)
        .fraAreaInfo.Caption = LoadResString(IDS_AREAINFO_CAP)
        .fraTargetInfo.Caption = LoadResString(IDS_TARGETINFO_CAP)
        .fraResults.Caption = LoadResString(IDS_RESULTS_CAP)
        .lblSiteName.Caption = LoadResString(IDS_SITENAME_CAP)
        .lblJobName.Caption = LoadResString(IDS_JOBNAME_CAP)
        .lblFumigationDate.Caption = LoadResString(IDS_FUMDATE_CAP)
        .lblLicFumigator.Caption = LoadResString(IDS_LICFUM_CAP)
        .lblTargetPests.Caption = LoadResString(IDS_TARGETPEST_CAP)
        
        .lblLifeStage.Caption = LoadResString(IDS_LIFESTAGE_CAP)
        .lblFumigationType.Caption = LoadResString(IDS_FUMTYPE_CAP)
        .lblPressureType.Caption = LoadResString(IDS_PRESTYPE_CAP)
        .lblCommodity.Caption = LoadResString(IDS_COMMODITY_CAP)
        .lblLoadFactor.Caption = LoadResString(IDS_LOADFACTOR_CAP)
        .cmdCalculate.Caption = LoadResString(IDS_CALCULATE_CAP)
        .cmdNextStep2.Caption = LoadResString(IDS_NEXTSTEP_CAP)
        
        'Set the tooltip text properties of the controls
        .txtSiteName.ToolTipText = LoadResString(IDS_TIP_SITE_NAME)
        .txtJobName.ToolTipText = LoadResString(IDS_TIP_JOB_NAME)
        .txtFumigationDate.ToolTipText = LoadResString(IDS_TIP_FUM_DATE)
        .txtLicFumigator.ToolTipText = LoadResString(IDS_TIP_LIC_FUMIGATOR)
        .lstTargetPests.ToolTipText = LoadResString(IDS_TIP_TARGET_PESTS)
        .cboCommodity.ToolTipText = LoadResString(IDS_TIP_COMMODITY)
        
        'PAT 02/10/2006 - Added new field
        .chkNoAdjSorp.Caption = LoadResString(IDS_ADJUSTED_SORP_CHKBOX)
        .cmdCalculate.ToolTipText = LoadResString(IDS_TIP_CALCULATE)
        
        
        'Retrieve the values from the database and display these in the UI
        .txtSiteName.Text = IIf(IsNull(rsJob("location_name")), "", rsJob("location_name"))
        .txtJobName.Text = IIf(IsNull(rsJob("name")), "", rsJob("name"))
        .txtFumigationDate.Text = IIf(IsNull(rsJob("fumigation_date")), "", rsJob("fumigation_date"))
        .txtLicFumigator.Text = IIf(IsNull(rsJob("fumigator_name")), "", rsJob("fumigator_name"))
            
        'Create the Job Pest recordset
        Call CreateJobPestRS
        
        'Initialize the options in the Target Pest list box
        nPestTotal = ProCalc.GetPestTotal
        .lstTargetPests.Clear
        For nCount = 0 To nPestTotal - 1
            .lstTargetPests.List(nCount) = ProCalc.GetPestString(nCount)
        Next
        If Not (rsJobPest.EOF And rsJobPest.BOF) Then
            bHasPest = True
            rsJobPest.MoveFirst
            While Not rsJobPest.EOF
                'We would not check the "deleted" field anymore.
                'If Not rsJobPest("deleted") Then
                '    .lstTargetPests.Selected(rsJobPest("pest_enum")) = True
                'End If
                .lstTargetPests.Selected(rsJobPest("pest_enum")) = True
                rsJobPest.MoveNext
            Wend
        End If
        
        'Initialize Life Stages drop-down control
        nLifeStageTotal = ProCalc.GetLifeStageTotal
        For nCount = 0 To nLifeStageTotal - 1
            .cboLifeStage.List(nCount) = ProCalc.GetLifeStageString(nCount)
        Next
        'Select the value from the Job recordset. Set the default to 0 if enum value
        'is less than 0.
        .cboLifeStage.ListIndex = IIf(rsJob("lifestage_enum") < 0, 0, rsJob("lifestage_enum"))
        
        'Initialize Fumigation Type drop-down control
        nFumTypeTotal = ProCalc.GetFumigationTypeTotal
        For nCount = 0 To nFumTypeTotal - 1
            .cboFumigationType.List(nCount) = ProCalc.GetFumigationTypeString(nCount)
        Next
        'Select the value from the Job recordset. Set the default to 0 if enum value
        'is less than 0.
        .cboFumigationType.ListIndex = IIf(rsJob("fumigation_type_enum") < 0, 0, _
            rsJob("fumigation_type_enum"))
        
        'Initialize Pressure Type (known as Treatment Type) drop-down control
        nPressTypeTotal = ProCalc.GetTreatmentTypeTotal
        For nCount = 0 To nPressTypeTotal - 1
            .cboPressureType.List(nCount) = ProCalc.GetTreatmentTypeString(nCount)
        Next
        'Select the value from the Job recordset. Set the default to 0 if enum value
        'is less than 0.
        .cboPressureType.ListIndex = IIf(rsJob("treatment_type_enum") < 0, 0, _
            rsJob("treatment_type_enum"))
            
        'Initialize Commodity drop-down control
        nCommodityTotal = ProCalc.GetCommodityTotal
        For nCount = 0 To nCommodityTotal - 1
            .cboCommodity.List(nCount) = ProCalc.GetCommodityString(nCount)
        Next
        'Select the value from the Job recordset. Set the default to 0 if enum value
        'is less than 0.
        .cboCommodity.ListIndex = IIf(rsJob("commodity_enum") < 0, 0, _
            rsJob("commodity_enum"))
            
        'Initialize Load Factor control
        .txtLoadFactor = rsJob("load_factor")

        'PAT 02/10/06 - Initialize Value of Checkbox
        If IsNull(rsJob("NoAdjSorp")) Or rsJob("NoAdjSorp") = 0 Then
            If gbChkEnabled Then
                '.chkNoAdjSorp.Enabled = True
                'Always disable
                .chkNoAdjSorp.Enabled = False
            Else
                .chkNoAdjSorp.Value = 0
                .chkNoAdjSorp.Enabled = False
            End If
        ElseIf rsJob("NoAdjSorp") = 1 Then
            .chkNoAdjSorp.Value = 1
        End If
    End With
    
    
    'Set the tooltip text property of the Results text box
    'This is to enable a multi-line tooltip text box.
    With mclsToolTip
        'Create the tooltip window.
        Call .Create(frmMain)
        
        'Set the tooltip's width so that it displays
        'multiline text and no tool's line length exceeds
        'roughly 240 pixels.
        .MaxTipWidth = 240
    
        'Show the tooltip for 20 seconds.
        .DelayTime(ttDelayShow) = 20000
    
        'Add the tooltip
        Call .AddTool(frmMain.rtfResults)
        
        'Set the text for Command1's tool.
        .ToolText(frmMain.rtfResults) = LoadResString(IDS_TIP_STRUCT_VOL) & vbCrLf & _
            LoadResString(IDS_TIP_AVG_CO) & vbCrLf & LoadResString(IDS_TIP_AVG_HLT) & _
            vbCrLf & LoadResString(IDS_TIP_AVG_CT) & vbCrLf & LoadResString(IDS_TIP_TOT_FUM) & _
            vbCrLf & LoadResString(IDS_TIP_WARNING)
            
            
        'CG 01/27/2004 Launch Changes - Modified the mouseover text for Life
                       'Stage, Load Factor and Fumigation Type
        'Add the tooltip
        Call .AddTool(frmMain.cboLifeStage)
        
        .ToolText(frmMain.cboLifeStage) = LoadResString(IDS_TOOLTIP_LIFESTAGE1) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_LIFESTAGE2) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_LIFESTAGE3) & vbCrLf & vbCrLf & _
                LoadResString(IDS_TOOLTIP_LIFESTAGE4) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_LIFESTAGE5) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_LIFESTAGE6) & vbCrLf & vbCrLf & _
                LoadResString(IDS_TOOLTIP_LIFESTAGE7) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_LIFESTAGE8) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_LIFESTAGE9) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_LIFESTAGE10) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_LIFESTAGE11) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_LIFESTAGE12) & vbCrLf
                
        'Add the tooltip
        Call .AddTool(frmMain.txtLoadFactor)
        .ToolText(frmMain.txtLoadFactor) = LoadResString(IDS_TOOLTIP_LOADFACTOR1) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_LOADFACTOR2) & vbCrLf '& _
                'LoadResString(IDS_TOOLTIP_LOADFACTOR3) & vbCrLf & _
                'LoadResString(IDS_TOOLTIP_LOADFACTOR4) & vbCrLf
        
        
          'Add the tooltip
        Call .AddTool(frmMain.cboFumigationType)
        .ToolText(frmMain.cboFumigationType) = LoadResString(IDS_TOOLTIP_FUMITYPE1) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_FUMITYPE2) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_FUMITYPE3) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_FUMITYPE4) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_FUMITYPE5) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_FUMITYPE6) & vbCrLf & _
                LoadResString(IDS_TOOLTIP_FUMITYPE7) & vbCrLf
        
        ' CG 2/6/2004 Add the tooltip
        Call .AddTool(frmMain.cboPressureType)
        .ToolText(frmMain.cboPressureType) = LoadResString(IDS_TIP_PRESS_TYPE) & vbCrLf & _
                LoadResString(IDS_TIP_PRESS_TYPE2) & vbCrLf & _
                LoadResString(IDS_TIP_PRESS_TYPE3) & vbCrLf & _
                LoadResString(IDS_TIP_PRESS_TYPE4) & vbCrLf & _
                LoadResString(IDS_TIP_PRESS_TYPE5)
        
    End With
    
    'Tell the app that is it done loading the form controls
    bIsLoading = False
    
    'Initialize the Area Info grid
    Call LoadAreaInfo
           
    'Initialize the Results text box
    Call InitResultsTextBox
    
    'Clean-up time!
    Set ProCalc = Nothing
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "InitializeControls", Err.Description)
    End If

End Sub

Private Sub CreateJobRS()
'**************************************************************
'   DESCRIPTION
'       This subroutine creates the Job recordset. This recordset
'       is always connected to the database so that any updates to
'       this recordset would be automatically committed to the
'       database. This recordset is available globally so that
'       other subroutines can access it.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim sJobSQL As String
    
    Set rsJob = New ADODB.Recordset
    
    sJobSQL = "select * from Job where deleted = False"
    rsJob.Open sJobSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    'It is assumed that there is always a record in the Job table
    rsJob.MoveFirst
    gsJobID = rsJob("id")
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "CreateJobRS", Err.Description)
    End If

End Sub

Private Sub CreateJobPestRS()
'**************************************************************
'   DESCRIPTION
'       This subroutine creates the Job Pest recordset.
'       This recordset is available globally so that
'       other subroutines can access it.
'
'   PARAMETERS
'          NONE
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'   10/21/02       Joseph Manyaga    JobPest table would no longer
'                                    have soft deletes. Therefore,
'                                    we no longer have to test the
'                                    deleted field.
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim sPestSQL As String
    
    'We need a list of all the pests even if these were deleted already. This will
    'be used when the user updates the Target Pests list box.
    'sPestSQL = "select * from JobPest where " & _
        "job_fkey = '" & gsJobID & "'" & " order by pest_enum, deleted"
    sPestSQL = "select * from JobPest where " & _
        "job_fkey = '" & gsJobID & "'" & " order by pest_enum"
        
    Set rsJobPest = New ADODB.Recordset
        
    rsJobPest.Open sPestSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "CreateJobPestRS", Err.Description)
    End If

End Sub

Private Sub LoadAreaInfo()
'**************************************************************
'   DESCRIPTION
'       This subroutine initializes the Area Info grid.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'      12/20/04     Yasmin Tan      added User Defined CT
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim nRow As Integer
    
    'Initialize the grid.
    GridWrap.Initialize "AreaInfo", frmMain.grdAreaInfo
    
    'Create the columns in the grid
    GridWrap.AddColumn "colAreaName", LoadResString(IDS_AREA_NAME_ASTERISK), colTypeText, flexAlignLeftTop
    GridWrap.AddColumn "colTemperature", LoadResString(IDS_TEMPERATURE), colTypeText, flexAlignLeftTop
    GridWrap.AddColumn "colEstHLT", LoadResString(IDS_EST_HLT), colTypeText, flexAlignLeftTop
    GridWrap.AddColumn "colPlanExposure", LoadResString(IDS_PLANNED_EXPOSURE), colTypeText, flexAlignLeftTop
    GridWrap.AddColumn "colAreaVolume", LoadResString(IDS_AREA_VOLUME), colTypeText, flexAlignLeftTop
    'This column is used in order to have a double line between the the Volume and Required
    'Fumigant columns.
    GridWrap.AddColumn "colSpace", " ", colTypeText, flexAlignLeftTop
    GridWrap.AddColumn "colFumigant", LoadResString(IDS_FUMIGANT_REQ), colTypeText, flexAlignLeftTop
    GridWrap.AddColumn "colCT", LoadResString(IDS_CT), colTypeText, flexAlignLeftTop
    '12/20/04-YYT added User defined CT in Grid
    GridWrap.AddColumn "colUserCT", LoadResString(IDS_USER_CT), colTypeText, flexAlignLeftTop
    GridWrap.AddColumn "colCo", LoadResString(IDS_CO), colTypeText, flexAlignLeftTop
    'This column is used to store the order/location of the record in the recordset.
    GridWrap.AddColumn "colOrderRS", " ", colTypeText, flexAlignLeftTop
    
    ' default column widths (will be overwritten if previous settings were saved)
    With frmMain.grdAreaInfo
        .ColWidth(0) = AREANAME_COLSIZE
        .ColWidth(1) = TEMP_COLSIZE
        .ColWidth(2) = ESTHLT_COLSIZE
        .ColWidth(3) = PLANEXP_COLSIZE
        .ColWidth(4) = AREAVOL_COLSIZE
        .ColWidth(5) = SPACE_COLSIZE
        .ColWidth(6) = FUMIGANT_COLSIZE
        .ColWidth(7) = CT_COLSIZE
        .ColWidth(8) = USER_CT_COLSIZE '12/20/04-YYT- Added User-defined CT, adjusted column #s
        .ColWidth(9) = CONC_COLSIZE
    
    End With
    
    
    'Format the columns according to the user's settings in the registry.
    GridWrap.LoadGridSettings GRID_SETTINGS_PATH
    
    frmMain.grdAreaInfo.ColHidden(COL_ORDER_RS) = True    'hide the Order column
    frmMain.grdAreaInfo.ColWidth(COL_SPACE) = 50         'set the size of the Space column
    frmMain.grdAreaInfo.Rows = GRID_MAX_ROWS              'Initialize the number of rows
    
    'Set the color of the read-only cells
    'For nRow = 1 To 49 Step 2
    '12/20/04 - YYT
    'gray out all columns from space up to target ct, then target concentration to end of columns
    'frmMain.grdAreaInfo.Cell(flexcpBackColor, 1, COL_SPACE, 100, COL_ORDER_RS) = &H8000000B
    frmMain.grdAreaInfo.Cell(flexcpBackColor, 1, COL_SPACE, 100, COL_CT) = &H8000000B
    frmMain.grdAreaInfo.Cell(flexcpBackColor, 1, COL_CO, 100, COL_ORDER_RS) = &H8000000B
    'Next
    
    
    
    'Identify the units for each column. Place the units in the flexcpData
    'attribute of the heading.
    Call GridIdentifyUnits
    
    'Fill the grid with the database values
    Call LoadAreaInfoRS
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "LoadAreaInfo", Err.Description)
    End If

End Sub
Private Sub GridIdentifyUnits()
'**************************************************************
'   DESCRIPTION
'       This subroutine sets the units for each column in the grid.
'       The units would be placed in the flexcpData attribute of the
'       first row, which contains the headings.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim dTemp As Double
    Dim sTempUnit As String
    Dim dVol As Double
    Dim sVolUnit As String
    Dim dFum As Double
    Dim sFumUnit As String
    
    With frmMain.grdAreaInfo
        'Call the UnitConvert function for dummy values in order to get the units
        'Temperature
        dTemp = 0#
        sTempUnit = ""
        Call UnitConvert(modeDisplay, unitTEMPERATURE, 1.5, dTemp, sTempUnit)
        .Cell(flexcpData, 0, COL_TEMPERATURE) = "°" & sTempUnit
        
        'Estimated HLT
        .Cell(flexcpData, 0, COL_EST_HLT) = " " & LoadResString(IDS_HOUR_ABBREV2)
        
        'Planned Exposure
        .Cell(flexcpData, 0, COL_PLAN_EXPOSURE) = " " & LoadResString(IDS_HOUR_ABBREV2)
        
        'Area Volume
        dVol = 0#
        sVolUnit = ""
        Call UnitConvert(modeDisplay, unitVOLUME, 1.5, dVol, sVolUnit)
        .Cell(flexcpData, 0, COL_VOLUME) = " " & sVolUnit
        
        'Required Fumigant
        dFum = 0#
        sFumUnit = ""
        Call UnitConvert(modeDisplay, unitWEIGHT, 1.5, dFum, sFumUnit)
        .Cell(flexcpData, 0, COL_FUMIGANT) = " " & sFumUnit
        
        'CT
        .Cell(flexcpData, 0, COL_CT) = " " & IIf(gbIsMetric, LoadResString(IDS_G_HR_PER_CUM), _
                            LoadResString(IDS_OZ_HR_PER_MCF))
                            
        '12/20/04-YYT - Added User Defined CT
        'USER-DEFINED CT
        .Cell(flexcpData, 0, COL_USER_CT) = " " & IIf(gbIsMetric, LoadResString(IDS_G_HR_PER_CUM), _
                            LoadResString(IDS_OZ_HR_PER_MCF))
                            
        'CO
        .Cell(flexcpData, 0, COL_CO) = " " & IIf(gbIsMetric, LoadResString(IDS_GRAMS_PER_CUBIC_METER), _
                            LoadResString(IDS_OZ_PER_MCF))
    End With
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "GridIdentifyUnits", Err.Description)
    End If

End Sub

Private Sub LoadAreaInfoRS()
'**************************************************************
'   DESCRIPTION
'       This subroutine creates and loads the Job Area recordset
'       into the Area Info grid.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'   12/20/04       Yasmin Tan       Added User-defined CT
'**************************************************************
    
    On Error GoTo Error_Handler
    
    'Declare the variables
    Dim sJobAreaSQL As String
    'Dim nCurrRow As Integer
    Dim dTemperature As Double
    Dim dVolume As Double
    Dim dFumigant As Double
    Dim dCT As Double
    Dim dUserCT As Double
    Dim dDose As Double
    Dim dCo As Double
    Dim dTempUserDefinedCT As Double
    Dim arrUserCT
    
    'Clear the grid of its contents
    frmMain.grdAreaInfo.Cell(flexcpData, 1, 0, GRID_MAX_ROWS - 1, 10) = ""
    frmMain.grdAreaInfo.Cell(flexcpText, 1, 0, GRID_MAX_ROWS - 1, 10) = ""
    
    'Reset the grid colors AND font
    frmMain.grdAreaInfo.Cell(flexcpForeColor, 1, COL_AREA_NAME, 100, COL_CO) = &H0
    frmMain.grdAreaInfo.Cell(flexcpBackColor, 1, COL_AREA_NAME, 100, COL_CO) = &H80000005
    '12/20/04 - YYT - put gray back color to read-only columns only excluding user defined ct
    'frmMain.grdAreaInfo.Cell(flexcpBackColor, 1, COL_SPACE, 100, COL_ORDER_RS) = &H8000000B
    frmMain.grdAreaInfo.Cell(flexcpBackColor, 1, COL_SPACE, 100, COL_CT) = &H8000000B
    frmMain.grdAreaInfo.Cell(flexcpBackColor, 1, COL_CO, 100, COL_ORDER_RS) = &H8000000B
    frmMain.grdAreaInfo.Cell(flexcpFontBold, 1, COL_AREA_NAME, 100, COL_ORDER_RS) = False
    
    'Initialize the string containing the list of Area Names
    sGridAreaName = ""
    
    'Check if the Job Area recordset is already open. Close it first before re-opening
    If rsJobArea.State Then
        rsJobArea.Close
    End If
    
    sJobAreaSQL = "select JobArea.* from JobArea where deleted = false and " & _
        "job_fkey = '" & gsJobID & "' order by row"
    
    'Open the recordset. This recordset is always open for updates.
    rsJobArea.Open sJobAreaSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    'Display the recordset in the grid if the recordset is not empty. Assign to flexcpText
    'the value and the corresponding units. Assign the value only to flexcpData.
    nCurrRow = 1        'the current row in the grid
    nRecordNumber = 1    'the current record number
    If Not rsJobArea.BOF And Not rsJobArea.EOF Then
        bHasJobArea = True  'set global boolean
        With frmMain.grdAreaInfo
            rsJobArea.MoveFirst
            'We are always sure that the user could not enter beyond the maximum number of rows.
            'Therefore, we don't need to test this condition. If in case that the records were
            'directly entered in the database and these exceed the maximum number, then the
            'error handler should handle this, rather than displaying incomplete records in
            'the user-interface.
            While Not rsJobArea.EOF
                'Display only if we are in the correct grid location.
                If rsJobArea("row") = nCurrRow Then
                    'area name
                    .Cell(flexcpData, nCurrRow, COL_AREA_NAME) = CStr(rsJobArea("area_name"))
                    .Cell(flexcpText, nCurrRow, COL_AREA_NAME) = rsJobArea("area_name")
                    sGridAreaName = sGridAreaName & "|" & CStr(rsJobArea("area_name")) & "|"
                                    
                    'temperature
                    dTemperature = 0#
                    Call UnitConvert(modeDisplay, unitTEMPERATURE, rsJobArea("temperature"), _
                        dTemperature)
                        
                    'JNM 11/13 - Changed nDecimal to 1
                    dTemperature = Round(dTemperature, IIf(gbIsMetric, 1, 0))
                    .Cell(flexcpData, nCurrRow, COL_TEMPERATURE) = FormatNumber(dTemperature, IIf(gbIsMetric, 1, 0))
                    .Cell(flexcpText, nCurrRow, COL_TEMPERATURE) = FormatNumber(dTemperature, IIf(gbIsMetric, 1, 0)) & _
                        .Cell(flexcpData, 0, COL_TEMPERATURE)
                                    
                    'estimated HLT
                    .Cell(flexcpData, nCurrRow, COL_EST_HLT) = FormatNumber(Round(rsJobArea("estimated_hlt"), 1), 1, , , vbFalse)
                    .Cell(flexcpText, nCurrRow, COL_EST_HLT) = FormatNumber(Round(rsJobArea("estimated_hlt"), 1), 1, , , vbFalse) & _
                        .Cell(flexcpData, 0, COL_EST_HLT)
                    
                    'planned exposure
                    .Cell(flexcpData, nCurrRow, COL_PLAN_EXPOSURE) = FormatNumber(Round(rsJobArea("planned_exposure"), 1), 1, , , vbFalse)
                    .Cell(flexcpText, nCurrRow, COL_PLAN_EXPOSURE) = FormatNumber(Round(rsJobArea("planned_exposure"), 1), 1, , , vbFalse) & _
                        .Cell(flexcpData, 0, COL_PLAN_EXPOSURE)
                        
                    'volume
                    dVolume = 0#
                    Call UnitConvert(modeDisplay, unitVOLUME, rsJobArea("volume"), dVolume)
                    'JNM 11/13 - Changed nDecimal to 0
                    dVolume = Round(dVolume, 0)
                    .Cell(flexcpData, nCurrRow, COL_VOLUME) = FormatNumber(dVolume, 0, , , vbTrue)
                    .Cell(flexcpText, nCurrRow, COL_VOLUME) = FormatNumber(dVolume, 0, , , vbTrue) & _
                        .Cell(flexcpData, 0, COL_VOLUME)
                    
                    'required fumigant
                    'Display the fumigant if this has already been calculated. Otherwise,
                    'display text saying that the required fumigant still needs to be
                    'calculated.
                    If rsJobArea("fumigant") >= 0# Then
                        dFumigant = 0#
                        Call UnitConvert(modeDisplay, unitWEIGHT, rsJobArea("fumigant"), dFumigant)
                        dFumigant = Round(dFumigant, 0)
                        .Cell(flexcpText, nCurrRow, COL_FUMIGANT) = FormatNumber(dFumigant, 0, , , vbFalse) & _
                            .Cell(flexcpData, 0, COL_FUMIGANT)
                    Else
                        .Cell(flexcpText, nCurrRow, COL_FUMIGANT) = LoadResString(IDS_NO_CALC)
                    End If
                    
                    'concentration time
                    'Display the CT if this has already been calculated. Otherwise,
                    'leave the field blank.
                    If rsJobArea("concentration_time") >= 0# Then
                        dCT = Round(rsJobArea("concentration_time"), 0)
                        dDose = Round(rsJobArea("dose"), 2) 'display this only if testing is active
                        If gbIsTestingActive Then
                            .Cell(flexcpText, nCurrRow, COL_CT) = FormatNumber(dCT, 0, , , vbFalse) & _
                                " (" & FormatNumber(dDose, 2, , , vbFalse) & ")" & .Cell(flexcpData, 0, COL_CT)
                        Else
                            .Cell(flexcpText, nCurrRow, COL_CT) = FormatNumber(dCT, 0, , , vbFalse) & _
                                .Cell(flexcpData, 0, COL_CT)
                        End If
                    End If
                    
                    '12/20/04 - YYT - Added user-defined CT
                    'user-defined concentration time
                    'Display it if it has already been populated, else, display 0.
                    dTempUserDefinedCT = IIf(IsNull(rsJobArea("user_defined_CT")), 0, rsJobArea("user_defined_CT"))
                    If rsJobArea("user_defined_CT") >= 0# Then
                        .Cell(flexcpData, nCurrRow, COL_USER_CT) = FormatNumber(Round(dTempUserDefinedCT, 0), 0, , , vbFalse)
                        .Cell(flexcpText, nCurrRow, COL_USER_CT) = FormatNumber(Round(dTempUserDefinedCT, 0), 0, , , vbFalse) & _
                            .Cell(flexcpData, 0, COL_USER_CT)
                    End If
                    
                    'initial concentration
                    'Display the Co if this has already been calculated. Otherwise,
                    'leave the field blank.
                    If rsJobArea("initial_concentration") >= 0# Then
                        dCo = Round(rsJobArea("initial_concentration"), 0)
                        .Cell(flexcpText, nCurrRow, COL_CO) = FormatNumber(dCo, 0, , , vbFalse) & .Cell(flexcpData, 0, COL_CO)
                    End If
                    
                    'recordset order
                    'Set this field to the current record number
                    .Cell(flexcpData, nCurrRow, COL_ORDER_RS) = nRecordNumber
                    
                    'set the background color to red if record has a warning
                    If rsJobArea("calculation_warning") <> "" Then
                        'red
                        .Cell(flexcpBackColor, nCurrRow, COL_AREA_NAME, nCurrRow, COL_CO) = &HFF& '&H8080FF
                        .Cell(flexcpForeColor, nCurrRow, COL_AREA_NAME, nCurrRow, COL_CO) = &HFFFFFF
                        .Cell(flexcpFontBold, nCurrRow, COL_AREA_NAME, nCurrRow, COL_CO) = True
                    End If
                    
                    'Increment the record number and move to the next record
                    nRecordNumber = nRecordNumber + 1
                    rsJobArea.MoveNext
                    
                End If
                'Increment nCurrRow in order to move to the next grid row
                nCurrRow = nCurrRow + 1
            Wend
        End With
    Else
        bHasJobArea = False
    End If
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "LoadAreaInfoRS", Err.Description)
    End If

End Sub

Private Sub InitResultsTextBox()
'**************************************************************
'   DESCRIPTION
'       This subroutine displays the text in the Results
'       text box control.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim sResults As String  'the string containing the results
    Dim nTotVolume As Double
    Dim nAvgCo As Double
    Dim nAvgHLT As Double
    Dim nAvgCT As Double
    Dim nAvgCnT As Double
    Dim nTotFumigant As Double
    
    Dim sVolUnit As String
    Dim sCoUnit As String
    Dim sHLTUnit As String
    Dim sCTUnit As String
    Dim sFumUnit As String
    Dim sResult1 As String
    Dim sResult2 As String
    Dim sResult3 As String
    Dim sWarningLabel As String
    
    Dim arrAreaWarnings(2) As String
    Dim sAreaWarnings As String
    Dim sWarnings As String
    
    sResults = ""
    sResult1 = ""
    sResult2 = ""
    sResult3 = ""
    sWarningLabel = ""
    sWarnings = ""
    
    
    'Create the Results string
    If Not IsNull(rsJob("calculation_error")) And rsJob("calculation_error") <> "" Then
        sResult1 = rsJob("calculation_error") & vbCrLf
        sResult2 = vbCrLf
    Else
        If rsJob("fumigant") < 0# Then
            sResult1 = LoadResString(IDS_NO_CALC) & vbCrLf
            sResult2 = vbCrLf
        Else
            nTotVolume = 0#
            Call UnitConvert(modeDisplay, unitVOLUME, rsJob("volume"), nTotVolume)
            nTotVolume = Round(nTotVolume, 0)
            
            nAvgCo = Round(rsJob("avg_initial_concentration"), 1)
            nAvgHLT = Round(rsJob("avg_estimated_hlt"), 1)
            nAvgCT = Round(rsJob("avg_concentration_time"), 0)
            nAvgCnT = Round(rsJob("avg_dose"), 2)
            nTotFumigant = IIf(gbIsMetric, Round(rsJob("fumigant"), 0), _
                Round(rsJob("fumigant") * POUNDS_PER_KG, 0))
            sVolUnit = frmMain.grdAreaInfo.Cell(flexcpData, 0, COL_VOLUME)
            sCoUnit = frmMain.grdAreaInfo.Cell(flexcpData, 0, COL_CO)
            sHLTUnit = frmMain.grdAreaInfo.Cell(flexcpData, 0, COL_EST_HLT)
            sCTUnit = frmMain.grdAreaInfo.Cell(flexcpData, 0, COL_CT)
            sFumUnit = frmMain.grdAreaInfo.Cell(flexcpData, 0, COL_FUMIGANT)
            
            'Assemble the first Results string
            sResult1 = LoadResString(IDS_TOTAL_FUMIGANT_CAP) & FormatNumber(nTotFumigant, 0, , , vbFalse) & sFumUnit & _
                " " & LoadResString(IDS_OR) & " " & Round(rsJob("cylinders"), 1) & " " & LoadResString(IDS_NUM_CYLINDERS) & _
                Space(SPACE_NUM) & LoadResString(IDS_TOT_STRUCTURE_VOL_CAP) & FormatNumber(nTotVolume, 0, , , vbTrue) & sVolUnit & _
                Space(SPACE_NUM) & LoadResString(IDS_AVG_CO_CAP) & FormatNumber(nAvgCo, 1, , , vbFalse) & sCoUnit & _
                Space(SPACE_NUM) & LoadResString(IDS_AVG_HLT_CAP) & FormatNumber(nAvgHLT, 1, , , vbFalse) & sHLTUnit & _
                Space(SPACE_NUM) & LoadResString(IDS_AVG_CT_CAP) & FormatNumber(nAvgCT, 0, , , vbFalse) & sCTUnit
            'Display CnT only if Testing is active
            If gbIsTestingActive Then
                sResult1 = sResult1 & Space(SPACE_NUM) & LoadResString(IDS_AVG_CNT_CAP) & _
                    FormatNumber(nAvgCnT, 2, , , vbFalse) & sCTUnit & vbCrLf & vbCrLf
            Else
                sResult1 = sResult1 & vbCrLf & vbCrLf
            End If
            
        End If
    End If
    
    sResult3 = String(228, "-") & vbCrLf
    sResults = sResult1
    
    'Create the Warnings string.
    If bHasJobArea Then
        rsJobArea.MoveFirst
        While Not rsJobArea.EOF
            If Len(rsJobArea("calculation_warning")) > 0 Then
                sAreaWarnings = ""
                arrAreaWarnings(1) = rsJobArea("calculation_warning")
                arrAreaWarnings(0) = rsJobArea("area_name") & ": "
                sAreaWarnings = Join(arrAreaWarnings, "")
                sAreaWarnings = Replace(sAreaWarnings, vbCrLf, vbCrLf & rsJobArea("area_name") & ": ")
                If sWarnings <> "" Then
                    sWarnings = sWarnings & vbCrLf
                End If
                sWarnings = sWarnings & sAreaWarnings & vbCrLf
            End If
            rsJobArea.MoveNext
        Wend
    End If
    
    sWarningLabel = LoadResString(IDS_WARNING_CAP) & vbCrLf
    If sWarnings <> "" Then
        sWarnings = sWarningLabel & sWarnings
        sResults = sResults & sResult3
    End If
                
    'Assemble the whole results string
    sResults = sResults & sWarnings
    
    'Display and format the Results string in the text box
    With frmMain.rtfResults
        'Display the value
        .TextRTF = sResults
        
        'Format the entire text
        .SelStart = 0
        .SelLength = Len(sResults)
        .SelFontSize = RTF_FONT_SIZE
        .SelFontName = RTF_FONT_NAME
        
        .SelStart = 0
        .SelLength = Len(LoadResString(IDS_TOTAL_FUMIGANT_CAP))
        .SelBold = True
        
        .SelStart = Len(LoadResString(IDS_TOTAL_FUMIGANT_CAP) & _
            FormatNumber(nTotFumigant, 0, , , vbFalse) & sFumUnit & _
            " or " & Round(rsJob("cylinders"), 1) & " " & LoadResString(IDS_NUM_CYLINDERS)) + _
            SPACE_NUM
        .SelLength = Len(LoadResString(IDS_TOT_STRUCTURE_VOL_CAP))
        .SelBold = True
        
        .SelStart = Len(LoadResString(IDS_TOTAL_FUMIGANT_CAP) & _
            FormatNumber(nTotFumigant, 0, , , vbFalse) & sFumUnit & _
            " or " & Round(rsJob("cylinders"), 1) & " " & LoadResString(IDS_NUM_CYLINDERS)) + _
            SPACE_NUM + Len(LoadResString(IDS_TOT_STRUCTURE_VOL_CAP) & _
            FormatNumber(nTotVolume, 0, , , vbFalse) & sVolUnit) + SPACE_NUM
        .SelLength = Len(LoadResString(IDS_AVG_CO_CAP))
        .SelBold = True
                
        .SelStart = Len(LoadResString(IDS_TOTAL_FUMIGANT_CAP) & _
            FormatNumber(nTotFumigant, 0, , , vbFalse) & sFumUnit & _
            " or " & Round(rsJob("cylinders"), 1) & " " & LoadResString(IDS_NUM_CYLINDERS)) + _
            SPACE_NUM + Len(LoadResString(IDS_TOT_STRUCTURE_VOL_CAP) & _
            FormatNumber(nTotVolume, 0, , , vbFalse) & sVolUnit) + SPACE_NUM + _
            Len(LoadResString(IDS_AVG_CO_CAP) & _
            FormatNumber(nAvgCo, 1, , , vbFalse) & sCoUnit) + SPACE_NUM
        .SelLength = Len(LoadResString(IDS_AVG_HLT_CAP))
        .SelBold = True
        
        .SelStart = Len(LoadResString(IDS_TOTAL_FUMIGANT_CAP) & _
            FormatNumber(nTotFumigant, 0, , , vbFalse) & sFumUnit & _
            " or " & Round(rsJob("cylinders"), 1) & " " & LoadResString(IDS_NUM_CYLINDERS)) + _
            SPACE_NUM + Len(LoadResString(IDS_TOT_STRUCTURE_VOL_CAP) & _
            FormatNumber(nTotVolume, 0, , , vbFalse) & sVolUnit) + SPACE_NUM + _
            Len(LoadResString(IDS_AVG_CO_CAP) & _
            FormatNumber(nAvgCo, 1, , , vbFalse) & sCoUnit) + SPACE_NUM + _
            Len(LoadResString(IDS_AVG_HLT_CAP) & _
            FormatNumber(nAvgHLT, 1, , , vbFalse) & sHLTUnit) + SPACE_NUM
        .SelLength = Len(LoadResString(IDS_AVG_CT_CAP))
        .SelBold = True
                
        'Format the Fumigant section
        '.SelStart = Len(sResult1)
        '.SelLength = Len(sResult2)
        '.SelBold = True
        
        If sWarnings <> "" Then
            'Format the Warning label
            '.SelStart = Len(sResult1) + Len(sResult2) + Len(sResult3)
            .SelStart = Len(sResult1) + Len(sResult3)
            .SelLength = Len(sWarningLabel)
            .SelBold = True
            .SelColor = &HFF&
            
            'JNM 11/14 - Format the Warnings text
            .SelStart = Len(sResult1) + Len(sResult3)
            .SelLength = Len(sWarnings)
            .SelColor = &HFF&
            
        End If
        .SelStart = 0
        .SelLength = 0
    End With
    
    bIsCleared = False
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "InitResultsTextBox", Err.Description)
    End If

End Sub


Private Function ValidateUpdateControls(ByVal sControlName As String) As Boolean
'**************************************************************
'   DESCRIPTION
'       This function validates the form controls. If the value is valid,
'       the value is saved in the database.
'
'   PARAMETERS
'           sControlName = the name of the Control that
'               needs to be validated
'
'   RETURNS - True - if value is valid
'             False - if value is invalid
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'       10/21/02   Joseph Manyaga    Removed soft-deletes for JobPest.
'                                    JobPests would be deleted first and then
'                                    selected values will be inserted.
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim nCount As Integer
    Dim sPestList As String
    Dim nLoadFactor As Long
    
    With frmMain
        Select Case sControlName
            'We don't need to validate the optional text boxes.
            Case "txtJobName"
                rsJob("name") = Trim(.txtJobName.Text)
                rsJob.Update
                ValidateUpdateControls = True
            Case "txtLicFumigator"
                rsJob("fumigator_name") = Trim(.txtLicFumigator.Text)
                rsJob.Update
                ValidateUpdateControls = True
            Case "txtSiteName"
                rsJob("location_name") = Trim(.txtSiteName.Text)
                rsJob.Update
                ValidateUpdateControls = True
            Case "txtFumigationDate"
                rsJob("fumigation_date") = Trim(.txtFumigationDate.Text)
                rsJob.Update
                ValidateUpdateControls = True
            Case "lstTargetPests"
                'If we are are in VALIDATE mode, then we should ensure that
                'the list box has at least one selected value. We are in VALIDATE
                'mode when the user is requesting for an operation like printing,
                'calculating, etc. except when exiting or when the user is still
                'editing the form.
                'If sValidateMode = "VALIDATE" And .lstTargetPests.SelCount = 0 Then
                '    MsgBox LoadResString(IDS_NO_PEST)
                '    .lstTargetPests.SetFocus
                '    ValidateUpdateControls = False
                'Else
                    sPestList = ""
                    If bHasPest Then
                        rsJobPest.MoveFirst
                        'No more soft deletes
                        While Not rsJobPest.EOF
                            'rsJobPest("deleted") = Not .lstTargetPests.Selected(rsJobPest("pest_enum"))
                            'rsJobPest.Update
                            'sPestList = sPestList & "|" & rsJobPest("pest_enum") & "|"
                            rsJobPest.Delete
                            rsJobPest.MoveNext
                        Wend
                        bHasPest = False    'JNM 04/14/03 - Reset the value of bHasPest.
                    End If
                    
                    nCount = 0
                    'Insert the new pests into the database.
                    For nCount = 0 To (nPestTotal - 1)
                        'If .lstTargetPests.Selected(nCount) And _
                        'InStr(sPestList, "|" & CStr(nCount) & "|") = 0 Then
                        '    rsJobPest.AddNew
                        '    rsJobPest("job_fkey") = gsJobID
                        '    rsJobPest("pest_enum") = nCount
                        '    rsJobPest("deleted") = False
                        '    rsJobPest.Update
                        'End If
                        If .lstTargetPests.Selected(nCount) Then
                            rsJobPest.AddNew
                            rsJobPest("job_fkey") = gsJobID
                            rsJobPest("pest_enum") = nCount
                            rsJobPest("deleted") = False
                            rsJobPest.Update
                            bHasPest = True    'JNM 04/14/03 - Reset the value of bHasPest.
                        End If
                    Next
                    ValidateUpdateControls = True
                'End If
            Case "cboFumigationType"
                rsJob("fumigation_type_enum") = .cboFumigationType.ListIndex
                rsJob.Update
                ValidateUpdateControls = True
            Case "cboLifeStage"
                rsJob("lifestage_enum") = .cboLifeStage.ListIndex
                rsJob.Update
                ValidateUpdateControls = True
            Case "cboPressureType"
                rsJob("treatment_type_enum") = .cboPressureType.ListIndex
                rsJob.Update
                ValidateUpdateControls = True
            Case "cboCommodity"
                rsJob("commodity_enum") = .cboCommodity.ListIndex
                rsJob.Update
                ValidateUpdateControls = True
            Case "txtLoadFactor"
                nLoadFactor = 0
                If Len(Trim(.txtLoadFactor.Text)) > 0 Then
                    If Not IsNumeric(.txtLoadFactor.Text) Then
                        'ReturnCurrTab TAB_DOSAGE_PLAN  - JNM AT FIX
                        MsgBox LoadResString(IDS_LOAD_FACTOR_ERROR), vbExclamation
                        .txtLoadFactor.SetFocus
                        ValidateUpdateControls = False
                    Else
                        nLoadFactor = CLng(Trim(.txtLoadFactor.Text))
                        'Test if value is an integer
                        If Not IsInteger(.txtLoadFactor.Text) Then
                            'ReturnCurrTab TAB_DOSAGE_PLAN - JNM AT FIX
                            MsgBox LoadResString(IDS_NOT_INTEGER), vbExclamation
                            .txtLoadFactor.SetFocus
                            ValidateUpdateControls = False
                        Else
                            If nLoadFactor < MIN_LOAD_FACTOR Or nLoadFactor > MAX_LOAD_FACTOR Then
                                'ReturnCurrTab TAB_DOSAGE_PLAN - JNM AT FIX
                                MsgBox LoadResString(IDS_LOAD_FACTOR_ERROR), vbExclamation
                                .txtLoadFactor.SetFocus
                                ValidateUpdateControls = False
                            Else
                                rsJob("load_factor") = nLoadFactor
                                rsJob.Update
                                ValidateUpdateControls = True
                            End If
                        End If
                    End If
                Else
                    rsJob("load_factor") = MIN_LOAD_FACTOR
                    frmMain.txtLoadFactor.Text = MIN_LOAD_FACTOR
                    rsJob.Update
                    ValidateUpdateControls = True
                End If
            Case "chkNoAdjSorp"
                rsJob("NoAdjSorp") = .chkNoAdjSorp.Value
                rsJob.Update
                ValidateUpdateControls = True
            'The case below is used to check if a calculation is required
            Case "Calculation"
                If HasMonitoringData And rsJob("fumigant") = -1 Then
                    ReturnCurrTab TAB_DOSAGE_PLAN
                    MsgBox LoadResString(IDS_CALC_FUMIGANT_REQ), vbExclamation
                    ValidateUpdateControls = False
                Else
                    ValidateUpdateControls = True
                End If
            Case Else
                ValidateUpdateControls = True
             
        End Select
    End With
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "ValidateUpdateControls", Err.Description)
    End If

End Function

Private Function HasMonitoringData() As Boolean
'**************************************************************
'   DESCRIPTION
'       This function determines if the Job has monitoring data
'       associated with it.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - True - if there are monitoring data
'             False - if there are no monitoring data
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim sMonitorPointDataSQL As String
    Dim rsMonitorPointData As New ADODB.Recordset
    
    sMonitorPointDataSQL = "select MonitorPointValue.* from MonitorPointValue, " & _
        "MonitorPoint where MonitorPointValue.monitorpoint_fkey = MonitorPoint.id " & _
        "and MonitorPointValue.deleted = False and MonitorPoint.deleted = False " & _
        "and MonitorPoint.job_fkey = '" & gsJobID & "'"
    rsMonitorPointData.Open sMonitorPointDataSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    If rsMonitorPointData.BOF And rsMonitorPointData.EOF Then
        HasMonitoringData = False
    Else
        HasMonitoringData = True
    End If
    
    'Clean-up
    rsMonitorPointData.Close
    Set rsMonitorPointData = Nothing
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "HasMonitoringData", Err.Description)
    End If

End Function

Private Function GridValidate(ByVal lRow As Long, ByVal lCol As Long) As Boolean
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
'      12/21/04     Yasmin Tan      Added Validation for user-defined CT
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim dTemperature As Double
    Dim dMinVol As Double
    Dim dMaxVol As Double
    Dim dVolume As Double
    Dim dExposure As Double
    Dim dHLT As Double
    Dim dMinTemp As Double
    Dim dMaxTemp As Double
    Dim sErrMsg As String
    Dim dUserCT As Double
    Dim sJobAreaSQL
    Dim arrTargetCT
    
    'Check if the Job Area recordset is already open. Close it first before re-opening
    If rsJobArea.State Then
        rsJobArea.Close
    End If
    
    sJobAreaSQL = "select JobArea.* from JobArea where deleted = false and " & _
        "job_fkey = '" & gsJobID & "' order by row"
    
    'Open the recordset. This recordset is always open for updates.
    rsJobArea.Open sJobAreaSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    nCurrRow = 1        'the current row in the grid
  
    With frmMain.grdAreaInfo
    
    sErrMsg = ""
    Select Case lCol
        Case COL_AREA_NAME
            If Len(Trim(.EditText)) = 0 Then
                MsgBox LoadResString(IDS_NO_AREA_NAME), vbExclamation
                GridValidate = False
            Else
                If .EditText = .Cell(flexcpData, lRow, lCol) Then
                    GridValidate = True
                Else
                    If InStr(sGridAreaName, "|" & Trim(.EditText) & "|") > 0 Then
                        MsgBox LoadResString(IDS_FUMPLAN_MSG_DUPLICATE_AREA), vbExclamation
                        GridValidate = False
                    Else
                        GridValidate = True
                    End If
                End If
            End If
        Case COL_TEMPERATURE
            dMinTemp = IIf(gbIsMetric, MIN_TEMP_METRIC, MIN_TEMP_ENGLISH)
            dMaxTemp = IIf(gbIsMetric, MAX_TEMP_METRIC, MAX_TEMP_ENGLISH)
            
            'Temperature should be numeric
            If Len(Trim(.EditText)) = 0 Or Not IsNumeric(Trim(.EditText)) Then
                MsgBox FormatSWParams(LoadResString(IDS_AREA_INFO_TEMP_NOT_A_NUMBER), _
                    dMinTemp & "|" & dMaxTemp & "|" & .Cell(flexcpData, 0, COL_TEMPERATURE)), vbExclamation
                GridValidate = False
            Else
                'If we are in English units, temperature should be a whole number
                dTemperature = Round(CDbl(Trim(.EditText)), 0)
                If Not gbIsMetric And (dTemperature - CDbl(Trim(.EditText))) <> 0 Then
                    MsgBox LoadResString(IDS_TEMP_WHOLE_NUM), vbExclamation
                    GridValidate = False
                Else
                    'Temperature should be within the allowable range
                    dTemperature = CDbl(Trim(.EditText))
                    If dTemperature < dMinTemp Or dTemperature > dMaxTemp Then
                        MsgBox FormatSWParams(LoadResString(IDS_AREA_INFO_TEMP_OUT_OF_RANGE), _
                            dMinTemp & "|" & dMaxTemp & "|" & .Cell(flexcpData, 0, COL_TEMPERATURE)), vbExclamation
                        GridValidate = False
                    Else
                        GridValidate = True
                    End If
                End If
            End If
        Case COL_EST_HLT
            'Estimated HLT should be numeric
            If Len(Trim(.EditText)) = 0 Or Not IsNumeric(Trim(.EditText)) Then
                MsgBox FormatSWParams(LoadResString(IDS_AREA_INFO_EST_HLT_NOT_A_NUMBER), _
                    FormatNumber(MIN_EST_HLT, 2, , , vbFalse) & "|" & FormatNumber(MAX_EST_HLT, 2, , , vbFalse)), vbExclamation
                GridValidate = False
            Else
                'Estimated HLT should be within the allowable range
                dHLT = CDbl(Trim(.EditText))
                If dHLT < MIN_EST_HLT Or dHLT > MAX_EST_HLT Then
                    MsgBox FormatSWParams(LoadResString(IDS_AREA_INFO_EST_HLT_OUT_OF_RANGE), _
                    FormatNumber(MIN_EST_HLT, 2, , , vbFalse) & "|" & FormatNumber(MAX_EST_HLT, 2, , , vbFalse)), vbExclamation
                    GridValidate = False
                Else
                    GridValidate = True
                End If
            End If
        Case COL_PLAN_EXPOSURE
            'Planned Exposure should be numeric
            If Len(Trim(.EditText)) = 0 Or Not IsNumeric(Trim(.EditText)) Then
                MsgBox FormatSWParams(LoadResString(IDS_AREA_INFO_PLANNED_EXPOSURE_NOT_A_NUMBER), _
                    FormatNumber(MIN_PLANNED_EXPOSURE, 2, , , vbFalse) & "|" & FormatNumber(MAX_PLANNED_EXPOSURE, 2, , , vbFalse)), vbExclamation
                GridValidate = False
            Else
                'Planned Exposure should be within the allowable range
                dExposure = CDbl(Trim(.EditText))
                If dExposure < MIN_PLANNED_EXPOSURE Or dExposure > MAX_PLANNED_EXPOSURE Then
                    MsgBox FormatSWParams(LoadResString(IDS_AREA_INFO_PLANNED_EXPOSURE_OUT_OF_RANGE), _
                    FormatNumber(MIN_PLANNED_EXPOSURE, 2, , , vbFalse) & "|" & FormatNumber(MAX_PLANNED_EXPOSURE, 2, , , vbFalse)), vbExclamation
                    GridValidate = False
                Else
                    GridValidate = True
                End If
            End If
        Case COL_VOLUME
            dMinVol = IIf(gbIsMetric, MIN_VOLUME_METRIC, MIN_VOLUME_ENGLISH)
            dMaxVol = IIf(gbIsMetric, MAX_VOLUME_METRIC, MAX_VOLUME_ENGLISH)
            
            'Volume should be numeric
            If Len(Trim(.EditText)) = 0 Or Not IsNumeric(Trim(.EditText)) Then
                MsgBox FormatSWParams(LoadResString(IDS_AREA_INFO_VOL_NOT_A_NUMBER), _
                    FormatNumber(dMinVol, nDecimal) & "|" & FormatNumber(dMaxVol, nDecimal) & "|" & .Cell(flexcpData, 0, COL_VOLUME)), vbExclamation
                GridValidate = False
            Else
                'If we are in English units, volume should be a whole number
                dVolume = Round(CDbl(Trim(.EditText)), 0)
                If Not gbIsMetric And (dVolume - CDbl(Trim(.EditText))) <> 0 Then
                    MsgBox LoadResString(IDS_VOL_WHOLE_NUM), vbExclamation
                    GridValidate = False
                Else
                    'Volume should be within the allowable range
                    dVolume = CDbl(Trim(.EditText))
                    If dVolume < dMinVol Or dVolume > dMaxVol Then
                        MsgBox FormatSWParams(LoadResString(IDS_AREA_INFO_VOL_OUT_OF_RANGE), _
                            FormatNumber(dMinVol, nDecimal) & "|" & FormatNumber(dMaxVol, nDecimal) & "|" & .Cell(flexcpData, 0, COL_VOLUME)), vbExclamation
                        GridValidate = False
                    Else
                        GridValidate = True
                    End If
                End If
            End If
        '12/21/04-YYT - Added validation for user-defined CT
        Case COL_USER_CT
            Do While Not rsJobArea.EOF And Not rsJobArea.BOF
                If rsJobArea("row") = lRow Then
                arrTargetCT = Split(.Cell(flexcpText, lRow, COL_CT), " ")
                
                  'User-defined CT should be numeric
                    'If Len(Trim(.EditText)) = 0 Or Not IsNumeric(Trim(.EditText)) Then
                    If Len(Trim(.EditText)) = 0 Then
                        GridValidate = True
                    ElseIf Not IsNumeric(Trim(.EditText)) Then
                        MsgBox LoadResString(IDS_AREA_INFO_USER_CT_NOT_A_NUMBER), vbExclamation
                        GridValidate = False
                    ElseIf .Cell(flexcpText, lRow, COL_CT) = Null Or .Cell(flexcpText, lRow, COL_CT) = "" Then
                        'Pat 4/28/2005: Check if there's a computed Target CT
                        MsgBox FormatSWParams(LoadResString(IDS_TARGETCT_NOT_PRESENT), _
                        MAX_VACUUM_USER_CT), vbExclamation
                        GridValidate = False

                    Else
                        '12/21/04-YYT-User-defined CT should have max of 200/1500 depending on pressure type
                        'Need to change this in the future/and ProfumeCalc as well because it is currently
                        'not allowing user to exceed max CT
                        'if pressure type is vacuum
                        dUserCT = CDbl(Trim(.EditText))
                        'if Vacuum
                        If ((frmMain.cboPressureType.ListIndex) = 1 And (dUserCT > MAX_VACUUM_USER_CT Or dUserCT < MIN_NORMAL_VACUUM_USER_CT)) Then
                            MsgBox FormatSWParams(LoadResString(IDS_AREA_INFO_USER_CT_OUT_OF_RANGE), _
                            MAX_VACUUM_USER_CT & "|" & MIN_NORMAL_VACUUM_USER_CT), vbExclamation
                            GridValidate = False
                        Else
                            'if Normal Compression
                            If ((frmMain.cboPressureType.ListIndex) = 0 And (dUserCT > MAX_NORMAL_USER_CT Or dUserCT < MIN_NORMAL_VACUUM_USER_CT)) Then
                                MsgBox FormatSWParams(LoadResString(IDS_AREA_INFO_USER_CT_OUT_OF_RANGE), _
                                MAX_NORMAL_USER_CT & "|" & MIN_NORMAL_VACUUM_USER_CT), vbExclamation
                                GridValidate = False
                            Else
                                GridValidate = True
                            End If
                        End If
                    End If

                End If
        
               
                
            rsJobArea.MoveNext
            Loop
        
        End Select
    End With
    
    'If sErrMsg <> "" Then
    '    ReturnCurrTab TAB_DOSAGE_PLAN
    '    MsgBox sErrMsg
    'End If
    
   
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "GridValidate", Err.Description)
    End If

End Function

Private Sub GridCellUpdate(ByVal lRow As Long, ByVal lCol As Long)
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
'      12/20/04     Yasmin Tan      Added User-Defined CT when updating values
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim nCount As Integer
    Dim dTemperature As Double
    Dim dVolume As Double
    Dim nCurrentRecord As Integer
    Dim bAreaOnly As Boolean
    
    bAreaOnly = False
    
    With frmMain.grdAreaInfo
    If .Cell(flexcpData, lRow, COL_ORDER_RS) = "" Then
        Call AddJobAreaRecord(lRow)
        nCurrentRecord = nRecordNumber - 1
    Else
        nCurrentRecord = CInt(.Cell(flexcpData, lRow, COL_ORDER_RS)) - 1
    End If
        
    'Go to the Job Area record which is being edited
    rsJobArea.MoveFirst
    For nCount = 1 To nCurrentRecord
        rsJobArea.MoveNext
    Next
    
    'Update the Job Area record
    Select Case lCol
        Case COL_AREA_NAME
            rsJobArea("area_name") = CStr(.Cell(flexcpData, lRow, lCol))
            bAreaOnly = True
        Case COL_TEMPERATURE
            dTemperature = 0#
            Call UnitConvert(modeSave, unitTEMPERATURE, CDbl(.Cell(flexcpData, lRow, lCol)), _
                dTemperature)
            rsJobArea("temperature") = Round(dTemperature, 2)
            If IsNull(rsJobArea("area_name")) Then
                rsJobArea("area_name") = GenerateAreaName(lRow, LoadResString(IDS_DEFAULT_AREA)) _
                    & ">"
            End If
        Case COL_EST_HLT
            rsJobArea("estimated_hlt") = CDbl(.Cell(flexcpData, lRow, lCol))
            If IsNull(rsJobArea("area_name")) Then
                rsJobArea("area_name") = GenerateAreaName(lRow, LoadResString(IDS_DEFAULT_AREA)) _
                    & ">"
            End If
        Case COL_PLAN_EXPOSURE
            rsJobArea("planned_exposure") = CDbl(.Cell(flexcpData, lRow, lCol))
            If IsNull(rsJobArea("area_name")) Then
                rsJobArea("area_name") = GenerateAreaName(lRow, LoadResString(IDS_DEFAULT_AREA)) _
                    & ">"
            End If
        Case COL_VOLUME
            dVolume = 0#
            Call UnitConvert(modeSave, unitVOLUME, CDbl(.Cell(flexcpData, lRow, lCol)), _
                dVolume)
            rsJobArea("volume") = Round(dVolume, 2)
            If IsNull(rsJobArea("area_name")) Then
                rsJobArea("area_name") = GenerateAreaName(lRow, LoadResString(IDS_DEFAULT_AREA)) _
                    & ">"
            End If
        '12/20/04 - YYT - Added User defined CT
        Case COL_USER_CT
            If .Cell(flexcpData, lRow, lCol) = "" Then
                .Cell(flexcpData, lRow, lCol) = 0#
            End If
            rsJobArea("user_defined_CT") = CDbl(.Cell(flexcpData, lRow, lCol))
            If IsNull(rsJobArea("area_name")) Then
                rsJobArea("area_name") = GenerateAreaName(lRow, LoadResString(IDS_DEFAULT_AREA)) _
                    & ">"
            End If
    End Select
    
    
    'Reset the calculated fields since a value has been changed. Don't do this however if
    'only the Area Name was changed.
    If Not bAreaOnly Then
        rsJobArea("fumigant") = -1
        rsJobArea("concentration_time") = -1
        rsJobArea("dose") = -1
        rsJobArea("initial_concentration") = -1
        rsJobArea("calculation_warning") = ""   'this was not included in the original app
    End If
    
    'Go forth and update
    rsJobArea.Update
    
    End With
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "GridCellUpdate", Err.Description)
    End If

End Sub
Private Function GenerateAreaName(ByVal lRow As Long, ByVal sAreaName As String) As String
'**************************************************************
'   DESCRIPTION
'       This subroutine generates a unique Area Name for the Job
'       Area record.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - a unique Area Name
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    sAreaName = sAreaName & CStr(lRow)
    
    If InStr(sGridAreaName, "|" & sAreaName & ">|") Then
        sAreaName = GenerateAreaName(lRow, sAreaName)
    End If
    
    GenerateAreaName = sAreaName
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "GenerateAreaName", Err.Description)
    End If

End Function



Private Sub AddJobAreaRecord(ByVal lRow As Long)
'**************************************************************
'   DESCRIPTION
'       This subroutine adds a new record in the Job Area table.
'
'   PARAMETERS
'           lRow - the row number of the cell that is being edited
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'   12/20/2004      Yasmin Tan          Added User-defined CT when adding a new record
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim sNewGUID As String
    Dim dDefaultTemperature As Double
    Dim dDefaultEstimatedHLT As Double
    Dim dDefaultPlannedExposure As Double
    Dim dDefaultAreaVolume As Double
    
    'Generate a new GUID for the new Job Area record
    sNewGUID = GetGUID
    
    'Set the default values for the new Job Area record
    dDefaultTemperature = DEFAULT_TEMPERATURE
    dDefaultEstimatedHLT = DEFAULT_HLT
    dDefaultPlannedExposure = DEFAULT_EXPOSURE
    dDefaultAreaVolume = DEFAULT_VOLUME
    
    'Create the new record
    rsJobArea.AddNew
    rsJobArea("id") = sNewGUID
    rsJobArea("job_fkey") = gsJobID
    rsJobArea("temperature") = dDefaultTemperature
    rsJobArea("estimated_hlt") = dDefaultEstimatedHLT
    rsJobArea("planned_exposure") = dDefaultPlannedExposure
    rsJobArea("volume") = dDefaultAreaVolume
    rsJobArea("user_defined_CT") = 0 '12/20/04-YYT-Added user-defined Ct. default to 0
    rsJobArea("row") = lRow
    rsJobArea.Update
    
    'Since we added a new record
    bIsCleared = False
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "AddJobAreaRecord", Err.Description)
    End If

End Sub
Private Sub ClearCalcResults()
'**************************************************************
'   DESCRIPTION
'       This subroutine clears the Results text box.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    'Update the Job record
    rsJob("fumigant") = -1
    rsJob("cylinders") = -1
    'Clear the warning. This was not included in the original application.
    rsJob("calculation_warning") = ""
    rsJob.Update
    
    frmMain.rtfResults.TextRTF = LoadResString(IDS_NO_CALC)
    
    
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "ClearCalcResults", Err.Description)
    End If
 
End Sub

Private Sub ClearAreaCalcResults()
'**************************************************************
'   DESCRIPTION
'       This subroutine resets the calculated fields values for the
'       Job Area records. This is called whenever any of the
'       following controls are updated:
'           cboTargetPests
'           cboFumigationType
'           cboLifeStage
'           cboCommodity
'           cboPressureType
'           txtLoadFactor
'
'   PARAMETERS
'           NONE
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Call ClearCalcResults
    
    'Update all the Job Areas!!!
    If bHasJobArea Then
        rsJobArea.MoveFirst
        While Not rsJobArea.EOF
            rsJobArea("fumigant") = 0
            rsJobArea("concentration_time") = 0
            rsJobArea("dose") = 0
            rsJobArea("initial_concentration") = 0
            rsJobArea("calculation_warning") = ""
            rsJobArea.Update
            rsJobArea.MoveNext
        Wend
    End If
    
    bIsCleared = True
    
    'Refresh the grid
    Call LoadAreaInfoRS
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "ClearAreaCalcResults", Err.Description)
    End If

End Sub

Private Sub DeleteJobArea(ByVal lRow As Long)
'**************************************************************
'   DESCRIPTION
'       This subroutine deletes the selected record in the Job Area
'       table. At the same time, this removes the row from the grid.
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
    
    'Determine if the current row has a Job Area record
    If frmMain.grdAreaInfo.Cell(flexcpData, lRow, COL_ORDER_RS) = "" Then
        'Current row has no Job Area record. Ergo, just remove the row.
        bDeleteRow = True
    Else
        nCurrentRecord = CInt(frmMain.grdAreaInfo.Cell(flexcpData, lRow, COL_ORDER_RS)) - 1
     
        'Go to the Job Area record which is being deleted
        rsJobArea.MoveFirst
        For nCount = 1 To nCurrentRecord
            rsJobArea.MoveNext
        Next
        
        'Do not delete the Job Area record if it has any associated
        'monitor point or introduction line or introduction history.
        If HasMonitoringPoint(rsJobArea("id")) Or HasIntroLine(rsJobArea("id")) Or _
        HasIntroHistory(rsJobArea("id")) Then
            bDeleteRow = False
        Else
           'Delete the record if the user agrees
           nDeleteOk = MsgBox(FormatSWParams(LoadResString(IDS_MSG_DELETE_GENERIC), _
               LoadResString(IDS_AREAINFO) & "|" & LoadResString(IDS_AREAINFO)), vbOKCancel)
           If nDeleteOk = vbOK Then
               rsJobArea("deleted") = True
               rsJobArea.Update
               bDeleteRow = True
               Call ClearCalcResults
           Else
               bDeleteRow = False
           End If
        End If
    End If
    
    If bDeleteRow Then
        'Update all Job Area records below the row
        rsJobArea.MoveLast
        While Not rsJobArea.BOF
            If rsJobArea("row") > lRow Then
                rsJobArea("row") = rsJobArea("row") - 1
                rsJobArea.Update
            End If
            rsJobArea.MovePrevious
        Wend
        'Refresh the grid
        Call LoadAreaInfoRS
    End If
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "DeleteJobArea", Err.Description)
    End If

End Sub

Private Function HasMonitoringPoint(sJobAreaId As String) As Boolean
'**************************************************************
'   DESCRIPTION
'       This function determines if the Job Area has monitoring points.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - True - if there are monitoring points
'             False - if there are no monitoring points
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim sMonitorPointSQL As String
    Dim rsMonitorPoint As New ADODB.Recordset
    
    sMonitorPointSQL = "SELECT * FROM MonitorPoint WHERE job_fkey = '" & gsJobID & "' AND " & _
        "jobarea_fkey = '" & sJobAreaId & "' AND deleted = FALSE"
    rsMonitorPoint.Open sMonitorPointSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    If rsMonitorPoint.BOF And rsMonitorPoint.EOF Then
        HasMonitoringPoint = False
    Else
        MsgBox LoadResString(IDS_MSG_NODELETE_AREA_MONITORPOINT), vbExclamation
        HasMonitoringPoint = True
    End If
    
    'Clean-up
    rsMonitorPoint.Close
    Set rsMonitorPoint = Nothing
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "HasMonitoringPoint", Err.Description)
    End If

End Function

Private Function HasIntroLine(sJobAreaId As String) As Boolean
'**************************************************************
'   DESCRIPTION
'       This function determines if the Job Area has associated
'       Introduction Line records.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - True - if there are Introduction Line records
'             False - if there are no Introduction Line records
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim sIntroLineSQL As String
    Dim rsIntroLine As New ADODB.Recordset
    
    sIntroLineSQL = "SELECT ShootingLine.* FROM ShootingLine " & _
        "where ShootingLine.jobarea_fkey = '" & sJobAreaId & "' and " & _
        "ShootingLine.deleted = FALSE and ShootingLine.job_fkey = '" & gsJobID & "'"
    rsIntroLine.Open sIntroLineSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    If rsIntroLine.BOF And rsIntroLine.EOF Then
        HasIntroLine = False
    Else
        MsgBox LoadResString(IDS_MSG_NODELETE_AREA_INTROLINE), vbExclamation
        HasIntroLine = True
    End If
    
    'Clean-up
    rsIntroLine.Close
    Set rsIntroLine = Nothing
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "HasIntroLine", Err.Description)
    End If

End Function

Private Function HasIntroHistory(sJobAreaId As String) As Boolean
'**************************************************************
'   DESCRIPTION
'       This function determines if the Job Area has associated
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
    Dim rsIntroHistory As New ADODB.Recordset
    
    sIntroHistorySQL = "SELECT IntroHistory.* FROM IntroHistory " & _
        "where IntroHistory.jobarea_fkey = '" & sJobAreaId & "' and " & _
        "IntroHistory.deleted = FALSE"
    rsIntroHistory.Open sIntroHistorySQL, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    If rsIntroHistory.BOF And rsIntroHistory.EOF Then
        HasIntroHistory = False
    Else
        MsgBox LoadResString(IDS_MSG_NODELETE_AREA_INTROHISTORY), vbExclamation
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

Public Sub Calculate(Optional bInForm As Boolean = True)
'**************************************************************
'   DESCRIPTION
'       This subroutine calculates, obviously...
'
'   PARAMETERS
'           bInForm - set to false if not called by the form itself
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'   1/3/2003       Jett Gamboa       added function to set End of Exposure date/time
'                                    if exposure times for all areas are the same
'  5/27/2003       Jett Gamboa       Used Planned Exposure * 60 to get End of Exposure
'  5/27/2003       Jett Gamboa       Added extra parameter so we can recalculate the Targets
'                                    when we force the exposure times to follow the actual
'                                    exposure duration.
'                                    Note: This is not the best way to do it, we are forced
'                                    to create a workaround so that we don't have the risk
'                                    of breaking the other pieces.
'  12/22/2004       Yasmin Tan       added logic to check if the user_defined CT = CT after
'                                    it was calculated from the Profume Calculator.
'  01/12/2004      Pat San Miguel    Added logic to check if the user has inputted
'                                    complete criteria so that it would compute in the
'                                    Profume Calculator the Adj HLT due to Sorption
'  12/27/2006      David Smith       Added code to change user entered CT to 36 for Rodents
'                                    This is per conversation with BP and only in effect
'                                    when rodents is the only pest
'  04/25/07        David Smith       Removed confirmation for users confirming use of space an commodity
'  02/06/09         Sebouh H         Remove logic for Commodity Calculation
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim sJobAreaSQL As String
    
    ' Only validate the form controls if we are actually in the form
    If bInForm Then
        'There should be a Job Area and at least one Pest
        If Not bHasJobArea Then
            MsgBox LoadResString(IDS_NO_CALC_AREA), vbExclamation
            Exit Sub
        Else
            If frmMain.lstTargetPests.SelCount = 0 Then
                MsgBox LoadResString(IDS_SELECT_PEST), vbExclamation
                frmMain.lstTargetPests.SetFocus
                Exit Sub
            End If
        End If
    
        'Save the current values of the drop-down controls. We need to do this
        'in case that a user does not pick any value which would cause an error
        'in the Profume Calculator. Ergo, just save the default value.
        Call ValidateUpdateControls("cboCommodity")
        Call ValidateUpdateControls("cboLifeStage")
        Call ValidateUpdateControls("cboFumigationType")
        Call ValidateUpdateControls("cboPressureType")
        
    Else
        
        ' We need to create the recordsets if we called calculate from the "outside"
        
        CreateJobPestRS
        
        ' Do not bother to recalculate if no pest was selected
        If rsJobPest.RecordCount = 0 Then
            rsJobPest.Close
            Set rsJobPest = Nothing
            Exit Sub
        End If
        
        sJobAreaSQL = "select JobArea.* from JobArea where deleted = false and " & _
            "job_fkey = '" & gsJobID & "' order by row"
        
        'Open the recordset. This recordset is always open for updates.
        rsJobArea.Open sJobAreaSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
        
        CreateJobRS
    
    End If
    
    'Create a new instance of the Profume Calculator
    Dim ProfumeCalc As New ProFumeCalculator
    
    'Declare the variables
    Dim dTotalProfume As Double
    Dim dTotalCylinders As Double
    Dim DTotalInitialConcentration As Double
    Dim dTotalConcentrationTime As Double
    Dim dTotalDose As Double
    Dim dTotalHlt As Double
    Dim dTotalVolume As Double
    Dim nTimeToAdd As Integer
    
    Dim dHLT As Double
    Dim dTemperature As Double
    Dim dExposure As Double
    Dim dVolume As Double
    Dim nFumigationType As Integer
    Dim nLifeStage As Integer
    Dim nAirPressure As Integer
    Dim nCommodity As Integer
    Dim dLoadFactor As Double 'Added 01/12/06-PAT
    
    Dim dProfume As Double
    Dim dCylinders As Double
    Dim dInitialConcentration As Double
    Dim dConcentrationTime As Double
    Dim dDose As Double
    Dim dUserDefinedCT As Double 'Added 12/22/04-YYT
    Dim dTempUserDefinedCT As Double
    Dim sWarning As String
    
    Dim sError As String
    Dim sTempError As String
    Dim bErrorOccured As Boolean
    
    Dim dtStartExposure As Date
    Dim dtEndExposure As Date
    
    Dim IsComplete As Integer 'Added 01/12/06-PAT
    Dim sMsg 'Added 01/26/06-PAT
    Dim intCounter 'Added 01/27/06-PAT
    Dim nNoAdjSorp 'Added 02/13/06-PAT
    
    'Initialize the variables
    dTotalProfume = 0#
    dTotalCylinders = 0#
    DTotalInitialConcentration = 0#
    dTotalConcentrationTime = 0#
    dTotalDose = 0#
    dTotalHlt = 0#
    dTotalVolume = 0#
    
    dProfume = 0#
    dCylinders = 0#
    dInitialConcentration = 0#
    dConcentrationTime = 0#
    dDose = 0#
    dUserDefinedCT = 0# '12/22/04-YYT-initialize user-defined target CT
    
    dHLT = 0#
    dTemperature = 0#
    dExposure = 0#
    dVolume = 0#
    nFumigationType = 0
    nLifeStage = 0
    nAirPressure = 0
    nCommodity = 0
    dLoadFactor = 0# '01/12/06-PAT-initialize Load Factor
    nNoAdjSorp = 0# '02/13/06-PAT-initialize Value
    
    sError = ""
    bErrorOccured = False
    
    rsJobArea.MoveFirst
    nFumigationType = rsJob("fumigation_type_enum")
    nLifeStage = rsJob("lifestage_enum")
    nAirPressure = rsJob("treatment_type_enum")
    nCommodity = rsJob("commodity_enum")
    dLoadFactor = rsJob("load_factor") '01/12/06-PAT-Get Load Factor value
    nNoAdjSorp = rsJob("NoAdjSorp") '02/13/06-PAT-this will indicate if we need to compute for the AdjHLT Sorption
    
    ' 1/10/03 Jett - change the End of Exposure to Start of Exposure + Exposure Time
    ' (specified in Dosage Plan) only if the exposure times for all areas are the same
    
    ' get the exposure times
    gnExposureTime = CheckExposureTimes()
    
    ' only do this if we have a common exposure time for all areas
    If gnExposureTime > 0 Then
    
        ' Jett 5/27/2003 We multiply the planned exposure time by 60 and add in minutes and
        ' not in hours. this makes it more precise
        nTimeToAdd = gnExposureTime * 60
    
        ' only update the End of Exposure if both Start Exposure and End Exposure already
        ' have values.
        If (Not IsNull(rsJob("startexp_datetime"))) And _
           (Not IsNull(rsJob("endexp_datetime"))) Then
            
            dtStartExposure = rsJob("startexp_datetime")
            
            ' add the Exposure time to Start of Exposure to get our new End Exposure date/time
            dtEndExposure = DateAdd("n", nTimeToAdd, dtStartExposure)
            
            rsJob("endexp_datetime") = dtEndExposure
            rsJob.Update
            
        End If
    
    End If
    intCounter = 1 'Initialize to 1
    
    'PAT 1/27/06 - Prompt User if has chosen SPACE fumigation
    '04/23/07 - David Smith  per bp request im removing popup
    
   ' If nFumigationType = 0 And nCommodity <> 0 And dLoadFactor <> 0 Then
   '     sMsg = MsgBox(LoadResString(IDS_FUMIGATION_SPACE_CHOSEN), vbYesNo)
   ' End If
   ' Setting sMsg to default to yes
   sMsg = vbYes
   
    
    
    While Not rsJobArea.EOF
        sWarning = ""
        dHLT = rsJobArea("estimated_hlt")
        dTemperature = rsJobArea("temperature")
        dExposure = rsJobArea("planned_exposure")
        dVolume = rsJobArea("volume")   'volume should never be 0, otherwise, there would be an error in the calculator
        dTempUserDefinedCT = IIf(IsNull(rsJobArea("user_defined_CT")), 0, rsJobArea("user_defined_CT"))
        dUserDefinedCT = dTempUserDefinedCT
        
'        If dUserDefinedCT <= CDbl(rsJobArea("concentration_time")) And bEnteredUserCT = True Then
'            MsgBox LoadResString(IDS_CALC_ERROR_SEE_WARNINGS), vbExclamation
'            Exit Sub
'        End If
            
        'SetPestsForCalculation
        'Be aware that our Job Pest recordset contains deleted pests as well
        'Dim sError As String
        rsJobPest.MoveFirst
        ProfumeCalc.ClearDosePest
        sTempError = ""


        '12/27/2006 David Smith Added User Defined CT for Rodent only selection
        If rsJobPest.RecordCount = 1 Then
            If rsJobPest("pest_enum") = 17 And dUserDefinedCT = 0 Then
                dUserDefinedCT = 36
                frmMain.grdAreaInfo.Cell(flexcpData, 1, 8) = dUserDefinedCT
                rsJobArea("user_defined_CT") = dUserDefinedCT
            End If
         End If
        
        While Not rsJobPest.EOF
            If Not rsJobPest("deleted") Then
            
                'SH 2/6/09 Calculation should continue, even if Fumigation Type is not selected as Commodity
                ' and even if Load Factor = 0
                ' ************* New Code *********************
                ' triple comment done by Sebouh '''
                
                IsComplete = 0
            
                ProfumeCalc.SetDosePest rsJobPest("pest_enum"), nFumigationType, _
                nLifeStage, nAirPressure, nCommodity, dHLT, dTemperature, _
                dExposure, dVolume, dLoadFactor, IsComplete, intCounter, sError
                If sError <> "" Then
                    sTempError = sError
                End If
                            
                ' *************
            
            
'''                'PAT 1/12/06 Profume Enhancement 2005-2006
'''                '-This will check if a Fumigation Type,Commodity and Load Factor
'''                ' has been chosen. In preparation to the computation of
'''                ' Adjustments to Expected Half Lost Time due to sorptions
'''
'''                If IsNull(nFumigationType) = False And nCommodity <> 0 And dLoadFactor <> 0 Then
'''                   'PAT 2/13/2006 - We need to check first if the user wanted to compute for the
'''                   '                the AdjHLT Sorption
'''                   If nNoAdjSorp = 0 Or IsNull(nNoAdjSorp) Then
'''                        If nFumigationType = 1 Then 'Fumigation Type = Commodity
'''                            IsComplete = 1 'PAT 1/12/06 - if true compute for the Adjustments to Expected Half Lost Time due to sorptions
'''                            'Unfortunately, we cannot get the return code (S_OK or whatever)
'''                            'from the Profume Calculator. Therefore, in order to check if
'''                            'we have an error, we check the value of sError.
'''                            ProfumeCalc.SetDosePest rsJobPest("pest_enum"), nFumigationType, _
'''                                nLifeStage, nAirPressure, nCommodity, dHLT, dTemperature, _
'''                                dExposure, dVolume, dLoadFactor, IsComplete, intCounter, sError
'''                            If sError <> "" Then
'''                                sTempError = sError
'''                            End If
'''                        ElseIf nFumigationType = 0 Then 'PAT 1/12/06 - Fumigation Type = Space
'''                            If sMsg = vbYes Then
'''                                IsComplete = 1
'''                                'Unfortunately, we cannot get the return code (S_OK or whatever)
'''                                'from the Profume Calculator. Therefore, in order to check if
'''                                'we have an error, we check the value of sError.
'''                                ProfumeCalc.SetDosePest rsJobPest("pest_enum"), nFumigationType, _
'''                                    nLifeStage, nAirPressure, nCommodity, dHLT, dTemperature, _
'''                                    dExposure, dVolume, dLoadFactor, IsComplete, intCounter, sError
'''                                If sError <> "" Then
'''                                    sTempError = sError
'''                                End If
'''                            Else
'''                                'MsgBox "Computation will not continue"
'''                                Exit Sub
'''                            End If
'''                        End If
'''                    Else
'''                        'PAT 2/13/06 - Eventhough the criteria are complete but the user opted to compute for the Adj HLT
'''                        'we are going to assume that the criteria is incomplete
'''                        IsComplete = 0
'''                        'Unfortunately, we cannot get the return code (S_OK or whatever)
'''                        'from the Profume Calculator. Therefore, in order to check if
'''                        'we have an error, we check the value of sError.
'''                        ProfumeCalc.SetDosePest rsJobPest("pest_enum"), nFumigationType, _
'''                            nLifeStage, nAirPressure, nCommodity, dHLT, dTemperature, _
'''                            dExposure, dVolume, dLoadFactor, IsComplete, intCounter, sError
'''                        If sError <> "" Then
'''                            sTempError = sError
'''                        End If
'''                    End If
'''                Else 'PAT 1/12/06 - Criteria is incomplete
'''                    IsComplete = 0 'PAT 1/12/06 - This will proceed to the normal flow of computation
'''                    If nCommodity = 0 And (dLoadFactor = 0 Or IsNull(dLoadFactor) = True) Then
'''                        'Unfortunately, we cannot get the return code (S_OK or whatever)
'''                        'from the Profume Calculator. Therefore, in order to check if
'''                        'we have an error, we check the value of sError.
'''                        ProfumeCalc.SetDosePest rsJobPest("pest_enum"), nFumigationType, _
'''                            nLifeStage, nAirPressure, nCommodity, dHLT, dTemperature, _
'''                            dExposure, dVolume, dLoadFactor, IsComplete, intCounter, sError
'''                        If sError <> "" Then
'''                            sTempError = sError
'''                        End If
'''                    ElseIf nCommodity = 0 Or IsNull(dLoadFactor) = True Or dLoadFactor = 0 Then
'''                        'MsgBox "Computation will not continue"
'''                        Exit Sub
'''                    End If
'''                End If 'PAT 1/12/06 - checking if the criteria is sufficient in computing for the HLT Sorption

            End If 'rsJobPest("deleted")

            rsJobPest.MoveNext
        Wend
            
        If sTempError <> "" Then
            bErrorOccured = True
            rsJobArea("fumigant") = 0#
            rsJobArea("concentration_time") = 0#
            rsJobArea("initial_concentration") = 0#
            rsJobArea("dose") = 0#
            rsJobArea("calculation_warning") = sError
            rsJobArea.Update
        Else    'meaning no error in SetDosePest
            '12/23/04-YYT-pass the userdefinedCT to the dConcentrationTime
            'this user_Defined_CT in the DB is saved as 0 when empty.
            'calculator will handle that when user-defined is 0, then compute target CT
            'else take this as target CT
            dConcentrationTime = 0
            dConcentrationTime = dUserDefinedCT
            
            'PerformCalculation
            ProfumeCalc.GetProfumeRequired dProfume, dCylinders, dInitialConcentration, _
                dConcentrationTime, dDose, sWarning
            'we cannot trap if the calculator errored out here because sWarning
            'can have a value even if there is no error. but anyway, after taking
            'a look at the Calculator code (during the design phase), we found out
            'that the Calculator would not return an error code in this part
            rsJobArea("fumigant") = dProfume
            rsJobArea("concentration_time") = dConcentrationTime
            rsJobArea("initial_concentration") = dInitialConcentration
            rsJobArea("dose") = dDose
            rsJobArea("calculation_warning") = sWarning
                   
            '12/23/04-YYT-Added this logic to trap the scenario wherein the ct is set to
            'default value of 200 or 1500, depending on treatment type.
            'this happens when treatment type cbo box is changed and user-defined CT is
            'not updated.
            
            If dConcentrationTime <> dUserDefinedCT And dUserDefinedCT > 0 Then
                rsJobArea("user_defined_CT") = dConcentrationTime
            End If
            
            rsJobArea.Update
                
            'Since we have updated the calculation, we need to update all
            'associated monitor points
            Call UpdateMonitorPoints(rsJobArea("id"))
                
            'Add up the values for the Job record
            dTotalProfume = dTotalProfume + dProfume
            dTotalCylinders = dTotalCylinders + dCylinders
            DTotalInitialConcentration = DTotalInitialConcentration + dInitialConcentration
            dTotalConcentrationTime = dTotalConcentrationTime + dConcentrationTime
            dTotalDose = dTotalDose + dDose
            dTotalHlt = dTotalHlt + dHLT
            dTotalVolume = dTotalVolume + dVolume
        End If
                              
        rsJobArea.MoveNext
        intCounter = intCounter + 1
    Wend
    
    Dim sResults As String
    sResults = ""
    
    'Format the results
    If bErrorOccured Then
        sResults = LoadResString(IDS_CALC_ERROR)
        MsgBox LoadResString(IDS_CALC_ERROR_SEE_WARNINGS), vbExclamation
        'bCalcRequired = True
    End If
    
    rsJob("calculation_error") = sResults
    
    rsJob("fumigant") = dTotalProfume
    rsJob("cylinders") = dTotalCylinders
    
    Dim lTotalAreas As Long
    lTotalAreas = rsJobArea.RecordCount
    rsJob("avg_concentration_time") = CDbl(dTotalConcentrationTime / lTotalAreas)
    rsJob("avg_initial_concentration") = CDbl(DTotalInitialConcentration / lTotalAreas)
    rsJob("avg_dose") = CDbl(dTotalDose / lTotalAreas)
    rsJob("avg_estimated_hlt") = CDbl(dTotalHlt / lTotalAreas)
    rsJob("volume") = dTotalVolume
    rsJob("calc_version") = ProfumeCalc.GetVersion
    rsJob.Update
          
    
    ' 5/27/03 Jett If we are in the form, update the grid and populate the results box,
    ' otherwise just cleanup and leave
    If bInForm Then
    
        'Refresh the grid
        Call LoadAreaInfoRS
        
        'Refresh the Results text box
        Call InitResultsTextBox
        
    Else
    
        ' 5/27/2003 Jett - If this was called from the "outside" clean up our objects
        rsJobArea.Close
        rsJob.Close
        rsJobPest.Close
        
        Set rsJobArea = Nothing
        Set rsJob = Nothing
        Set rsJobPest = Nothing
        
    End If
    
    'Clean up
    Set ProfumeCalc = Nothing
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "Calculate", Err.Description)
    End If

End Sub

'**************************************************************
'   DESCRIPTION
'       This subroutine updates the monitor point records associated
'       with the Job Area. This is called during Calculation.
'
'   PARAMETERS
'           NONE
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
 Private Sub UpdateMonitorPoints(sJobAreaId As String)
   
    On Error GoTo Error_Handler
    
    Dim sMonitorPointsSQL As String
    Dim rsMonitorPoints As New ADODB.Recordset
    
    'JNM 11/14 - removed Inactive column
    sMonitorPointsSQL = "SELECT MonitorPoint.* FROM MonitorPoint WHERE " & _
        "MonitorPoint.deleted = FALSE AND MonitorPoint.jobarea_fkey = '" & _
        sJobAreaId & "' AND MonitorPoint.job_fkey = '" & _
        gsJobID & "'"    ' AND MonitorPoint.inactive = FALSE"
    
    rsMonitorPoints.Open sMonitorPointsSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    If rsMonitorPoints.BOF And rsMonitorPoints.EOF Then
        'Leave!!!
        Exit Sub
    End If
    
    'Iterate through all the monitor points and recalc
    
    Dim MonitorExpert As New clsMonitorExpert
    
    MonitorExpert.SetJobArea gsJobID, sJobAreaId
    
    'Get the start introduction and start exposure date/time
    Dim dtStartIntro As Date
    Dim dtStartExposure As Date
    Dim dtEndExposure As Date
    Dim dtDefault As Date
    
    dtDefault = 0#
    
    dtStartIntro = IIf(IsNull(rsJob("startintro_datetime")), dtDefault, _
        rsJob("startintro_datetime"))
    dtStartExposure = IIf(IsNull(rsJob("startexp_datetime")), dtDefault, _
        rsJob("startexp_datetime"))
    dtEndExposure = IIf(IsNull(rsJob("endexp_datetime")), dtDefault, _
        rsJob("endexp_datetime"))
    
    rsMonitorPoints.MoveFirst
    While Not rsMonitorPoints.EOF
        MonitorExpert.UpdateMonitorPointStatus rsMonitorPoints("id"), dtStartIntro, _
            dtStartExposure, dtEndExposure
        rsMonitorPoints.MoveNext
    Wend
    
    'Update the Job Area status
    '05/15/2003 Jett - added dates in call to UpdateJobAreaStatus
    MonitorExpert.UpdateJobAreaStatus gsJobID, sJobAreaId, dtStartIntro, dtStartExposure, dtEndExposure
    
    'Clean-up
    rsMonitorPoints.Close
    Set rsMonitorPoints = Nothing
    Set MonitorExpert = Nothing
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(MODULE_NAME, "UpdateMonitorPoints", Err.Description)
    End If

End Sub


Public Function IsInteger(sValue As String) As Boolean
'**************************************************************
'   DESCRIPTION
'       This function determines if sValue is a whole number
'       or not.
'
'   PARAMETERS
'           sValue = the value to be tested
'
'   RETURNS - TRUE -if value is a whole number
'             FALSE - if value is not a whole number
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'      10/25/02   JN Manyaga         Instead of testing if sValue is equal
'                                    to lValue, the difference of both the
'                                    values are computed. An integer should
'                                    have a difference of 0.
'**************************************************************
    
    On Error GoTo Error_Handler
    
    Dim lValue As Long
    Dim dValue As Double
    Dim dDiff As Double
    
    lValue = CLng(sValue)
    dValue = CDbl(sValue)
    dDiff = lValue - dValue
    
    If dDiff <> 0 Then
        IsInteger = False
    Else
        IsInteger = True
    End If
    
Error_Handler:
    If Err.Number <> 0 Then
        IsInteger = False
    End If
End Function




Public Sub Load_DosagePlan()

    Dim bPrevHasChanged As Boolean

    'We need to store the original value of gbHasChanged so that we would
    'be able to protect this from the changes due to the form initialization.
    bPrevHasChanged = gbHasChanged
        
    gbCanChangeTabs = True
    
    Call InitializeControls
    
    'Reset the value of gbHasChanged
    gbHasChanged = bPrevHasChanged
    
    'set the focus on the first control ; JNM 11/04 - removed the lines below so that
    'multi-line tooltip class can function
    'Me.Show
    'txtSiteName.SetFocus

End Sub


Public Sub Unload_DosagePlan()

Dim Cancel As Boolean

'JNM 05/02/03 - Since we are already in one form, the Validate event would
'   automatically be called. Ergo, no need to do these steps.
'    Dim element As Variant
'    Dim arrControl(10) As String
'
'    Dim Cancel As Boolean
'
'    arrControl(0) = "txtSiteName"
'    arrControl(1) = "txtJobName"
'    arrControl(2) = "txtFumigationDate"
'    arrControl(3) = "txtLicFumigator"
'    arrControl(4) = "lstTargetPests"
'    arrControl(5) = "cboFumigationType"
'    arrControl(6) = "cboLifeStage"
'    arrControl(7) = "cboPressureType"
'    arrControl(8) = "cboCommodity"
'    arrControl(9) = "txtLoadFactor"
'
'    'Validate all the controls before unloading the form
'    If frmMain.grdAreaInfo.EditWindow <> 0 Then
'        Cancel = True
'    Else
'        For Each element In arrControl()
'            If ValidateUpdateControls(element) Then
'                Cancel = False
'            Else
'                Cancel = True
'                Exit For
'            End If
'        Next
'    End If
    
    
    If Cancel = False Then
        Cancel = Not ValidateUpdateControls("Calculation")
    End If
    
    If Cancel Then
        gbCanChangeTabs = False
    Else
        ' call the DosagePlan cleanup routine (used to be in Form Unload)
        ModuleCleanUp
    End If

End Sub
'PAT 2/13/06 - This would check if there's a need to compute for the Adj HLT Sorption
Public Sub EnableAdjHLTSorptionChkBox()
    If gbChkCommodity = True And gbChkLoadFactor = True And frmMain.txtLoadFactor <> "0" Then
        'frmMain.chkNoAdjSorp.Enabled = True
        'Always false
        frmMain.chkNoAdjSorp.Enabled = False
        gbChkEnabled = True
    Else
        frmMain.chkNoAdjSorp.Enabled = False
        frmMain.chkNoAdjSorp.Value = 0
        gbChkEnabled = False
    End If
End Sub


'========================================================================================================
' -------------------- Handlers moved from the form ---------------
'========================================================================================================

Public Sub cboCommodity_Click_Handler()
    If Not bIsLoading And Not bIsCleared Then
        Call ClearAreaCalcResults
    End If
    gbHasChanged = True
    'PAT 2/10/06 - This would check if a commodity has been chosen.
    If frmMain.cboCommodity.ListIndex = 0 Then
        gbChkCommodity = False
    Else
        gbChkCommodity = True
    End If
End Sub



Public Sub cboCommodity_Validate_Handler(Cancel As Boolean)
    If Not bRSClosed Then
        'Call the function that updates the database
        ValidateUpdateControls "cboCommodity"
    End If
End Sub

Public Sub cboFumigationType_Click_Handler()
    If Not bIsLoading And Not bIsCleared Then
        Call ClearAreaCalcResults
    End If
    gbHasChanged = True
End Sub


Public Sub cboFumigationType_Validate_Handler(Cancel As Boolean)
    If Not bRSClosed Then
        'Call the function that updates the database
        ValidateUpdateControls "cboFumigationType"
    End If
End Sub

Public Sub cboLifeStage_Click_Handler()
    If Not bIsLoading And Not bIsCleared Then
        Call ClearAreaCalcResults
    End If
    
    gbHasChanged = True
End Sub

Public Sub cboLifeStage_Validate_Handler(Cancel As Boolean)
    If Not bRSClosed Then
        'Call the function that updates the database
        ValidateUpdateControls "cboLifeStage"
    End If

End Sub
Public Sub chkNoAdjSorp_Validate_Handler(Cancel As Boolean)
    If Not bRSClosed Then
        'Call the function that updates the database
        ValidateUpdateControls "chkNoAdjSorp"
    End If

End Sub

Public Sub cboPressureType_Click_Handler()
    If Not bIsLoading And Not bIsCleared Then
        Call ClearAreaCalcResults
    End If
    gbHasChanged = True
End Sub

Public Sub cboPressureType_Validate_Handler(Cancel As Boolean)
    If Not bRSClosed Then
        'Call the function that updates the database
        ValidateUpdateControls "cboPressureType"
    End If

End Sub

Public Sub cmdCalculate_Click_Handler()
    Call Calculate
    gbHasChanged = True
End Sub


Public Sub grdAreaInfo_AfterEdit_Handler(ByVal Row As Long, ByVal Col As Long)
    'Refresh the grid!!!
    Call LoadAreaInfoRS
End Sub


Public Sub grdAreaInfo_BeforeUserResize_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Do not allow the user to change the width of the Space column
    If Col = 5 Then
        Cancel = True
    End If
End Sub

Public Sub grdAreaInfo_KeyDown_Handler(KeyCode As Integer, Shift As Integer)
    'If user pressed the Delete button, delete the Job Area record
    If KeyCode = 46 Then
        Call DeleteJobArea(frmMain.grdAreaInfo.Row)
    End If
End Sub

Public Sub grdAreaInfo_MouseDown_Handler(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton And frmMain.grdAreaInfo.MouseRow > 0 Then
        'Select the cell
        'MsgBox mnuDelete.Visible
        
        frmMain.grdAreaInfo.Select frmMain.grdAreaInfo.MouseRow, frmMain.grdAreaInfo.MouseCol
        
        frmMain.PopupMenu frmMain.mnuDosagePlanContext
        
    End If

End Sub

Public Sub grdAreaInfo_MouseMove_Handler(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With frmMain.grdAreaInfo
        Select Case .MouseCol
            Case COL_AREA_NAME
                .ToolTipText = LoadResString(IDS_TIP_AREA_NAME)
            Case COL_TEMPERATURE
                .ToolTipText = LoadResString(IDS_TIP_TEMPERATURE)
            Case COL_EST_HLT
                .ToolTipText = LoadResString(IDS_TIP_EST_HLT)
            Case COL_PLAN_EXPOSURE
                .ToolTipText = LoadResString(IDS_TIP_EXPOSURE)
            Case COL_VOLUME
                .ToolTipText = LoadResString(IDS_TIP_VOLUME)
            Case COL_FUMIGANT
                .ToolTipText = LoadResString(IDS_TIP_FUM_REQUIRED)
            Case COL_CT
                .ToolTipText = LoadResString(IDS_TIP_TARGET_CT)
            Case COL_USER_CT '12/20/04-YYT-Added user-defined CT tooltip
                .ToolTipText = LoadResString(IDS_TIP_USER_TARGET_CT)
            Case COL_CO
                .ToolTipText = LoadResString(IDS_TIP_CO)
        End Select
    End With

End Sub

Public Sub grdAreaInfo_StartEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With frmMain
        'We need this code in order to make columns 5 and above Read-Only,
        '12/21/04 - YYT Make Column 8 (user-defined CT) as editable
        If (Row < 1) Or (Col > 4 And Col <> 8) Then
            Cancel = True
        Else
            Cancel = False
            'Remove the units once the user starts to edit the cell.
            .grdAreaInfo.Cell(flexcpText, Row, Col) = .grdAreaInfo.Cell(flexcpData, Row, Col)
            If Col = 0 Then
                .grdAreaInfo.EditMaxLength = 50
            Else
                .grdAreaInfo.EditMaxLength = 0
            End If
        End If
    End With
    
    'The line below would ensure that the tab does not change if the user enter enters
    'an invalid value and then clicks on another tab
    ReturnCurrTab TAB_DOSAGE_PLAN

End Sub

Public Sub grdAreaInfo_ValidateEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With frmMain
        If GridValidate(Row, Col) Then
            'Store the new value in flexcpData and flexcpText
            .grdAreaInfo.Cell(flexcpData, Row, Col) = Trim(.grdAreaInfo.EditText)
        
            'And let there be units!!! The units are hidden in the flexcpData attribute
            'of the grid heading.
            .grdAreaInfo.Cell(flexcpText, Row, Col) = Trim(.grdAreaInfo.EditText) & _
                .grdAreaInfo.Cell(flexcpData, 0, Col)
        
            GridCellUpdate Row, Col
                
            Call ClearCalcResults
        
            Cancel = False
        Else
            Cancel = True
        End If
        
        gbHasChanged = True
        
        'JNM 05/02/03 - PT fix - remove this since we are already in one form
        'JNM 10/23 - Focus should return to the grid if value is invalid
        'If Cancel Then
        '    frmMain.grdAreaInfo.SetFocus
        'End If
    End With

End Sub

Public Sub lstTargetPests_GotFocus_Handler()
    Call ClearAreaCalcResults
    gbHasChanged = True
End Sub

Public Sub lstTargetPests_Validate_Handler(Cancel As Boolean)
    If Not bRSClosed Then
        'Call the function that updates the database
        ValidateUpdateControls "lstTargetPests"
    End If
End Sub


Public Sub mnuDosagePlanDelete_Click_Handler()

    Call DeleteJobArea(frmMain.grdAreaInfo.Row)

End Sub


Public Sub rtfResults_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'rtfResults.ToolTipText = LoadResString(IDS_TIP_AVG_CO) & vbCrLf & _
    LoadResString(IDS_TIP_AVG_HLT) & vbCrLf & LoadResString(IDS_TIP_AVG_CT) & _
    LoadResString(IDS_TIP_TOT_FUM) & vbCrLf & LoadResString(IDS_TIP_WARNING)
End Sub

Public Sub txtFumigationDate_Change_Handler()
    gbHasChanged = True
End Sub




Public Sub txtJobName_Change_Handler()
    gbHasChanged = True
End Sub



Public Sub txtLicFumigator_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtLoadFactor_Change_Handler()
    If Not bIsLoading And Not bIsCleared Then
        Call ClearAreaCalcResults
    End If
    gbHasChanged = True
    'PAT 2/10/06 - This would check if there's a Load Factor.
    If IsNull(frmMain.txtLoadFactor.Text) Or frmMain.txtLoadFactor.Text = "0" Or frmMain.txtLoadFactor.Text = "" Then
        gbChkLoadFactor = False
    Else
        gbChkLoadFactor = True
    End If
End Sub

'Public Sub txtLoadFactor_LostFocus_Handler()
'    ' do not bother checking if this was already validated in the Validate event
'
'    DoEvents
'
'    'JNM 04/24/03 - AT Fix - Check first if the recordsets are still open
'    If Not bRSClosed Then
'        If Not bInvalid Then
'            If Not ValidateUpdateControls("txtLoadFactor") Then
'                frmMain.txtLoadFactor.SetFocus
'            End If
'        Else
'            frmMain.txtLoadFactor.SetFocus
'        End If
'    End If
'
'End Sub

Public Sub txtLoadFactor_Validate_Handler(Cancel As Boolean)

    bInvalid = False
    
    If Not bRSClosed Then
        If Not ValidateUpdateControls("txtLoadFactor") Then
            gbCanChangeTabs = False
            'frmMain.txtLoadFactor.SetFocus   JNM - no need
            bInvalid = True
            gbCanExit = False
            'JNM 04/24/03 - AT Fix
            Cancel = True
        End If
        'JNM 05/02/03 - PT fix - Ignore the error if we are saving
        If gbIsSaving Then
            Cancel = False
        End If
    End If

End Sub

Public Sub txtSiteName_Change_Handler()
    gbHasChanged = True
End Sub



Public Sub txtSiteName_Validate_Handler(Cancel As Boolean)
'JNM 04/24/03 - AT Fix
    If Not bRSClosed Then
        ValidateUpdateControls "txtSiteName"
    End If
End Sub

Public Sub txtJobName_Validate_Handler(Cancel As Boolean)
'JNM 04/24/03 - AT Fix
    If Not bRSClosed Then
        ValidateUpdateControls "txtJobName"
    End If
End Sub

Public Sub txtLicFumigator_Validate_Handler(Cancel As Boolean)
'JNM 04/24/03 - AT Fix
    If Not bRSClosed Then
        ValidateUpdateControls "txtLicFumigator"
    End If
End Sub

Public Sub txtFumigationDate_Validate_Handler(Cancel As Boolean)
'JNM 04/24/03 - AT Fix
    If Not bRSClosed Then
        ValidateUpdateControls "txtFumigationDate"
    End If
End Sub

