Attribute VB_Name = "modMonitorStatus"
'**************************************************************
'   FORM:  frmMonitorStatus  (MonitorStatus)
'
'   DESCRIPTION
'       This form provides the status monitoring functionality for Profume.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/20/2002       Malou Espaldon    Initial Revision
' 10/18/2002       Joseph Manyaga    Added a code in the Unload event that
'                                    would close and destroy open recordset
'                                    objects.
'                                    Modified the subroutines that
'                                    display the values on the grid.
' 04/09/2003       Joseph Manyaga    Added double vertical lines in the grid
'                                    Added a code that would highlight the
'                                    rows that have Incompatible status.
' 06/03/2003       Malou Espaldon    changed dtStartIntro variable to dtStartIntroduction for
'                                    consistency and to fix a computation problem becauuse
'                                    this variable is known in other modules as dtStartIntroduction and
'                                    in other, dtStartIntro
'
'**************************************************************

Option Explicit
'--Private Constants
Private Const frmName = "frmMonitorStatus"
Private Const fCheckStrtExpDtTm = "CheckStartExposureDateTime"
Private Const fCheckStrtIntDtTm = "CheckStartIntroDateTime"
Private Const fCheckEndExpDtTm = "CheckEndExposureDateTime"
Private Const fRecalcValues = "RecalcValues"
Private Const fMain = "Form_Load"
Private Const fGetMonitorDataInputDates = "GetMonitorDataInputDates"
Private Const fUnload = "Form_Unload"
Private Const fPopulateStatus = "PopulateStatus"
Private Const fGetMonitorValues = "GetMonitorValues"
Private Const fPopulateMonitorPoint = "PopulateMonitorPoint"
Private Const fPopulateAllPoint = "PopulateAllPoint"
Private Const fdtpEndExposureDate2_Change = "dtpEndExposureDate2_Change"
Private Const fdtpEndExposureTime2_Change = "dtpEndExposureTime2_Change"
Private Const fdtpStartExposureDate2_Change = "dtpStartExposureDate2_Change"
Private Const fdtpStartExposureTime2_Change = "dtpStartExposureTime2_Change"
Private Const fdtpStartIntroDate2_Change = "dtpStartIntroDate2_Change"
Private Const fdtpStartIntroTime2_Change = "dtpStartIntroTime2_Change"

Dim MStatGrid As New clsGridWrapper
Dim rsMonitorPoint As New ADODB.Recordset
Dim rsMonitorValues As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset

'MME 6/3/03 Commenting out next line, replacing variable with dtStartIntroduction
'Dim dtStartIntro As Date
Dim dtStartIntroduction As Date

Dim dtStartExposure As Date
Dim dtEndExposure As Date

'JNM 05/12/03
Dim bDateTimeStatusClick As Boolean

Const AREANAME_COLSIZE = 2265
Const MONPOINT_COLSIZE = 1320
Const LASTENTRY_COLSIZE = 960
Const ELAPSED_COLSIZE = 1080
Const HLT_COLSIZE = 750
Const STARTHLT_COLSIZE = 840
Const ENDHLT_COLSIZE = 960
Const CTACH_COLSIZDE = 1110
Const PROJ_COLSIZE = 960
Const STATUS_COLSIZE = 3615

'JNM 04/04/03 - FE 2003 - Defined the constants below in order to
'   make column adjustments easier. If a new column is to be inserted
'   there would be no need to go over the code in order to adjust the
'   column numbers.
Private Const COL_AREANAME As Integer = 0
Private Const COL_MONPOINT As Integer = 1
Private Const COL_LASTENTRY As Integer = 2
Private Const COL_ELAPSED As Integer = 3
Private Const COL_SPACE1 As Integer = 4
Private Const COL_HLT As Integer = 5
Private Const COL_STARTHLT As Integer = 6
Private Const COL_ENDHLT As Integer = 7
Private Const COL_SPACE2 As Integer = 8
Private Const COL_CTACH As Integer = 9
Private Const COL_PROJ As Integer = 10
Private Const COL_STATUS As Integer = 11
Private Const COL_TESTCTACH As Integer = 12


' SUBROUTINE : dtpEndExposureDate2_Change()
'
' DESCRIPTION: Executes when user changes the End Exposure DATE.
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
'*******************************************************
Public Sub dtpEndExposureDate2_Change_Handler()
    
    On Error GoTo Error_Handler
    
    'JNM 05/12/03 - PT fix - Set the booleans...
    gbHasChanged = True
    bDateTimeStatusClick = False
    
    Exit Sub

Error_Handler:
    
    If Err.Number <> 0 Then
         Call psAbort(frmName, fdtpEndExposureDate2_Change, Err.Description)
    End If
    
End Sub





'*******************************************************
' SUBROUTINE : dtpEndExposureTime2_Change()
'
' DESCRIPTION: Executes when user changes the End Exposure TIME
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
'*******************************************************
Public Sub dtpEndExposureTime2_Change_Handler()

    On Error GoTo Error_Handler
    
    'JNM 05/12/03 - PT fix - Set the booleans...
    gbHasChanged = True
    bDateTimeStatusClick = False
    
    Exit Sub
    
Error_Handler:

    If Err.Number <> 0 Then
         Call psAbort(frmName, fdtpEndExposureTime2_Change, Err.Description)
    End If
    
End Sub


'*******************************************************
' SUBROUTINE : dtpStartExposureTime2_Change()
'
' DESCRIPTION: Executes when user changes the Start Exposure TIME
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
'*******************************************************
Public Sub dtpStartExposureDate2_Change_Handler()
    On Error GoTo Error_Handler
    'JNM 05/12/03 - PT fix - Set the booleans...
    gbHasChanged = True
    bDateTimeStatusClick = False
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(frmName, fdtpStartExposureDate2_Change, Err.Description)
    End If
End Sub




'*******************************************************
' SUBROUTINE : dtpStartExposureTime2_Change()
'
' DESCRIPTION: Executes when user changes the Start Exposure TIME
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
'*******************************************************

Public Sub dtpStartExposureTime2_Change_Handler()
On Error GoTo Error_Handler
    'JNM 05/12/03 - PT fix - Set the booleans...
    gbHasChanged = True
    bDateTimeStatusClick = False
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(frmName, fdtpStartExposureTime2_Change, Err.Description)
    End If
End Sub


'*******************************************************
' SUBROUTINE : dtpStartIntroDate2_Change()
'
' DESCRIPTION: Executes when user changes the Start Introduction DATE.
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
'*******************************************************
Public Sub dtpStartIntroDate2_Change_Handler()
    On Error GoTo Error_Handler
    'JNM 05/12/03 - PT fix - Set the booleans...
    gbHasChanged = True
    bDateTimeStatusClick = False
Error_Handler:
     If Err.Number <> 0 Then
        Call psAbort(frmName, fdtpStartIntroDate2_Change, Err.Description)
    End If
End Sub



'*******************************************************
' SUBROUTINE : Recalc values
'
' DESCRIPTION: This subroutine recalculate values after a change in the Start Intro Date/Time,
'              Start Exposure Date/Time, or End Exposure Date/Time.
'              Calls the following functions from Monitor Data Input module:
'                   SetModeFromTimes, CalculateMonDataInput
'              After calculation, this subroutine will reload the form to show updated data.
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
' REVISION HISTORY
' Revision:    0
' Author:      Gamboa, Maria Cecilia (NAI1686)
' Date:        October 12, 2002
' Description: Commented-out the portion that sets the connection to nothing
'
' Revision:    1
' Author:      Espaldon, Marilou (NAI1389)
' Date:        6/3/03
' Description: Changed parameters being passed to function SetModesFromTimes.

'*******************************************************
Private Sub RecalcValues()
   Dim sMonitorValueID As String
    Dim dtDateTime As Date
    Dim dfumiConc As Double
    Dim dInterConc As Double
    Dim sNote As String
    Dim eMode As Integer
    Dim sMontPtGuid As String
    Dim sSQL As String
    
    On Error GoTo Error_Handler
    
    
    '--Update Modes
    'MME 6/6/03 Changing parameters passed to the variables containing the date values
    '           instead of the date control values.
    'SetModesFromTimes frmMain.dtpStartIntroDate2.Value, frmMain.dtpStartExposureDate2.Value, _
    '        frmMain.dtpEndExposureDate2.Value
    
    SetModesFromTimes dtStartIntroduction, dtStartExposure, dtEndExposure
    
    
    GetMonitorPoint ' JNM - added this line
    GetMonitorValues
    
    '--Recalc
    If Not (rsMonitorValues.EOF And rsMonitorValues.BOF) Then
        CalculateMonDataInput
    End If
    
    '-Save grid settings and set open objects to nothing
     MStatGrid.SaveGridSettings "Software\DAS\Profume\Settings"
     rsMonitorValues.Close
     rsMonitorPoint.Close
     Set rsMonitorValues = Nothing
     Set rsMonitorPoint = Nothing
    
     ' TODO: Jett could this be the culprit? Why do we have to reload the whole page?
     ' Form_Load
     Unload_MonitorStatus
     Load_MonitorStatus
     
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(frmName, fRecalcValues, Err.Description)
    End If
End Sub

' SUBROUTINE : dtpStartIntroTime2_Change()
'
' DESCRIPTION: Executes when user changes the Start Introduction TIME.
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
'*******************************************************

Public Sub dtpStartIntroTime2_Change_Handler()
    On Error GoTo Error_Handler
    'JNM 05/12/03 - PT fix - Set the booleans...
    gbHasChanged = True
    bDateTimeStatusClick = False
Error_Handler:
     If Err.Number <> 0 Then
        Call psAbort(frmName, fdtpStartIntroTime2_Change, Err.Description)
    End If
End Sub


' SUBROUTINE : Form_Load()
'
' DESCRIPTION: Executes upon load of this form, hence this will be the very first subroutine executed.
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25,  2002
' Revision:    1
' Author:      Martinez, Amana Karen G. (NAI1499)
' Date:        March 31, 2003
'
' Revision:    2
' Author:      Manyaga, Joseph N. (NAI1227)
' Date:        April 4, 2003
' Change:      Added 2 new columns in order to create double vertical lines
'              in the grid.
' Revision:    3
' Author:      Martinez, Amana Karen (NAI1499)
' Date:        April 24, 2003
' Change:      Added a code for the new control
'*******************************************************
Public Sub Load_MonitorStatus()
    
    Dim sSQL
    
    On Error GoTo Error_Handler
    
    'JNM 05/15/03 - PT fix - Set this to true since we know that the dates
    '   are valid during Load
    bDateTimeStatusClick = True
    gbCanChangeTabs = True
    
    With frmMain
        ' 03/31/03 - AM Removed the existing labels and added the following code for Fumiguide Enhancements 2003
        ' Load the Resource Strings in the Appropiate controls
        .lblStartIntro2.Caption = LoadResString(IDS_START_INTRO)
        .lblStartExpo2.Caption = LoadResString(IDS_START_EXPOSURE)
        .lblEndExpo2.Caption = LoadResString(IDS_END_EXPOSURE)
        .lblSIDate2.Caption = LoadResString(IDS_DATE)
        .lblSITime2.Caption = LoadResString(IDS_TIME)
        .lblSEDate2.Caption = LoadResString(IDS_DATE)
        .lblSETime2.Caption = LoadResString(IDS_TIME)
        .lblEEDate2.Caption = LoadResString(IDS_DATE)
        .lblEETime2.Caption = LoadResString(IDS_TIME)
        ' 03/31/03 - AM Removed the existing labels and added the code above for Fumiguide Enhancements 2003
        
        'JNM 05/12/03 - Pt fix - create this new button
        .cmdValidDateStatus.Caption = LoadResString(IDS_LBL_CMD_VALIDDATE)
        
        '-- Set tooltip text for date/time labels
        'JNM 10/26/02 - The controls should have the tooltip text
        .dtpStartIntroDate2.ToolTipText = LoadResString(IDS_MSTOOLTIP_START_INTRODUCTION)
        .dtpStartIntroTime2.ToolTipText = LoadResString(IDS_MSTOOLTIP_START_INTRODUCTION)
        .dtpStartExposureDate2.ToolTipText = LoadResString(IDS_MSTOOLTIP_START_EXPOSURE)
        .dtpStartExposureTime2.ToolTipText = LoadResString(IDS_MSTOOLTIP_START_EXPOSURE)
        .dtpEndExposureDate2.ToolTipText = LoadResString(IDS_MSTOOLTIP_END_EXPOSURE)
        .dtpEndExposureTime2.ToolTipText = LoadResString(IDS_MSTOOLTIP_END_EXPOSURE)
        
        '04/24/03 - AM Added the following for Assembly Test Defect 8495 of Fumiguide Enhancement
        'set the button text
        .cmdNextStep7.Caption = LoadResString(IDS_BTN_NEXTSTEP)
        '04/24/03 - AM Added the following for Assembly Test Defect 8495 of Fumiguide Enhancement
    End With
    
    '--Start Initialize Grid
    MStatGrid.Initialize "MonitorStatus", frmMain.grdMonitorStatus
    'Pat 3/3/2004 - Make strings constant
    MStatGrid.AddColumn "colAreaName", LoadResString(IDS_COLHEADER1_STATUS), colTypeText, flexAlignGeneral
    MStatGrid.AddColumn "colMonPoint", LoadResString(IDS_COLHEADER2_STATUS), colTypeText, flexAlignLeftTop
    MStatGrid.AddColumn "colTimeLastEntry", LoadResString(IDS_COLHEADER3_STATUS), colTypeText, flexAlignGeneral
    MStatGrid.AddColumn "colElapsedTime", LoadResString(IDS_COLHEADER4_STATUS), colTypeText, flexAlignGeneral
    
    'JNM 04/04/03 - Fum. Enhancements 2003 - Added extra column in order to create
    '   double vertical line.
    MStatGrid.AddColumn "colSpace1", "", colTypeText, flexAlignGeneral
    
    MStatGrid.AddColumn "colHLT", LoadResString(IDS_COLHEADER5_STATUS), colTypeTime, flexAlignGeneral
    MStatGrid.AddColumn "colStartTimeHLT", LoadResString(IDS_COLHEADER6_STATUS), colTypeText, flexAlignGeneral
    MStatGrid.AddColumn "colEndTimeHLT", LoadResString(IDS_COLHEADER7_STATUS), colTypeText, flexAlignGeneral
    
    'JNM 04/04/03 - Fum. Enhancements 2003 - Added extra column in order to create
    '   double vertical line.
    MStatGrid.AddColumn "colSpace2", "", colTypeText, flexAlignGeneral
    
    MStatGrid.AddColumn "colCTAchieved", LoadResString(IDS_COLHEADER8_STATUS), colTypeText, flexAlignGeneral
    MStatGrid.AddColumn "colProjected", LoadResString(IDS_COLHEADER9_STATUS), colTypeText, flexAlignGeneral
    MStatGrid.AddColumn "colStatus", LoadResString(IDS_COLHEADER10_STATUS), colTypeText, flexAlignGeneral
    
    If gbIsTestingActive Then
        MStatGrid.AddColumn "colCnt", "Cnt/Dose", colTypeText, flexAlignGeneral
    End If
       
    With frmMain
        ' set the default widths (will be overridden if settings were previously saved
        .grdMonitorStatus.ColWidth(COL_AREANAME) = AREANAME_COLSIZE
        .grdMonitorStatus.ColWidth(COL_MONPOINT) = MONPOINT_COLSIZE
        .grdMonitorStatus.ColWidth(COL_LASTENTRY) = LASTENTRY_COLSIZE
        .grdMonitorStatus.ColWidth(COL_ELAPSED) = ELAPSED_COLSIZE
        .grdMonitorStatus.ColWidth(COL_SPACE1) = 50   'JNM 04/04/03 - Size for space column
        .grdMonitorStatus.ColWidth(COL_HLT) = HLT_COLSIZE
        .grdMonitorStatus.ColWidth(COL_STARTHLT) = STARTHLT_COLSIZE
        .grdMonitorStatus.ColWidth(COL_ENDHLT) = ENDHLT_COLSIZE
        .grdMonitorStatus.ColWidth(COL_SPACE2) = 50   'JNM 04/04/03 - Size for space column
        .grdMonitorStatus.ColWidth(COL_CTACH) = CTACH_COLSIZDE
        .grdMonitorStatus.ColWidth(COL_PROJ) = PROJ_COLSIZE
        .grdMonitorStatus.ColWidth(COL_STATUS) = STATUS_COLSIZE
    
        MStatGrid.LoadGridSettings "Software\DAS\Profume\Settings"
        
        '--set Grid Properties
        .grdMonitorStatus.AutoSizeMode = flexAutoSizeRowHeight
        .grdMonitorStatus.WordWrap = True
        .grdMonitorStatus.Editable = flexEDNone
        .grdMonitorStatus.Rows = 1
        .grdMonitorStatus.BackColor = READ_ONLY_COLOR
        .grdMonitorStatus.GridColor = GRID_COLOR
        
        .grdMonitorStatus.AllowSelection = False
        .grdMonitorStatus.AllowBigSelection = False
        
    End With
                   
    '-- Get input dates
    GetMonitorDataInputDates
          
    PopulateStatus
    
    frmMain.grdMonitorStatus.AutoSize COL_STATUS   'wrap the status column
        
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(frmName, fMain, Err.Description)
    End If
End Sub

'*******************************************************
' SUBROUTINE : GetMonitorDataInputDates
'
' DESCRIPTION: Retrieves from Job table the following date/time info:
'               >startintro_datetime (Start Introduction)
'               >startExp_datetime (Start Exposure)
'               >EndExp_Datetime (End Exposure)
'              These dates are the ones displayed on the form as "Start Introduction Date/Time",
'              Start Exposure Date/Time, and End Exposure Date/Time
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
'
' Revision:    1
' Author:      Gamboa, Maria Cecilia (NAI1686)
' Date:        October 11, 2002
' Description: Added code to handle null values and disable the controls
'               if the value is null
'*******************************************************
Private Sub GetMonitorDataInputDates()
   Dim sSQL As String
   
   On Error GoTo Error_Handler
   
   sSQL = "SELECT startintro_datetime, startexp_datetime, endexp_datetime " & _
            "FROM Job where id = '" & gsJobID & "'"
   
   rsJob.Open sSQL, gdbConn, adOpenDynamic, adLockReadOnly
     
    With frmMain
        If Not (rsJob.EOF And rsJob.BOF) Then
             rsJob.MoveFirst
             '--Display the dates and times in the corresponding controls
             If IsNull(rsJob("startintro_datetime")) Then
                 .dtpStartIntroDate2.Enabled = False
                 .dtpStartIntroTime2.Enabled = False
             Else
                 .dtpStartIntroDate2.Enabled = True
                 .dtpStartIntroTime2.Enabled = True
                 .dtpStartIntroDate2.Value = rsJob("startintro_datetime")
                 .dtpStartIntroTime2.Value = rsJob("startintro_datetime")
             End If
             
             If IsNull(rsJob("startexp_datetime")) Then
                 .dtpStartExposureDate2.Enabled = False
                 .dtpStartExposureTime2.Enabled = False
             Else
                 .dtpStartExposureDate2.Enabled = True
                 .dtpStartExposureTime2.Enabled = True
                 .dtpStartExposureDate2.Value = rsJob("startexp_datetime")
                 .dtpStartExposureTime2.Value = rsJob("startexp_datetime")
             End If
             
             If IsNull(rsJob("endexp_datetime")) Then
                 .dtpEndExposureDate2.Enabled = False
                 .dtpEndExposureTime2.Enabled = False
             Else
                 .dtpEndExposureDate2.Enabled = True
                 .dtpEndExposureTime2.Enabled = True
                 .dtpEndExposureDate2.Value = rsJob("endexp_datetime")
                 .dtpEndExposureTime2.Value = rsJob("endexp_datetime")
             End If
        End If
   End With
   
Error_Handler:
   If Err.Number <> 0 Then
        Call psAbort(frmName, fGetMonitorDataInputDates, Err.Description)
   End If
End Sub


Public Sub Unload_MonitorStatus()

    On Error GoTo Error_Handler
    
    If Not CanExitMonStatus Then
        gbCanChangeTabs = False
    End If
    
    MStatGrid.SaveGridSettings "Software\DAS\Profume\Settings"
    'Optimize! Destroy!
    Set MStatGrid = Nothing
        
    'Destory the recordset objects if these are still open
    If rsMonitorValues.State = adStateOpen Then
        rsMonitorValues.Close
        Set rsMonitorValues = Nothing
    End If
    
    If rsMonitorPoint.State = adStateOpen Then
        rsMonitorPoint.Close
        Set rsMonitorPoint = Nothing
    End If
    
    If rsJob.State = adStateOpen Then
        rsJob.Close
        Set rsJob = Nothing
    End If
        
Error_Handler:
   If Err.Number <> 0 Then
        Call psAbort(frmName, fUnload, Err.Description)
   End If
End Sub


'*******************************************************
' SUBROUTINE : PopulateStatus
'
' DESCRIPTION: This subroutine populates the grid with monitoring points details
'              Calls PopulateMonitorPoint and PopulateAllPoint
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
' Revision:    1
' Author:      Gamboa, Maria Cecilia (NAI1686)
' Date:        October 14, 2002
' Description: Commented-out all statements that closes the connection
' Revision:    2
' Author:      Manyaga, Joseph
' Date:        October 16, 2002
' Description: Commented-out all statements that displays the message boxes
' Revision:    3
' Author:      Manyaga, Joseph
' Date:        October 18, 2002
' Description: Changed the overall processing and flow of
'   the subroutine. This was a major change as an AT fix.
'*******************************************************
Private Sub PopulateStatus()
    Dim nMonPtsInArea As Integer
    Dim sLastAreaID As String
    Dim nItem As Integer
               
    On Error GoTo Error_Handler
    
    'JNM - Create the MonitorPoint recordset
    GetMonitorPoint
       
    'JNM - check if we have MonitorPoints
    If (rsMonitorPoint.EOF And rsMonitorPoint.BOF) Then
        rsMonitorPoint.Close
        Set rsMonitorPoint = Nothing
        Exit Sub
    End If
    
    rsMonitorPoint.MoveFirst
     
    '--initialize Current Area ID
    sLastAreaID = rsMonitorPoint("jobarea_fkey")
    nMonPtsInArea = 0   'the number of monitor points for the area
    nItem = 1           'the grid row number
    
    'JNM - insert data in one row for each MonitorPoint
    While Not rsMonitorPoint.EOF
        '--make sure that the records are arranged in the same order as the monitoring points are displayed in the monitor plan
            'JNM - check if we are still in the same area
            If rsMonitorPoint("jobarea_fkey") = sLastAreaID Then
                PopulateMonitorPoint sLastAreaID, nItem
                nMonPtsInArea = nMonPtsInArea + 1
            Else
            '--(It means we have a new area)
                If nMonPtsInArea > 1 Then
                    'JNM - If the Area has more than 1 Mon Pt, insert a status for the Area
                    PopulateAllPoint sLastAreaID, nItem
                    nItem = nItem + 1
                Else
                    'JNM - Since there was only one monitor point, we need to display the area status
                    '--PopulatePrevMonitorPointStatus
                    PopulatePrevMonitorPointStatus nItem
                End If
                'JNM  - Display the current mon pt
                sLastAreaID = rsMonitorPoint("jobarea_fkey")
                PopulateMonitorPoint sLastAreaID, nItem
                nMonPtsInArea = 1
            End If
            nItem = nItem + 1
        rsMonitorPoint.MoveNext
    Wend
                
    '--May need to add the last monitoring point's area status
    If nMonPtsInArea > 1 Then
        PopulateAllPoint sLastAreaID, nItem
    Else
        PopulatePrevMonitorPointStatus nItem
    End If
    
    ' 11/13/02 TSG - added this so no rows are selected
    frmMain.grdMonitorStatus.Select 0, 0
    
    '--Close all
    Set rsJob = Nothing
        
Error_Handler:
   If Err.Number <> 0 Then
        Call psAbort(frmName, fPopulateStatus, Err.Description)
   End If
   
End Sub


Private Sub GetMonitorPoint()
'*******************************************************
' SUBROUTINE : GetMonitorPoint
'
' DESCRIPTION: Puts data onto recordset rsMonitorPoint.
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Manyaga, Joseph N.  (NAI1227)
' Date:        October 18, 2002
'*******************************************************
'We need all the MonitorPoints that match our Job, plus the Area name that
'matches each MonitorPoint

    Dim sSQL As String
    
    On Error GoTo Error_Handler
    
'   5/14/2003 Jett - removed the ORDER BY statement so that the Areas and Monitor Points
'   appear as they are listed in the Dosage Plan and Monitor Plan tabs
'    sSQL = "SELECT JobArea.area_name, MonitorPoint.* FROM MonitorPoint, JobArea " & _
'       "WHERE JobArea.id = MonitorPoint.jobarea_fkey AND MonitorPoint.job_fkey = '" & _
'       gsJobID & "' AND MonitorPoint.deleted = FALSE ORDER BY JobArea.area_name, " & _
'       "MonitorPoint.name"
'
'    6/03/2003 MME - Changed sequel to order records by row number (row field) which will
'                    let the order of the Areas in the Monitor Status tab coincide with the order
'                    in the Fumigation Dosage Plan.

'    sSQL = "SELECT JobArea.area_name, MonitorPoint.* FROM MonitorPoint, JobArea " & _
'       "WHERE JobArea.id = MonitorPoint.jobarea_fkey AND MonitorPoint.job_fkey = '" & _
'       gsJobID & "' AND MonitorPoint.deleted = FALSE"
       
       
    sSQL = "SELECT Jobarea.row,JobArea.area_name, MonitorPoint.* FROM MonitorPoint, JobArea " & _
       "WHERE JobArea.id = MonitorPoint.jobarea_fkey AND MonitorPoint.job_fkey = '" & _
       gsJobID & "' AND MonitorPoint.deleted = FALSE ORDER BY JobArea.row"
    
    If rsMonitorPoint.State = adStateOpen Then
        rsMonitorPoint.Close
    End If
    
    rsMonitorPoint.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
Error_Handler:
       If Err.Number <> 0 Then
            Call psAbort(frmName, "GetMonitorPoint", Err.Description)
       End If

End Sub


'*******************************************************
' SUBROUTINE : GetMonitorValues
'
' DESCRIPTION: Puts data onto recordsets rsMonitorPoint and rsMonitorValues.
'              rsMonitorValues contains the records processed and from which the details
'              displayed on the grid came from.
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
' Revision:    0
' Author:      Manyaga, Joseph N.  (NAI1227)
' Date:        October 16, 2002
' Description  Changed the SQL statement used in retrieving the
'              monitor points.
'*******************************************************
Private Sub GetMonitorValues()
    Dim sSQL
    
    On Error GoTo Error_Handler
         
    '--open recordset
    'rsMonitorPoint.Open sSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
        
    'Create the MonitorPointValue recordset if there are Monitor Points
    If Not (rsMonitorPoint.EOF And rsMonitorPoint.BOF) Then
        sSQL = "SELECT MonitorPoint.*, MonitorPointValue.* " & _
                 "FROM MonitorPointValue, MonitorPoint " & _
                 "WHERE MonitorPointValue.deleted = FALSE " & _
                 "AND MonitorPoint.id = MonitorPointValue.monitorpoint_fkey " & _
                 "ORDER BY MonitorPoint.jobarea_fkey ASC, MonitorPointValue.date_time ASC, " & _
                 "MonitorPointValue.mode ASC, MonitorPoint.name ASC "
                 
        If rsMonitorValues.State = adStateOpen Then
            rsMonitorValues.Close
        End If
        
        rsMonitorValues.Open sSQL, gdbConn, adOpenDynamic, adLockReadOnly
        
    End If
        
Error_Handler:
   If Err.Number <> 0 Then
        Call psAbort(frmName, fGetMonitorValues, Err.Description)
   End If
    
End Sub


'*******************************************************
' SUBROUTINE : PopulateMonitorPoint
'
' DESCRIPTION: This subroutine populates a row on the grid with details for a given
'              Monitoring Point.
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
' Revision:    1
' Author:      Manyaga, Joseph (NAI1227)
' Date:        October 18, 2002
' Description: Changed the flow of the subroutine.
'
' Revision:    2
' Author:      Manyaga, Joseph (NAI1227)
' Date:        April 4, 2003
' Description: Fum. Enhancements 2003 - Adjusted the column numbers due
'              to double vertical line.
'
' 7/30/2003    Jett - SIMS #2099998 - fixed unit of measure displayed for CnT, CT AChieved, and Projectd CT
'*******************************************************
Private Sub PopulateMonitorPoint(sLastAreaID As String, nRow As Integer)
    Dim sSQL As String
    Dim I As Integer
    
    On Error GoTo Error_Handler
   
    '--add a new row
    frmMain.grdMonitorStatus.AddItem ("")
    I = nRow 'JNM - changed this : grdMonitorStatus.Rows - 1
    frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_AREANAME) = rsMonitorPoint("area_name")
    frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_MONPOINT) = rsMonitorPoint("name")
    If rsMonitorPoint("last_entry") > 0 Then 'JNM - added this condition
        frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_LASTENTRY) = FormatFumDate(rsMonitorPoint("last_entry"))
    End If
    If rsMonitorPoint("elapsed_time") >= 0 Then 'JNM - added this condition
        frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_ELAPSED) = Format(Round(rsMonitorPoint("elapsed_time"), 1)) & " " & LoadResString(IDS_HOUR_ABBREV2)
    End If
    If rsMonitorPoint("hlt") > 0# Then  'JNM - added this condition
        frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_HLT) = Format(Round(rsMonitorPoint("hlt"), 1)) & " " & LoadResString(IDS_HOUR_ABBREV2)
        frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_STARTHLT) = FormatFumDate(rsMonitorPoint("start_time"))
        frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_ENDHLT) = FormatFumDate(rsMonitorPoint("end_time"))
    End If
    If rsMonitorPoint("achieved_concentration_time") >= 0 Then  'JNM - added this condition
        frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_CTACH) = Format(Round(rsMonitorPoint("achieved_concentration_time"), 0)) & " " & _
                                                      IIf(gbIsMetric, LoadResString(IDS_G_HR_PER_CUM), _
                                                      LoadResString(IDS_OZ_HR_PER_MCF))
        If gbIsTestingActive Then
            frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_TESTCTACH) = Format(Round(rsMonitorPoint("achieved_dose"), 0)) & " " & _
                                                      IIf(gbIsMetric, LoadResString(IDS_G_HR_PER_CUM), _
                                                      LoadResString(IDS_OZ_HR_PER_MCF))
        End If
    End If
    'If rsMonitorPoint("status") <> 4 And rsMonitorPoint("projected_concentration_time") >= 0 Then   'JNM - added this condition
    If rsMonitorPoint("projected_concentration_time") >= 0 Then   'JNM - added this condition
        frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_PROJ) = Format(Round(rsMonitorPoint("projected_concentration_time"), 0)) & " " & _
                                                      IIf(gbIsMetric, LoadResString(IDS_G_HR_PER_CUM), _
                                                      LoadResString(IDS_OZ_HR_PER_MCF))
    End If
                                           
Error_Handler:
   If Err.Number <> 0 Then
        Call psAbort(frmName, fPopulateMonitorPoint, Err.Description)
   End If
End Sub


'*******************************************************
' SUBROUTINE : PopulateAllPoint
'
' DESCRIPTION: This subroutine populates a row on the grid containing summary of all monitoring
'              points read for an area
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
'
' Revision:    0
' Author:      Manyaga, Joseph N. (NAI1227)
' Date:        April 4, 2003
' Description: Adjusted column numbers due to double vertical lines
'              Added a code that would highlight rows for Incompatible status.
'
' 7/30/2003 Jett - SIMS 2099998: used correct unit of measure for CT Achieved, Projected CT, CnT
'*******************************************************
Private Sub PopulateAllPoint(sAreaId As String, nRow As Integer)    'JNM - added the 2 parameters
'--This function provides the area summary lines in the monitoring status grid box.  The summary row will be highlighted

    On Error GoTo Error_Handler
    
    Dim strJobId As String
    Dim sSQL As String
    Dim I As Integer
    Dim sStatus As String
    Dim rsJobArea As New ADODB.Recordset
        
    sSQL = "SELECT * " & _
             "FROM JobArea " & _
             "WHERE JobArea.deleted = FALSE " & _
             "AND JobArea.id= '" & sAreaId & "' and job_fkey = '" & gsJobID & "'"
    
    rsJobArea.Open sSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    rsJobArea.MoveFirst
    '--put items in the status list control
    frmMain.grdMonitorStatus.AddItem ("")
    I = nRow
    frmMain.grdMonitorStatus.Select I, 0, I, IIf(gbIsTestingActive, COL_TESTCTACH, COL_STATUS)
    frmMain.grdMonitorStatus.FillStyle = flexFillRepeat
   
    ' 11/15/02 TSG - set text to bold for all area sumary
    frmMain.grdMonitorStatus.Cell(flexcpFontBold, nRow, 0, nRow, frmMain.grdMonitorStatus.Cols - 1) = True
   
    frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_AREANAME) = rsJobArea("area_name")
    frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_MONPOINT) = LoadResString(IDS_ALL) & " " & rsJobArea("area_name")
    '--time of last entry and elapsed time are both empty
    If rsJobArea("hlt") > 0# Then
       frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_HLT) = Format(Round(rsJobArea("hlt"), 1)) & " " & LoadResString(IDS_HOUR_ABBREV2)
    End If
    '--start time and end time of hlt are both empty
    If (rsJobArea("achieved_concentration_time") >= 0) Then
        'JNM - added the code below
        frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_CTACH) = Format(Round(rsJobArea("achieved_concentration_time"), 0)) & " " & _
                                                      IIf(gbIsMetric, LoadResString(IDS_G_HR_PER_CUM), _
                                                      LoadResString(IDS_OZ_HR_PER_MCF))
        If gbIsTestingActive Then
            frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_TESTCTACH) = Format(Round(rsJobArea("achieved_dose"), 0)) & " " & _
                                                      IIf(gbIsMetric, LoadResString(IDS_G_HR_PER_CUM), _
                                                      LoadResString(IDS_OZ_HR_PER_MCF))
        End If
    End If
    
    '--Insert Projected Concentration Time(CT)
    If (rsJobArea("status") <> 4 And rsJobArea("projected_concentration_time") >= 0) Then
        frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_PROJ) = Format(Round(rsJobArea("projected_concentration_time"), 0)) & " " & _
                                                      IIf(gbIsMetric, LoadResString(IDS_G_HR_PER_CUM), _
                                                      LoadResString(IDS_OZ_HR_PER_MCF))
    End If
       
    '--Write out the JobArea status information
    sStatus = GetJobAreaStatus(rsJobArea, rsJob)
    frmMain.grdMonitorStatus.Cell(flexcpText, I, COL_STATUS) = sStatus
       
    ' 12/16/02 Jett - set text to white on red background if we have a bad elapsed times status
    ' 1/3/03 Jett - also set text if target cannot be achieved
    'JNM 04/09/03 - Incompatable will also be highlighted since this does not also achieve
    '   the target.
    'TSG 05/19/03 - Also set if missing concentration readings
    If rsJobArea("status") = STATUS_BAD_ELAPSED_TIMES Or rsJobArea("status") = STATUS_BELOW_TARGET Or rsJobArea("status") = STATUS_TARGET_DOSE_INCOMPATABLE Or rsJobArea("status") = STATUS_MISSING_CONC_READING Then
        frmMain.grdMonitorStatus.Cell(flexcpFontBold, nRow, 0, nRow, frmMain.grdMonitorStatus.Cols - 1) = True
        frmMain.grdMonitorStatus.Cell(flexcpBackColor, nRow, 0, nRow, frmMain.grdMonitorStatus.Cols - 1) = &HFF& '&H8080FF
        frmMain.grdMonitorStatus.Cell(flexcpForeColor, nRow, 0, nRow, frmMain.grdMonitorStatus.Cols - 1) = &HFFFFFF
    End If
       
    rsJobArea.Close
    Set rsJobArea = Nothing
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(frmName, fPopulateAllPoint, Err.Description)
    End If
End Sub

'*******************************************************
' SUBROUTINE : PopulatePrevMonitorPointStatus
'
' DESCRIPTION: Sets the status of a single-monitor point Area
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Joseph Manyaga (NAI1227)
' Date:        October 18, 2002
'
' Revision:    1
' Author:      Joseph Manyaga (NAI1227)
' Date:        April 4, 2003
' Description: Adjusted column numbers due to double vertical lines
'              Added code that would highlight rows for Incompatible status.
'*******************************************************
Private Sub PopulatePrevMonitorPointStatus(nRow As Integer)
    Dim myBookMark
    Dim bEOF As Boolean
    Dim rsJobArea As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo Error_Handler
    
    bEOF = False
    'Set the bookmark for the current record
    If rsMonitorPoint.EOF Then
        bEOF = True
    Else
        myBookMark = rsMonitorPoint.Bookmark
    End If
    
    'Go to the most recent active monitor point
    'Do
    '    rsMonitorPoint.MovePrevious
    'Loop While Not rsMonitorPoint.BOF And rsMonitorPoint("inactive")
    
    rsMonitorPoint.MovePrevious
        
    'If we are at the top of the file, then no monitoring points are active before our current pt
    If Not rsMonitorPoint.BOF Then
        sSQL = "SELECT * FROM JobArea WHERE JobArea.deleted = FALSE AND " & _
            "JobArea.job_fkey = '" & gsJobID & "' AND " & _
            "JobArea.id = '" & rsMonitorPoint("jobarea_fkey") & "'"
            
        rsJobArea.Open sSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
        
        If rsJobArea.BOF And rsJobArea.EOF Then
            frmMain.grdMonitorStatus.Cell(flexcpText, nRow - 1, COL_STATUS) = ""
        Else
            rsJobArea.MoveFirst
            frmMain.grdMonitorStatus.Cell(flexcpText, nRow - 1, COL_STATUS) = GetJobAreaStatus(rsJobArea, rsJob)
            
            'JNM 11/15 - set the Area font to bold
            frmMain.grdMonitorStatus.Cell(flexcpFontBold, nRow - 1, 0, nRow - 1, frmMain.grdMonitorStatus.Cols - 1) = True
        End If
        
        'JNM 04/09/04 - Using Jett's code, highlight text if target cannot be achieved.
        If rsJobArea("status") = STATUS_BAD_ELAPSED_TIMES Or rsJobArea("status") = STATUS_BELOW_TARGET Or rsJobArea("status") = STATUS_TARGET_DOSE_INCOMPATABLE Then
            frmMain.grdMonitorStatus.Cell(flexcpFontBold, nRow - 1, 0, nRow - 1, frmMain.grdMonitorStatus.Cols - 1) = True
            frmMain.grdMonitorStatus.Cell(flexcpBackColor, nRow - 1, 0, nRow - 1, frmMain.grdMonitorStatus.Cols - 1) = &HFF& '&H8080FF
            frmMain.grdMonitorStatus.Cell(flexcpForeColor, nRow - 1, 0, nRow - 1, frmMain.grdMonitorStatus.Cols - 1) = &HFFFFFF
        End If
        
    End If
    
    'Go back to the current record
    If bEOF Then
        rsMonitorPoint.MoveLast
        rsMonitorPoint.MoveNext
    Else
        rsMonitorPoint.Bookmark = myBookMark
    End If
    
    If rsJobArea.State = adStateOpen Then
        rsJobArea.Close
        Set rsJobArea = Nothing
    End If

Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(frmName, "PopulatePrevMonitorPointStatus", Err.Description)
    End If

End Sub

Public Sub grdMonitorStatus_BeforeUserResize_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Do not allow the user to change the width of the Space column
    If Col = COL_SPACE1 Or Col = COL_SPACE2 Then
        Cancel = True
    End If
End Sub


'*******************************************************
' SUBROUTINE : grdMonitorStatus_MouseMove
'
' DESCRIPTION: Sets ToolTip Texts for the grid columns
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Espaldon, Marilou M. (NAI1389)
' Date:        September 25, 2002
'*******************************************************

Public Sub grdMonitorStatus_MouseMove_Handler(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    '--Note that the next step will just be executed in case there is an error in this part.
    '--This function will be displaying tooltip texts only and an error here should not be considered
    '--as a cause for program termination.
    
    On Error Resume Next
    
    With frmMain.grdMonitorStatus
        Select Case .MouseCol
            Case COL_AREANAME
                .ToolTipText = LoadResString(IDS_MSTOOLTIP_AREA_NAME)
            Case COL_MONPOINT
                .ToolTipText = LoadResString(IDS_MSTOOLTIP_MONITORING_POINT)
            Case COL_LASTENTRY
                .ToolTipText = LoadResString(IDS_MSTOOLTIP_TIME_OF_LAST)
            Case COL_ELAPSED
                .ToolTipText = LoadResString(IDS_MSTOOLTIP_ELAPSED)
            Case COL_HLT
                .ToolTipText = LoadResString(IDS_MSTOOLTIP_HLT)
            Case COL_STARTHLT
                .ToolTipText = LoadResString(IDS_MSTOOLTIP_START_TIME_OF_HLT)
            Case COL_ENDHLT
                .ToolTipText = LoadResString(IDS_MSTOOLTIP_END_TIME_OF_HLT)
            Case COL_CTACH
                .ToolTipText = LoadResString(IDS_MSTOOLTIP_CT_ACHIEVED)
            Case COL_PROJ
                .ToolTipText = LoadResString(IDS_MSTOOLTIP_PROJECTED)
            Case COL_STATUS
                .ToolTipText = LoadResString(IDS_MSTOOLTIP_STATUS_RECOMMENDATIONS)
                
        End Select
    End With
End Sub


'*******************************************************
' SUBROUTINE : ReinitDateTime
'
' DESCRIPTION: Stores the date and time values of all the
'              dtPicker controls.
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Manyaga, Joseph (NAI1227)
' Date:        October 18, 2002
'*******************************************************
Private Sub ReinitDateTime()

    On Error GoTo Error_Handler
    
    'JNM - combine the values of the 2 controls
    dtStartIntroduction = DateValue(frmMain.dtpStartIntroDate2.Value) & " " & TimeValue(frmMain.dtpStartIntroTime2.Value)
    dtStartExposure = DateValue(frmMain.dtpStartExposureDate2.Value) & " " & _
        TimeValue(frmMain.dtpStartExposureTime2.Value)
    dtEndExposure = DateValue(frmMain.dtpEndExposureDate2.Value) & " " & _
        TimeValue(frmMain.dtpEndExposureTime2.Value)
        
Error_Handler:
        If Err.Number <> 0 Then
            Call psAbort(frmName, "ReinitDateTime", Err.Description)
        End If

End Sub


'*************************************************************
'   DESCRIPTION
'   Handles the Click handler of the cmdValidDateStatus button
'
'   Input Parameters: None
'
'   RETURNS - True if the dates are valid
'             False if the dates are invalid
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  May 12, 2003     Joseph Manyaga    Initial Revision
'  May 26, 2003    Jett Gamboa       Forced change the planned Exposure Time
'  June 3, 2003    Malou Espaldon    Removed variable declaration for first three
'                                    variables because these variables are already
'                                    declared at the top of this module.
'**************************************************************
Public Sub cmdValidDateStatus_Click_Handler()
    '--MME Deleted the following variables
    'Dim dtStartIntroduction As Date
    'Dim dtStartExposure As Date
    'Dim dtEndExposure As Date
    
    Dim dActualExposureTime As Double
    
    On Error GoTo Error_Handler
    
    dtStartIntroduction = DateValue(frmMain.dtpStartIntroDate2) & " " & _
        TimeValue(frmMain.dtpStartIntroTime2)
    dtStartExposure = DateValue(frmMain.dtpStartExposureDate2) & " " & _
        TimeValue(frmMain.dtpStartExposureTime2)
    dtEndExposure = DateValue(frmMain.dtpEndExposureDate2) & " " & _
        TimeValue(frmMain.dtpEndExposureTime2)
        
    Select Case ValidateJobDates(dtStartIntroduction, dtStartExposure, dtEndExposure)
        Case "Start Introduction"
            frmMain.dtpStartIntroDate2.SetFocus
        Case "Start Exposure"
            frmMain.dtpStartExposureDate2.SetFocus
        Case Else
        
            dActualExposureTime = DateDiff("n", dtStartExposure, dtEndExposure) / 60
            '--JG 6/3/2003 removed rounding off
            'dActualExposureTime = Round(dActualExposureTime, 1)
            dActualExposureTime = dActualExposureTime
            
            ' Check if Actual Exposure Time is the same as the planned exposure
            ' time. If they are not the same, force the planned exposure to be
            ' the same as the actual exposure
            If dActualExposureTime <> gnExposureTime Then
                MsgBox LoadResString(IDS_EXPOSURE_TIME_FORCED), vbInformation
                UpdatePlannedExposure dActualExposureTime
                
                ' Call the Dosage Plan calculate method
                Calculate False
            End If
        
            'Set the click flag to true
            bDateTimeStatusClick = True
            
            'Recalculate
            RecalcValues
            
    End Select
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(frmName, "cmdValidDateStatus_Click_Handler", Err.Description)
    End If

End Sub

Public Function CanExitMonStatus() As Boolean

    On Error GoTo Error_Handler
    
    If Not bDateTimeStatusClick Then
        MsgBox LoadResString(IDS_DATETIME_NOCLICK)
        CanExitMonStatus = False
    Else
        CanExitMonStatus = True
    End If
    
Error_Handler:
    If Err.Number <> 0 Then
        Call psAbort(frmName, "CanExitMonStatus", Err.Description)
    End If
    
End Function




