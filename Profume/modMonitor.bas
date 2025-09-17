Attribute VB_Name = "modMonitor"
'------------------------------------------------------
'*
'* Filename: modMonitor.bas
'*
'* DESCRIPTION
'* -----------
'* This code will contain the functions and subroutines
'* to be referenced by the Monitor Data Input and
'* Monitor Status windows of the Fumiguide (ProFume) Application.
'*
'* Copyright (c) 2002 Accenture.  All rights reserved.
'-----------------------------------------------------
'*
'* REVISION HISTORY
'* Revision:    0
'* Author:      Gamboa, Maria Cecilia <nai1686>
'* Date:        September 25, 2002
'*
'* 4/20/2003    Jett    Cleaned up memory leaks and inefficient recordset declarations
'* 4/21/2003    Jett    display calculation time in status bar if SPEEDTEST = 1 (in project properties)
'* 3/18/2004    Jett    fixed UpdateMonitorPointStatus so it only does GetAreaStatus once for every area
'-----------------------------------------------------

Private Enum Phase
    Start_Intro = 1
    Start_Exp = 2
    End_Exp = 3
End Enum

Private Const MODULE_NAME = "modMonitor"
'OPTIMIZE is not used -- Dim rsMonitorPoint As ADODB.Recordset
'OPTIMIZE is not used -- Dim rsAreaId As New ADODB.Recordset
Dim sUpdateSQL As String


'===========================================================================================================
' the following variables were copied from the (original) frmMonitorInput form
'===========================================================================================================

'Dim hwndHelp As Long
Dim thisGrid As New clsGridWrapper
Private Const GRIDSIZE = 3000
Private Const GRID_COLOR = "&H80000008"
Private Const BACK_COLOR_ALTERNATE = "&H80000016"
Private Const READ_ONLY_COLOR = "&H8000000B"
Private Const NOTES_MAX = 50
Private Const INTERSCAN = -1
Private Const sDate = "DATE"
Private Const sTime = "TIME"

Dim rsMonitorPoint As ADODB.Recordset           ' Jett 4/20/03 removed As New Declaration
' Dim rsMonPtVal As New ADODB.Recordset
' Dim rsConcPattern As New ADODB.Recordset      ' Jett 4/20/03 never used
' Dim rsDelList As New ADODB.Recordset
' Dim rsMonitorValues As New ADODB.Recordset    ' Jett 4/20/03 declared in validate edit (only used there)
' Dim rsMonitorId As New ADODB.Recordset        ' Jett 4/20/03 declared in validate edit and paste handler
' Dim rs As New ADODB.Recordset                 ' Jett 4/20/03 declared in functions where used
' Dim rsNewData As New ADODB.Recordset          ' Jett 4/20/03 declared in pasteclick where it is used
' Dim rsUpdateData As New ADODB.Recordset       ' Jett 4/20/03 declared in pasteclick where it is used

Dim bRecordDirty As Boolean
Dim bIsTableEmpty As Boolean
Dim bIsDatesEmpty As Boolean
Dim bHasCalled As Boolean
Dim bTimeChanged As Boolean
Dim bDateChanged As Boolean
Dim bdtpChanged As Boolean

Dim dtGridDate As Date
Dim dtIntroStarted As Date
Dim dtExposureStarted As Date
Dim dtExposureEnded As Date
Dim sMonPtValueId As String
Dim sSQL As String
Dim sPipeList As String
Dim sGridDateTime As String
Dim nEnabled As Integer
Dim nRow As Integer
Dim nDateCtr As Integer
Dim nTimeCtr As Integer
Dim dValue As Double
Dim I As Integer
Dim nDecimal As Integer
Dim mode As Integer
Dim sConcUnits As String
Dim sHighlight As String
Dim sDefaultMonPtID As String
Dim sDefaultMonPt As String

'JNM 05/12/03 - This will flag if the Accept Date:Time button has been clicked
Dim bDateTimeClick As Boolean

Const DATE_COLSIZE = 1260
Const TIME_COLSIZE = 1365
Const MPNAME_COLSIZE = 2805
Const CONC_COLSIZE = 1005
Const HLT_COLSIZE = 1320
Const NOTES_COLSIZE = 3570
 
Public Enum gridColumnNumber
    colDate = 0
    colTime = 1
    colMonitorPoint = 2
    colConcentration = 3
    colHLT = 4
    colNotes = 5
    colID = 6
End Enum

'===========================================================================================================
' end of the shitty form variable copying
'===========================================================================================================

Option Explicit


'*******************************************************
' DESCRIPTION: Sets the Modes based on the Date/Time that has been set
'               in the Start Intro, Start Expo and End Expo
'   INPUTS:     date & time the Intoduction Started
'               date & time the Exposure Started
'               date & time the Exposure Ended
'   Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Gamboa, Maria Cecilia <nai1686>
' Date:        September 25, 2002
'*******************************************************
'Change: rsMonitorValues As Recordset input parameter has been removed
Public Sub SetModesFromTimes(dtIntroStarted As Date, dtExposureStarted As Date, _
            dtExposureEnded As Date)

    Dim eMode As Integer
    Dim sUpdateSQL As String
    
    On Error GoTo Error_Handler
       
    sUpdateSQL = "UPDATE MonitorPointValue " _
                    & "SET Mode = " & MDM_PRE_INTRO & " " _
                    & "Where date_time < CDate('" & dtIntroStarted & "') and Deleted = False"
    
    'Execute the SQL
    gdbConn.BeginTrans
    gdbConn.Execute sUpdateSQL
    gdbConn.CommitTrans
    
    sUpdateSQL = "UPDATE MonitorPointValue " _
                    & "SET Mode = " & MDM_PRE_EXPOSURE & " " _
                    & "Where date_time >= Cdate('" & dtIntroStarted & "') and date_time < CDate('" & dtExposureStarted & "') and Deleted = False"
    
    'Execute the SQL
    gdbConn.BeginTrans
    gdbConn.Execute sUpdateSQL
    gdbConn.CommitTrans
    
    sUpdateSQL = "UPDATE MonitorPointValue " _
                    & "SET Mode = " & MDM_EXPOSURE & " " _
                    & "Where date_time >= Cdate('" & dtExposureStarted & "') and date_time < CDate('" & dtExposureEnded & "') and Deleted = False"
                    
    'Execute the SQL
    gdbConn.BeginTrans
    gdbConn.Execute sUpdateSQL
    gdbConn.CommitTrans
                    
    sUpdateSQL = "UPDATE MonitorPointValue " _
                    & "SET Mode = " & MDM_PRE_AERATION & " " _
                    & "Where date_time >= CDate('" & dtExposureEnded & "') and Deleted = False"
    
    'Execute the SQL
    gdbConn.BeginTrans
    gdbConn.Execute sUpdateSQL
    gdbConn.CommitTrans
    
    Exit Sub

Error_Handler:

    Call psAbort(MODULE_NAME, "SetModesFromTimes", Err.Description)

End Sub




'*******************************************************
' DESCRIPTION: Determines the Concentration pattern
'              Used in determining of the concentration has been decreastring
'   INPUTS:   rsConcPattern as recordset
'             Number of Phases enabled (Set this to 3 for the Monitor Status)
'   Output:     string
'
' REVISION HISTORY
' Revision:    0
' Author:      Gamboa, Maria Cecilia <nai1686>
' Date:        September 30, 2002
'
' Jett 04/20/2003 Commented out this function. This is never called
'*******************************************************
'Private Function GetPeakPattern(rsConcPattern As Recordset, dConcentration As Double) As String
'
'    Dim sPeak As String
'    ' We will work our way through the recordset
'    ' When we find the first drop in value then we know we have achieved the last peak
'
'    On Error GoTo Error_Handler
'
'     sPeak = ""
'     rsConcPattern.MoveFirst
'
'    ' Compare the Concentration Value and date entered with those that are in the recordset
'
'       While Not rsConcPattern.EOF
'
'            If rsConcPattern("date_time") <= CDate(frmMain.grdMonitorDataInput.Cell(flexcpText, frmMain.grdMonitorDataInput.Row, 1)) Then
'            ' check the concentration value
'                If rsConcPattern("fumiscope_conc") > dConcentration Then
'                        'The current concentration reading is less than a previous peak,
'                        'so the HLT can be calculated. Set sPeak to PP_DECREAstring
'                        sPeak = "PP_DECREASING"
'                End If
'                GetPeakPattern = sPeak
'                Exit Function
'            End If
'             rsConcPattern.MoveNext
'        Wend
'
'    Exit Function
'Error_Handler:
'    Call psAbort(MODULE_NAME, "GetPearPattern", Err.Description)
'
'End Function


'*******************************************************
' DESCRIPTION: Validation for the Concentration Time
'              **Intro start time cannot be after a concentration reading,
'              and a concentration reading cannot be before an Intro start.
'
'   INPUTS:   dtDate as Date set in the grid
'             LOzMCF as the Concentration Reading
'             eMode as the mode
'             bDelete
'             rsMonitorValues as recordset
'   Output:     Boolean
'
' REVISION HISTORY
' Revision:    0
' Author:      Gamboa, Maria Cecilia <nai1686>
' Date:        September 30, 2002
'*******************************************************
Private Function ValidateConcentrationTime(LOzMCF As Double, eMode As String) As Boolean

 Dim varBookMark As Variant
 Dim bRetval As Boolean
 Dim bEOF As Boolean

On Error GoTo Error_Handler

' Initialize the value of bRetVal and set it to TRUE
bRetval = True
    If (eMode = MDM_PRE_INTRO) And LOzMCF > 0 Then
                MsgBox LoadResString(IDS_MONDATA_MSG_INTRO_AFTER_CONCENTRATION), vbExclamation
                bRetval = False
                Exit Function
    End If

   ValidateConcentrationTime = bRetval

Exit Function
Error_Handler:
    Call psAbort(MODULE_NAME, "ValidateConcentrationTime", Err.Description)

End Function


'*******************************************************
' DESCRIPTION: Handles the valid data and calls the expert monitor methods that
'               will be used for the calculations

'   INPUTS:
'   Output:
'
' REVISION HISTORY
' Revision:    0
' Author:      Gamboa, Maria Cecilia <nai1686>
' Date:        September 30, 2002
' Revision:    1
' Author:      Martinez, Amana Karen G <nai1499>
' Date:        April 2, 2003
' Reason:      Fumiguide Enhancements 2003, Coding Phase
'*******************************************************
Public Sub CalculateMonDataInput(Optional sMonitorValueID As String, _
                             Optional sDateTime As Date, _
                             Optional sFumiConc As Double, _
                             Optional sInterConc As Double, _
                             Optional sNote As String, _
                             Optional eMode As Integer, _
                             Optional sMonPtGuid As String)

    Dim sMonPtValGuid As String
    Dim sMonPtIds As String
    Dim sSQL As String
    ' OPTIMIZE Dim rs As New ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim rsDateTime As ADODB.Recordset
       
    On Error GoTo Error_Handler
    
    frmMain.sbStatusBar.Panels(1).Text = "Calculating ..."
    
    ' performance test if debug flag is set
    #If SPEEDTEST = 1 Then
        PerfInit
        PerfStart
    #End If
    
    Set rs = New ADODB.Recordset
    Set rsDateTime = New ADODB.Recordset
    
    ' 11/15/2002 TSG - change mouse cursor to hourglass to indicate calculation in progress
    Screen.MousePointer = vbHourglass
    
    'Get the Monitor Point id of the row that is being calculated at
    'store it in the variable sEditingRecordID
    
    If sMonitorValueID <> "" Then
    
        Call UpdateMonValRecord(sMonitorValueID, _
                            sDateTime, _
                            sFumiConc, _
                            sInterConc, _
                            sNote, _
                            eMode, _
                            sMonPtGuid)
    End If
    
    sSQL = "Select startintro_datetime, startexp_datetime, endexp_datetime from Job where id = '" & gsJobID & "'"
        
    rsDateTime.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
        
    If Not (rsDateTime.EOF And rsDateTime.BOF) Then
        If IsNull(rsDateTime("startintro_datetime")) Or _
            IsNull(rsDateTime("startexp_datetime")) Or IsNull(rsDateTime("endexp_datetime")) Then
                 rsDateTime.Close
                Set rsDateTime = Nothing
                
                ' 11/15/02 TSG - change cursor back to normal
                Screen.MousePointer = vbDefault
                
                ' Jett 4/3/03 cleanup before we exit
                Set rs = Nothing
                Set rsDateTime = Nothing
                
                frmMain.sbStatusBar.Panels(1).Text = "OK"
                
                Exit Sub
        End If
    End If
    
    rsDateTime.Close
    Set rsDateTime = Nothing
    
    If Not NeedsCalculation() Then
    
        'Update the monitor point values, and the monitor point status. Create a comma-delimited string
        'AM 4/2/2003 - Added to following to limit the record set in getting the monitor point where monitoring point value was modified, For Fumiguide Enhancement 2003
        If sMonPtGuid <> "" Then
            sSQL = "Select distinct(MonitorPoint.id)  as id from MonitorPoint, MonitorPointValue where MonitorPoint.id = MonitorPointValue.monitorpoint_fkey and MonitorPointValue.deleted = FALSE and MonitorPoint.Id = '" & sMonPtGuid & "'"
        Else
           'AM 4/2/2003 - Added the codes above to limit the record set in getting the monitor point where monitoring point value was modified, For Fumiguide Enhancement 2003
           ' sSQL = "Select distinct(MonitorPoint.id)  as id from MonitorPoint, MonitorPointValue where MonitorPoint.id = MonitorPointValue.monitorpoint_fkey and MonitorPointValue.deleted = FALSE"
           
           ' Jett 3/12/04 - SIMS #2121768 fixed
           sSQL = "SELECT MonitorPoint.id FROM MonitorPoint WHERE id IN (SELECT distinct(monitorpoint_fkey) FROM MonitorPointValue WHERE deleted = FALSE) ORDER BY jobarea_fkey"

        End If
        
        'OPTIMIZE: rs.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
        rs.Open sSQL, gdbConn, adOpenForwardOnly, adLockReadOnly, adCmdText
        'Gather up all of the monitoring points and add them in the array
         
         If Not (rs.BOF And rs.EOF) Then
         rs.MoveFirst
         
            While Not rs.EOF
                If sMonPtIds = "" Then
                    'Add the Monitoring point to string
                    sMonPtIds = rs("id")
                Else
                    sMonPtIds = sMonPtIds & ", " & rs("id")
                End If
                rs.MoveNext
            Wend
        End If
        
        'Update the last time entry for the monitor point and pass the array as an argument
        Call UpdateMonitorPointLastTimeEntry(sMonPtIds)
        
        Call UpdateMonitorPointStatus(sMonPtIds)
        
        ' OPTIMIZE - (moved the code before the end if)
        rs.Close
        Set rs = Nothing
    
    End If
    
    ' 11/15/02 TSG - change cursor back to normal
    Screen.MousePointer = vbDefault
    
    frmMain.sbStatusBar.Panels(1).Text = "OK"
    
    ' update status bar with performance test results if debug flag is set
    #If SPEEDTEST = 1 Then
        PerfFinish
        frmMain.sbStatusBar.Panels(1).Text = "Last calculation took " & PerfElapsed & " ms"
    #End If
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "CalculateMonDataInput", Err.Description)
    
End Sub


Private Sub UpdateMonValRecord(sMonitorPointValueId As String, _
                                           dtDateTime As Date, _
                                           dfumiConc As Double, _
                                           dInterConc As Double, _
                                           sNote As String, _
                                           eMode As Integer, _
                                           sMonPtFKey As String)
    On Error GoTo Error_Handler
    
    ' Save the existing record
    sUpdateSQL = "UPDATE MonitorPointValue SET date_time = '" & dtDateTime & "', " _
                             & "fumiscope_conc = " & FormatNumber(dfumiConc, 2) & ", " _
                             & " interscan_conc = " & dInterConc _
                             & ", general_note = '" & sNote & "'" _
                             & ", mode = " & eMode _
                             & ", monitorpoint_fkey = '" & sMonPtFKey & "' " & _
                             " WHERE id = '" & sMonitorPointValueId & "'"
    
    gdbConn.BeginTrans
    gdbConn.Execute sUpdateSQL
    gdbConn.CommitTrans
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "UpdateMonValRecord", Err.Description)
    
End Sub


Private Sub UpdateMonitorPointStatus(asMonPtId As String)

    ' Declare the ff. Variables:  sMonPtId as string, sSql as string, nTotal as integer, nIndex as Integer, nAreaIndex
    ' Get the array size and store it in nTotal
    ' Call the expertMonitor class
    
    Dim ExpertMonitor As clsMonitorExpert
    Dim arrMonPt() As String
    Dim arrAreaId() As String
    Dim nIndex As Integer
    Dim bFoundArea As Boolean
    Dim nAreaIndex As Integer
    Dim sMonPtId As String
    Dim sAreaId As String
    Dim dtIntroStarted As Date
    Dim dtExposureStarted As Date
    Dim dtExposureEnded As Date
    Dim asAreaId As String
    Dim sSQL As String
    Dim sOldAreaID As String
    
    ' OPTIMIZE move from global to local this is the only place where it is used
    Dim rsMonPtVal As ADODB.Recordset
    Dim rsMonPt As ADODB.Recordset
    Dim rsDates As ADODB.Recordset
    
    On Error GoTo Error_Handler
            
    'OPTIMIZE
    Set rsMonPtVal = New ADODB.Recordset
    Set rsMonPt = New ADODB.Recordset
    Set rsDates = New ADODB.Recordset
    
    Set ExpertMonitor = New clsMonitorExpert
      
    sSQL = "Select startintro_datetime, startexp_datetime, endexp_datetime from Job where id = '" & gsJobID & "'"
    'OPTIMIZING rsDates.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    rsDates.Open sSQL, gdbConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDates.EOF And rsDates.BOF) Then
        dtIntroStarted = rsDates("startintro_datetime")
        dtExposureStarted = rsDates("startexp_datetime")
        dtExposureEnded = rsDates("endexp_datetime")
    End If
    
    rsDates.Close
    Set rsDates = Nothing
    
    
    sSQL = "SELECT MonitorPoint.id as MonPtID, MonitorPoint.name, MonitorPoint.description, " _
    & "MonitorPoint.job_fkey, MonitorPoint.status, MonitorPoint.hlt, MonitorPoint.last_entry, " _
    & "MonitorPoint.elapsed_time, MonitorPoint.jobarea_fkey, MonitorPoint.deleted, " _
    & "MonitorPoint.start_time, MonitorPoint.end_time, MonitorPoint.achieved_concentration_time, " _
    & "MonitorPoint.achieved_dose, MonitorPoint.projected_concentration_time, MonitorPoint.projected_time, " _
    & "MonitorPoint.future_target_concentration, MonitorPoint.future_concentration, " _
    & "MonitorPoint.order_entered, JobArea.id, JobArea.temperature, " _
    & "JobArea.estimated_hlt, JobArea.volume, JobArea.planned_exposure, JobArea.fumigant, " _
    & "JobArea.calculation_warning, JobArea.job_fkey, JobArea.hlt, JobArea.concentration_time, " _
    & "JobArea.initial_concentration, JobArea.dose, JobArea.achieved_concentration_time, " _
    & "JobArea.achieved_dose, JobArea.actual_hlt, JobArea.projected_concentration_time, " _
    & "JobArea.projected_time, JobArea.mean_ntc, JobArea.status, JobArea.add_profume, " _
    & "JobArea.order_entered, JobArea.area_name, JobArea.row, JobArea.deleted " _
    & "FROM MonitorPoint, JobArea WHERE MonitorPoint.deleted = FALSE " _
    & " AND JobArea.id = MonitorPoint.jobarea_fkey " _
    & " ORDER BY JobArea.id, MonitorPoint.id"
                 
    rsMonPt.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    arrMonPt() = Split(asMonPtId, ",")
        
    sOldAreaID = ""
        
    For nIndex = 0 To UBound(arrMonPt)
        sMonPtId = Trim$(arrMonPt(nIndex))
    
        ' Get the area id from the monitor point array in the doc
        If Not (rsMonPt.EOF And rsMonPt.BOF) Then
          rsMonPt.MoveFirst
         
          Do While Not rsMonPt.EOF
             
            If sMonPtId = rsMonPt("MonPtID") Then
                
                    ' We found our monitoring point, thus our area
                    bFoundArea = False

                    If InStr(1, asAreaId, rsMonPt("jobarea_fkey")) Then
                        bFoundArea = True
                    End If
                    
                    If (Not bFoundArea) Then
                        ' 3/18/04 Jett SIMS #2121768 we do not need to recalculate for the
                        ' area everytime it finds an MP belonging to that area so we just keep
                        ' track of the area it finds
                        'JETT If asAreaId = "" Then
                           asAreaId = rsMonPt("jobarea_fkey")
                        'JETT Else
                        'JETT    asAreaId = asAreaId & ", " & rsMonPt("jobarea_fkey")
                        'JETT End If
                
                    ' Exit Do
                    End If
                    
                    Exit Do
                    
                End If
            rsMonPt.MoveNext
        Loop
        End If
        
        
         ExpertMonitor.SetJobArea gsJobID, rsMonPt("jobarea_fkey")
                
         ' Need to iterate through all of the monitor point values associated with
         ' the monitor point and update their hlt value
                 
         sSQL = "SELECT MonitorPoint.id, MonitorPoint.name, MonitorPoint.description, " _
         & "MonitorPoint.job_fkey, MonitorPoint.status, MonitorPoint.hlt, MonitorPoint.last_entry, " _
         & "MonitorPoint.elapsed_time, MonitorPoint.jobarea_fkey, MonitorPoint.deleted, " _
         & "MonitorPoint.start_time, MonitorPoint.end_time, MonitorPoint.achieved_concentration_time, " _
         & "MonitorPoint.achieved_dose, MonitorPoint.projected_concentration_time, MonitorPoint.projected_time, " _
         & "MonitorPoint.future_target_concentration, MonitorPoint.future_concentration, MonitorPoint.order_entered, " _
         & "MonitorPointValue.id as MonPtValID, MonitorPointValue.date_time, MonitorPointValue.fumiscope_conc, " _
         & "MonitorPointValue.interscan_conc, MonitorPointValue.general_note, MonitorPointValue.mode, MonitorPointValue.hlt_pt2pt, " _
         & "MonitorPointValue.monitorpoint_fkey, MonitorPointValue.deleted FROM MonitorPointValue, MonitorPoint " _
         & "WHERE MonitorPointValue.deleted = FALSE AND MonitorPointValue.monitorpoint_fkey = '" & sMonPtId & "' AND " _
         & "MonitorPoint.id = MonitorPointValue.monitorpoint_fkey " _
         & "ORDER BY MonitorPointValue.date_time ASC"
        
         ' Execute the SQL and store the data in the rsMonPtVal recordset
         ' OPTIMIZE rsMonPtVal.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
         rsMonPtVal.Open sSQL, gdbConn, adOpenForwardOnly, adLockReadOnly, adCmdText
                
         If Not (rsMonPtVal.EOF And rsMonPtVal.BOF) Then
            rsMonPtVal.MoveFirst
                
            While (Not rsMonPtVal.EOF)
               ExpertMonitor.UpdateMonPtValuePt2PtHlt sMonPtId, _
                                 rsMonPtVal("MonPtValID"), _
                                 dtIntroStarted, _
                                 dtExposureStarted, _
                                 dtExposureEnded
                
               rsMonPtVal.MoveNext
            Wend
         End If
                 
        rsMonPtVal.Close
         
         ' Alright, we have updated all of the monitor point values,now we have to
         ' update the monitor point itself
         
         ExpertMonitor.UpdateMonitorPointStatus sMonPtId, dtIntroStarted, dtExposureStarted, dtExposureEnded

        ' calculate our job area only once per monitoring point
        ' only get called once for single-MP areas MP and multi-MP areas
        If sOldAreaID <> asAreaId Then
            ExpertMonitor.UpdateJobAreaStatus gsJobID, asAreaId, dtIntroStarted, dtExposureStarted, dtExposureEnded
            sOldAreaID = asAreaId
        End If
     Next
     
    ' Update the jobarea status
    ' 5/15/2003 Jett - added the dates so we can use them in the special processing for
    ' areas with multiple monitoring points
    'JETT ExpertMonitor.UpdateJobAreaStatus gsJobID, asAreaId, dtIntroStarted, dtExposureStarted, dtExposureEnded
    
    ' Updating records straight at the db invalidates the recordset
    rsMonPt.Close
    'rsMonPtVal.Close
    Set rsMonPt = Nothing
    Set rsMonPtVal = Nothing
    
    ' OPTIMIZE
    Set ExpertMonitor = Nothing

Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "UpdateMonitorPointStatus", Err.Description)
    
End Sub


Private Sub UpdateMonitorPointLastTimeEntry(asMonPtId As String)
    ' Define the following variables
    Dim sMonPtId As String
    Dim sSQL As String
    Dim sUpdateSQL As String
    Dim dtLastEntry As Date
    Dim nTotal As Integer
    Dim nIndex As Integer
    Dim arrMonPt() As String
    Dim rsMonPtVal As ADODB.Recordset
    
    On Error GoTo Error_Handler
    
    Set rsMonPtVal = New ADODB.Recordset
    
    ' Get the size of the array and store it in the variable nTotal
    ' Get the monitor point values for this monitor point
    ' Iterate through the values in the array
    arrMonPt() = Split(asMonPtId, ",")
    
       For nIndex = 0 To UBound(arrMonPt)
       
          sMonPtId = Trim$(arrMonPt(nIndex))
    
            'JNM 04/26/03 - AT Fix - Consider interscan_conc gt 0 only
            sSQL = "SELECT * FROM MonitorPointValue, MonitorPoint " _
                & "WHERE MonitorPointValue.deleted = FALSE AND " _
                & "MonitorPointValue.monitorpoint_fkey = '" & sMonPtId & "' " _
                & "AND MonitorPoint.id = MonitorPointValue.monitorpoint_fkey " _
                & "AND (MonitorPointValue.fumiscope_conc >= 0 OR " _
                & "MonitorPointValue.interscan_conc > 0) " _
                & "ORDER BY MonitorPointValue.date_time DESC"
    
            dtLastEntry = 0#
    
    ' OPTIMIZE we do not define new from the start so we close this
    'If rsMonPtVal.State Then
    '    rsMonPtVal.Close
    'End If
    
        'Execute the SQL and store the data in rsMonPtVal recordset
        rsMonPtVal.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
        If (rsMonPtVal.BOF And rsMonPtVal.EOF) Then
            dtLastEntry = 0#
        Else
            'Since we are in descending order, the last concentration entry will be first
            rsMonPtVal.MoveFirst
            dtLastEntry = rsMonPtVal("date_time")
        End If
    
    rsMonPtVal.Close
    'Now update the monitor point
     sUpdateSQL = "UPDATE MonitorPoint SET last_entry = '" & dtLastEntry & "' WHERE id = '" & sMonPtId & "'"
    'Execute the SQL to update the records
    
    gdbConn.BeginTrans
    gdbConn.Execute sUpdateSQL
    gdbConn.CommitTrans
                                                         
    Next
    
    Set rsMonPtVal = Nothing
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "UpdateMonitorPointLastTimeEntry", Err.Description)
    
End Sub


Private Function GetCurrentTime() As String

    Const FUNCTION_NAME = "GetCurrentTime()"
    Dim sCurrentTime As String

On Error GoTo Error_Handler

    ' Get the current time but do not get the seconds field so it defaults
    ' to 0
    sCurrentTime = Hour(Now()) & ":" & Minute(Now())

    GetCurrentTime = sCurrentTime
    
Exit Function
Error_Handler:
    Call psAbort(MODULE_NAME, FUNCTION_NAME, Err.Description)
        
End Function







'========================================================================================================
' -------------------- Handlers moved from the form ---------------
'========================================================================================================














   




'**************************************************************
'   DESCRIPTION
'   This subroutine creates the pipe-delimited string which contains the
'   list of monitor point names (can be found in the Monitor Point name
'   column in the grid) and populates the Monitor Point Name list box
'
'   Input Parameters: (None)
'
'   RETURNS - (None)
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Private Sub CreateMonitorPtList()

    On Error GoTo Error_Handler

    ' Set the MultiSelect property of the Listbox to 2 (Extended). In extended selection, you can also use the Shift key to select ranges of items.
    ' Define the variables to be used, sMonPt as String
    Dim sMonitorPt As String
    Dim sMonitorPtId As String
    
    sMonitorPt = ""
    sPipeList = ""
    sMonitorPtId = ""
    
    ' Jett 4/18/03: clear everything first before repopulating
    frmMain.lstDataFilter.Clear
    
    If Not (rsMonitorPoint.EOF And rsMonitorPoint.BOF) Then
    
        rsMonitorPoint.MoveFirst
        sDefaultMonPtID = rsMonitorPoint("id")
        sDefaultMonPt = rsMonitorPoint("name")
        frmMain.lstDataFilter.AddItem LoadResString(IDS_MONPOINTS_FILTER_ALL)
        
        While Not rsMonitorPoint.EOF
        
            If sPipeList = "" Then
                sPipeList = rsMonitorPoint("name") & "|"
            Else
                sPipeList = sPipeList & rsMonitorPoint("name") & "|"
            End If
            
            frmMain.lstDataFilter.AddItem rsMonitorPoint("name")
            rsMonitorPoint.MoveNext
            
        Wend
        
    End If

    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "CreateMonitorPtList", Err.Description)

End Sub


'**************************************************************
'   DESCRIPTION
'   This subroutine retrieves the start intro, start exposure and end exposure
'   in the Job table and displays these values (if there are any)
'
'   Input Parameters: (None)
'
'   RETURNS - (None)
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
' 11/26/2002       Jett Gamboa      Defaulted Start Exposure to Start Intro + 1 hour and
'                                   End Exposure to Start Exposure + 24 hours
'  5/27/2003       Jett Gamboa      Defaulted End Exposure to Start Exposure + Planned Exposure time
'**************************************************************
Private Sub DisplayDateTime()

    Dim rsDateTime As ADODB.Recordset
    Dim nTimeToAdd As Integer
    
    Set rsDateTime = New ADODB.Recordset
On Error GoTo Error_Handler

    'Retrieve the Date & Time from the Job Table
    bIsDatesEmpty = False
    sSQL = "Select startintro_datetime, startexp_datetime, endexp_datetime from Job where id = '" & gsJobID & "'"
    
    rsDateTime.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    With frmMain
        If Not (rsDateTime.EOF And rsDateTime.BOF) Then
            If IsNull(rsDateTime("startintro_datetime")) Then
                .dtpStartIntroDate.Value = Now()
                .dtpStartIntroTime.Value = GetCurrentTime()
                nEnabled = 3
            Else
                .dtpStartIntroDate.Value = rsDateTime("startintro_datetime")
                .dtpStartIntroTime.Value = rsDateTime("startintro_datetime")
            End If
            
             If IsNull(rsDateTime("startexp_datetime")) Then
                 .dtpStartExposureDate.Value = DateAdd("h", 1, Now())
                 .dtpStartExposureTime.Value = DateAdd("h", 1, GetCurrentTime())
                 'JNM 05/09/03 - PT fix - This is now enabled since we already have
                 '  the button.
'                 .dtpStartExposureDate.Enabled = False
'                 .dtpStartExposureTime.Enabled = False
                 nEnabled = 2
            Else
                .dtpStartExposureDate.Value = rsDateTime("startexp_datetime")
                .dtpStartExposureTime.Value = rsDateTime("startexp_datetime")
            End If
            
            If IsNull(rsDateTime("endexp_datetime")) Then
                ' Jett 5/27/03 - if there was no end of exposure date/time yet,
                ' default it to Start Exposure + Planned Exposure Time (we are doing a
                ' + 1 because default Start Exposure is Start Intro + 1 hour)
                
                ' We add the number of minutes this time
                nTimeToAdd = (gnExposureTime * 60) + 60
                
                .dtpEndExposureDate.Value = DateAdd("n", nTimeToAdd, Now())
                .dtpEndExposureTime.Value = DateAdd("n", nTimeToAdd, GetCurrentTime())
                'JNM 05/09/03 - PT fix - This is now enabled since we already have
                 '  the button.
'                .dtpEndExposureDate.Enabled = False
'                .dtpEndExposureTime.Enabled = False
                nEnabled = 1
            Else
                .dtpEndExposureDate.Value = rsDateTime("endexp_datetime")
                .dtpEndExposureTime.Value = rsDateTime("endexp_datetime")
            End If
         
        End If
    
    End With
    
    If IsNull(rsDateTime("startintro_datetime")) Or _
    IsNull(rsDateTime("startexp_datetime")) Or IsNull(rsDateTime("endexp_datetime")) Then
        bIsDatesEmpty = True
    End If
    
    rsDateTime.Close
    Set rsDateTime = Nothing

    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "DisplayDateTime", Err.Description)

End Sub


'**************************************************************
'   DESCRIPTION
'   Retrieves and displays the Monitor Data Input grid data.
'
'   Input Parameters: (None)
'
'   RETURNS - (None)
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'  4/2/2003        Amana Martinez   Fumiguide Enhancements 2003
'**************************************************************
Private Sub DisplayGridData()
    Dim sSQL As String
    Dim intRecCount As Integer
    Dim rsGridData As New ADODB.Recordset

On Error GoTo Error_Handler
    
    bIsTableEmpty = False
    
    I = 1
    dValue = 0#
    
    With frmMain
    
        ' 05/16/2003 Jett - we do not need this when we load. Only other instance
        ' it is called is when we click an item in the filter (but we clean out the
        ' grid already
        'Clear the grid of its contents
        '.grdMonitorDataInput.Cell(flexcpData, 1, 0, GRIDSIZE - 1, 6) = ""
        '.grdMonitorDataInput.Cell(flexcpText, 1, 0, GRIDSIZE - 1, 6) = ""
        
        'JNM 11/15 - Changed the sorting of the records on the grid
        'sSQL = "SELECT MonitorPointValue.id as Id, MonitorPointValue.date_time, " _
                & "MonitorPointValue.fumiscope_conc, MonitorPointValue.interscan_conc, " _
                & "MonitorPointValue.general_note, mode, monitorpoint_fkey, hlt_pt2pt " _
                & "FROM MonitorPointValue " _
                & "where MonitorPointValue.deleted = false " _
                & "ORDER by MonitorPointValue.date_time DESC"
        sSQL = "SELECT MonitorPointValue.id as Id, MonitorPointValue.date_time, name, " _
                & "MonitorPointValue.fumiscope_conc, MonitorPointValue.interscan_conc, " _
                & "MonitorPointValue.general_note, mode, monitorpoint_fkey, hlt_pt2pt " _
                & "FROM MonitorPointValue, MonitorPoint " _
                & "where MonitorPointValue.deleted = false and monitorpoint_fkey = MonitorPoint.id " _
                & "ORDER by name, date_time ASC"
    
        ' Execute the SQL and store the data in rsGridData
        rsGridData.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
        ' Make sure that the recordset is not empty
        If Not (rsGridData.EOF And rsGridData.BOF) Then
        
            ' Move to the first record in the recordset
            rsGridData.MoveFirst
            
            While Not (rsGridData.EOF)
                'i = rsGridData("row")
                .grdMonitorDataInput.Cell(flexcpText, I, 0) = IIf(IsNull(rsGridData("date_time")), "", rsGridData("date_time"))
                .grdMonitorDataInput.Cell(flexcpText, I, 1) = IIf(IsNull(rsGridData("date_time")), "", rsGridData("date_time"))
                .grdMonitorDataInput.Cell(flexcpText, I, 2) = DisplayMonitorName(IIf(IsNull(rsGridData("monitorpoint_fkey")), "", rsGridData("monitorpoint_fkey")))
                .grdMonitorDataInput.Cell(flexcpData, I, 3) = CStr(IIf(IfLessZero(rsGridData("fumiscope_conc")), "", rsGridData("fumiscope_conc")))
                .grdMonitorDataInput.Cell(flexcpText, I, 3) = IIf(IfLessZero(rsGridData("fumiscope_conc")), "", rsGridData("fumiscope_conc"))           '& " " & sConcUnits) AM 4/2/03 - Commented this for Fumiguide Enhancements 2003
                'JNM 11/13 - Changed decimal from nDecimal to 1
                'dValue = Round(rsGridData("hlt_pt2pt"), 1)
                'Pat 2/4/04 - Used Roundoff function to have consistency in rounding off numbers
                '              SIMS Ticket 2116768
                dValue = RoundOff(rsGridData("hlt_pt2pt"), 1)
                .grdMonitorDataInput.Cell(flexcpText, I, 4) = IIf(IfZero(CStr(dValue)), "", dValue)
                .grdMonitorDataInput.Cell(flexcpText, I, 5) = IIf(IsNull(rsGridData("general_note")), "", rsGridData("general_note"))
'                If sHighlight = rsGridData("id") Then
'                   .grdMonitorDataInput.Select I, 0, I, 5
'                   'JNM 11/15 - show the highlighted cell
'                   .grdMonitorDataInput.ShowCell I, 1
'                End If
                .grdMonitorDataInput.Cell(flexcpText, I, 6) = rsGridData("id")
                I = I + 1
                ' Move to next record in the recordset
                    rsGridData.MoveNext
            Wend
            
            ' 11/26/02 TSG - scroll to the row after the last one with value after filtering
            .grdMonitorDataInput.TopRow = IIf(I - 21 > 0, I - 21, 1)
            .grdMonitorDataInput.Select I, 1
            
        Else
        
            bIsTableEmpty = True
            
        End If
        
    End With
    
    rsGridData.Close
    Set rsGridData = Nothing

    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "DisplayGridData", Err.Description)

End Sub


'**************************************************************
'   DESCRIPTION
'   Retrieves and displays the Monitor Point Name for the current record
'
'   Input Parameters: String (Monitor Point Id)
'
'   RETURNS - String (Monitor Point Name)
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Private Function DisplayMonitorName(sMonPtNameId As String) As String

    Dim rs As ADODB.Recordset

    On Error GoTo Error_Handler
    
    Set rs = New ADODB.Recordset

    If sMonPtNameId = "" Or IsNull(sMonPtNameId) Then
        DisplayMonitorName = ""
    Else
       sSQL = "Select name from MonitorPoint where id = '" & sMonPtNameId & "' and deleted = False"
       rs.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
       If rs.RecordCount = 0 Then
            DisplayMonitorName = ""
        Else
            rs.MoveFirst
            DisplayMonitorName = rs("Name")
        End If
        rs.Close
    End If

    Set rs = Nothing

Exit Function
Error_Handler:
    Call psAbort(MODULE_NAME, "DisplayMonitorName", Err.Description)

End Function


'**************************************************************
'   DESCRIPTION
'   Checks if the data is less than or equal to 0.
'
'   Input Parameters: String
'
'   RETURNS - (Boolean) Returns True/False.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Private Function IfZero(sData) As Boolean

On Error GoTo Error_Handler

    IfZero = False
    
    If sData <= 0 Or IsNull(sData) Then
        IfZero = True
    End If

Exit Function
Error_Handler:
    Call psAbort(MODULE_NAME, "IfZero", Err.Description)

End Function


'**************************************************************
'   DESCRIPTION
'   Checks if the data is less than zero.
'
'   Input Parameters: String
'
'   RETURNS - (Boolean) Returns True/False.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Private Function IfLessZero(sData) As Boolean

On Error GoTo Error_Handler

    IfLessZero = False
    
    If sData < 0 Or IsNull(sData) Then
        IfLessZero = True
    End If


Exit Function
Error_Handler:
    Call psAbort(MODULE_NAME, "IfLessZero", Err.Description)

End Function


'**************************************************************
'   DESCRIPTION
'   Handles the deletion of the selected record. After the deletion,
'   calls the function responsible for the re-calculation
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Private Sub OnRowDelete()
    'Define the variables to be used
    Dim sDelMsg As String
    Dim varBookMark As Variant
    Dim asMonPtId As String
    Dim asRows As String
    Dim intDelCount As Integer
    Dim lSelected As Long
    Dim lFirstRow As Long
    Dim lFirstCol As Long
    Dim lLastRow As Long
    Dim lLastCol As Long
    Dim iCtr As Integer
    Dim bAnyDeleted As Boolean
    Dim sCtr As String
    Dim iNewRow As Integer
    Dim rsMonId As ADODB.Recordset
    Dim rsDateTime As ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo Error_Handler
    
    Set rsDateTime = New ADODB.Recordset
    
    iCtr = 0
    lSelected = 0
    asRows = ""
    
    ' Attempt deletion of all selected items in the list view
    ' Retrieve the contents of the data that has been selected for deletion
    lFirstRow = frmMain.grdMonitorDataInput.Row
    lLastRow = frmMain.grdMonitorDataInput.RowSel
    
    If lFirstRow > lLastRow Then
        'User selected rows from bottom - up. Switch them.
        lSelected = lLastRow
        lLastRow = lFirstRow
    Else
        lSelected = lFirstRow
    End If
    
    If frmMain.grdMonitorDataInput.Cell(flexcpText, lSelected, 6) <> "" Then
            If (lSelected > 0) Then
            ' Iterate through the list of selected items for deletion
                While Not lSelected > lLastRow
                    'Create an array of row numbers to be deleted
                    If asRows = "" Then
                        asRows = "'" & frmMain.grdMonitorDataInput.Cell(flexcpText, lSelected, 6) & "'"
                        asMonPtId = "'" & frmMain.grdMonitorDataInput.Cell(flexcpText, lSelected, gridColumnNumber.colMonitorPoint) & "'"
                    Else
                        asRows = asRows & ", " & "'" & frmMain.grdMonitorDataInput.Cell(flexcpText, lSelected, 6) & "'"
                        asMonPtId = asMonPtId & ", " & "'" & frmMain.grdMonitorDataInput.Cell(flexcpText, lSelected, gridColumnNumber.colMonitorPoint) & "'"
                    End If
                lSelected = lSelected + 1
                iCtr = iCtr + 1
                Wend
            
                bAnyDeleted = False
                sCtr = iCtr & " " & LoadResString(IDS_ROW)
                sDelMsg = FormatSWParams(LoadResString(IDS_MSG_DELETE_GENERIC), sCtr & "|" & LoadResString(IDS_ROW))
                ' Prompt user before deletion
                If (MsgBox(sDelMsg, vbOKCancel) = vbOK) Then
                                                 
                        bAnyDeleted = True
                    
                        ' Delete (soft) from DB
                        sUpdateSQL = "UPDATE MonitorPointValue SET deleted = TRUE " _
                                      & "Where id in (" & asRows & ")"
                        
                        
                        gdbConn.Execute sUpdateSQL
                        
                End If
            
            ' Don 't need to refresh anything if nothing was deleted
            If (bAnyDeleted) Then
               
            ' Jett 4/20/03: removed this line, only used because we were doing a delete when another
            ' recordset was open using the same table.
            ' Re-Calculate the remaining data
            'If rsMonitorValues.State Then
            '    rsMonitorValues.Close
            'End If

            Set rsMonId = New ADODB.Recordset
            
            sUpdateSQL = "Select distinct(id) from MonitorPoint where name in (" & asMonPtId & ")"
            rsMonId.Open sUpdateSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
            asMonPtId = ""
            If Not (rsMonId.BOF And rsMonId.EOF) Then
                rsMonId.MoveFirst
                While Not rsMonId.EOF
                    If asMonPtId = "" Then
                        asMonPtId = rsMonId("id")
                    Else
                        asMonPtId = asMonPtId & ", " & rsMonId("id")
                    End If
                    rsMonId.MoveNext
                Wend
                
            End If
            
            sSQL = "Select startintro_datetime, startexp_datetime, endexp_datetime from Job where id = '" & gsJobID & "'"
    
            rsDateTime.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
            If Not (rsDateTime.EOF And rsDateTime.BOF) Then
                If IsNull(rsDateTime("startintro_datetime")) Or _
                    IsNull(rsDateTime("startexp_datetime")) Or IsNull(rsDateTime("endexp_datetime")) Then
                        rsDateTime.Close
                        Set rsDateTime = Nothing
                        Call lstDataFilter_Click_Handler
                        gbHasChanged = True
                        
                        Set rsDateTime = Nothing
                        Exit Sub
                End If
            End If

            rsDateTime.Close
            Set rsDateTime = Nothing
            
             'Update the last time entry for the monitor points
              UpdateMonitorPointLastTimeEntry (asMonPtId)
            
             'Update the monitor point values, and the monitor point status
              UpdateMonitorPointStatus asMonPtId
              
              CalculateMonDataInput
            
            ' Repopulate the recordsets and make sure the filter is still retained
             Call lstDataFilter_Click_Handler
             gbHasChanged = True
             
            End If
        End If
    Else
    MsgBox LoadResString(IDS_MONDATA_MSG_NO_ITEMS_SELECTED), vbExclamation
    End If
        
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "OnRowDelete", Err.Description)
     
End Sub


'**************************************************************
'   DESCRIPTION
'   Validates the concentration reading that has been added and checks
'   if all necessary data has been provided.
'
'   Input Parameters: Monitor Point Id as String
'                     Concentration Reading as String
'                     sNotes As String
'                     Monitor Point Value Id as String
'
'   RETURNS - (Boolean) Returns True/False.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Private Function ValidateRecordAdd(sMonPt As String, sConReading As String, _
            sNotes As String) As Boolean

    Dim lfumiConc As Double
    Dim sFumiConc As String
    Dim dPpm As Double
    Dim dInterConc As Double
    Dim lPpm As Long
    Dim sPeakPattern As String
    
On Error GoTo Error_Handler

    ' Validate field entries and write to recordset, creating a new record if necessary
    ' Fumiscope Concentration
    If (Len(Trim$(sConReading)) <= 0) Then
        MsgBox LoadResString(IDS_MONDATA_MSG_CONCENTRATION_NUMERIC), vbExclamation
        'Concentration must be an integer between 0 and 500.
        ValidateRecordAdd = False
        Exit Function
    End If
    
    
    If (Len(Trim$(sConReading)) > 0 And Not IsNumeric(sConReading)) Then
        MsgBox LoadResString(IDS_MONDATA_MSG_CONCENTRATION_NUMERIC), vbExclamation
        'Concentration must be an integer between 0 and 500.
        ValidateRecordAdd = False
        Exit Function
    End If
    
    If Len(Trim$(sConReading)) > 0 Then
        
'        If (Len(sConReading) > 0 And Int(CDbl(sConReading)) <> CDbl(sConReading)) Then
'            MsgBox LoadResString(IDS_MONDATA_MSG_CONCENTRATION_NUMERIC), vbExclamation
'            ValidateRecordAdd = False
'            Exit Function
'        End If
        
        lfumiConc = CDbl(sConReading)
        
        If lfumiConc < 0 Then
            MsgBox LoadResString(IDS_MONDATA_MSG_CONCENTRATION_NUMERIC), vbExclamation
            ValidateRecordAdd = False
            Exit Function
        End If
    
        If lfumiConc > 500 Then
           MsgBox LoadResString(IDS_MONDATA_MSG_CONCENTRATION_500_LESS), vbExclamation
            'Concentration cannot be greater than 500
             'TO DO: Set the focus to the Concentration column
            ValidateRecordAdd = False
            Exit Function
        End If
        
        If lfumiConc > 128 Then
               MsgBox LoadResString(IDS_MONDATA_MSG_CONCENTRATION_OVER_128), vbExclamation
               'Warning: Concentration exceeding 128
        End If
           ValidateRecordAdd = True
    Else
        ' set the value of sOzMCF = 0 and sFumiConc = -1
        sFumiConc = -1
        lfumiConc = 0
        ValidateRecordAdd = False
    End If
     
Exit Function
Error_Handler:
    Call psAbort(MODULE_NAME, "ValidateRecordAdd", Err.Description)

End Function


'**************************************************************
'   DESCRIPTION
'   This subroutine calls the grid wrapper and build the grid.
'
'   Input Parameters: None
'
'   RETURNS: None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'  4/2/2003        Amana Martinez   Fumiguide Enhancements 2003
'**************************************************************
Private Sub PaintGrid()

    With frmMain.grdMonitorDataInput
    
        'call the grid wrapper
        thisGrid.Initialize MODULE_NAME, frmMain.grdMonitorDataInput
        
        'Date Column
        thisGrid.AddColumn "colDate", LoadResString(IDS_MONDATA_DATE_COL), colTypeDate, flexAlignGeneral
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        
        'Time Column
        thisGrid.AddColumn "colTime", LoadResString(IDS_MONDATA_DATE_TIME), colTypeTime, flexAlignGeneral
        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        
        'Monitor Point Name Column
        thisGrid.AddColumn "colMonitorPointName", LoadResString(IDS_MONDATA_MON_PT_NAME_COL), colTypeCombo, flexAlignGeneral
        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
        
        'Concentration Column
        sConcUnits = IIf(gbIsMetric, LoadResString(IDS_CONCENTRATION_METRIC), LoadResString(IDS_CONCENTRATION_ENGLISH))
        thisGrid.AddColumn "colConcentratrion", LoadResString(IDS_CONCENTRATION_COL) & vbCr & sConcUnits, colTypeText, flexAlignRightCenter 'flexAlignGeneral AM 4/2/2003 - Commented this and Added flexAlignRightCenter for Fumiguide Enhancements 2003
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        .AutoSize 0, 3
        .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
        
        'HLT Column
        thisGrid.AddColumn "colHlt", LoadResString(IDS_HLT_COL) & vbCr & LoadResString(IDS_HLT_COL2), colTypeText, flexAlignRightCenter 'flexAlignGeneral AM 4/2/2003 - Commented this and Added flexAlignRightCenter for Fumiguide Enhancements 2003
        .AutoSize 0, 4
        .Cell(flexcpAlignment, 0, 4) = flexAlignCenterCenter
        
        'Note Column
        thisGrid.AddColumn "colNotes", LoadResString(IDS_NOTE), colTypeText, flexAlignGeneral
        .Cell(flexcpAlignment, 0, 5) = flexAlignCenterCenter
        
        ' Hidden Column
        ' This column is used to store the Monitor Value id of the record in the recordset.
        thisGrid.AddColumn "colId", " ", colTypeText, flexAlignLeftTop
        .ColHidden(gridColumnNumber.colID) = True
                
        .ColDataType(0) = flexDTDate
        .ColDataType(1) = flexDTDate
        '.ColDataType(3) = flexDTDouble
        '.ColFormat(0) = "Medium Date"
        '.ColFormat(1) = "Medium Time"
        .ColComboList(0) = "Dummy" ' just to show the down-arrow
        .ColComboList(1) = "Dummy" ' just to show the down-arrow
            
        ' put default column size (will be overriden if settings were previously saved)
        .ColWidth(0) = DATE_COLSIZE
        .ColWidth(1) = TIME_COLSIZE
        .ColWidth(2) = MPNAME_COLSIZE
        .ColWidth(3) = CONC_COLSIZE
        .ColWidth(4) = HLT_COLSIZE
        .ColWidth(5) = NOTES_COLSIZE
        
        thisGrid.LoadGridSettings GRID_SETTINGS_PATH
        
        .Rows = GRIDSIZE
        .GridColor = GRID_COLOR
        .SelectionMode = flexSelectionByRow
        .AllowSelection = True
        .FocusRect = flexFocusNone
        
        For nRow = 1 To (GRIDSIZE - 1)
            .Cell(flexcpData, nRow, 0) = -1
        Next
        
        'For nRow = 1 To (GRIDSIZE - 1) Step 2
        '.Cell(flexcpBackColor, nRow, 4) = READ_ONLY_COLOR
        'Next
        
        .Cell(flexcpBackColor, 1, 4, GRIDSIZE - 1, 4) = READ_ONLY_COLOR
    
    End With
End Sub


'**************************************************************
'   DESCRIPTION
'   This subroutine saves the date/time entered in the grid.
'
'   Input Parameters: String (specify if the change is made in the SDATE or
'                     STIME column
'
'   RETURNS - (Boolean) Returns True/False.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'  3/15/2003       Jett Gamboa      BUGFIX: changed calls to TimeValue(Now()) to GetCurrentTime()
'  5/28/2003       Malou Espaldon   ENHANCEMENT: Passed parameters to CalculateMonDataInput so that
'                                   recalculation will only be done for the row where date/time
'                                   was changed and not for all the monitor points within the area.
'**************************************************************
Private Sub SaveDateTime(sData As String)

    Dim iRow As Long
    Dim rs As ADODB.Recordset
    Dim rsMonitorID As ADODB.Recordset
    Dim sMonitorPoint As String
    Dim sgrdDate As String
    Dim sgrdTime As String
    Dim dtGrid As Date
    
    On Error GoTo Error_Handler

    Set rs = New ADODB.Recordset
    Set rsMonitorID = New ADODB.Recordset
    'JNM 04/23/03 - AT Fix - Get the current row number from the heading
    'iRow = frmMain.grdMonitorDataInput.Row
    iRow = frmMain.grdMonitorDataInput.Cell(flexcpData, 0, 0)
    
    If frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 6) = "" Then
    
        If sData = sDate Then
                sGridDateTime = DateValue(frmMain.dtpDummyDate.Value) & " " & GetCurrentTime()
                frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 1) = GetCurrentTime()
        Else
                sGridDateTime = DateValue(Now) & " " & TimeValue(frmMain.dtpDummyTime.Value)
                frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 0) = DateValue(Now)
        End If

        'if the row has never been updated or is empty
        frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 6) = GetGUID()
        sSQL = "Insert into MonitorPointValue (id, date_time) VALUES ('" & frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 6) & "', '" & sGridDateTime & "')"
    
    Else
        
        sSQL = "Select date_time from MonitorPointValue where id = '" & frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 6) & "'"
        rs.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
        If rs.RecordCount = 0 Then
            If sData = sDate Then
                sGridDateTime = DateValue(frmMain.dtpDummyDate.Value) & " " & GetCurrentTime()
            Else
                sGridDateTime = DateValue(Now) & " " & TimeValue(frmMain.dtpDummyTime.Value)
            End If
        Else
        
            If sData = sDate Then
                sGridDateTime = DateValue(frmMain.dtpDummyDate.Value) & " " & TimeValue(rs("date_time"))
            Else
                sGridDateTime = DateValue(rs("date_time")) & " " & TimeValue(frmMain.dtpDummyTime.Value)
            End If
            
        End If
        sSQL = "Update MonitorPointValue Set date_time = '" & sGridDateTime & "'" _
               & "where id = '" & frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 6) & "'"
        gbHasChanged = True
        rs.Close
    End If
    
    
    gdbConn.Execute sSQL
    
    'Call the SetMode to determine the mode of the data that has been provided
    SetMode frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 6), CDate(sGridDateTime)
    
   
    If Trim$(frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 3)) <> "" And Trim$(frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 3)) <> "0" Then
        sHighlight = frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 6)
    '--Begin mme 5/29/03 passed parameters to CalculateMonDataInput.  Originally, the said
    '--subroutine is being called in this function with no parameters passed.
        '--Get other values
        sMonitorPoint = frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 2)
        sSQL = "Select id from MonitorPoint where name = '" & sMonitorPoint & "' and deleted = FALSE"
        rsMonitorID.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
        dtGrid = DateValue(sGridDateTime) & " " & TimeValue(sGridDateTime)
        CalculateMonDataInput sHighlight, dtGrid, _
            CDbl(frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 3)), INTERSCAN, _
            frmMain.grdMonitorDataInput.Cell(flexcpText, iRow, 5), mode, _
            rsMonitorID("id")
        Call lstDataFilter_Click_Handler
        rsMonitorID.Close
    '--End mme 5/29/03
    End If
    
    Set rs = Nothing

    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "SaveDateTime", Err.Description)

End Sub


'**************************************************************
'   DESCRIPTION
'   Using the date and time provided, compare the data provided in
'   the Data and Time fields for Start Intro/Start Exposure/End Exposure and
'   compare it with the date and time entered in the grid
'
'   Input Parameters: Monitor Point Valie Id as String
'                     Date/Time in the Grid as Date
'
'   RETURNS - (None)
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Private Sub SetMode(sMonPtValId As String, sGridDateTime As Date)

    Dim rsDateTime As ADODB.Recordset
    
    On Error GoTo Error_Handler
    
    Set rsDateTime = New ADODB.Recordset

    sSQL = "Select startintro_datetime, startexp_datetime, endexp_datetime from Job where id = '" & gsJobID & "'"
    rsDateTime.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Not (rsDateTime.EOF And rsDateTime.BOF) Then
    
        If IsNull(rsDateTime("startintro_datetime")) Or _
        IsNull(rsDateTime("startexp_datetime")) Or IsNull(rsDateTime("endexp_datetime")) Then
            rsDateTime.Close
            Set rsDateTime = Nothing
            Exit Sub
        End If
        
        dtIntroStarted = rsDateTime("startintro_datetime")
        dtExposureStarted = rsDateTime("startexp_datetime")
        dtExposureEnded = rsDateTime("endexp_datetime")
        
            'Check if we are adding before Start Introduction
            ' - - Set the mode for the new record to 1 (Pre-Intro)
            If sGridDateTime < dtIntroStarted Then
                mode = MDM_PRE_INTRO
            End If
         
        'Check if we are adding before Start Introduction and before Start Exposure
        ' - - Set the mode for the new record to 2 (Pre-Exposure)
         If (sGridDateTime >= dtIntroStarted And sGridDateTime < dtExposureStarted) Then
            mode = MDM_PRE_EXPOSURE
         End If
                         
        ' Check if we are adding after Start Exposure and before End Introduction
        ' - -  Set the mode for the new record to 3 (Exposure)
        
         If (sGridDateTime >= dtExposureStarted And sGridDateTime < dtExposureEnded) Then
              mode = MDM_EXPOSURE
         End If
                
                
         ' Check if we are adding after End Exposure
        ' - -  Set the mode for the new record to 4 (Pre-Aeration)
               
         If (sGridDateTime >= dtExposureEnded) Then
            mode = MDM_PRE_AERATION
         End If
                    
        ' TO DO: checking if there is already an existing record..then update.
        ' Else, Insert a new record, Call GetGuid
        
                    
         sUpdateSQL = "UPDATE MonitorPointValue " _
                       & "SET Mode = " & mode & " " _
                       & "Where id = '" & sMonPtValId & "'"
               
         'Execute the SQL
         
         gdbConn.Execute sUpdateSQL
         
        
    End If
    rsDateTime.Close
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "SetMode", Err.Description)
    
End Sub



'**************************************************************
'   DESCRIPTION
'   This event is called once the form has already been displayed.
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'  4/22/2003       Joseph Manyaga   AT - fix declare function as public
'**************************************************************
Public Sub Form_Activate_Handler()


    On Error GoTo Error_Handler
    
    If Not bHasCalled Then
        If (bIsDatesEmpty And bIsTableEmpty) Then
           If MsgBox(LoadResString(IDS_MONDATA_FIRST_TIME), vbExclamation) = vbOK Then
                'JNM 05/09/03 - PT fix - We can enable this due to the button
                'frmMain.grdMonitorDataInput.Enabled = False
                'JNM 05/09/03 - But we should instead call the Validate Dates function
                '   immediately so that the user can immediately use the grid without
                '   any hassle
                cmdValidDate_Click_Handler
                frmMain.dtpStartIntroDate.SetFocus
            End If
            bHasCalled = True
        End If
    End If
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "Form_Activate", Err.Description)
    
End Sub


'**************************************************************
'   DESCRIPTION
'   Loads the form
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub Load_MonitorInput()
      
On Error GoTo Error_Handler

    'Initialize the global variables.
    nDecimal = 2
    bHasCalled = False
    bTimeChanged = False
    bDateChanged = False
    bdtpChanged = False
    gbCanChangeTabs = True
    
    'JNM 05/12/03 - Tell the app that the date and time are valid
    bDateTimeClick = True
        
    With frmMain
        ' Load the Resource Strings in the Appropiate controls
        .lblStartIntro.Caption = LoadResString(IDS_START_INTRO)
        .lblStartExpo.Caption = LoadResString(IDS_START_EXPOSURE)
        .lblEndExpo.Caption = LoadResString(IDS_END_EXPOSURE)
        .lblSIDate.Caption = LoadResString(IDS_DATE)
        .lblSITime.Caption = LoadResString(IDS_TIME)
        .lblSEDate.Caption = LoadResString(IDS_DATE)
        .lblSETime.Caption = LoadResString(IDS_TIME)
        .LblEEDate.Caption = LoadResString(IDS_DATE)
        .lblEETime.Caption = LoadResString(IDS_TIME)
        .lblMonDataInput.Caption = LoadResString(IDS_MONITOR_DATA_INPUT)
        .lblMonDataInput.FontBold = True
        .lblDataFilter.Caption = LoadResString(IDS_DATA_FILTER)
        .cmdNextStep6.Caption = LoadResString(IDS_BTN_NEXTSTEP)
        'JNM 05/09/03 - Pt fix - create this new button
        .cmdValidDate.Caption = LoadResString(IDS_LBL_CMD_VALIDDATE)
              
        .dtpStartIntroDate.ToolTipText = LoadResString(IDS_TIP_START_INTRO)
        .dtpStartIntroTime.ToolTipText = LoadResString(IDS_TIP_START_INTRO)
        .dtpStartExposureDate.ToolTipText = LoadResString(IDS_TIP_START_EXPOSURE)
        .dtpStartExposureTime.ToolTipText = LoadResString(IDS_TIP_START_EXPOSURE)
        .dtpEndExposureDate.ToolTipText = LoadResString(IDS_TIP_END_EXPOSURE)
        .dtpEndExposureTime.ToolTipText = LoadResString(IDS_TIP_END_EXPOSURE)
        .lstDataFilter.ToolTipText = LoadResString(IDS_TIP_DATA_FILTER)
        .dtpDummyDate = DateValue(Now())
        .dtpDummyTime = GetCurrentTime()
        
        'JNM 05/09/03 - PT fix - enable the list box
        .lstDataFilter.Enabled = True
        
        PaintGrid
        
        ' Jett 4/20/03 - OPTIMIZE: instantiated recordset here instead of in the module declaration
        'If rsMonitorPoint.State Then
        '    rsMonitorPoint.Close
        'End If
        
        Set rsMonitorPoint = New ADODB.Recordset
        
        sSQL = "Select id, name, jobarea_fkey from MonitorPoint where deleted = false"
        rsMonitorPoint.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
        
        If (rsMonitorPoint.EOF And rsMonitorPoint.BOF) Then
            'JNM 05/09/03 - PT fix - everything should be enabled already since
            '   we know that the dates are valid due to the button that was
            '   added to the form
'            ' Disable all controls
'            .dtpStartIntroDate.Enabled = False
'            .dtpStartIntroTime.Enabled = False
'            .dtpStartExposureDate.Enabled = False
'            .dtpStartExposureTime.Enabled = False
'            .dtpEndExposureDate.Enabled = False
'            .dtpEndExposureTime.Enabled = False
'            .lstDataFilter.Enabled = False
'            .grdMonitorDataInput.Enabled = False
            'cmdNextTab.Enabled = False
            MsgBox LoadResString(IDS_MONDATA_MSG_SAVE_NO_MONITOR_POINTS), vbExclamation
        Else
            'Initialize the boolean variables
            'Call the function that displays the time
            DisplayDateTime
            'Call the fucntion that displays the data in the grid
            DisplayGridData
            'Create the Data Filter list
            CreateMonitorPtList
            
            .grdMonitorDataInput.ColComboList(2) = sPipeList
    
        End If
        
    End With
    
    ' TODO: check if this belongs here
    'Form_Activate_Handler  JNM 04/22/03 - AT fix - We will be calling this in modmain.

    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "Form_Load", Err.Description)

End Sub


'**************************************************************
'   DESCRIPTION
'   Handles the keydown event of the date picker
'
'   Input Parameters: KeyCode, Shift
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub dtpDummyDate_KeyDown_Handler(KeyCode As Integer, Shift As Integer)

    On Error GoTo Error_Handler

    ' close date picker when user hits escape or return
    Select Case KeyCode
        Case vbKeyEscape
            frmMain.grdMonitorDataInput = frmMain.dtpDummyDate.Tag
            frmMain.dtpDummyDate.Visible = False
        Case vbKeyReturn
            frmMain.dtpDummyDate.Visible = False
    End Select

    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "dtpDummyDate_KeyDown", Err.Description)
    
End Sub


'**************************************************************
'   DESCRIPTION
'   Lost Focus event for the date picker
'
'   Input Parameters: KeyCode, Shift
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'  4/28/2003       Joseph Manyaga   Transferred code from Validate event.
'                                   Also modified the code to get correct row number.
'**************************************************************

Public Sub dtpDummyDate_LostFocus_Handler()

    Dim dCon As Double
    Dim rsValMode As ADODB.Recordset
    Dim sSQL As String
    Dim dtPreviousValue As Date
    Dim dGridDate As Date
    Dim nCurrRow As Integer
    
    On Error GoTo Error_Handler
    
    With frmMain
    
        'JNM 04/28/03 - AT fix - Store the row and col numbers at the back of the heading.
        '   This will be used later after the grid is refreshed.
        .grdMonitorDataInput.Cell(flexcpData, 0, 1) = .grdMonitorDataInput.Row
        .grdMonitorDataInput.Cell(flexcpData, 0, 2) = .grdMonitorDataInput.Col
        
        'JNM 04/28 - Get the current row
        nCurrRow = .grdMonitorDataInput.Cell(flexcpData, 0, 0)
    
        If .grdMonitorDataInput.Cell(flexcpText, nCurrRow, 6) <> "" Then
        
            dtPreviousValue = .grdMonitorDataInput.Cell(flexcpText, nCurrRow, 1)
            .grdMonitorDataInput.Cell(flexcpText, nCurrRow, 0) = .dtpDummyDate.Value
            dtIntroStarted = DateValue(.dtpStartIntroDate.Value) & " " & TimeValue(.dtpStartIntroTime.Value)
            dGridDate = DateValue(.grdMonitorDataInput.Cell(flexcpText, nCurrRow, 0)) & " " & TimeValue(.grdMonitorDataInput.Cell(flexcpText, nCurrRow, 1))
        
            If .grdMonitorDataInput.Cell(flexcpData, nCurrRow, gridColumnNumber.colConcentration) = "" Then
                dCon = 0
            Else
                dCon = CDbl(.grdMonitorDataInput.Cell(flexcpData, nCurrRow, gridColumnNumber.colConcentration))
                Set rsValMode = New ADODB.Recordset
                sSQL = "Select mode from MonitorPointValue where id = '" & .grdMonitorDataInput.Cell(flexcpText, nCurrRow, 6) & "'"
                rsValMode.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
                mode = rsValMode("mode")
            End If
        
        Else
            .grdMonitorDataInput.Cell(flexcpText, nCurrRow, 0) = .dtpDummyDate.Value
        End If
        
        If (dGridDate <= dtIntroStarted) And dCon > 0 Then
                
            If Not ValidateConcentrationTime(dCon, MDM_PRE_INTRO) Then
                'grdMonitorDataInput.Cell(flexcpData, Row, Col) = "Exit"
                .dtpDummyDate.Value = dtPreviousValue
                .grdMonitorDataInput.Cell(flexcpText, nCurrRow, 0) = dtPreviousValue
                frmMain.dtpDummyDate.Visible = False
                Exit Sub
            End If
            
        End If
        
    End With
    
    SaveDateTime (sDate)

    ' hide date picker when user is done with it
    frmMain.dtpDummyDate.Visible = False
    'SaveDateTime (sDate)
    
    'JNM 04/23/03 - AT Fix - In order to prevent multiple row selection
    DoEvents
    'JNM 04/23/03 - Select the row clicked by the user
    frmMain.grdMonitorDataInput.Select frmMain.grdMonitorDataInput.Cell(flexcpData, 0, 1), frmMain.grdMonitorDataInput.Cell(flexcpData, 0, 2)
    
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "dtpDummyDate_LostFocus", Err.Description)
    
End Sub


'**************************************************************
'   DESCRIPTION
'   AfterEdit event of the grid
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'  10/15/2002      Cel Gamboa       Added code for the highlights
'  4/2/2003        Amana Martinez   Fumiguide Enhancement 2003 Coding Phase
'**************************************************************
Public Sub grdMonitorDataInput_AfterEdit_Handler(ByVal Row As Long, ByVal Col As Long)

    On Error GoTo Error_Handler

    If Row > 0 And Col = gridColumnNumber.colConcentration Then
        If frmMain.grdMonitorDataInput.EditText <> "" Then
            sHighlight = frmMain.grdMonitorDataInput.Cell(flexcpText, Row, 6)
            'JNM 11/15 - Commented-out the line below in order to prevent resorting and
            '   hard-coded the units
            'AM 4/2/03 - Remove the comment below to reload the values within the grid for Fumiguide Enhancement 2003
            Call lstDataFilter_Click_Handler
            'AM 4/2/03 - Remove the comment below to reload the values within the grid for Fumiguide Enhancement 2003
            'AM 4/2/03 - Added the following to keep the cursor to the last selected record for Fumiguide Enhancement 2003
            frmMain.grdMonitorDataInput.Select Row, Col
            'AM 4/2/03 - Added the following to keep the cursor to the last selected record for Fumiguide Enhancement 2003
            frmMain.grdMonitorDataInput.Cell(flexcpText, Row, Col) = _
                frmMain.grdMonitorDataInput.Cell(flexcpData, Row, Col) '& " " & sConcUnits AM 4/2/2003 - Commented this for Fumiguide Enhancement 2003
        End If
    End If
    
    gbHasChanged = True
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "grdMonitorDataInput_AfterEdit", Err.Description)
End Sub


'**************************************************************
'   DESCRIPTION
'   Before Scroll event of the grid
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************

Public Sub grdMonitorDataInput_BeforeScroll_Handler(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    On Error GoTo Error_Handler

    ' don't scroll while editing dates
    If frmMain.dtpDummyDate.Visible Then Cancel = True
     ' don't scroll while editing time
    If frmMain.dtpDummyTime.Visible Then Cancel = True

    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "grdMonitorDataInput_BeforeScroll", Err.Description)
    
End Sub


'**************************************************************
'   DESCRIPTION
'   Mouse move event of the grid
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub grdMonitorDataInput_MouseMove_Handler(Button As Integer, Shift As Integer, x As Single, Y As Single)

    On Error GoTo Error_Handler

    With frmMain.grdMonitorDataInput
        Select Case .MouseCol
            Case 0
                .ToolTipText = LoadResString(IDS_TIP_DATE_COL)
            Case 1
                .ToolTipText = LoadResString(IDS_TIP_TIME_COL)
            Case 2
                .ToolTipText = LoadResString(IDS_TIP_MON_PT_NAME_COL)
            Case 3
                .ToolTipText = LoadResString(IDS_TIP_CONCENTRATION_COL)
            Case 4
                .ToolTipText = LoadResString(IDS_TIP_HLT_COL)
            Case 5
                .ToolTipText = LoadResString(IDS_TIP_NOTE_COL)
        End Select
    End With

    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "grdMonitorDataInput_MouseMove", Err.Description)
    
End Sub


'**************************************************************
'   DESCRIPTION
'   Keydown event of the time picker
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub dtpDummyTime_KeyDown_Handler(KeyCode As Integer, Shift As Integer)
 
    On Error GoTo Error_Handler

    ' close date picker when user hits escape or return
    Select Case KeyCode
        Case vbKeyEscape
            frmMain.grdMonitorDataInput = frmMain.dtpDummyTime.Tag
            frmMain.dtpDummyTime.Visible = False
        Case vbKeyReturn
            frmMain.grdMonitorDataInput = frmMain.dtpDummyTime.Tag
            frmMain.dtpDummyTime.Visible = False
            SaveDateTime (sTime)
    End Select
   
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "dtpDummyTime_KeyDown", Err.Description)
    
End Sub


'**************************************************************
'   DESCRIPTION
'   Lost focus event of the time picker
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub dtpDummyTime_LostFocus_Handler()

    On Error GoTo Error_Handler

'JNM 04/23/03 - AT Fix - Transferred entired code from dtpDummyTime_Change_Handler()
'   here so that saving would be done only after the user clicks another control.
'   Modified the code so that it would get current row which is hidden in the heading.
'*****************************************************************************
    
    Dim dtPreviousValue As Date
    Dim dGrdConcentration As Double
    Dim rsValMode As ADODB.Recordset
    Dim sSQL As String
    Dim dGridTime As Date
    Dim nCurrRow As Integer
    
    With frmMain
        'JNM 04/23/03 - AT fix - Store the row and col numbers at the back of the heading.
        '   This will be used later after the grid is refreshed.
        .grdMonitorDataInput.Cell(flexcpData, 0, 1) = .grdMonitorDataInput.Row
        .grdMonitorDataInput.Cell(flexcpData, 0, 2) = .grdMonitorDataInput.Col
        
        nCurrRow = frmMain.grdMonitorDataInput.Cell(flexcpData, 0, 0)
        
        If .grdMonitorDataInput.Cell(flexcpText, nCurrRow, 6) <> "" Then
        
            dtPreviousValue = .grdMonitorDataInput.Cell(flexcpText, nCurrRow, 0)
            'JNM 04/23 - Place the value in the cell, not the grid
            '.grdMonitorDataInput.Text = .dtpDummyTime.Value
            .grdMonitorDataInput.Cell(flexcpText, nCurrRow, 1) = .dtpDummyTime.Value
            
            dtIntroStarted = DateValue(.dtpStartIntroDate.Value) & " " & TimeValue(.dtpStartIntroTime.Value)
            dGridTime = DateValue(.grdMonitorDataInput.Cell(flexcpText, nCurrRow, 0)) & " " & TimeValue(.grdMonitorDataInput.Cell(flexcpText, nCurrRow, 1))
        
            If .grdMonitorDataInput.Cell(flexcpData, nCurrRow, gridColumnNumber.colConcentration) = "" Then
                dGrdConcentration = 0
             Else
                dGrdConcentration = CDbl(.grdMonitorDataInput.Cell(flexcpData, nCurrRow, gridColumnNumber.colConcentration))
                
                Set rsValMode = New ADODB.Recordset
                
                sSQL = "Select mode from MonitorPointValue where id = '" & .grdMonitorDataInput.Cell(flexcpText, nCurrRow, 6) & "'"
                rsValMode.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
                mode = rsValMode("mode")
                
                rsValMode.Close
                Set rsValMode = Nothing
                
            End If
            
        Else
            .grdMonitorDataInput.Cell(flexcpText, nCurrRow, 1) = .dtpDummyTime.Value
        End If
        
        If (dGridTime <= dtIntroStarted) And dGrdConcentration > 0 Then
                
                If Not ValidateConcentrationTime(dGrdConcentration, MDM_PRE_INTRO) Then
                       'grdMonitorDataInput.Cell(flexcpData, Row, Col) = "Exit"
                       .dtpDummyTime.Value = dtPreviousValue
                       .grdMonitorDataInput.Cell(flexcpText, nCurrRow, 1) = dtPreviousValue
                       frmMain.dtpDummyTime.Visible = False
                       Exit Sub
                End If
        End If
        
    End With
    
    SaveDateTime (sTime)
    'bTimeChanged = True
'*****************************************************************************
    ' hide date picker when user is done with it
    frmMain.dtpDummyTime.Visible = False
    
    'JNM 04/23/03 - AT Fix - In order to prevent multiple row selection
    DoEvents
    'JNM 04/23/03 - Select the row clicked by the user
    frmMain.grdMonitorDataInput.Select frmMain.grdMonitorDataInput.Cell(flexcpData, 0, 1), frmMain.grdMonitorDataInput.Cell(flexcpData, 0, 2)
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "dtpDummyTime_LostFocus", Err.Description)
    
End Sub


'**************************************************************
'   DESCRIPTION
'   Before user resize event of the grid
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub grdMonitorDataInput_BeforeUserResize_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  
    On Error GoTo Error_Handler

    ' don't resize columns while editing dates
    If frmMain.dtpDummyDate.Visible Then Cancel = True
    If frmMain.dtpDummyTime.Visible Then Cancel = True

    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "grdMonitorDataInput_BeforeUserResize", Err.Description)
    
End Sub


'**************************************************************
'   DESCRIPTION
'   Keydown event of the grid
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub grdMonitorDataInput_KeyDown_Handler(KeyCode As Integer, Shift As Integer)

    On Error GoTo Error_Handler
    
    If KeyCode = vbKeyDelete Then
        Call OnRowDelete
    End If
        
    If Shift = 2 And KeyCode = 67 Then
        ' grdMonitorDataInput.Select grdMonitorDataInput.Row, gridColumnNumber.colDate, grdMonitorDataInput.RowSel, gridColumnNumber.colMonitorPoint
        Call mnuMonitorInputCopy_Click_Handler
    End If
    If Shift = 2 And KeyCode = 86 Then
        Call mnuMonitorInputPaste_Click_Handler
    End If
    
    ' Add key code for copy and paste
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "grdMonitorDataInput_KeyDown", Err.Description)

End Sub


'**************************************************************
'   DESCRIPTION
'   Mouse down event of the grid
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub grdMonitorDataInput_MouseDown_Handler(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim sCopied As Variant

    On Error GoTo Error_Handler

    sCopied = Clipboard.GetText(vbCFText)
    
    'JNM 04/25/03 - AT Fix - Display pop-up only if grid is not being edited
    If frmMain.grdMonitorDataInput.EditWindow = 0 Then
        'JNM 11/04 - Do not display the pop-up menu if selected row is the heading.
        If Button = 2 And frmMain.grdMonitorDataInput.Row > 0 Then
            frmMain.grdMonitorDataInput.Select frmMain.grdMonitorDataInput.Row, gridColumnNumber.colDate, frmMain.grdMonitorDataInput.RowSel, gridColumnNumber.colMonitorPoint
            'check if there is already an existing pop-up, if so release the old one.
            If IsEmpty(sCopied) Then
                frmMain.PopupMenu frmMain.mnuMonitorInputContext
                frmMain.mnuMonitorInputPaste.Enabled = False
            Else
                frmMain.PopupMenu frmMain.mnuMonitorInputContext
                frmMain.mnuMonitorInputPaste.Enabled = True
            End If
        End If
    End If

    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "grdMonitorDataInput_MouseDown_Handler", Err.Description)
    
End Sub


'**************************************************************
'   DESCRIPTION
'   Start edit event of the grid
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub grdMonitorDataInput_StartEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    On Error GoTo Error_Handler

    'JNM 04/23/03 - AT Fix - We need to store the current row number at the back
    '   of the heading so that other functions can access it (especially the
    '   functions that use the date-time picker
    frmMain.grdMonitorDataInput.Cell(flexcpData, 0, 0) = frmMain.grdMonitorDataInput.Row
        
    If Col = gridColumnNumber.colHLT And Row > 0 Then
        Cancel = True
    End If
    
    With frmMain
        ' if this is a date column, edit it with the date picker control
        If Col = gridColumnNumber.colDate And Row > 0 Then
            
            ' we'll handle the editing ourselves
            Cancel = True
            
            ' position date picker control over cell
           .dtpDummyDate.Move 3000 + .grdMonitorDataInput.CellLeft, 2290 + .grdMonitorDataInput.CellTop, 5 + .grdMonitorDataInput.CellWidth, 35 + .grdMonitorDataInput.CellHeight
            
            ' initialize value, save original in tag in case user hits escape
            If .grdMonitorDataInput <> "" Then
                .dtpDummyDate.Value = .grdMonitorDataInput
                .dtpDummyDate.Tag = .grdMonitorDataInput
            End If
            
            ' show and activate date picker control
            .dtpDummyDate.Visible = True
            .dtpDummyDate.SetFocus
            
            ' make it drop down the calendar
            SendKeys "{f4}"
            
        End If
        
        If Col = gridColumnNumber.colTime And Row > 0 Then
            
            ' we'll handle the editing ourselves
            Cancel = True
            
            ' position date picker control over cell
            .dtpDummyTime.Move 3000 + .grdMonitorDataInput.CellLeft, 2290 + .grdMonitorDataInput.CellTop, 5 + .grdMonitorDataInput.CellWidth, 35 + .grdMonitorDataInput.CellHeight
            
            ' initialize value, save original in tag in case user hits escape
            If .grdMonitorDataInput <> "" Then
                .dtpDummyTime.Value = .grdMonitorDataInput
                .dtpDummyTime.Tag = .grdMonitorDataInput
            End If
            
            ' show and activate date picker control
            .dtpDummyTime.Visible = True
            .dtpDummyTime.SetFocus
            
            ' make it drop down the calendar
            SendKeys "{f4}"
            
        End If
        
        If Col = gridColumnNumber.colConcentration And Row > 0 Then
             'Remove the units once the user starts to edit the cell.
             If CStr(.grdMonitorDataInput.Cell(flexcpData, Row, Col)) = "Exit" Then
                 .grdMonitorDataInput.Cell(flexcpData, Row, Col) = ""
                 Cancel = True
             Else
                 .grdMonitorDataInput.Cell(flexcpText, Row, Col) = CStr(.grdMonitorDataInput.Cell(flexcpData, Row, Col))
                 .grdMonitorDataInput.EditMaxLength = 6
             End If
             
        End If
        
        If Col = gridColumnNumber.colNotes And Row > 0 Then
            .grdMonitorDataInput.EditMaxLength = 50
        End If
        
    End With

    ReturnCurrTab TAB_MONITOR_INPUT
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "grdMonitorDataInput_StartEdit", Err.Description)

End Sub



'**************************************************************
'   DESCRIPTION
'   Validate edit event of the grid
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************

Public Sub grdMonitorDataInput_ValidateEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim sMonitorPoint As String
    Dim dConcentration As Double
    Dim sNotes As String
    Dim sMonPtValId As String
    Dim sConReading As String
    Dim sgrdDate As String
    Dim sgrdTime As String
    Dim dtGrid As Date
    Dim sDateTime As String
    
    Dim rsMonitorValues As ADODB.Recordset
    Dim rsMonitorID As ADODB.Recordset
    
    On Error GoTo Error_Handler
    
    Set rsMonitorID = New ADODB.Recordset
    Set rsMonitorValues = New ADODB.Recordset
    
    With frmMain
        ' Jett 04/20/03: removed this, it will never be open now that the recordset is
        ' declared from within
        'If rsMonitorValues.State Then
        '    rsMonitorValues.Close
        'End If
        
        If Col = gridColumnNumber.colMonitorPoint Then
            bRecordDirty = True
            sMonPtValId = .grdMonitorDataInput.Cell(flexcpText, Row, 6)
            sMonitorPoint = .grdMonitorDataInput.EditText
            
            If sMonitorPoint <> "" Then
            
                  If Trim$(sMonPtValId) <> "" Then
                    sSQL = "Select * from MonitorPointvalue where id = '" & sMonPtValId & "'"
                    rsMonitorValues.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
                    sHighlight = sMonPtValId
                    
                    sSQL = "Select id from MonitorPoint where name = '" & sMonitorPoint & "' and deleted = false"
                    
                    rsMonitorID.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
                    rsMonitorValues("monitorpoint_fkey") = rsMonitorID("id")
                    rsMonitorValues.Update
                    rsMonitorID.Close
                    rsMonitorValues.Close
                    If Trim$(.grdMonitorDataInput.Cell(flexcpText, Row, 3)) <> "" And Trim$(.grdMonitorDataInput.Cell(flexcpText, Row, 3)) <> "0" Then
                        CalculateMonDataInput
                        Call lstDataFilter_Click_Handler
                    End If
                    
                   
                  Else
                    sMonPtValId = GetGUID()
                    .grdMonitorDataInput.Cell(flexcpText, Row, 0) = DateValue(Now)
                    .grdMonitorDataInput.Cell(flexcpText, Row, 1) = GetCurrentTime()
                    sDateTime = .grdMonitorDataInput.Cell(flexcpText, Row, 0) & " " & .grdMonitorDataInput.Cell(flexcpText, Row, 1)
                    
                    sSQL = "Select id from MonitorPoint where name = '" & sMonitorPoint & "' and deleted = false"
                    
                    rsMonitorID.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
                    sMonitorPoint = rsMonitorID("id")
                    
                    sSQL = "Insert into MonitorPointValue (id, monitorpoint_fkey, date_time) VALUES ('" & sMonPtValId & "', '" & sMonitorPoint & "', '" & sDateTime & "')"
                    gdbConn.Execute sSQL
                    .grdMonitorDataInput.Cell(flexcpText, Row, 6) = sMonPtValId
                    rsMonitorID.Close
                  End If
            End If
            
        End If
    
        If Col = gridColumnNumber.colNotes Then
            bRecordDirty = True
            sMonPtValId = .grdMonitorDataInput.Cell(flexcpText, Row, 6)
            sNotes = Trim$(.grdMonitorDataInput.EditText)
            
            If Len(sNotes) > NOTES_MAX Then
                MsgBox LoadResString(IDS_MONDATA_MSG_NOTE_50_CHAR), vbExclamation
                Cancel = True
                
                Set rsMonitorValues = Nothing
                Set rsMonitorID = Nothing
                Exit Sub
            End If
            
            If sNotes <> "" Then
                If sMonPtValId <> "" Then
                    sSQL = "Select * from MonitorPointvalue where id = '" & sMonPtValId & "'"
                    
                    rsMonitorValues.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
                
                    rsMonitorValues("general_note") = sNotes
                    rsMonitorValues.Update
                    rsMonitorValues.Close
                Else
                    sMonPtValId = GetGUID()
                    'Set the default values
                    .grdMonitorDataInput.Cell(flexcpText, Row, 0) = DateValue(Now)
                    .grdMonitorDataInput.Cell(flexcpText, Row, 1) = GetCurrentTime()
                    sDateTime = .grdMonitorDataInput.Cell(flexcpText, Row, 0) & " " & .grdMonitorDataInput.Cell(flexcpText, Row, 1)
                    .grdMonitorDataInput.Cell(flexcpText, Row, 2) = sDefaultMonPt
                    sSQL = "Insert into MonitorPointValue (id, date_time, monitorpoint_fkey, general_note) VALUES ('" & sMonPtValId & "', '" & sDateTime & "', '" & sDefaultMonPtID & "', '" & sNotes & "')"
                    
                    gdbConn.Execute sSQL
                    .grdMonitorDataInput.Cell(flexcpText, Row, 6) = sMonPtValId
                End If
            End If
        End If
    
        If Col = gridColumnNumber.colConcentration Then
                  
                bRecordDirty = True
                sMonPtValId = .grdMonitorDataInput.Cell(flexcpText, Row, 6)
                sConReading = Trim$(.grdMonitorDataInput.EditText)
                'If sConReading <> "" Then
                    'Make sure that the Date, Time are populated.
                    sgrdDate = .grdMonitorDataInput.Cell(flexcpText, Row, gridColumnNumber.colDate)
                    sgrdTime = .grdMonitorDataInput.Cell(flexcpText, Row, gridColumnNumber.colTime)
                    
                    If sgrdDate = "" Then
                        .grdMonitorDataInput.Cell(flexcpText, Row, 0) = DateValue(Now)
                        sgrdDate = DateValue(Now)
                    End If
                    
                    If sgrdTime = "" Then
                        .grdMonitorDataInput.Cell(flexcpText, Row, 1) = GetCurrentTime()
                        sgrdTime = GetCurrentTime()
                    End If
                    
                    If IsNull(.grdMonitorDataInput.Cell(flexcpText, Row, gridColumnNumber.colMonitorPoint)) Or .grdMonitorDataInput.Cell(flexcpText, Row, gridColumnNumber.colMonitorPoint) = "" Then
                        .grdMonitorDataInput.Cell(flexcpText, Row, 2) = sDefaultMonPt
                    End If
                    
                    If ValidateRecordAdd(.grdMonitorDataInput.Cell(flexcpText, Row, gridColumnNumber.colMonitorPoint), sConReading, .grdMonitorDataInput.Cell(flexcpText, Row, gridColumnNumber.colNotes)) = True Then
                        dtGrid = DateValue(sgrdDate) & " " & TimeValue(sgrdTime)
                        sMonitorPoint = .grdMonitorDataInput.Cell(flexcpTextDisplay, Row, gridColumnNumber.colMonitorPoint)
                        
                        sSQL = "Select id from MonitorPoint where name = '" & sMonitorPoint & "' and deleted = FALSE"
                        rsMonitorID.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
                        
                        If Trim$(sMonPtValId) = "" Then
                            sMonPtValId = GetGUID()
                            sSQL = "Insert into MonitorPointValue (id, date_time, monitorpoint_fkey) VALUES ('" & sMonPtValId & "','" & dtGrid & "','" & rsMonitorID("Id") & "' )"
                            gdbConn.Execute sSQL
                            .grdMonitorDataInput.Cell(flexcpText, Row, 6) = sMonPtValId
        
                        End If
                        
                        SetMode sMonPtValId, dtGrid
                        
                        If sConReading = "" Then
                            dConcentration = 0
                        Else
                            dConcentration = CDbl(sConReading)
                        End If
                        
                            If mode = MDM_PRE_INTRO And dConcentration > 0 Then
            
                                If Not ValidateConcentrationTime(dConcentration, MDM_PRE_INTRO) Then
                                    .grdMonitorDataInput.Cell(flexcpData, Row, Col) = "Exit"
                                    Cancel = True
                                    rsMonitorID.Close
                                    
                                    Set rsMonitorValues = Nothing
                                    Set rsMonitorID = Nothing
                                    Exit Sub
                                End If
                            End If
                        
                        .grdMonitorDataInput.Cell(flexcpData, Row, Col) = sConReading
                        
                        ' Jett 04/20/03: removed this. the recordset will not be open anymore
                        'If rsMonitorValues.State Then
                        '    rsMonitorValues.Close
                        'End If
                        
                        ' If the data are valid, call the CalculateMonDataInput to perform the calculations
                        CalculateMonDataInput sMonPtValId, dtGrid, dConcentration, INTERSCAN, .grdMonitorDataInput.Cell(flexcpText, Row, gridColumnNumber.colNotes), mode, rsMonitorID("id")
                                        
                        rsMonitorID.Close
                        
                    Else
                        .grdMonitorDataInput.Cell(flexcpData, Row, Col) = ""
                        .grdMonitorDataInput.Cell(flexcpText, Row, Col) = ""
                        Cancel = True
                    End If
                'End If
            End If
            
            If Cancel Then
                .grdMonitorDataInput.SetFocus
            End If
    End With
    
    Set rsMonitorValues = Nothing
    Set rsMonitorID = Nothing

    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "grid_ValidateEdit", Err.Description)
       
End Sub


'**************************************************************
'   DESCRIPTION
'   Handles the on click event of the Data Filter listbox
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'  4/2/2003        Amana Martinez   Fumiguide Enhancement 2003, Coding Phase
'  5/12/2004       Jett Gamboa      changed checking of "ALL" option in filter
'**************************************************************
Public Sub lstDataFilter_Click_Handler()

    Dim sSelected As String
    Dim I As Long
    Dim sSQLFilter As String
    Dim j As Integer
    Dim intRec As Integer
    
    Dim rsFilterData As ADODB.Recordset
    
    On Error GoTo Error_Handler
    
    Set rsFilterData = New ADODB.Recordset
    
    sSelected = ""
    dValue = 0#
    
    For I = 0 To frmMain.lstDataFilter.ListCount - 1
        If frmMain.lstDataFilter.Selected(I) Then
            If sSelected = "" Then
                sSelected = "'" & frmMain.lstDataFilter.List(I) & "'"
            Else
                sSelected = sSelected & " , " & "'" & frmMain.lstDataFilter.List(I) & "'"
            End If
        End If
    Next
    
    ' 5/12/04 JETT LOCALIZATION
    ' Changed the way we check if the "ALL" item was selected. Now that
    ' we are localizing, we we should not check for the word ALL but instead
    ' check if the user selected the first item on the list.
    
    ' If sSelected = "'ALL'" Or sSelected = "" Then
    If frmMain.lstDataFilter.Selected(0) Or sSelected = "" Then
    
        frmMain.grdMonitorDataInput.Rows = GRIDSIZE
        
        'clear the grid of its contents
        frmMain.grdMonitorDataInput.Cell(flexcpData, 1, 0, GRIDSIZE - 1, 6) = ""
        frmMain.grdMonitorDataInput.Cell(flexcpText, 1, 0, GRIDSIZE - 1, 6) = ""
        DisplayGridData
        
    Else
    
        ' 11/18/02 TSG - changed sort order when user has decided to filter results
        sSQLFilter = "SELECT MonitorPointValue.id as Id, MonitorPointValue.date_time, " _
        & "MonitorPointValue.fumiscope_conc, MonitorPointValue.interscan_conc, " _
        & "MonitorPointValue.general_note, mode, monitorpoint_fkey, hlt_pt2pt, name " _
        & "FROM MonitorPointValue, MonitorPoint where MonitorPointValue.monitorpoint_fkey = MonitorPoint.id " _
        & "and MonitorPoint.name in (" & sSelected & ") and MonitorPointValue.deleted = FALSE " _
        & "ORDER by name, date_time ASC"
        
        rsFilterData.Open sSQLFilter, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
        
        j = 1
        
        If Not (rsFilterData.EOF And rsFilterData.BOF) Then
           frmMain.grdMonitorDataInput.Cell(flexcpData, 1, 0, GRIDSIZE - 1, 6) = ""
           frmMain.grdMonitorDataInput.Cell(flexcpText, 1, 0, GRIDSIZE - 1, 6) = ""
            intRec = rsFilterData.RecordCount
            rsFilterData.MoveFirst
            While Not (j > intRec)
                    frmMain.grdMonitorDataInput.Cell(flexcpText, j, 0) = IIf(IsNull(rsFilterData("date_time")), "", rsFilterData("date_time"))
                    frmMain.grdMonitorDataInput.Cell(flexcpText, j, 1) = IIf(IsNull(rsFilterData("date_time")), "", rsFilterData("date_time"))
                    frmMain.grdMonitorDataInput.Cell(flexcpText, j, 2) = IIf(IsNull(rsFilterData("name")), "", rsFilterData("name"))
                    frmMain.grdMonitorDataInput.Cell(flexcpData, j, 3) = CStr(IIf(IfLessZero(rsFilterData("fumiscope_conc")), "", rsFilterData("fumiscope_conc")))
                    frmMain.grdMonitorDataInput.Cell(flexcpText, j, 3) = IIf(IfLessZero(rsFilterData("fumiscope_conc")), "", rsFilterData("fumiscope_conc"))  '& " " & sConcUnits) AM 4/2/03 - Commented this for Fumiguide Enhancement 2003
                    dValue = Round(rsFilterData("hlt_pt2pt"), nDecimal)
                    frmMain.grdMonitorDataInput.Cell(flexcpText, j, 4) = IIf(IfZero(CStr(dValue)), "", dValue)
                    frmMain.grdMonitorDataInput.Cell(flexcpText, j, 5) = IIf(IsNull(rsFilterData("general_note")), "", rsFilterData("general_note"))
                    If sHighlight = rsFilterData("Id") Then
                        frmMain.grdMonitorDataInput.Select j, 0, j, 5
                    End If
                    frmMain.grdMonitorDataInput.Cell(flexcpText, j, 6) = rsFilterData("Id")
                    j = j + 1
                    rsFilterData.MoveNext
                Wend
        Else
            frmMain.grdMonitorDataInput.Cell(flexcpData, 1, 0, GRIDSIZE - 1, 6) = ""
            frmMain.grdMonitorDataInput.Cell(flexcpText, 1, 0, GRIDSIZE - 1, 6) = ""
        End If
        
        rsFilterData.Close
        Set rsFilterData = Nothing
        
        ' 11/26/02 TSG - scroll to the row after the last one with value after filtering
        frmMain.grdMonitorDataInput.TopRow = IIf(j - 21 > 0, j - 21, 1)
        frmMain.grdMonitorDataInput.Select j, 1
        
    End If

    Exit Sub
    
Error_Handler:
    Call psAbort(MODULE_NAME, "lstDataFilter", Err.Description)

End Sub


'**************************************************************
'   DESCRIPTION
'   Handles the copying
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub mnuMonitorInputCopy_Click_Handler()
    
    Dim sData As String
    Dim lRow As Long
    Dim lLastRow As Long
    Dim iCtr As Integer
    Dim sRowNum As String
    
    On Error GoTo Error_Handler

    frmMain.grdMonitorDataInput.Select frmMain.grdMonitorDataInput.Row, gridColumnNumber.colDate, frmMain.grdMonitorDataInput.RowSel, gridColumnNumber.colMonitorPoint

    lRow = frmMain.grdMonitorDataInput.Row
    lLastRow = frmMain.grdMonitorDataInput.RowSel
    iCtr = lRow
        
    While Not iCtr > lLastRow
        If (frmMain.grdMonitorDataInput.Cell(flexcpText, iCtr, 6) = "") Then
            sRowNum = "# " & iCtr
            MsgBox FormatSWParams(LoadResString(IDS_COPY_NON_EMPTY), sRowNum), vbExclamation
            Exit Sub
        End If
        iCtr = iCtr + 1
    Wend
        
    Clipboard.Clear
    sData = frmMain.grdMonitorDataInput.Clip
    Clipboard.SetText sData, vbCFText
    
    Exit Sub
    
Error_Handler:
    Call psAbort(MODULE_NAME, "mnuMonitorInputCopy_Click_Handler", Err.Description)
    
End Sub


'**************************************************************
'   DESCRIPTION
'   Handles the pasting
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub mnuMonitorInputPaste_Click_Handler()
       
    Dim vPasteRow As Variant
    Dim vPasteCell As Variant
    Dim nRow As Integer
    Dim vDate As Variant
    Dim vTime As Variant
    Dim vMonPtName As Variant
    Dim x As Integer
    Dim sMonPVId As String
    Dim sGridID As String
    Dim bCalculate As Boolean
    
    Dim rsNewData As ADODB.Recordset
    Dim rsUpdateData As ADODB.Recordset
    Dim rsMonitorID As ADODB.Recordset
            
    On Error GoTo Error_Handler
    
    Set rsNewData = New ADODB.Recordset
    Set rsUpdateData = New ADODB.Recordset
    Set rsMonitorID = New ADODB.Recordset
         
    bCalculate = False
         
    vPasteRow = Split(Clipboard.GetText(vbCFText), vbCr)
    nRow = frmMain.grdMonitorDataInput.Row - 1
    
    sSQL = "Select id, date_time, monitorpoint_fkey from MonitorPointvalue"
    rsNewData.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
            
    For x = 0 To UBound(vPasteRow)
        nRow = nRow + 1
        vPasteCell = Split(vPasteRow(x), vbTab)
            If UBound(vPasteCell) < gridColumnNumber.colMonitorPoint Then
                MsgBox LoadResString(IDS_PASTE_DATA_INVALID), vbExclamation
                rsNewData.Close
                Set rsNewData = Nothing
                Set rsUpdateData = Nothing
                Set rsMonitorID = Nothing
                Exit Sub
            End If
            
        vDate = vPasteCell(gridColumnNumber.colDate)
        vTime = vPasteCell(gridColumnNumber.colTime)
        vMonPtName = vPasteCell(gridColumnNumber.colMonitorPoint)
        sMonPVId = frmMain.grdMonitorDataInput.Cell(flexcpText, nRow, gridColumnNumber.colID)
           
        If vDate <> "" Or vTime <> "" Or vMonPtName <> "" Then
                           
                If sMonPVId = "" Then
                
                    rsNewData.AddNew
                    sGridID = GetGUID()
                    rsNewData("id") = sGridID
                    frmMain.grdMonitorDataInput.Cell(flexcpText, nRow, gridColumnNumber.colID) = sGridID
                    If vMonPtName <> "" Then
                        sSQL = "Select id from MonitorPoint where name = '" & vMonPtName & "' and deleted = false"
                        rsMonitorID.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
                            If Not (rsMonitorID.BOF And rsMonitorID.EOF) Then
                                rsNewData("monitorpoint_fkey") = rsMonitorID("id")
                            End If
                        rsMonitorID.Close
                    End If
                    
                    'JNM 05/08/03 - PT fix - sometimes the date has no time part
                    If vDate <> "" And vTime <> "" Then
                        If IsDate(vDate) Then
                            If IsDate(vTime) Then
                                rsNewData("date_time") = DateValue(vDate) & " " & TimeValue(vTime)
                            Else
                                MsgBox LoadResString(IDS_PASTE_DATA_INVALID), vbExclamation
                                rsNewData.Close
                                Set rsNewData = Nothing
                                Set rsUpdateData = Nothing
                                Set rsMonitorID = Nothing
                                Exit Sub
                            End If
                        Else
                            MsgBox LoadResString(IDS_PASTE_DATA_INVALID), vbExclamation
                            rsNewData.Close
                            Set rsNewData = Nothing
                            Set rsUpdateData = Nothing
                            Set rsMonitorID = Nothing
                            Exit Sub
                        End If
                    Else
                        If vDate <> "" Then
                            If IsDate(vDate) Then
                                rsNewData("date_time") = vDate
                            Else
                                MsgBox LoadResString(IDS_PASTE_DATA_INVALID), vbExclamation
                                rsNewData.Close
                                Set rsNewData = Nothing
                                Set rsUpdateData = Nothing
                                Set rsMonitorID = Nothing
                                Exit Sub
                            End If
                        End If
                        
                        If vTime <> "" Then
                            If IsDate(vTime) Then
                                rsNewData("date_time") = vTime
                            Else
                                MsgBox LoadResString(IDS_PASTE_DATA_INVALID), vbExclamation
                                rsNewData.Close
                                Set rsNewData = Nothing
                                Set rsUpdateData = Nothing
                                Set rsMonitorID = Nothing
                                Exit Sub
                            End If
                        End If
                    End If
                        
                    rsNewData.Update
                    
                 Else
                 
                        sSQL = "Select id, date_time, monitorpoint_fkey, fumiscope_conc from MonitorPointvalue where id = '" & sMonPVId & "'"
                        rsUpdateData.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText

                        If vMonPtName <> "" Then
                            sSQL = "Select id from MonitorPoint where name = '" & vMonPtName & "' and deleted = false"
                            rsMonitorID.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
                                If Not (rsMonitorID.BOF And rsMonitorID.EOF) Then
                                rsUpdateData("monitorpoint_fkey") = rsMonitorID("id")
                                    If Not IsNull(rsUpdateData("fumiscope_conc")) Then
                                        bCalculate = True
                                    End If
                                End If
                            rsMonitorID.Close
                        End If
                        
                        'JNM 05/08/03 - PT fix - sometimes the date has no time part
                        If vDate <> "" And vTime <> "" Then
                            If IsDate(vDate) Then
                                If IsDate(vTime) Then
                                    rsUpdateData("date_time") = DateValue(vDate) & " " & TimeValue(vTime)
                                Else
                                    MsgBox LoadResString(IDS_PASTE_DATA_INVALID), vbExclamation
                                    rsUpdateData.Close
                                    
                                    Set rsUpdateData = Nothing
                                    Set rsNewData = Nothing
                                    Set rsMonitorID = Nothing
                                    Exit Sub
                                End If
                            Else
                                MsgBox LoadResString(IDS_PASTE_DATA_INVALID), vbExclamation
                                rsUpdateData.Close
                                Set rsUpdateData = Nothing
                                Set rsNewData = Nothing
                                Set rsMonitorID = Nothing
                                Exit Sub
                            End If
                        Else
                            If vDate <> "" Then
                                If IsDate(vDate) Then
                                    rsUpdateData("date_time") = vDate
                                Else
                                    MsgBox LoadResString(IDS_PASTE_DATA_INVALID), vbExclamation
                                    rsUpdateData.Close
                                    Set rsUpdateData = Nothing
                                    Set rsNewData = Nothing
                                    Set rsMonitorID = Nothing
                                    Exit Sub
                                End If
                            End If
                            
                            If vTime <> "" Then
                                If IsDate(vTime) Then
                                    rsUpdateData("date_time") = vTime
                                Else
                                    MsgBox LoadResString(IDS_PASTE_DATA_INVALID), vbExclamation
                                    rsUpdateData.Close
                                    
                                    Set rsUpdateData = Nothing
                                    Set rsNewData = Nothing
                                    Set rsMonitorID = Nothing
                                    Exit Sub
                                End If
                            End If
                        End If
                        
                        rsUpdateData.Update
                        rsUpdateData.Close
                 End If
                     
            End If
            
            'grdMonitorDataInput.Select nRow, gridColumnNumber.colDate, nRow, gridColumnNumber.colMonitorPoint
            'grdMonitorDataInput.Clip = vPasteRow(x)
            gbHasChanged = True

    Next
    
    frmMain.grdMonitorDataInput.Select frmMain.grdMonitorDataInput.Row, gridColumnNumber.colDate, nRow, gridColumnNumber.colMonitorPoint
    frmMain.grdMonitorDataInput.Clip = Clipboard.GetText(vbCFText)
    rsNewData.Close
    
    Set rsNewData = Nothing
    Set rsMonitorID = Nothing
    Set rsUpdateData = Nothing
    
    If bCalculate Then
        CalculateMonDataInput
        Call lstDataFilter_Click_Handler
    End If

    'Call lstDataFilter_Click
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "mnuMonitorInputPaste_Click_Handler", Err.Description)
End Sub


'**************************************************************
'   DESCRIPTION
'   Handles the deletion
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub mnuMonitorInputDelete_Click_Handler()

On Error GoTo Error_Handler

    OnRowDelete

    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "mnuMonitorInputDelete_Click_Handler", Err.Description)
    
End Sub


'**************************************************************
'   DESCRIPTION
'   Form Clean-up subroutine
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************
Private Sub FormCleanUp()
    
    On Error GoTo Error_Handler
    
    'Save the grid preferences
    thisGrid.SaveGridSettings GRID_SETTINGS_PATH
    
    'Destroy the grid instance
        Set thisGrid = Nothing
        Set rsMonitorPoint = Nothing
        ' Set rsFilterData = Nothing    Jett 4/20/03 OPTIMIZE: don't need this here
        ' Set rsMonPtVal = Nothing      Jett 4/20/03 OPTIMIZE: moved to functions where it is used
        ' Set rsConcPattern = Nothing   Jett 4/20/03 OPTIMIZE: this variable was never used
        'Set rsDelList = Nothing        Jett 4/20/03 OPTIMIZE: never used
        'Set rsMonitorValues = Nothing  Jett 4/20/03 OPTIMIZE: declared in validate edit only
        'Set rsMonitorId = Nothing      Jett 4/20/03 OPTIMIZE: declared in validate edit and paste click only
        ' Set rs = Nothing              Jett 4/20/03 OPTIMIZE: moved to where it was used
        ' Set rsNewData = Nothing
        ' Set rsUpdateData = Nothing    Jett 4/20/03 OPTIMIZE: declared in paste click
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "FormCleanUp", Err.Description)
    
End Sub


'**************************************************************
'   DESCRIPTION
'   Unloads the form
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/26/2002       Cel Gamboa       Initial Revision
'**************************************************************

Public Sub Unload_MonitorInput()
    
    'Clean-up and validate all the controls before leaving
    On Error GoTo Error_Handler
    
    'JNM 05/12/03 - Tell the user that he must click the button
    If Not CanExitMonInput Then
        gbCanChangeTabs = False
    End If
    bDateTimeClick = True
    
    'Validate all the controls before unloading the form
    If frmMain.grdMonitorDataInput.EditWindow <> 0 Then
       ' Cancel = True
       gbCanChangeTabs = False
    End If
    
    'If gbCanChangeTabs Then
    '    cmdValidDate_Click_Handler
    'End If
    
    'JNM 04/23/03 - AT Fix - This is just a waste of time. The data in the grid
    '   is automatically saved once it loses focus.
    'If bTimeChanged Then
        'SaveDateTime (sTime)   JNM
    'End If
    'If bDateChanged Then
        'SaveDateTime (sDate)   JNM
    'End If
    
    'If Not Cancel Then
    If Not gbCanChangeTabs Then
        Call FormCleanUp
    End If
    
    Exit Sub
Error_Handler:
    Call psAbort(MODULE_NAME, "Form_Unload", Err.Description)
    
End Sub



'*************************************************************
'   DESCRIPTION
'   Handles the Click handler of the cmdValidDate button
'
'   Input Parameters: None
'
'   RETURNS - True if the dates are valid
'             False if the dates are invalid
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  May 9, 2003     Joseph Manyaga    Initial Revision
'  May 26, 2003    Jett Gamboa       Forced change the planned Exposure Time
'**************************************************************
Public Sub cmdValidDate_Click_Handler()
    Dim dtStartIntroduction As Date
    Dim dtStartExposure As Date
    Dim dtEndExposure As Date
    
    Dim dActualExposureTime As Double
    
    On Error GoTo Error_Handler
    
    dtStartIntroduction = DateValue(frmMain.dtpStartIntroDate) & " " & _
        TimeValue(frmMain.dtpStartIntroTime)
    dtStartExposure = DateValue(frmMain.dtpStartExposureDate) & " " & _
        TimeValue(frmMain.dtpStartExposureTime)
    dtEndExposure = DateValue(frmMain.dtpEndExposureDate) & " " & _
        TimeValue(frmMain.dtpEndExposureTime)
        
    Select Case ValidateJobDates(dtStartIntroduction, dtStartExposure, dtEndExposure)
        Case "Start Introduction"
            frmMain.dtpStartIntroDate.SetFocus
        Case "Start Exposure"
            frmMain.dtpStartExposureDate.SetFocus
        Case Else
            'Set the click flag to true
            bDateTimeClick = True
            
            dActualExposureTime = DateDiff("n", dtStartExposure, dtEndExposure) / 60
            '--JG 6/03/03 removed rounding off
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
            
            'Enable the other controls since the dates are valid
            EnableMDIControls True
            
            'Reset the modes
            SetModesFromTimes dtStartIntroduction, dtStartExposure, dtEndExposure
            
            If Not (bIsTableEmpty) Then
                'Recalculate
                CalculateMonDataInput
                    
                'Refresh the grid
                Call lstDataFilter_Click_Handler
            End If
    End Select
    
    Exit Sub
    
Error_Handler:
    Call psAbort(MODULE_NAME, "cmdValidDate_Click_Handler", Err.Description)

End Sub

'**************************************************************
'   DESCRIPTION
'   Enables or disables the controls on Monitor Status
'
'   Input Parameters: True if enable
'                     False if disable
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  May 9, 2003     Joseph Manyaga    Initial Revision
'**************************************************************
Public Sub EnableMDIControls(bEnable As Boolean)
    
    On Error GoTo Error_Handler
    
    'frmMain.lstDataFilter.Enabled = bEnable
    'frmMain.grdMonitorDataInput.Enabled = bEnable
    
    'Change the font color so that the users will know that the grid has been disabled
    With frmMain.grdMonitorDataInput
        
        If bEnable Then
            .Cell(flexcpForeColor, 1, 0, .Rows - 1, .Cols - 1) = 0
        Else
            .Cell(flexcpForeColor, 1, 0, .Rows - 1, .Cols - 1) = &H80000011
        End If
        
    End With
    
    Exit Sub
    
Error_Handler:
    Call psAbort(MODULE_NAME, "EnableMDIControls", Err.Description)
    
End Sub


Public Sub dtpEndExposureDate_Change_Handler()
    'we have to disable the grid
    EnableMDIControls False
    gbHasChanged = True
    bDateTimeClick = False
End Sub

Public Sub dtpEndExposureTime_Change_Handler()
    'we have to disable the grid
    EnableMDIControls False
    gbHasChanged = True
    bDateTimeClick = False
End Sub

Public Sub dtpStartExposureDate_Change_Handler()
    'we have to disable the grid
    EnableMDIControls False
    gbHasChanged = True
    bDateTimeClick = False
End Sub

Public Sub dtpStartExposureTime_Change_Handler()
    'we have to disable the grid
    EnableMDIControls False
    gbHasChanged = True
    bDateTimeClick = False
End Sub

Public Sub dtpStartIntroDate_Change_Handler()
    'we have to disable the grid
    EnableMDIControls False
    gbHasChanged = True
    bDateTimeClick = False
End Sub

Public Sub dtpStartIntroTime_Change_Handler()
    'we have to disable the grid
    EnableMDIControls False
    gbHasChanged = True
    bDateTimeClick = False
End Sub

Public Sub grdMonitorDataInput_BeforeMouseDown_Handler(ByVal Button As _
Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, _
Cancel As Boolean)
    
    On Error GoTo Error_Handler
    
    If Not bDateTimeClick Then
        MsgBox LoadResString(IDS_DATETIME_NOCLICK)
        Cancel = True
    Else
        Cancel = False
    End If
    
    Exit Sub
    
Error_Handler:
    Call psAbort(MODULE_NAME, "grdMonitorDataInput_BeforeMouseDown_Handler", Err.Description)
    
End Sub

Public Function CanExitMonInput() As Boolean

    On Error GoTo Error_Handler
    
    If Not bDateTimeClick Then
        MsgBox LoadResString(IDS_DATETIME_NOCLICK)
        CanExitMonInput = False
    Else
        CanExitMonInput = True
    End If
    
    Exit Function
    
Error_Handler:
    Call psAbort(MODULE_NAME, "CanExitMonInput", Err.Description)
    
End Function

        
'**************************************************************
'   DESCRIPTION
'   Updates the planned exposure
'
'   Input Parameters:
'   dActualExposure - the value to set planned exposure to
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  May 25, 2003    Jett Gamboa       Initial Revision
'  April 6, 2004   Jett Gamboa       replaced commas to dot (to handle european formats)
'**************************************************************
Public Sub UpdatePlannedExposure(dActualExposure As Double)

    Dim strSQL As String
    Dim sActualExposure As String
          
    On Error GoTo Error_Handler
    
    sActualExposure = Replace(dActualExposure, ",", ".")
    
    strSQL = "UPDATE JobArea SET planned_exposure = " & sActualExposure & " WHERE deleted = False"
    
    gdbConn.Execute strSQL
    
    Exit Sub
    
Error_Handler:
    Call psAbort(MODULE_NAME, "CheckPlannedExposure", Err.Description)
    
End Sub


