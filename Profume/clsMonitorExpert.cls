VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonitorExpert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------
'*
'* Filename: clsMonitorExpert.cls
'*
'* DESCRIPTION
'* -----------
'* This code will contain the functions and subroutines
'* to be referenced by the different modules and
'* components of the Fumiguide (ProFume) Application.
'*
'* Copyright (c) 2002 Accenture.  All rights reserved.
'-----------------------------------------------------
'*
'* REVISION HISTORY
'* Revision:    0
'* Author:      Gamboa, Maria Cecilia <nai1686>
'* Date:        September 12, 2002
'*
'* 4/20/2003    Jett    Cleaned up memory leaks and inefficient recordset declarations
'* 4/20/2003    Jett    Added debug output (compiled only if EXPERTDEBUGMODE in Project properties
'*                      is set to 1
'* 5/15/2003    Jett    Added code in GetAreaStatus to handle cases where an area has more than
'                       one monitoring point. In this special case, we have to use the average of
'                       the concentration readings. We are now enforcing that the fumigator have
'                       concentration readings taken at the same time for all the monitoring
'                       points.
'* 5/29/2003    Jett    Used StdRound whenever we used our the average
'* 3/18/2004    Jett    Removed use of StdRound when dealing with Concentration. Also changed
'*                      variables that store concentration readings to be of type Double rather
'*                      than long (we need the extra precision) SIMS #2121768
'* 12/29/2004   Yasmin  pass user-defined CT to the profume calculator Getprofumerequired.
'-----------------------------------------------------

Option Explicit
Private Const CLSNAME = "Expert Monitor Class"
Private Const MSETMONITOREXPERT = "SetMonitorExpert"
Private Const MSETJOBAREA = "SetJobArea"
Private Const MGETPOINTSTATUS = "GetPointStatus"
Private Const MGETAREASTATUS = "GetAreaStatus"
Private Const MPOPULATEPESTDATA = "PopulatePestData"
Private Const MGETFIRSTCONCENTRATION = "GetFirstConcentration"
Private Const MGETLASTCONCENTRATION = "GetLastConcentration"
Private Const MSETCONCENTRATION = "SetConcentration"
Private Const MGETPOINT2POINTHLT = "GetPoint2PointHlt"
Private Const MGETHIGHESTMODE = "GetHighestMode"
Private Const MUPDATEJOBAREARECORD = "UpdateJobAreaRecord"
Private Const MUPDATEMONITORPOINTSTATUS = "UpdateMonitorPointStatus"
Private Const MUPDATEMONPTVALUEPT2PTHLT = "UpdateMonPtValuePt2PtHlt"
Private Const MGETACTUALEXPOSURETIME = "GetActualExposureTime"
Private Const MUPDATEJOBAREASTATUS = "UpdateJobAreaStatus"

Dim dpCalc As DPCALCLib.ProFumeCalculator

Private Type PointStatusUDT
    achieved_concentration_time As Double
    achieved_dose As Double
    elapsed_time As Double
    future_target_concentration As Double
    future_concentration As Double
    hlt As Double
    projected_concentration_time As Double
    projected_time As Double
    last_entry As Date         ' dtLastPeakDateTime ****
    start_time As Date         ' dtHltDateTime
    strMsg As String
End Type

Private Type AreaStatusUDT
    achieved_concentration_time As Double
    achieved_dose As Double
    elapsed_time As Double
    future_target_concentration As Double
    future_concentration As Double
    hlt As Double
    projected_concentration_time As Double
    projected_time As Double
    mean_ntc As Double
    add_profume As Double
    status As Double
    strMsg As String
End Type

Dim AreaStatus As AreaStatusUDT
Dim PointStatus As PointStatusUDT
' this converts all commas to period
Private Function DecVal(dNumber)

    DecVal = Replace(dNumber, ",", ".")

End Function

'**************************************************************
'   DESCRIPTION
'       The job and selected pests records are determined though this method.
'       They are then sent to the PopulatePestData function. Also, the
'       SetJobArea function makes use of the GetProFumeRequired function
'       and sets the worst pest in the calculator
'
'   Input Parameters:     strJobId (String) -  Job Id for the record.
'                         strAreaId (String) -  Area Id for the record.

'   RETURNS - (Boolean) Returns True/False.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'  4/20/2003       Jett Gamboa      cleaning up memory leaks
' 12/29/2004       Yasmin Tan       get user-defined CT and pass it on to GetProfumeRequired.
'**************************************************************
Public Function SetJobArea(strJobId As String, strAreaID As String) As Boolean

    ' Declare the variables
    Dim bReturn As Boolean
    Dim rsJob As ADODB.Recordset
    Dim rsJobArea As ADODB.Recordset
    Dim rsJobPest As ADODB.Recordset
    Dim strSQL As String
    Dim dProfumeReq As Double
    Dim dCylinders As Double
    Dim dCo As Double
    Dim dCT As Double
    Dim dDose As Double
    Dim strMsg As String
    Dim strTest As Variant
    Dim dTempUserCT As Double
    
    On Error GoTo Error_Handler
    
    Set rsJob = New ADODB.Recordset
    Set rsJobArea = New ADODB.Recordset
    Set rsJobPest = New ADODB.Recordset
    
    bReturn = True
    
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "clsMonitorExpert: SetJobArea"
        Debug.Print "for Area Id " & strAreaID
    #End If
    
    ' Get the Job recordset (rsJob). Set the SQL statement to be used
    strSQL = "SELECT * FROM Job WHERE Job.id = " & "'" & strJobId & "'"
    
    ' Execute the SQL and store its contents in rsJob
    'OPTIMIZE: rsJob.Open strSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    rsJob.Open strSQL, gdbConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' Check if the recordset is empty.
    ' If the recordset is empty, return FALSE and exit the method
    If (rsJob.EOF And rsJob.BOF) Then
    
        'OPTIMIZE memory leaks
        rsJob.Close
        
        Set rsJob = Nothing
        Set rsJobArea = Nothing
        Set rsJobPest = Nothing
        
        SetJobArea = False
        Exit Function
    End If
    
    'JNM 04/26/03 - AT Fix -Why the heck are we using this SQL statement?!
    '   We need to only get the SPECIFIC job area record!!!!!!!!!!
    ' Set the SQL Statement to be used.
    'strSQL = "SELECT * FROM JobArea WHERE JobArea.job_fkey = " & "'" & strJobId & "'"
    strSQL = "SELECT * FROM JobArea WHERE JobArea.id = " & "'" & strAreaID & "'"
        
    ' Execute the SQL and store its contents in rsJobArea
    'OPTIMIZE rsJobArea.Open strSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    rsJobArea.Open strSQL, gdbConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' Check if the recordset is empty.
    ' If the recordset is empty, return FALSE and exit the method
    If (rsJobArea.EOF And rsJobArea.BOF) Then
        
        'OPTIMIZE memory leaks
        rsJobArea.Close
        Set rsJob = Nothing
        Set rsJobArea = Nothing
        Set rsJobPest = Nothing
        
        SetJobArea = False
        Exit Function
    End If
    
    ' Get the selected pests for the job
    strSQL = "SELECT * FROM jobpest WHERE job_fkey = " & "'" & strJobId & "'"
    
    ' Execute the SQL and store its contents in rsJobPest
    'OPTIMIZE rsJobPest.Open strSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    rsJobPest.Open strSQL, gdbConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' Check if the recordset is empty, if it is return FALSE (SetJobArea = False)
    If (rsJobPest.EOF And rsJobPest.BOF) Then
        bReturn = False
    End If
    
    ' Get the information from the job, selected pests records and send it to dpcalc (Verify if the boolean variable bReturn is TRUE and PopulatePestData is False)
    If (bReturn And Not PopulatePestData(rsJob, rsJobArea, rsJobPest)) Then
    
        'OPTIMIZE memory leaks
        rsJobPest.Close
        Set rsJob = Nothing
        Set rsJobArea = Nothing
        Set rsJobPest = Nothing
    
        SetJobArea = False
        Exit Function
    End If
    
    ' Initialize values for dpCalc
     dProfumeReq = 0
     dCylinders = 0
     dCo = 0
     dCT = 0
     dDose = 0
     strMsg = ""
       
    '12/29/04 - YYT - get user-defined CT from DB, pass this as Target CT and calculate fumigant and initial concentration
    'if user-defined CT is 0, calculator assumes that there was no user-defined CT indicated, and it will compute
    'for the target CT.
    dTempUserCT = IIf(IsNull(rsJobArea("user_defined_CT")), 0, rsJobArea("user_defined_CT"))
    dCT = dTempUserCT
    
    dpCalc.GetProfumeRequired dProfumeReq, dCylinders, dCo, dCT, dDose, strMsg
    SetJobArea = bReturn
    
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "          DPCALC: GetProFumeRequired " & dProfumeReq & ", " & dCylinders & ", " & dCo & ", " & dCT & ", " & dDose & ", " & strMsg; ""
    #End If
    
    'OPTIMIZE memory leaks
    rsJob.Close
    rsJobArea.Close
    rsJobPest.Close
    
    Set rsJob = Nothing
    Set rsJobArea = Nothing
    Set rsJobPest = Nothing
    
    Exit Function
Error_Handler:
    Call psAbort(CLSNAME, MSETJOBAREA, Err.Description)

End Function

'**************************************************************
'   DESCRIPTION
'       After the point-to-point half-loss-time has been determined for the
'       last concentration/time value, the status of the monitor point needs
'       to be determined.  Again, the existing concentration/time values must
'       be cleared from the calculator.  All of the concentration/time values
'       that occur during exposure have to be added to the calculator.
'       These values are used to determine how many doses have been achieved
'       for the fumigation as of the last concentration/time value.
'
'   Input Parameters: lpMonitorPointId (String)
'                  dtStartIntro (Date)
'                  dtStartExposure (Date)
'                  dtEndExposure (Date)
'                  lpPointStatus (String)
'   Returns:           (Boolean) Returns True/False
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'  5/28/2003       Malou Espaldon   Profume Enhancements 2003
'                  Forced Mean Projected CT to always round up if decimal part is
'                  exactly .5.  VB uses the Computer Rule when rounding off numbers.
'
'**************************************************************
Public Function GetPointStatus(lpMonitorPointId As String, _
                            dtStartIntro As Date, _
                            dtStartExposure As Date, _
                            dtEndExposure As Date, _
                            Optional bAreaAverage As Boolean = False) As Boolean
                            
    Dim strSQL As String
    Dim lConcFirst As Double
    Dim lConcLast As Double
    Dim bFirstExp As Boolean
    Dim bLastExp As Boolean
    Dim LastPeakDateTime As Date
    Dim HltDateTime As Date
    Dim dtSpan As Double
    Dim dElapsedTime As Double
    Dim estimated_hlt As Double
    Dim dAchievedCT As Double
    Dim dAchievedDose As Double
    Dim dHLT As Double
    Dim dProjectedCT As Double
    Dim sProjectedCT As String
    Dim dProjectedTime As Double
    Dim dFutureTargetConcentration As Double
    Dim dFutureConcentration As Double
    Dim bstrMsg As String
    Dim strJobId As String
    
    Dim rsMonPtVal As ADODB.Recordset
    Dim rsJobArea As ADODB.Recordset
    
    On Error GoTo Error_Handler
    
    Set rsMonPtVal = New ADODB.Recordset
    Set rsJobArea = New ADODB.Recordset
    
    ' Clear the concentrations in dpcalc
    Call dpCalc.ClearConcentrations
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "clsMonitorExpert: GetPointStatus"
        Debug.Print "Getting Monitor Point Status for " & lpMonitorPointId
        Debug.Print "          DPCALC: ClearConcentrations"
    #End If
    
    ' 05/15/2003 Jett - If we are getting the status of an area with monitoring points
    ' use a different SQL statement to retrieve the average of the concentration readings
    If bAreaAverage Then
        strSQL = "SELECT MonitorPointValue.date_time, Avg(MonitorPointValue.fumiscope_conc) AS fumiscope_conc, MonitorPointValue.mode, -1 AS hlt_pt2pt " & _
                 "FROM MonitorPoint INNER JOIN MonitorPointValue ON MonitorPoint.id = MonitorPointValue.monitorpoint_fkey " & _
                 "WHERE (((MonitorPointValue.deleted) = False) And ((MonitorPoint.deleted) = False) And ((MonitorPointValue.fumiscope_conc) > -1)) " & _
                 "AND MonitorPoint.jobarea_fkey = '" & lpMonitorPointId & "' " & _
                 "GROUP BY MonitorPointValue.date_time, MonitorPointValue.mode " & _
                 "ORDER BY MonitorPointValue.date_time, Avg(MonitorPointValue.fumiscope_conc) DESC"
    Else
        ' Get all the records associated with the monitor point
        strSQL = "SELECT * FROM MonitorPointValue, MonitorPoint WHERE MonitorPointValue.deleted = FALSE And MonitorPointValue.monitorpoint_fkey = '" & lpMonitorPointId & "' AND MonitorPoint.id = MonitorPointValue.monitorpoint_fkey ORDER BY MonitorPointValue.date_time ASC, MonitorPointValue.fumiscope_conc DESC"
    End If
    
    ' Execute the SQL and store its contents in rsMonPtVal
    'OPTIMIZE rsMonPtVal.Open strSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    ' this recordset is only being read so we do a forward only
    rsMonPtVal.Open strSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText
    
    ' Check if there are any records returned.
    ' If the recordset is empty set GetPointStatus to True and exit the method.
    If (rsMonPtVal.BOF And rsMonPtVal.EOF) Then
        'OPTIMIZE memory leaks
        rsMonPtVal.Close
        Set rsMonPtVal = Nothing
        Set rsJobArea = Nothing
        
        GetPointStatus = False
        Exit Function
    End If
    
    lConcFirst = 0
    bFirstExp = False
    
    ' Iterate through the records and find the first concentration in the exposure
    ' mode
    
    Call GetFirstConcentration(rsMonPtVal, dtStartIntro, dtStartExposure, lConcFirst, bFirstExp)
    
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "Adding Concentration readings for MonitorPoint " & lpMonitorPointId
        Debug.Print "Setting First Concentration"
    #End If
    
    ' Add this value to dpcalc.  This should be the first item sent to dpcalc
    If (Not SetConcentration(lConcFirst, dtStartExposure)) Then
        'OPTIMIZE memory leaks
        rsMonPtVal.Close
        Set rsMonPtVal = Nothing
        Set rsJobArea = Nothing
        
        GetPointStatus = False
        'Add the exposure concentration/time values to dpcalc ****
        Exit Function
    End If
    
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "Setting concentration readings"
    #End If
    
    rsMonPtVal.MoveFirst
    While Not rsMonPtVal.EOF
    
    If (rsMonPtVal("mode") = MDM_EXPOSURE And rsMonPtVal("fumiscope_conc") >= 0) Then
    
        If Not bFirstExp Then
        ' Add this value to dpcalc.  This should be the first item sent to dpcalc
        
        '--MME 5/28.  called our Standard Rounding function
             If (Not SetConcentration(rsMonPtVal("fumiscope_conc"), rsMonPtVal("date_time"))) Then
             
                'OPTIMIZE memory leaks
                rsMonPtVal.Close
                Set rsMonPtVal = Nothing
                Set rsJobArea = Nothing
             
                GetPointStatus = False
                Exit Function
             End If
        ' Get the elapsed exposure time
            dtSpan = DateDiff("s", dtStartExposure, rsMonPtVal("date_time"))
            PointStatus.elapsed_time = CDbl(dtSpan / 3600)  '***3600 seconds/hour
        Else
             bFirstExp = False
        End If
    End If
    rsMonPtVal.MoveNext
    Wend
    
    ' Jett 5/29/03 - added extra parameter because we modified GetLastConcentration
    Call GetLastConcentration(rsMonPtVal, dtStartExposure, dtEndExposure, lConcLast, bLastExp, bAreaAverage, lConcFirst)
    If bLastExp Then
    
        #If EXPERTDEBUGMODE = 1 Then
          Debug.Print "Setting last concentration"
        #End If
        
        'Add the end exposure concentration, if we have reached that point in the fumigation
        If (Not SetConcentration(lConcLast, dtEndExposure)) Then
            GetPointStatus = False
            
            Set rsJobArea = Nothing
            Set rsMonPtVal = Nothing
            
            Exit Function
        End If
        
        lConcLast = "0"
        bLastExp = False '****
        
        ' Get the elapsed exposure time since we are at the end of exposure
        dtSpan = DateDiff("s", dtStartExposure, dtEndExposure)
        'dElapsedTime = dtSpan.GetTotalSeconds / 3600   '3600 seconds/hour ****
        PointStatus.elapsed_time = CDbl(dtSpan / 3600)
    End If
    
    ' Clean up all objects used
    ' Get the monitor point status from dpcalc
    LastPeakDateTime = 0
    HltDateTime = 0
    strJobId = gsJobID
    
    ' Set the SQL Statement to be used.
    ' JETT 3/8/04 SIMS #2121224 - the old sql statement was always returning the estimated
    ' HLT of the first area. This results in erroneous status/recommendations if the HLT
    ' for areas are not the same as the first one
    'strSQL = "SELECT estimated_hlt FROM JobArea WHERE JobArea.job_fkey = " & "'" & strJobId & "'"
    
    ' If we called this to get the Point Status for an area with multiple monitoring points
    ' lpMonitorPoint actually contains the Area id so we use a different SQL statement
    If bAreaAverage Then
        strSQL = "SELECT estimated_hlt " & _
                 "FROM JobArea " & _
                 "WHERE JobArea.id = '" & lpMonitorPointId & "'"
    Else
        strSQL = "SELECT estimated_hlt " & _
                 "FROM JobArea, MonitorPoint " & _
                 "WHERE MonitorPoint.jobarea_fkey = JobArea.id AND " & _
                 "MonitorPoint.id = " & "'" & lpMonitorPointId & "'"
    End If
    
    ' Execute the SQL and store its contents in rsJobArea
    rsJobArea.Open strSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    ' Check if the recordset is empty.
    ' If the recordset is empty, return FALSE and exit the method
    If (rsJobArea.EOF And rsJobArea.BOF) Then
        estimated_hlt = -1
    Else
        estimated_hlt = CDbl(rsJobArea("estimated_hlt"))
    End If
    
    dAchievedCT = PointStatus.achieved_concentration_time
    dAchievedDose = PointStatus.achieved_dose
    dHLT = PointStatus.hlt
    dProjectedCT = PointStatus.projected_concentration_time
    dProjectedTime = PointStatus.projected_time
    dFutureTargetConcentration = PointStatus.future_target_concentration
    dFutureConcentration = PointStatus.future_concentration
    dElapsedTime = PointStatus.elapsed_time
    
    dpCalc.GetMonitorPointStatus estimated_hlt, _
                                    dAchievedCT, _
                                    dAchievedDose, _
                                    dHLT, _
                                    dProjectedCT, _
                                    dProjectedTime, _
                                    dFutureTargetConcentration, _
                                    dFutureConcentration, _
                                    dElapsedTime, _
                                    LastPeakDateTime, _
                                    HltDateTime, _
                                    bstrMsg
                                    
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "          DPCALC: GetMonitorPointStatus " & estimated_hlt & ", " & _
                     dAchievedCT & ", " & dHLT & ", " & dProjectedCT & ", " & dProjectedTime & ", " & _
                     dFutureTargetConcentration & ", " & dFutureConcentration & ", " & dElapsedTime & _
                     ", " & LastPeakDateTime & ", " & HltDateTime & ", " & bstrMsg
    #End If
    
    PointStatus.achieved_concentration_time = dAchievedCT
    PointStatus.achieved_dose = dAchievedDose
    PointStatus.hlt = dHLT
    sProjectedCT = dProjectedCT
    
    If sProjectedCT = "-1.#IND" Then
        dProjectedCT = PointStatus.projected_concentration_time
    End If
    
    PointStatus.projected_concentration_time = dProjectedCT
    PointStatus.projected_time = dProjectedTime
    PointStatus.future_target_concentration = dFutureTargetConcentration
    PointStatus.future_concentration = dFutureConcentration
    PointStatus.elapsed_time = dElapsedTime
    PointStatus.last_entry = LastPeakDateTime
    PointStatus.start_time = HltDateTime
    
    If (bstrMsg <> "") Then
        ' Show the error string that was returned in bstrMsg
        MsgBox bstrMsg, vbExclamation
        
        'OPTIMIZE memory leaks
        rsMonPtVal.Close
        rsJobArea.Close
        Set rsMonPtVal = Nothing
        Set rsJobArea = Nothing
        
        GetPointStatus = False
        Exit Function
    End If
    
    GetPointStatus = True
    
    'OPTIMIZE memory leaks
    rsMonPtVal.Close
    rsJobArea.Close
    Set rsMonPtVal = Nothing
    Set rsJobArea = Nothing
    
    Exit Function
Error_Handler:
    Call psAbort(CLSNAME, MGETPOINTSTATUS, Err.Description)
End Function

'**************************************************************
' Description:
'   After all of the monitor point status' has been determined for an area,
'   the area status needs to be determined.  The existing area monitor point
'   status values must be cleared from the calculator.
'      Input Parameters: strJobId (String) - Job Id for the Record
'                        strAreaId (String) - Area Id for the Record
'                        strAreaStatus (String) - Area Status of the record.
'      Output:           (Boolean) Returns True/False
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'  5/15/2003       Jett Gamboa      added date parameters, we need them for our special
'                                   handling of multiple monitoring points in an area. Also added
'                                   special handling for Areas with multiple monitoring points.
'  7/30/2003       Jett Gamboa      SIMS #2100090: Added a call to SetJobArea before calling GetAreaStatus
'                                   The way the program flows, SetJobArea is called everytime we try to determine
'                                   the monitor point status. Among other things, SetJobArea specifieds the volume
'                                   for the area. Since we process all monitor point status before Area status, the
'                                   volume for the last monitor point process will be used for the AddGas calculation
'                                   in dpcalc. This caused incorrect results. We need to make a call so that the correct
'                                   area volume is set.
'**************************************************************

Public Function GetAreaStatus(strJobId As String, _
                              strAreaID As String, _
                              dtStartIntro As Date, dtStartExposure As Date, dtEndExposure As Date) As Boolean
    ' Declare all variables to be used
    Dim rsMonPt As ADODB.Recordset
    Dim strSQL As String
    Dim nMonitorPointsAdded As Integer
    Dim rsJobAreas As ADODB.Recordset
    Dim strMsg As String
    Dim bstrMsg As String
    Dim dAchievedCT As Double
    Dim dAchieved_dose As Double
    Dim dElapsed_time As Double
    Dim dMeanNTC As Double
    Dim dAddProfume As Double
    Dim dHLT As Double
    Dim dProjectedCT As Double
    Dim dProjectedTime As Double
    Dim lStatus As Long
    Dim dConTime As Double
    
    On Error GoTo Error_Handler
    
    Set rsJobAreas = New ADODB.Recordset
    Set rsMonPt = New ADODB.Recordset
    
    ' Clear the monitor points in dpcalc
    Call dpCalc.ClearAreaMonitorPoints
     
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "clsMonitorExpert: GetAreaStatus"
        Debug.Print "Getting Area Status for Area id " & strAreaID
        Debug.Print "          DPCALC: ClearAreaMonitorPoints"
    #End If
     
    ' Get all the records associated with the monitor point
    strSQL = "SELECT * FROM MonitorPoint WHERE MonitorPoint.deleted = FALSE " _
             & "And MonitorPoint.jobarea_fkey = '" & strAreaID & "' " _
             & "And MonitorPoint.job_fkey = '" & strJobId & "'"
    ' And JobArea.id = MonitorPoint.jobarea_fkey
    
    'Execute the SQL and store its contents in rsMonPt
    rsMonPt.Open strSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    'Check if there are any records returned
    If (rsMonPt.BOF And rsMonPt.EOF) Then
       GetAreaStatus = True
       Set rsJobAreas = Nothing
       Set rsMonPt = Nothing
       Exit Function
    End If
    
    nMonitorPointsAdded = 0
    
    ' 5/15/2003 Jett: New approach if there are multiple monitoring points for an area
    ' If there are more than 1 monitoring points then we must do our special processing
    ' to average.
    If rsMonPt.RecordCount > 1 Then
    
        ' Call the function to check if we have complete readings for all monitor
        ' points.
        If CheckMultipleMonPts(strAreaID, rsMonPt.RecordCount) = False Then
        
        #If EXPERTDEBUGMODE = 1 Then
            Debug.Print "** Missing Concentration Readings for this Area"
            Debug.Print "** We will not calculate Area status"
        #End If
       
            ' This is the only time we force dpcalc to return a specific status.
            ' We have to do this because this is where we validate and not in
            ' dpcalc.
            lStatus = STATUS_MISSING_CONC_READING
        
            ' just to be consistent we still have dpcalc return the status
            dpCalc.GetAreaStatus dConTime, _
                     dAchievedCT, _
                     dAchieved_dose, _
                     dHLT, _
                     dProjectedCT, _
                     dProjectedTime, _
                     dMeanNTC, _
                     lStatus, _
                     dAddProfume, _
                     strMsg
                     
            AreaStatus.achieved_concentration_time = dAchievedCT
            AreaStatus.achieved_dose = dAchieved_dose
            AreaStatus.hlt = dHLT
            AreaStatus.projected_concentration_time = dProjectedCT
            AreaStatus.projected_time = dProjectedTime
            AreaStatus.add_profume = dAddProfume
            AreaStatus.projected_time = dProjectedTime
            AreaStatus.mean_ntc = dMeanNTC
            AreaStatus.add_profume = dAddProfume
            AreaStatus.status = lStatus
                     
            GetAreaStatus = True
            
            ' cleanup and let's get outta here
            rsMonPt.Close
            
            Set rsMonPt = Nothing
            Set rsJobAreas = Nothing
            
            ' note that we do not call SetAreaMonitorPoints. By doing this dpcalc
            ' will know that it cannot calculate the area status
            Exit Function
        
        End If
    
    
        #If EXPERTDEBUGMODE = 1 Then
            Debug.Print "** Special handling for areas with more than 1 monitoring point"
            Debug.Print "** Calling GetPointStatus to use the Average concentration reading"
        #End If
    
        ' Initialize the PointStatus structure members, GetPointStatus will return the
        ' status of our "averaged" monitor point concentration readings here.
        PointStatus.achieved_concentration_time = -1
        PointStatus.achieved_dose = -1
        PointStatus.elapsed_time = -1
        PointStatus.future_target_concentration = -1
        PointStatus.future_concentration = -1
        PointStatus.hlt = -1
        PointStatus.projected_concentration_time = -1
        PointStatus.projected_time = -1
        PointStatus.last_entry = 0
        PointStatus.start_time = 0
        PointStatus.strMsg = ""
    
        ' Call GetPointStatus and pass true for the last variable because we want to
        ' average the concentration readings for the areas
        GetPointStatus strAreaID, dtStartIntro, dtStartExposure, dtEndExposure, True
        
        #If EXPERTDEBUGMODE = 1 Then
            Debug.Print "** Setting AreaMonitorPoint with the result of our averages"
        #End If
        
        dpCalc.SetAreaMonitorPoint PointStatus.achieved_concentration_time, _
                                   PointStatus.achieved_dose, _
                                   PointStatus.elapsed_time, _
                                   PointStatus.future_target_concentration, _
                                   PointStatus.future_concentration, _
                                   PointStatus.hlt, _
                                   PointStatus.projected_concentration_time, _
                                   PointStatus.projected_time, _
                                   bstrMsg
                                   
        ' If there was something returned in our string (this shouldn't happen), we need
        ' to let the user know that this abnormal condition occurred
        If (bstrMsg <> "") Then
            MsgBox bstrMsg, vbExclamation
            GetAreaStatus = False
            Set rsJobAreas = Nothing
            Set rsMonPt = Nothing
            Exit Function
        End If
    
        nMonitorPointsAdded = 1
        
    Else
   
        ' Iterate through the records and gather their information
        ' 5/15/2003 Jett - We will never iterate through this loop more than once, we already
        ' handle multiple monitoring points above :)
        rsMonPt.MoveFirst
        While Not rsMonPt.EOF
                
            ' Check if there is an achieved concentration-time calc'd for the monitor point
            If (rsMonPt("achieved_concentration_time") >= 0) Then
            
                ' Add the monitor point in order to calc the area status
                dpCalc.SetAreaMonitorPoint rsMonPt("achieved_concentration_time"), _
                                        rsMonPt("achieved_dose"), _
                                        rsMonPt("elapsed_time"), _
                                        rsMonPt("future_target_concentration"), _
                                        rsMonPt("future_concentration"), _
                                        rsMonPt("hlt"), _
                                        rsMonPt("projected_concentration_time"), _
                                        rsMonPt("projected_time"), _
                                        bstrMsg
                                        
                #If EXPERTDEBUGMODE = 1 Then
                    Debug.Print "SetAreaMonitorPoint for Monitor Point " & rsMonPt("id")
                    Debug.Print "          DPCALC: SetAreaMonitorPoint " & rsMonPt("achieved_concentration_time") & ", " & _
                                        rsMonPt("achieved_dose") & ", " & _
                                        rsMonPt("elapsed_time") & ", " & _
                                        rsMonPt("future_target_concentration") & ", " & _
                                        rsMonPt("future_concentration") & ", " & _
                                        rsMonPt("hlt") & ", " & _
                                        rsMonPt("projected_concentration_time") & ", " & _
                                        rsMonPt("projected_time") & ", " & _
                                        bstrMsg
    
                #End If
                
                If (bstrMsg <> "") Then
                    MsgBox bstrMsg, vbExclamation
                    GetAreaStatus = False
                    Set rsJobAreas = Nothing
                    Set rsMonPt = Nothing
                    Exit Function
                End If
                 
                nMonitorPointsAdded = nMonitorPointsAdded + 1
                
            End If
            
            rsMonPt.MoveNext
            
        Wend
        
    End If
    
    If (nMonitorPointsAdded <= 0) Then
    
      ' OPTIMIZE memory leaks
      rsMonPt.Close
      Set rsMonPt = Nothing
      Set rsJobAreas = Nothing
    
      ' There were no monitor points so no status check will be done
      GetAreaStatus = True
                
      Exit Function '***
    End If
      
    ' Get target concentration-time from the JobArea table
    strSQL = "SELECT * FROM JobArea WHERE JobArea.job_fkey = '" & strJobId & "' " _
             & "And JobArea.id = '" & strAreaID & "' " _
             & "AND JobArea.deleted = FALSE"
    
    ' Execute the SQL and store its contents in rsJobAreas
    'OPTIMIZE rsJobAreas.Open strSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    rsJobAreas.Open strSQL, gdbConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' Check if there are any records returned
    If (rsJobAreas.BOF And rsJobAreas.EOF) Then
    
       'OPTIMIZE memory leaks
       rsJobAreas.Close
       Set rsJobAreas = Nothing
    
       GetAreaStatus = True
       Exit Function '***
    End If
       
    ' There should only be 1 JobArea record
    ' Move to the first record in the recordset
    'OPTIMIZE rsJobAreas.MoveFirst
    
    lStatus = rsJobAreas("status")
    dAchievedCT = rsJobAreas("achieved_concentration_time")
    dAchieved_dose = rsJobAreas("achieved_dose")
    dMeanNTC = rsJobAreas("mean_ntc")
    dAddProfume = rsJobAreas("add_profume")
    dHLT = rsJobAreas("hlt")
    dProjectedCT = rsJobAreas("projected_concentration_time")
    dProjectedTime = rsJobAreas("projected_time")
    dConTime = rsJobAreas("concentration_time")
    
    ' 7/30/03 Jett - calling the SetJobArea method so that we can set the the volume of the area
    ' for the area we are currently processing. (Of course, SetJobArea does other things that we
    ' do not care about at this point. We are forced to do that to implement a simple fix.
    ' SIMS #2100090
    Me.SetJobArea gsJobID, strAreaID
    
    dpCalc.GetAreaStatus dConTime, _
                         dAchievedCT, _
                         dAchieved_dose, _
                         dHLT, _
                         dProjectedCT, _
                         dProjectedTime, _
                         dMeanNTC, _
                         lStatus, _
                         dAddProfume, _
                         strMsg
                         
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "          DPCALC: GetAreaStatusX " & dConTime & ", " & _
                         dAchievedCT & ", " & _
                         dAchieved_dose & ", " & _
                         dHLT & ", " & _
                         dProjectedCT & ", " & _
                         dProjectedTime & ", " & _
                         dMeanNTC & ", " & _
                         lStatus & ", " & _
                         dAddProfume & ", " & _
                         strMsg
    #End If
                         
    AreaStatus.achieved_concentration_time = dAchievedCT
    AreaStatus.achieved_dose = dAchieved_dose
    AreaStatus.hlt = dHLT
    AreaStatus.projected_concentration_time = dProjectedCT
    AreaStatus.projected_time = dProjectedTime
    AreaStatus.add_profume = dAddProfume
    AreaStatus.projected_time = dProjectedTime
    AreaStatus.mean_ntc = dMeanNTC
    AreaStatus.add_profume = dAddProfume
    AreaStatus.status = lStatus
                         
    ' Get the area status from the dpCalc
    If (strMsg <> "") Then
        MsgBox strMsg, vbExclamation
        
        ' OPTIMIZE Memory leaks
        rsMonPt.Close
        rsJobAreas.Close
        
        Set rsMonPt = Nothing
        Set rsJobAreas = Nothing

        
        GetAreaStatus = False
        Exit Function
    End If
    
    ' Destroy all objects that have been used.
    'OPTIMIZE
    rsMonPt.Close
    rsJobAreas.Close
    
    Set rsMonPt = Nothing
    Set rsJobAreas = Nothing
    
    Exit Function
Error_Handler:
    Call psAbort(CLSNAME, MSETMONITOREXPERT, Err.Description)
End Function


'**************************************************************
'This method iterates through the pests and places them into the
'calculator Since the Calculator only works in metric, we will be
'using the SetDosePest function to convert this values.
'Input Parameters:  rsJob (Recordset)
'                             rsJobArea (Recordset)
'                             rsJobPest (Recordset)
'Output                : (Boolean) Returns True/False
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'**************************************************************

' OPTIMIZE: no need to, recordsets are being passed to this function
Public Function PopulatePestData(rsJob As Recordset, rsJobArea As Recordset, rsJobPest As Recordset) As Boolean

    Dim bstrMsg As String
    Dim IsComplete As Integer 'PAT 1/13/05 - Added
    Dim intCounter As Integer 'PAT 1/27/05 - Added
    On Error GoTo Error_Handler
    
    'First we need to clear the existing pest list in the calculator
    Call dpCalc.ClearDosePest
    
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "clsMonitorExpert: PopulatePestData"
        Debug.Print "          DPCALC: ClearDosePest"
    #End If
    
    ' Move to the first record in rsJobPest Recordset
    rsJobPest.MoveFirst
    
    ' Since have some data, let's iterate through the pests and place them
    ' into the calculator.
    While Not rsJobPest.EOF
    'Only work with the record if the delete flag
    ' has not been set
    
          If (Not rsJobPest("deleted")) Then
            ' Calculator only works in metric, ensure all values are metric
            dpCalc.SetDosePest rsJobPest("pest_enum"), _
                              rsJob("fumigation_type_enum"), _
                              rsJob("lifestage_enum"), _
                              rsJob("treatment_type_enum"), _
                              rsJob("commodity_enum"), _
                              rsJobArea("estimated_hlt"), _
                              rsJobArea("temperature"), _
                              rsJobArea("planned_exposure"), _
                              rsJobArea("volume"), _
                              rsJob("Load_Factor"), _
                              IsComplete, _
                              intCounter, _
                              bstrMsg
                              
            #If EXPERTDEBUGMODE = 1 Then
                Debug.Print "          DPCALC: SetDosePest " & rsJobPest("pest_enum") & ", " & _
                              rsJob("fumigation_type_enum") & ", " & _
                              rsJob("lifestage_enum") & ", " & _
                              rsJob("treatment_type_enum") & ", " & _
                              rsJob("commodity_enum") & ", " & _
                              rsJobArea("estimated_hlt"), _
                              rsJobArea("temperature") & ", " & _
                              rsJobArea("planned_exposure") & ", " & _
                              rsJobArea("volume") & ", " & _
                              rsJob("Load_Factor") & ", " & _
                              IsComplete & ", " & _
                              intCounter & ", " & _
                              bstrMsg
            #End If
                              
            If bstrMsg <> "" Then
               MsgBox bstrMsg, vbExclamation
               PopulatePestData = False
               Exit Function
            End If
            
        End If
        
        rsJobPest.MoveNext
    
    Wend
    
    PopulatePestData = True
    ' Destroy all objects that have been used.
    
    Exit Function
Error_Handler:
    Call psAbort(CLSNAME, MPOPULATEPESTDATA, Err.Description)

End Function

'**************************************************************
' Description:
'There are several ways the first concentration at the start of exposure
'may be retrieved from the monitor point's concentration/time record set.
'First, a concentration may be entered at the start of exposure.
'Second, concentrations may be entered during the introduction of ProFume
'and after the start of exposure.  Third, no concentration was entered
'during the introduction of ProFume, however, a concentration was added
'after the start of exposure.  The last two items require interpolation
'in order to determine what the concentration value would be at the start
'of exposure.
'Input Parameters:       rsMonPtVal (Recordset)
'                                   dtStartIntro (Date)
'                                   dtStartExposure (Date)
'                                   lConcentration  (Long)
'                                   bFirstExp (Boolean)
'Output:                       None
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'  5/29/2003       Jett Gamboa      Changed concentration to double
'**************************************************************

'OPTIMIZE no need to
Public Sub GetFirstConcentration(rsMonPtVal As Recordset, _
                                   dtStartIntro As Date, _
                                   dtStartExposure As Date, _
                                   lConcentration As Double, _
                                   bFirstExp As Boolean)
    Dim lConcExp As Double ' jett 3/15/04 should be double from long
    Dim dtDateTimeExp As Date
    Dim lConcPrev As Double ' jett 3/15/04 should be double from long
    Dim dtDateTimePrev As Date
    Dim bFoundPrev As Boolean
    Dim dtspanExp As Long
    Dim dSeconds As Double
    Dim dtspanStart As Long
    Dim dTimeRatio As Double
    
    On Error GoTo Error_Handler
    
    'Initialize the values of the variables
    lConcExp = 0
    dtDateTimeExp = 0
    lConcPrev = 0
    dtDateTimePrev = 0
    bFoundPrev = False
    
    ' Move to the first record in the recordset.
    rsMonPtVal.MoveFirst
    
    Do While Not rsMonPtVal.EOF
    ' Iterate through the records and find the first concentration in the exposure
    ' mode. Check the mode and Fumiscope Concentration
    
    If (rsMonPtVal("mode") = MDM_EXPOSURE And rsMonPtVal("fumiscope_conc") >= 0) Then
        
        ' Get the current concentration and date/time
        lConcExp = rsMonPtVal("fumiscope_conc")
        dtDateTimeExp = rsMonPtVal("date_time")
        
        'Check the time we may already be at the value for starting exposure
             If (rsMonPtVal("date_time") = dtStartExposure And rsMonPtVal("fumiscope_conc") >= 0) Then
                lConcPrev = rsMonPtVal("fumiscope_conc")
                dtDateTimePrev = rsMonPtVal("date_time")
                bFoundPrev = True
                bFirstExp = True
             End If
             
        ' Move to the previous value
        rsMonPtVal.MovePrevious
    
        ' Get the previous concentration and date/time
        If Not bFoundPrev Then
           Do While Not rsMonPtVal.BOF
        
            ' Get previous monitor point value where we have a fumiscope value
              If rsMonPtVal("fumiscope_conc") >= 0 Then
                lConcPrev = rsMonPtVal("fumiscope_conc")
                dtDateTimePrev = rsMonPtVal("date_time")
                bFoundPrev = True
                Exit Do
              End If
            ' Move to the previous record
              rsMonPtVal.MovePrevious
              Loop
        End If
        Exit Do
    
    End If
    
    ' Move to the next record in the recordset
    rsMonPtVal.MoveNext
    
    Loop
    
    ' If the previous value was not found, then use the intro date/time
    ' for the previous date/time
    
       If (Not bFoundPrev) Then
    
        ' The concentration should be 0.0 at the start of ProFume introduction
        ' and the previous concentration has been initialized to 0.0 so we are ok
          lConcExp = 0
          dtDateTimePrev = dtStartIntro
       End If
       
    ' Get the (estimated) concentration/time at the start of exposure using a straight
    ' line interprelation between the exposure date/time and the previous date/time
    
    'Get the difference between the two dates (in seconds)
    If (dtDateTimeExp = "12:00:00 AM" Or dtDateTimePrev = "12:00:00 AM") Then
    dtspanExp = 0
    Else
    dtspanExp = DateDiff("s", dtDateTimePrev, dtDateTimeExp)
    End If
    
    'dtspanExp = dtDateTimeExp - dtDateTimePrev
    
    ' Get the total hour difference for the Exposure Span Date
       dSeconds = dtspanExp / 3600
       
       If (dSeconds = 0) Then
             lConcentration = lConcExp
       Else
          dtspanStart = DateDiff("s", dtDateTimePrev, dtStartExposure)
          'dtspanStart = dtStartExposure - dtDateTimePrev
          ' Get the Time Ratio by Dividing the Span Start Total Seconds by the Span Exposure total Seconds
          dTimeRatio = (dtspanStart / 3600) / (dtspanExp / 3600)
          lConcentration = ((lConcExp - lConcPrev) * dTimeRatio) + lConcPrev
       End If
    
    Exit Sub
Error_Handler:
    Call psAbort(CLSNAME, MGETFIRSTCONCENTRATION, Err.Description)

End Sub

'**************************************************************
'There are several ways the last concentration at the end of exposure may
'be retrieved from the monitor point's concentration/time record set.
'First, a concentration may be entered at the end of exposure.
'Second, no concentration was entered after the end
'of exposure, however, a concentration was added during exposure.
'Last, no HLT exists for the concentration during the exposure.
'Item two  require extrapolation in order to determine what
'the concentration value would be at the end of exposure.
'Input Parameters:  rsMonPtVal (Recordset)
'                   dtEndExposure (Date)
'                   lConcentration (Long)
'                   bLastExp (Boolean)
'
'Output:                       None
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'  4/2/2003        Amana Martinez   Fumiguide Enhancements 2003
'                                   Removed the original logic in the code
'                                   And replace it with a new one
'  4/16/2003       J. Manyaga       AT fixes
'  5/29/2003       Jett Gamboa      Added some extra parameters to the function. Added code to
'                                   calculate for the HLT of the reading prior to end of exposure
'                                   so we can extrapolate the concentration reading at end of
'                                   exposure.
'  9/17/2003       Jett Gamboa      Fix for SIMS #2104694. Added check if concentration prior to
'                                   End of Exposure is 0. If it is, set End Exposure concentration
'                                   to 0. We cannot extrapolate because we cannot get the ln of 0.
'**************************************************************

' OPTIMIZE no need to, recordset is created in calling function
Public Sub GetLastConcentration(rsMonPtVal As Recordset, _
                                 dtStartExposure As Date, _
                                 dtEndExposure As Date, _
                                 lConcentration As Double, _
                                 bLastExp As Boolean, _
                                 bAreaAverage As Boolean, _
                                 lConcFirst As Double)
    Dim lConcExp As Double ' jett 3/15/04 should be double from long
    Dim dtDateTimeExp As Date
    Dim lConcPrev As Double ' jett 3/15/04 should be double from long
    Dim dtDateTimePrev As Date
    Dim dHLT As Double
    Dim dHLTTemp As Double
    Dim dtspanExp As Long
    Dim dSeconds As Long
    Dim dTime As Double
    Dim bFoundPrev As Boolean
    
    On Error GoTo Error_Handler
     
    ' Check the end exposure date/time
    ' If 0.0 or max double, then we haven't reached the end of exposure
    ' (start of pre-aeration)
    'JNM 04/21/03 - AT Fix - This condition will be true only if the user has
    '   not yet entered the End Exposure Date. But the original use for this code
    '   is to check if there are points on or beyond the End Exposure time. Therefore,
    '   in order to still achieve the same intention, we will just check the mode of the
    '   last point.
       'If (dtEndExposure = 0) Then
       '     bLastExp = False
       '     Exit Sub
       'End If
       rsMonPtVal.MoveLast
       If rsMonPtVal("mode") <> MDM_PRE_AERATION Then
            bLastExp = False
            Exit Sub
       End If
       
       'JNM 04/16 - If we have passed the line above, this means that there are points on or
       '    beyond the End Exposure time. Up next, we have to get the actual conc. at
       '    end exposure time or we have
       '    to extraplolate it using the last pt. directly before end exposure time.
           
    ' Initialize the Variables to be used
    lConcExp = 0
    dtDateTimeExp = 0
    lConcPrev = 0
    dtDateTimePrev = 0
    dHLT = 0
    dHLTTemp = 0
    bFoundPrev = False
    
    'Move to the first record in the recordset
    rsMonPtVal.MoveFirst
    
    Do While Not rsMonPtVal.EOF
    
        ' Iterate through the records and find the last concentration in the
        ' Exposure mode
    
        'JNM 04/21/03 - This is the first pt. on or beyond End Exposure time that
        '   has a concentration.
        If (rsMonPtVal("mode") = MDM_PRE_AERATION And rsMonPtVal("fumiscope_conc") >= 0) Then
        ' Check the time we may already be at the value for end exposure
            'JNM 04/21/03 - AT Fix - Use = instead of >= for dtEndExposure.
            '   We need the conc. at the End Exposure Pt., not beyond it.
            If (rsMonPtVal("date_time") = dtEndExposure And rsMonPtVal("fumiscope_conc") >= 0) Then
                lConcPrev = rsMonPtVal("fumiscope_conc")
                dtDateTimePrev = rsMonPtVal("date_time")
                bFoundPrev = True
            End If
     
            ' Move to the previous value
            rsMonPtVal.MovePrevious
            
            ' Get the previous concentration, pt-pt hlt and date/time
            If Not bFoundPrev Then
                Do While Not rsMonPtVal.BOF
                
                    ' *************************************************************
                    
                    If bAreaAverage Then
                    
                        #If EXPERTDEBUGMODE = 1 Then
                            Debug.Print "       [SPECIAL HLT COMPUTATION] Determining HLT of averaged concentration readings"
                        #End If
                    
                        'Here 's the fun part, since we don't have the point to point HLT for
                        'the averaged multiple monitor point, we need to calculate it, quick
                        
                        Dim dpcalc2 As DPCALCLib.ProFumeCalculator
                        
                        ' We don't really need the the three, but we need a place to store the stuff
                        ' from dpcalc when getting the Point to Point HLT
                        Dim dtLastPeak As Date
                        Dim dtHLTDate As Date
                        Dim sMsg As String
                        
                        Set dpcalc2 = New DPCALCLib.ProFumeCalculator
                        
                        dpcalc2.ClearConcentrations
                                        
                        rsMonPtVal.MovePrevious
                        
                        If rsMonPtVal.BOF Then
                            ' If we are at the start, there is definitely no HLT so we move back
                            ' to the record we were in
                            dHLTTemp = -1
                            rsMonPtVal.MoveNext
                        Else
                            ' if the previous reading was exactly at start of exposure OR prior
                            ' to the exposure, use the Concentration that was determined in
                            ' GetFirstConcentration (called by GetPointStatus)
                            If rsMonPtVal("date_time") = dtStartExposure Or _
                               rsMonPtVal("mode") < MDM_EXPOSURE Then
                                dpcalc2.SetConcentration lConcFirst, dtStartExposure, sMsg
                                
                                #If EXPERTDEBUGMODE = 1 Then
                                    Debug.Print "       [SPECIAL HLT COMPUTATION] Setting Concentration " & lConcFirst & " @ " & dtStartExposure & " @Start Exposure"
                                #End If
                            
                            Else
                                ' If it was just some other point within the exposure period,
                                ' set the concentration and date/time
                                dpcalc2.SetConcentration rsMonPtVal("fumiscope_conc"), rsMonPtVal("date_time"), sMsg
                                
                                #If EXPERTDEBUGMODE = 1 Then
                                    Debug.Print "       [SPECIAL HLT COMPUTATION] Setting Concentration " & rsMonPtVal("fumiscope_conc") & " @ " & rsMonPtVal("date_time")
                                #End If
                            End If
                            
                            ' go to back to the record we were originally working on
                            rsMonPtVal.MoveNext
                            dpcalc2.SetConcentration rsMonPtVal("fumiscope_conc"), rsMonPtVal("date_time"), sMsg
                            
                            #If EXPERTDEBUGMODE = 1 Then
                                Debug.Print "       [SPECIAL HLT COMPUTATION] Setting Concentration " & rsMonPtVal("fumiscope_conc") & " @ " & rsMonPtVal("date_time")
                            #End If
                            
                            ' Get the Point to Point HLT for the point prior to this
                            dpcalc2.GetPt2PtHlt dHLTTemp, dtLastPeak, dtHLTDate, sMsg
                            
                            #If EXPERTDEBUGMODE = 1 Then
                                Debug.Print "       [SPECIAL HLT COMPUTATION] dpcalc.GetPt2PtHlt returned " & dHLTTemp
                            #End If
                            
                        End If
                        
                        Set dpcalc2 = Nothing
                        
                    Else
                    
                        dHLTTemp = rsMonPtVal("hlt_pt2pt")
                    
                    End If
                    
                    ' *************************************************************
                
                    ' Get previous monitor point value where we have a fumiscope value
                    'If (rsMonPtVal("fumiscope_conc") >= 0 And rsMonPtVal("hlt_pt2pt") >= 0) Then
                    If (rsMonPtVal("fumiscope_conc") >= 0 And dHLTTemp >= 0) Then
                        lConcPrev = rsMonPtVal("fumiscope_conc")
                        dtDateTimePrev = rsMonPtVal("date_time")
                        'dHLT = rsMonPtVal("hlt_pt2pt")
                        dHLT = dHLTTemp
                        bFoundPrev = True
                        Exit Do
                    End If
                    
                    ' Move to the previous record in the recordset
                    rsMonPtVal.MovePrevious
                Loop
            End If
            
            Exit Do
    
        End If
            
        rsMonPtVal.MoveNext
    
    Loop
    
    If rsMonPtVal.EOF Then
        'JNM 04/26/03 - meaning there is no conc. beyond end exposure
        bLastExp = False
        Exit Sub
    End If
        
    ' If we did not find the previous value, return with no concentration
    ' and let the calling function know that no concentration was found.
       
    If Not bFoundPrev Then
     ' Check if the last exposure time has an hlt value where we can
     ' predict what the concentration would be at end of exposure
       
       If (rsMonPtVal.BOF Or rsMonPtVal.EOF) Then
             ' Move to the last record in the recordset
             rsMonPtVal.MoveLast
     
             'Iterate though the records in the recordset
             Do While Not rsMonPtVal.BOF
             
                 ' Find a record that has a concentration in it
                 If (rsMonPtVal("fumiscope_conc") >= 0) Then
                     Exit Do
                 End If
     
             ' Move to the previous record
             rsMonPtVal.MovePrevious
     
            Loop
       End If
          
       If Not rsMonPtVal.BOF Then
             'JNM 04/21/03 - AT Fix - Removed extra condition.
             'If (rsMonPtVal("hlt_pt2pt") >= 0) And (rsMonPtVal("mode") = MDM_PRE_AERATION) Then
             
             'MME 6/03/03- Added condition to make sure that program realizes that the concentration
             '             entered is really 0 not 0.001
             
             'If (rsMonPtVal("hlt_pt2pt") >= 0) Then
             If (rsMonPtVal("hlt_pt2pt") >= 0) And (rsMonPtVal("fumiscope_conc") <> 0) Then
                lConcPrev = rsMonPtVal("fumiscope_conc")
                dtDateTimePrev = rsMonPtVal("date_time")
                dHLT = rsMonPtVal("hlt_pt2pt")
                'JNM 04/21/03 - AT Fix - We'll set this at the end of the function
                'bLastExp = True
             Else
                 lConcExp = 0 '-- Set it to zero (0) L.
                 bLastExp = False
                 Exit Sub
             End If
       Else ' the BOF has been reached, thus no concentration was found
         ' The concentration should be 0 at the end of exposure
          lConcExp = 0 '-- Set it to zero (0) L.
          bLastExp = False
          'Exit Do
          'JNM 04/21/03 - AT fix - We exit.
          Exit Sub
       End If
    
    End If
       
    ' Get the (estimated) concentration/time at the end of exposure using
    '   the point to point hlt of the previous value
    'If bLastExp = False Then
    '    dtspanExp = 0
    'Else
    '    dtspanExp = DateDiff("s", dtDateTimePrev, dtEndExposure)
    'End If
    
    'JNM 04/16/03 - We always need the date difference.
    dtspanExp = DateDiff("s", dtDateTimePrev, dtEndExposure)
    
    ' Get the Total Seconds for the Span Exposure
    dSeconds = dtspanExp
    
    ' Check if the Total Seconds is less than or equal to 0
    If (dSeconds <= 0) Then
               lConcentration = lConcPrev
    Else
    ' If the HLT (dHLT) is equal to 0
        If dHLT = 0 Then
    ' Set the Concentration to the Previous Concentration Value
            lConcentration = lConcPrev
        Else
    '                                  ln(2) * (T-To)
    '        NewConc = exp  [ ln(Co) - -------------- ]
    '                                       HLT
            dTime = (Log(2) * dSeconds / 3600) / dHLT     '**** 3600 seconds/hour
            
            ' 9/17/03 Jett - SIMS #2104694 added check to see if previous concentration is 0.
            ' if it is set extrapolated concentration to 0. (we cannot get the natural log
            ' of 0
            If lConcPrev = 0 Then
                lConcentration = 0
            Else
                lConcentration = (Exp(Log(lConcPrev) - dTime))
            End If
       End If
    
    End If
    
    'JNM 04/21/03 - AT fix - We always need to return the value of bLastExp
    'If bFoundPrev = True Then
        bLastExp = True
    'End If

Exit Sub
Error_Handler:
    Call psAbort(CLSNAME, MGETLASTCONCENTRATION, Err.Description)

End Sub

'**************************************************************
'
'This method calls the SetConcentration function that is in the Profume
'Calculator.
'   Input Parameters:  lConcentration (Long)
'                                  dtDateTime (Date)
'   Output          : (Boolean) Returns True/False
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'**************************************************************


Public Function SetConcentration(lConcentration As Double, dtDateTime As Date) As Boolean

    Dim bstrMsg As String
    On Error GoTo Error_Handler
    
    ' Call the SetConcentration function that is in the Profume Calc
    dpCalc.SetConcentration lConcentration, dtDateTime, bstrMsg
    
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "          DPCALC: SetConcentration " & lConcentration & ", " & dtDateTime & ", " & bstrMsg
    #End If
    
       If bstrMsg <> "" Then
        MsgBox bstrMsg, vbExclamation
        SetConcentration = False
        Exit Function
       End If
     
    'Return a value of True for the method.
    SetConcentration = True
    
    Exit Function
Error_Handler:
Call psAbort(CLSNAME, MSETCONCENTRATION, Err.Description)

End Function

'**************************************************************
'Description:
'
'When a user adds a new time/concentration data value, the half loss time of
'that point needs to be determined by the calculator.  In order to do this
'the current and previous (if available) time/concentration values must be
'added to the calculator using SetConcentration().  Before adding the values
'to the calculator ensure the existing concentration values have been cleared
'out, otherwise they will also be used to determine the HLT.
'After the concentration/time data has been added to the calculator, the current
'concentration/time's HLT may be determined.  Please note dHltPt2Pt value has
'been initialized to -1.0.  0 is a valid HLT.  This occurs when working with
'vacuum chambers.  If an HLT value cannot be determined by the calculator,
'which occurs, when the concentrations are increasing (during ProFume
'introduction) then the calculator returns -1.0.
'
'      Input Parameters:  lMonitorPointId (String)
'                                    lMonPtValueId (String)
'                                    dtStartIntro (Date)
'                                    dtStartExposure (Date)
'                                    dtEndExposure (Date)
'                                    dHltPt2Pt (Date)
'                              strMsg (String)
'               Output :   Boolean) Returns True/False
'
'  REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'**************************************************************

Public Function GetPoint2PointHlt(lMonitorPointId As String, _
                                    lMonPtValueId As String, _
                                    dtStartIntro As Date, _
                                    dtStartExposure As Date, _
                                    dtEndExposure As Date, _
                                    dHltPt2Pt As Double, _
                                    strMsg As String) As Boolean

    Dim rsMonPtVal As ADODB.Recordset
    Dim strSQL As String
    Dim bFound As Boolean
    Dim strMonPtValueId As String
    Dim varMonPtValLast As Variant
    Dim bFoundPrev As Boolean
    Dim lConcFirst As Double
    Dim bFirstExp As Boolean
    Dim LastPeakDateTime As Date
    Dim HltDateTime As Date
    Dim bstrMsg As String
    
    On Error GoTo Error_Handler
    
    Set rsMonPtVal = New ADODB.Recordset
    
    ' Clear the concentrations using the function in the dpcalc
    Call dpCalc.ClearConcentrations
    
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "clsMonitorExpert: GetPoint2PointHLT"
        Debug.Print "Getting Pt2PtHLT for MonitorPointValue id " & lMonPtValueId
        Debug.Print "          DPCALC: ClearConcentrations"
    #End If
    
    ' Get all the records associated with the monitor point
    ' Getting the concentrations in the reverse order
    strSQL = "SELECT MonitorPointValue.id as MonValId, date_time, fumiscope_conc, interscan_conc, general_note, " _
            & "mode, hlt_pt2pt, monitorpoint_fkey, MonitorPointValue.row, MonitorPointValue.deleted, MonitorPoint.id as MonPtId, " _
            & "name, description, job_fkey, status, hlt, last_entry, elapsed_time, jobarea_fkey, MonitorPoint.deleted, start_time, " _
            & "end_time, achieved_concentration_time, achieved_dose, projected_concentration_time, projected_time, future_target_concentration, " _
            & "future_concentration, order_entered, MonitorPoint.row FROM MonitorPointValue, MonitorPoint " _
            & "WHERE MonitorPointValue.deleted = FALSE AND " _
            & "MonitorPointValue.monitorpoint_fkey = '" & lMonitorPointId & "' " _
            & "AND MonitorPoint.id = MonitorPointValue.monitorpoint_fkey " _
            & "ORDER BY MonitorPointValue.date_time ASC"
    
    ' Execute the SQL and store its contents in rsMonPtVal
    rsMonPtVal.Open strSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    ' Check if there are any records returned.
    ' If the recordset is empty, return FALSE
    If (rsMonPtVal.BOF And rsMonPtVal.EOF) Then
        GetPoint2PointHlt = False
        Set rsMonPtVal = Nothing
        Exit Function
    End If
    
    ' Add the exposure concentration/time values to dpcalc
    ' Set the bFound variable/flag to False
    bFound = False
    
    ' Move to the first record in the rsMonPtVal recordset.
    rsMonPtVal.MoveFirst
    
    strMonPtValueId = lMonPtValueId
    
    ' Iterate/Loop though the records in the recordset.
    Do While Not rsMonPtVal.EOF
    
    ' Find the monitoring point we are interested in
    ' Compare the value of the Monitoring Point Value we have with
    ' that of the data in the recordset
       If strMonPtValueId = rsMonPtVal("MonValId") Then
        ' Now that the value has been found, check if it is valid
            
            If ((rsMonPtVal("mode") = MDM_EXPOSURE Or rsMonPtVal("mode") = MDM_PRE_AERATION And rsMonPtVal("date_time") >= dtStartExposure) And rsMonPtVal("fumiscope_conc") >= 0) Then
            ' Ensure this is NOT the start of exposure, since we can't do a point to point calculation with only one point or the first point after the end of exposure time
                
                If (rsMonPtVal("date_time") = dtStartExposure Or rsMonPtVal("date_time") > dtEndExposure) Then
                   GetPoint2PointHlt = True
                   Set rsMonPtVal = Nothing
                   Exit Function  '****
                Else
                   bFound = True
                   ' We found the point
                   varMonPtValLast = rsMonPtVal.Bookmark
                End If
                
             End If
            
            Exit Do
            
        End If
    rsMonPtVal.MoveNext
    Loop
    
    If (Not bFound) Then
        ' The monitor point value is not a candidate for an hlt calc
        ' set the (double) variable to -1.0
        dHltPt2Pt = -1#
        GetPoint2PointHlt = True
        
        ' OPTIMIZE memory leaks
        rsMonPtVal.Close
        Set rsMonPtVal = Nothing
        
        Exit Function  '****
    End If
    
    ' Now that the record has been found, we need to get the previous valid record
    rsMonPtVal.Bookmark = varMonPtValLast
    bFoundPrev = False
    rsMonPtVal.MovePrevious
       
    Do While Not rsMonPtVal.BOF
        
        If (rsMonPtVal("mode") = MDM_EXPOSURE And rsMonPtVal("fumiscope_conc") >= 0) Then
            ' Now the previous point has been found
             bFoundPrev = True
            
            ' Call the SetConcentration function. Add this value to dpcalc.  This should be the first item sent to dpcalc
             If (Not SetConcentration(rsMonPtVal("fumiscope_conc"), rsMonPtVal("date_time"))) Then
                    GetPoint2PointHlt = False
                    
                    ' OPTIMIZE memory leaks
                    rsMonPtVal.Close
                    Set rsMonPtVal = Nothing
                    
                    Exit Function  '****
             End If
             
             Exit Do
        
        End If
    
        rsMonPtVal.MovePrevious
    Loop
    
    ' We have not found the record, so get the first concentration
    ' in the exposure mode
    
    If (Not bFoundPrev) Then
    
        lConcFirst = 0
        bFirstExp = False
        ' Call the Get First Concentration Subroutine
        Call GetFirstConcentration(rsMonPtVal, dtStartIntro, dtStartExposure, lConcFirst, bFirstExp)
    
        ' Add this value to dpcalc.  This should be the first item sent to dpcalc
               
        If (Not SetConcentration(lConcFirst, dtStartExposure)) Then
             GetPoint2PointHlt = False
            
            ' OPTIMIZE memory leaks
            rsMonPtVal.Close
            Set rsMonPtVal = Nothing
            
            Exit Function  '****
        End If
    
    End If
    
    ' Finally, the record can be added.
    rsMonPtVal.Bookmark = varMonPtValLast
    
    If (Not SetConcentration(rsMonPtVal("fumiscope_conc"), rsMonPtVal("date_time"))) Then
        GetPoint2PointHlt = False
        
        ' OPTIMIZE memory leaks
        rsMonPtVal.Close
        Set rsMonPtVal = Nothing
        
        Exit Function  '****
    End If
    
    ' Get the halflife from dpcalc
    ' Call the function GetPt2PtHlt (from dpcalc)
    dpCalc.GetPt2PtHlt dHltPt2Pt, LastPeakDateTime, HltDateTime, bstrMsg
    
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "          DPCALC: GetPt2PtHlt " & dHltPt2Pt & ", " & LastPeakDateTime & ", " & HltDateTime & ", " & bstrMsg
    #End If
    
    If (bstrMsg <> "") Then
        MsgBox bstrMsg, vbExclamation
        GetPoint2PointHlt = False
        
        ' OPTIMIZE memory leaks
        rsMonPtVal.Close
        Set rsMonPtVal = Nothing
        
        Exit Function
    End If
    
    ' OPTIMIZE memory leaks
    rsMonPtVal.Close
    Set rsMonPtVal = Nothing
    
    GetPoint2PointHlt = True
    
    Exit Function
Error_Handler:
    Call psAbort(CLSNAME, MGETPOINT2POINTHLT, Err.Description)

End Function

'**************************************************************
'Description:
'
'This method determines the highest mode that has been reached for the job.
'It returns an integer that corresponds to the mode.
'
'Input Parameters:       lMonPtId (String)
'         Output :       Integer
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'  4/20/2003       Jett Gamboa      Optimized recordset usage
'**************************************************************


Public Function GetHighestMode(lMonPtId As String) As Integer

    Dim nReturn As Integer
    Dim strSQL As String
    Dim rsMonPtVal As ADODB.Recordset
    
    On Error GoTo Error_Handler
    
    Set rsMonPtVal = New ADODB.Recordset
    
    ' Intialize the value of the variable, nReturn
    nReturn = -1
    
    ' Get all the records associated with the monitor point
    strSQL = "SELECT * FROM MonitorPointValue, MonitorPoint " _
             & "WHERE MonitorPointValue.deleted = FALSE AND " _
             & "MonitorPointValue.monitorpoint_fkey = '" & lMonPtId & "'" _
             & "AND MonitorPoint.id = MonitorPointValue.monitorpoint_fkey " _
             & "ORDER BY MonitorPointValue.mode DESC"
    
    ' Execute the SQL and store its contents in rsMonPtVal
    'OPTIMIZE rsMonPtVal.Open strSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    rsMonPtVal.Open strSQL, gdbConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' Check if there are any records returned
       If (rsMonPtVal.BOF And rsMonPtVal.EOF) Then
            GetHighestMode = nReturn
            
            ' OPTIMIZE memory leaks
            rsMonPtVal.Close
            Set rsMonPtVal = Nothing
            
            Exit Function
       End If
    
    ' Since we are in descending order, the highest value will be the first record
          rsMonPtVal.MoveFirst
          nReturn = rsMonPtVal("mode")
          GetHighestMode = nReturn
    
    ' OPTIMIZE memory leaks
    rsMonPtVal.Close
    Set rsMonPtVal = Nothing
    
    Exit Function
Error_Handler:
    Call psAbort(CLSNAME, MGETHIGHESTMODE, Err.Description)

End Function

'**************************************************************
'Description:
'The UpdateJobAreaRecord method updates the percent dose and achieved dose
'in status and hlt (last peak) in the Job Area table.
'
'       Input Parameters:     strJobId (String)
'                             strAreaId (String)
'                 Output:     None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'**************************************************************

Public Sub UpdateJobAreaRecord(strJobId As String, strAreaID As String) ' areastatus

    On Error GoTo Error_Handler
    
    Dim strAchievedCT As String
    Dim strAchievedDose As String
    Dim strHLT As String
    Dim strProjectedCT As String
    Dim strProjectedTime As String
    Dim strMeanNTC As String
    Dim strStatus As String
    Dim strAddProFume As String
    Dim strUpdateSql As String
    
    strAchievedCT = AreaStatus.achieved_concentration_time
    strAchievedDose = AreaStatus.achieved_dose
    strHLT = AreaStatus.hlt
    strProjectedCT = AreaStatus.projected_concentration_time
    strProjectedTime = AreaStatus.projected_time
    strMeanNTC = AreaStatus.mean_ntc
    strStatus = AreaStatus.status
    strAddProFume = AreaStatus.add_profume
    
    ' Update the percent dose achieved, achieved dose in status and hlt (last peak) in the JobArea table
    strUpdateSql = "UPDATE JobArea SET achieved_concentration_time = " & DecVal(strAchievedCT) & ", " _
                    & "achieved_dose = " & DecVal(strAchievedDose) & ", hlt = " & DecVal(strHLT) & ", " _
                    & "projected_concentration_time = " & DecVal(strProjectedCT) & ", " _
                    & "projected_time = " & DecVal(strProjectedTime) & ", mean_ntc = " & DecVal(strMeanNTC) & ", " _
                    & "status = " & strStatus & ", add_profume = " & DecVal(strAddProFume) & " " _
                    & "WHERE deleted = FALSE AND job_fkey = '" & strJobId & "' AND id = '" & strAreaID & "' "
                    
    ' Execute the SQL
    gdbConn.Execute strUpdateSql
    
    Exit Sub
Error_Handler:
    Call psAbort(CLSNAME, MUPDATEJOBAREARECORD, Err.Description)

End Sub

'**************************************************************
'Description:
' This method updates the percent dose achieved, achieved dose in status
' and hlt (last peak) in the Monitor Point table
'
' Input Parameters: strMonitorPointId (String)
'
'           Output: None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'**************************************************************

' OPTIMIZED
Public Sub UpdateMonitorPointRecord(strMonitorPointId As String)
   
    Dim strUpdateSql As String
    Dim strHistorySql As String
    'OPTIMIZE don't need this --- Dim rsUpdate As New ADODB.Recordset
    'OPTIMIZE don't need this --- Dim rsInsert As New ADODB.Recordset
    Dim rsJobKey As ADODB.Recordset
    Dim dElapsedTime As Double
    Dim dAchievedCT As Double
    Dim dAchievedDose As Double
    Dim dProjectedCT As Double
    Dim dProjectedTime As Double
    Dim dFutureTC As Double
    Dim dFutureC As Double
    Dim dHLT As Double
    Dim dtsTime As Date
    Dim dteTime As Date
    Dim strStatus As String
    Dim strSQL As String
    Dim strID As String
    Dim dlast_entry As Date
    Dim strJobArea As String
    
    On Error GoTo Error_Handler
    
    Set rsJobKey = New ADODB.Recordset
    
    dAchievedDose = PointStatus.achieved_dose
    dElapsedTime = PointStatus.elapsed_time
    dAchievedCT = PointStatus.achieved_concentration_time
    dProjectedCT = PointStatus.projected_concentration_time
    dProjectedTime = PointStatus.projected_time
    dFutureTC = PointStatus.future_target_concentration
    dFutureC = PointStatus.future_concentration
    dHLT = PointStatus.hlt
    dtsTime = PointStatus.start_time  ' LastPeakStartTime
    dteTime = PointStatus.last_entry    ' End time
    
    strStatus = "Dose Achieved =" & dAchievedDose & "mg-hrs/liter"
    
    'OPTIMIZE don't need this --- rsUpdate.Open "MonitorPoint", gdbConn
    ' Update the percent dose achieved, achieved dose in status and hlt (last peak) in the monitor point table
    strUpdateSql = "UPDATE MonitorPoint SET elapsed_time = " & DecVal(dElapsedTime) & ", " _
                    & "achieved_concentration_time = " & DecVal(dAchievedCT) & " " _
                    & ", achieved_dose = " & DecVal(dAchievedDose) & " " _
                    & ", projected_concentration_time = " & DecVal(dProjectedCT) & " " _
                    & ", projected_time = " & DecVal(dProjectedTime) & " " _
                    & ", future_target_concentration = " & DecVal(dFutureTC) & " " _
                    & ", future_concentration = " & DecVal(dFutureC) & " " _
                    & ", hlt = " & DecVal(dHLT) & " " _
                    & ", start_time = '" & dteTime & "' " _
                    & ", end_time = '" & dtsTime & "' "
     
    ' This will be saved when the testing attribute is set to true
       If (gbIsTestingActive) Then
            strUpdateSql = strUpdateSql + ", status = '" & strStatus & "' "
       End If
       
       strUpdateSql = strUpdateSql + " WHERE id = '" & strMonitorPointId & "' "
     
    ' Execute the SQL
    gdbConn.Execute strUpdateSql
    
    strSQL = "Select last_entry, jobarea_fkey from MonitorPoint where id = '" & strMonitorPointId & "' "
    
    ' Execute the SQL and store its contents in rsJob
    'OPTIMIZE rsJobKey.Open strSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    rsJobKey.Open strSQL, gdbConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not (rsJobKey.EOF And rsJobKey.BOF) Then
          strJobArea = rsJobKey("jobarea_fkey")
          If IsNull(rsJobKey("last_entry")) Then
            dlast_entry = "12:00:00 AM"
          Else
            dlast_entry = rsJobKey("last_entry")
          End If
    ' Insert the Monitor Point Data in the LogHistory table.
    ' The following data should be stored in the table: Area name, Monitoring Point,
    ' Time of Last Entry, Elapsed Time, HLT, Start time of HLT, End Time of HLT,
    ' CT Achieved, Projected, Status/Recommendation.
    'OPTIMIZE don't need this --- rsInsert.Open "LogHistory", gdbConn
    
    strID = GetGUID()
    
    ' 12/23/2004 Jett - Commented out this code. Based on discussions with Brian
    ' Schneider, the LogHistory table is never used. This only slows us down and
    ' causes the data file to be huge.
    
'    strHistorySql = "INSERT INTO LogHistory (id , " _
'                        & "jobarea_fkey, " _
'                        & "monitorpoint_fkey, " _
'                        & "elapsed_time, " _
'                        & "ct_achieved, " _
'                        & "projected_ct, " _
'                        & "hlt, " _
'                        & "hlt_start_time, " _
'                        & "hlt_end_time, " _
'                        & "status"
            
'    If Not IsNull(rsJobKey("last_entry")) Then
'        strHistorySql = strHistorySql & ", lastentry_time) "
'    Else
'        strHistorySql = strHistorySql & ")"
'    End If
        
'    strHistorySql = strHistorySql & "VALUES ('" & strID & "', '" _
                  & strJobArea & "', '" _
                  & strMonitorPointId & "', " _
                  & DecVal(dElapsedTime) & ", " _
                  & DecVal(dAchievedCT) & ", " _
                  & DecVal(dProjectedCT) & ", " _
                  & DecVal(dHLT) & ", '" _
                  & dteTime & "' , '" _
                  & dtsTime & "' , '" _
                  & strStatus & "'"
      
'      If Not IsNull(rsJobKey("last_entry")) Then
'            strHistorySql = strHistorySql & ", '" & dlast_entry & "')"
'        Else
'            strHistorySql = strHistorySql & ")"
'        End If
     
    ' Execute the SQL
'    gdbConn.Execute strHistorySql
    
    End If
    
    ' OPTIMIZE move this outside the loop so we still do clean up
    rsJobKey.Close
    Set rsJobKey = Nothing
    
    
    Exit Sub
Error_Handler:
    Call psAbort(CLSNAME, "UpdateMonitorPointRecord", Err.Description)

End Sub

'**************************************************************
' Description:
' The UpdateMonitorPointStatus method updates the percent dose, achieved
' dose in status and hlt (last peak) in the Monitor Point table.
'
'       Input Parameters:   strMonitorPointId (String)
'                           dtStartIntro (Date)
'                           dtStartExposure (Date)
'                           dtEndExposure (Date)
'                 Output:   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'**************************************************************

' OPTIMIZE no need to
Public Sub UpdateMonitorPointStatus(strMonitorPointId As String, dtStartIntro As Date, dtStartExposure As Date, dtEndExposure As Date)

    Dim strMsg As String
    
    On Error GoTo Error_Handler
    ' Initialize the structure so we know where we started before the call
    'Use the Used Defined Type PointStatus
       PointStatus.achieved_concentration_time = -1
       PointStatus.achieved_dose = -1
       PointStatus.elapsed_time = -1
       PointStatus.future_target_concentration = -1
       PointStatus.future_concentration = -1
       PointStatus.hlt = -1
       PointStatus.projected_concentration_time = -1
       PointStatus.projected_time = -1
       PointStatus.last_entry = 0       ' dtLastPeakDateTime ****
       PointStatus.start_time = 0       ' dtHltDateTime
       PointStatus.strMsg = ""
       
       If dtStartExposure = DBL_MAX Then
            ' The start of exposure has not been set yet, so
            ' update the monitoring points with initialized values
            Call UpdateMonitorPointRecord(strMonitorPointId) ' rsPointStatus
       Else
             If (Not GetPointStatus(strMonitorPointId, dtStartIntro, dtStartExposure, dtEndExposure)) Then ' rsPointStatus
                ' Display the error string strMsg
                 Call UpdateMonitorPointRecord(strMonitorPointId)
             Else
                ' Update the percent dose achieved dose in status and hlt (last peak) in the monitor point table
                 Call UpdateMonitorPointRecord(strMonitorPointId) ' rsPointStatus
             End If
    End If
    
    Exit Sub
Error_Handler:
    Call psAbort(CLSNAME, MUPDATEMONITORPOINTSTATUS, Err.Description)

End Sub

'**************************************************************
' Description:
' This method updates the record for a new HLT (Half Loss Time) value
' in the Monitor Point Value Table.
'
'        Input Parameters:  strMonitorPointId (String)
'                           strMonPtValueId (String)
'                           dtStartIntro (Date)
'                           dtStartExposure (Date)
'                           dtEndExposure (Date)
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'**************************************************************

Public Sub UpdateMonPtValuePt2PtHlt(strMonitorPointId As String, strMonPtValueId As String, dtStartIntro As Date, dtStartExposure As Date, dtEndExposure As Date)
    Dim dHltPt2Pt As Double
    Dim strUpdateSql As String
    ' OPTIMIZE Do not need this --- Dim rsUpdate As New ADODB.Recordset
    Dim strMsg As String
    
    On Error GoTo Error_Handler
    
    ' Initialize variables
    dHltPt2Pt = -1#
    
    If (Not GetPoint2PointHlt(strMonitorPointId, strMonPtValueId, dtStartIntro, dtStartExposure, dtEndExposure, dHltPt2Pt, strMsg)) Then
        MsgBox strMsg, vbExclamation
    End If
       
    ' Update the record for a new hlt value
    
    strUpdateSql = "UPDATE MonitorPointValue " _
                 & "SET hlt_pt2pt = " & DecVal(dHltPt2Pt) & " " _
                 & "WHERE id = '" & strMonPtValueId & "'"
    
    ' Execute the SQL
    gdbConn.Execute strUpdateSql
    
    Exit Sub
Error_Handler:
    Call psAbort(CLSNAME, MUPDATEMONPTVALUEPT2PTHLT, Err.Description)

End Sub

'**************************************************************
' Description:
' This method determines the Actual Exposure time, which is based on the
' time span that occurred between the time of first exposure and the time
' of Pre-Aeration.
'      Input Parameters:    strJobGuid (String)
'                           dtTimeSpan (Date)
'               Output:     (Boolean) Returns True/False
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'**************************************************************

' OPTIMIZE
Public Function GetActualExposureTime(strJobGuid As String, dtTimeSpan As Long) As Boolean

    Dim bRetval As Boolean
    Dim strSQL As String
    Dim rsStartEndVals As ADODB.Recordset
    Dim dtExposureStart As Date
    Dim dtExposureEnd As Date
    
    On Error GoTo Error_Handler
    ' Initialize variables
    
    Set rsStartEndVals = New ADODB.Recordset
    
    bRetval = False
    ' Select the first mode E value and the first mode Pre-A value into the recordset.
    strSQL = "SELECT top 1 MonitorPointValue.*, MonitorPoint.* " _
            & "FROM MonitorPointValue inner join " _
            & "(MonitorPoint inner join Job on Job.id = MonitorPoint.job_fkey) " _
            & "on MonitorPointValue.monitorpoint_fkey = MonitorPoint.id " _
            & "WHERE MonitorPointValue.deleted = false and " _
            & "MonitorPointValue.mode = " & MDM_EXPOSURE & " and " _
            & "Job.id = '" & strJobGuid & "'" _
            & "order by date_time " _
            & "UNION ALL " _
            & "(SELECT top 1 MonitorPointValue.*, MonitorPoint.* " _
            & " FROM MonitorPointValue inner join " _
            & "(MonitorPoint inner join Job on Job.id = MonitorPoint.job_fkey) " _
            & "on MonitorPointValue.monitorpoint_fkey = MonitorPoint.id " _
            & "WHERE MonitorPointValue.deleted = false and " _
            & "MonitorPointValue.mode = " & MDM_PRE_AERATION & " and " _
            & "Job.id = '" & strJobGuid & "'" _
            & "order by date_time)"
    
    ' Execute the SQL and store the records in rsStartEndVals
    rsStartEndVals.Open strSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    'rsStartEndVals.Open strSQL, gdbConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' Make sure there are records in the recordset.
    If ((Not rsStartEndVals.BOF) And (Not rsStartEndVals.EOF)) Then
        ' We may have more than 2 records, however, all we care about are the first and last ones since they will have the bounding date/time values.
        rsStartEndVals.MoveFirst
        dtExposureStart = rsStartEndVals("date_time")
        rsStartEndVals.MoveLast
        dtExposureEnd = rsStartEndVals("date_time")
        ' Make sure both dates are valid.
        If (IsDate(dtExposureStart) And IsDate(dtExposureEnd)) Then
            ' And make sure the end is after the beginning.
            If (dtExposureEnd > dtExposureStart) Then
                ' And if so, then get the time span.
                dtTimeSpan = DateDiff("s", dtExposureStart, dtExposureEnd)
                bRetval = True
            End If
        End If
    End If
    
    ' Return bRetVal
    GetActualExposureTime = bRetval
    
    ' OPTIMIZE
    rsStartEndVals.Close
    Set rsStartEndVals = Nothing
    
    Exit Function
Error_Handler:
    Call psAbort(CLSNAME, MGETACTUALEXPOSURETIME, Err.Description)

End Function

'**************************************************************
' Description:
' This method updates/re-initializes the values of the Job Area Status'
' and references the UpdateJobAreaRecord method which updates the
' Job Area table.
'       Input Parameters:  strJobId (String)
'                                    astrAreaId (Array)
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       Cel Gamboa       Initial Revision
'  5/15/2003       Jett Gamboa      added parameters for dates (will use them to handle
'                                   multiple Monitoring Points in GetAreaStatus
'**************************************************************
Public Sub UpdateJobAreaStatus(strJobId As String, astrAreaId As String, dtStartIntro As Date, dtStartExposure As Date, dtEndExposure As Date)

    'astrAreadId is an array. So I need to pass an array. But how?
    ' how about a comma-delimited string?
    Dim nIndex As Integer
    Dim arrArea() As String
    Dim strAreaID As String
    Dim strMsg As String
    
    On Error GoTo Error_Handler
    ' Declare and Initialize variables
    
    arrArea() = Split(astrAreaId, ",")
    
       For nIndex = 0 To UBound(arrArea)
          strAreaID = Trim$(arrArea(nIndex))
          AreaStatus.achieved_concentration_time = -1
          AreaStatus.achieved_dose = -1
          AreaStatus.hlt = -1
          AreaStatus.projected_concentration_time = -1
          AreaStatus.projected_time = -1
          AreaStatus.mean_ntc = -1
          AreaStatus.status = 0
          AreaStatus.add_profume = -1
          AreaStatus.strMsg = ""
          If (Not GetAreaStatus(strJobId, strAreaID, dtStartIntro, dtStartExposure, dtEndExposure)) Then ' AreaStatus
           ' Display the error string strMsg
            If strMsg <> "" Then
                MsgBox strMsg, vbExclamation
            End If
          End If
     
         Call UpdateJobAreaRecord(strJobId, strAreaID) ' AreaStatus
        Next
    
    Exit Sub
Error_Handler:
    Call psAbort(CLSNAME, MUPDATEJOBAREASTATUS, Err.Description)
End Sub


'**************************************************************
' Description:
' This private method checks if areas that have multiple monitoring points have
' concentration readings for all monitoring points
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  5/16/2003       Jett Gamboa       Initial Revision
'**************************************************************
Private Function CheckMultipleMonPts(strAreaID As String, NumPoints As Integer)

    Dim strSQL As String
    Dim rsMonitorPoints As ADODB.Recordset
    
    On Error GoTo Error_Handler
       
    strSQL = "SELECT MonitorPointValue.date_time, Count(MonitorPointValue.fumiscope_conc) AS MPReadings " & _
             "FROM MonitorPoint INNER JOIN MonitorPointValue ON MonitorPoint.id = MonitorPointValue.monitorpoint_fkey " & _
             "WHERE (((MonitorPointValue.deleted) = False) And ((MonitorPoint.deleted) = False)) " & _
             "AND MonitorPointValue.fumiscope_conc > -1 " & _
             "AND MonitorPoint.jobarea_fkey = '" & strAreaID & "' " & _
             "GROUP BY MonitorPointValue.date_time "
             
    Set rsMonitorPoints = New ADODB.Recordset
             
    rsMonitorPoints.Open strSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    While Not rsMonitorPoints.EOF
    
        ' Check if the number of readings we have for this date/time is
        ' the less than the number of monitoring points. If it is then
        ' return false to indicate that we do not have enough readings
        If Val(rsMonitorPoints("MPReadings").Value) < NumPoints Then
        
            CheckMultipleMonPts = False
            rsMonitorPoints.Close
            
            Set rsMonitorPoints = Nothing
            
            Exit Function
            
        End If
        
        rsMonitorPoints.MoveNext
    
    Wend
    
    CheckMultipleMonPts = True

    Exit Function
Error_Handler:
    Call psAbort(CLSNAME, MUPDATEJOBAREASTATUS, Err.Description)

End Function



'**************************************************************
' Description:
' This method is called whenever the Monitor Expert Class instance
' is instantiated.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  4/20/2003       Jett Gamboa       Initial Revision
'**************************************************************
Private Sub Class_Initialize()
    
    Set dpCalc = New DPCALCLib.ProFumeCalculator
          
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print vbCrLf & vbCrLf
        Debug.Print "============================================================="
        Debug.Print "Class Monitor Initializing " & Now()
    #End If
End Sub


'**************************************************************
' Description:
' This method is called whenever the Monitor Expert Class instance
' is destroyoed.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  4/20/2003       Jett Gamboa       Initial Revision
'**************************************************************
Private Sub Class_Terminate()
        
    ' ... and this single line, my friends, fixed that biggest memory leak in the Fumiguide
    Set dpCalc = Nothing
    
    #If EXPERTDEBUGMODE = 1 Then
        Debug.Print "Monitor Expert Terminating " & Now()
        Debug.Print "============================================================="
    #End If
    
End Sub


'**************************************************************
' Description:
' Rounds the number based on standard convention (if decimal is
' greater than .5 round up otherwise round down
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  5/28/2003       Malou Espaldon    Initial Revision
'**************************************************************
Private Function StdRound(ByVal dNum As Double) As Long

    Dim dConcMPCT As Double

    dConcMPCT = dNum
    StdRound = IIf(dConcMPCT - Int(dConcMPCT) >= 0.5, Int(dConcMPCT) + 1, dConcMPCT)

End Function
