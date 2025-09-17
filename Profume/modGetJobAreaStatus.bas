Attribute VB_Name = "modGetJobAreaStatus"
Option Explicit

Const MODULE_NAME = "modGetJobAreaStatus()"

Public Function GetJobAreaStatus(ByRef rsJobArea As ADODB.Recordset, _
ByVal rsJob As ADODB.Recordset) As String
'**************************************************************
'   DESCRIPTION
'       This function determines the status of the current Job Area.
'
'   PARAMETERS
'           rsJobArea = the Job Area recordset
'
'   RETURNS - the Status of the Job Area
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'   11/18/02       Jett Gamboa       Changed status string for pre-exposure and post-exposure
'   11/19/02       Jett Gamboa       Moved code from Initialize Job Mode so that we get the mode
'                                    of each area and not from the entire job.
'   12/12/02       Jett Gamboa       added new status BAD_ELAPSED_TIMES
'   04/09/03       Joseph Manyaga    Added a condition that would check if there
'                                    is a Projected Time value for Incompatable status.
'                                    Added a condition that would check if there is
'                                    still an hour left before the end Exposure Time using
'                                    the function HasTimeLeft.
'**************************************************************

    On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "GetJobAreaStatus()"
    
    Dim strErrorDesc As String
    
    Dim sReturn As String
    Dim sStatus As String
    Dim sWeightUnits As String
    Dim sStartOfExposureMPSQL As String
    Dim dtExposureStarted As Date
    Dim dtComplete As Date
    Dim dDays As Double
    Dim rsStartOfExposureMP As New ADODB.Recordset
    Dim dAddProfume As Double
    Dim sAddProfumeUnit As String
    Dim dMeanNTC As Double
    Dim sMeanNTCUnit As String
    
    Dim sMonitorPointValueSQL
    Dim rsMonitorPointValue As New ADODB.Recordset
    
    Dim nAreaJobMode As Integer
    
    'JNM 05/08/03 - PT fix - Added this in order to store the date-time of the last MP value
    Dim dtLastDateTime As Date
    
    dAddProfume = 0#
    sAddProfumeUnit = ""
    dMeanNTC = 0#
    sMeanNTCUnit = ""
    sReturn = ""
    
    ' Check if we had a missing concentration reading status for this area. If we did then
    ' do not proceed with the other checks
    If rsJobArea("status") = STATUS_MISSING_CONC_READING Then
    
        sStatus = LoadResString(IDS_MISSING_READINGS)
    
    Else
    
        'Get the most recent Monitor Point Value for this area
        sMonitorPointValueSQL = "select top 1 * from MonitorPointValue " & _
            "inner join MonitorPoint on MonitorPointValue.monitorpoint_fkey = MonitorPoint.id " & _
            "where MonitorPointValue.deleted = false and MonitorPoint.jobarea_fkey = '" & _
            rsJobArea("id") & "' order by MonitorPointValue.mode desc"
            
        rsMonitorPointValue.Open sMonitorPointValueSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
        
        ' get the most recent status for the area
        If rsMonitorPointValue.BOF And rsMonitorPointValue.EOF Then
            nAreaJobMode = MDM_UNKNOWN
            dtLastDateTime = 0#
        Else
            rsMonitorPointValue.MoveFirst
            nAreaJobMode = IIf(IsNull(rsMonitorPointValue("mode")), "", rsMonitorPointValue("mode"))
            
            'JNM 05/08/03 - PT fix - Added this in order to store the date-time of the last MP value
            dtLastDateTime = IIf(IsNull(rsMonitorPointValue("date_time")), "", rsMonitorPointValue("date_time"))
        End If
        
        
        sWeightUnits = IIf(gbIsMetric, LoadResString(IDS_KILOGRAMS), LoadResString(IDS_POUNDS))
        
        
        If rsJobArea("status") <> 4 And rsJobArea("projected_time") >= 0 Then
                dDays = rsJobArea("projected_time") / 24
                dtComplete = rsJob("startexp_datetime") + dDays
                'JNM 04/06/03 - Fumiguide Enh. 2003 - Removed the line below. This statement
                '   has been incorporated with Below Target and Incompatible statements.
        '        sReturn = FormatSWParams(LoadResString(IDS_EST_TIME_OF_COMPLETION), FormatFumDate(dtComplete))
        End If
        
        ' Set the appropriate status string
        Select Case nAreaJobMode
            Case MDM_PRE_EXPOSURE
                sStatus = LoadResString(IDS_INTRODUCTION_BEGUN)
            
            'JNM 05/08/03 - PT fix - Transferred mode MDM_PRE_AERATION to another case
            Case MDM_EXPOSURE
                'Determine the status of the Area.
                Select Case rsJobArea("status")
                    Case STATUS_ON_TARGET
                        sStatus = LoadResString(IDS_FUMIGATION_ON_TARGET)
                        
                    Case STATUS_ABOVE_TARGET
                        If (rsJobArea("projected_time") < 0#) Then
                            sStatus = LoadResString(IDS_FUMIGATION_ABOVE_TARGET_NO_TIME)
                        Else
                            sStatus = FormatSWParams(LoadResString(IDS_FUMIGATION_ABOVE_TARGET), _
                                CStr(Round(rsJobArea("projected_time"), 2)) & "|" & _
                                FormatFumDate(dtComplete))
                        End If
                        
                    Case STATUS_BELOW_TARGET
                        'JNM 04/09/03 - Fumiguide Enh. 2003 - Added condition that would check
                        '   if there is an hour left before the End Exposure time. A different
                        '   message would be displayed if this is the case.
                        If HasTimeLeft(rsJobArea("id"), rsJob("endexp_datetime")) Then
                            Call UnitConvert(modeDisplay, unitWEIGHT, rsJobArea("add_profume"), _
                                dAddProfume, sAddProfumeUnit)
                            dAddProfume = FormatNumber(Round(dAddProfume, 2), 2)
                            dMeanNTC = Round(rsJobArea("mean_ntc"), 0)
                            sMeanNTCUnit = IIf(gbIsMetric, LoadResString(IDS_GRAMS_PER_CUBIC_METER), _
                                LoadResString(IDS_OZ_PER_MCF))
                            If rsJobArea("projected_time") < 0# Then
                                sStatus = FormatSWParams(LoadResString(IDS_FUMIGATION_BELOW_TARGET_NO_TIME), _
                                    CStr(dAddProfume) & "|" & sAddProfumeUnit & "|" & _
                                    CStr(dMeanNTC) & "|" & sMeanNTCUnit)
                            Else
                                sStatus = FormatSWParams(LoadResString(IDS_FUMIGATION_BELOW_TARGET), _
                                    CStr(dAddProfume) & "|" & sAddProfumeUnit & "|" & _
                                    CStr(dMeanNTC) & "|" & sMeanNTCUnit & "|" & _
                                    CStr(Round(rsJobArea("projected_time"), 2)) & "|" & _
                                    FormatFumDate(dtComplete))
                            End If
                        Else
                            sStatus = LoadResString(IDS_NOT_CALC_ADDGAS)
                        End If
                        
                    Case STATUS_ACHIEVED_TARGET
                        sStatus = LoadResString(IDS_TARGET_ACHIEVED)
                        
                    Case STATUS_INCREASING_CONCENTRATION
                        sStatus = LoadResString(IDS_EXPOSURE_IN_PROGRESS)
                        
                    Case STATUS_UNKNOWN
                        sStatus = LoadResString(IDS_EXPOSURE_IN_PROGRESS)
                        
                    Case STATUS_TARGET_DOSE_INCOMPATABLE
                        'JNM 04/09/03 - Fumiguide Enh. 2003 - Added condition that would check
                        '   if there is an hour left before the End Exposure time. A different
                        '   message would be displayed if this is the case.
                        If HasTimeLeft(rsJobArea("id"), rsJob("endexp_datetime")) Then
                            Call UnitConvert(modeDisplay, unitWEIGHT, rsJobArea("add_profume"), _
                                dAddProfume, sAddProfumeUnit)
                            dAddProfume = FormatNumber(Round(dAddProfume, 2), 2)
                            dMeanNTC = Round(rsJobArea("mean_ntc"), 0)
                            sMeanNTCUnit = IIf(gbIsMetric, LoadResString(IDS_GRAMS_PER_CUBIC_METER), _
                                LoadResString(IDS_OZ_PER_MCF))
                            'JNM 04/06/03 - Fumiguide Enh. 2003 - Added a condition to consider
                            '   the Projected Time value.
                            If rsJobArea("projected_time") < 0# Then
                                sStatus = FormatSWParams(LoadResString(IDS_TARGET_DOSE_INCOMPATABLE_NO_TIME), _
                                    CStr(dAddProfume) & "|" & sAddProfumeUnit & "|" & _
                                    CStr(dMeanNTC) & "|" & sMeanNTCUnit)
                            Else
                                sStatus = FormatSWParams(LoadResString(IDS_TARGET_DOSE_INCOMPATABLE), _
                                    CStr(Round(rsJobArea("projected_time"), 2)) & "|" & _
                                    FormatFumDate(dtComplete) & "|" & _
                                    CStr(dAddProfume) & "|" & sAddProfumeUnit & "|" & _
                                    CStr(dMeanNTC) & "|" & sMeanNTCUnit)
                            End If
                        Else
                            sStatus = LoadResString(IDS_NOT_CALC_ADDGAS)
                        End If
                            
                    Case STATUS_BAD_ELAPSED_TIMES
                        ' clear the sReturn which has the estimated time of completion
                        ' we do cannot present that information if elapsed times are not within
                        ' 30 minutes of each other
                        sReturn = ""
                        sStatus = LoadResString(IDS_FUMIGATION_BAD_ELAPSED_TIMES)
                End Select
            
            'JNM 05/08/03 - PT fix - Created this new case to handle cases wherein exposure
            '   period has ended
            Case MDM_PRE_AERATION
                Select Case rsJobArea("status")
                    Case STATUS_ON_TARGET
                        sStatus = LoadResString(IDS_FUMIGATION_ON_TARGET_ENDEXP)
                        
                    Case STATUS_ABOVE_TARGET, STATUS_ACHIEVED_TARGET
                        sStatus = LoadResString(IDS_TARGET_ACHIEVED)
                        
                    Case STATUS_BELOW_TARGET, STATUS_TARGET_DOSE_INCOMPATABLE, STATUS_INCREASING_CONCENTRATION, STATUS_UNKNOWN
                        sStatus = LoadResString(IDS_CT_NOT_ACHIEVED_ENDEXP)
                    
                    Case STATUS_BAD_ELAPSED_TIMES
                        ' clear the sReturn which has the estimated time of completion
                        ' we do cannot present that information if elapsed times are not within
                        ' 30 minutes of each other
                        sReturn = ""
                        sStatus = LoadResString(IDS_FUMIGATION_BAD_ELAPSED_TIMES)
                End Select
        End Select
        
        ' 11/18/02 TSG - Added extra line if area has a reading entered after end of exposure date/time
        If nAreaJobMode = MDM_PRE_AERATION Then
            'JNM 05/08/03 - PT fix - added this condition in order to not display the message
            '   if the last MP value is at end exposure
            If dtLastDateTime > CDate(rsJob("endexp_datetime")) Then
                sStatus = sStatus & vbCrLf & LoadResString(IDS_NO_CALC_AFTER_EXPOSURE)
            End If
        End If
        
        rsMonitorPointValue.Close
        Set rsMonitorPointValue = Nothing
        
    End If
    
    'Check if a calc has been performed on the dosage plan page.
    If rsJobArea("fumigant") < 0# Then
        sStatus = LoadResString(IDS_FP_CALC_REQ)
    End If
    
    'Assemble the whole string
    sReturn = sReturn & sStatus
    
    'I shall return...
    GetJobAreaStatus = sReturn
    
    Exit Function

Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Function


'**************************************************************
'   Function HasTimeLeft(JobAreaID As String, EndExposure As String) As Boolean
'   DESCRIPTION
'   Checks if there is at least 30 minutes left to end of exposure. This is only
'   called to determine if there is enough time to add gas. (Calculator assumes that there
'   is a one hour lag time for add gas option.
'
'   PARAMETERS
'   JobAreaID - The job area id that we are looking at
'   EndExposure - The end exposure date/time as a a string
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  4/02/2003       Jett Gamboa       Initial Revision (Addendum for 1st project)
'  7/01/2003       Jett Gamboa       Changed comparison from 1 hour (60 minutes) to 30 minutes
'**************************************************************
Function HasTimeLeft(JobAreaID As String, EndExposure As String) As Boolean

    Dim rsLatestReading As ADODB.Recordset
    Dim strSQL As String
    Dim intTimeLeft As Integer
    Dim strLatestTime As String
    Dim strErrorDesc As String
    
    Const FUNCTION_NAME = "HasTimeLeft()"
    
    On Error GoTo Error_Handler
    
    Set rsLatestReading = New ADODB.Recordset
    
    HasTimeLeft = False
    
    'JNM o4/25/03 - AT Fix - Include deleted clause
    strSQL = "SELECT Max(date_time) AS LatestEntry " & _
             "FROM MonitorPoint AS MonPt, MonitorPointValue AS MonPtVal " & _
             "WHERE MonPt.deleted = false AND MonPtVal.deleted = false AND " & _
             "MonPtVal.monitorpoint_fkey=MonPt.id AND MonPtVal.fumiscope_conc>-1 " & _
             "AND MonPt.jobarea_fkey = '" & JobAreaID & "'"
             
    rsLatestReading.Open strSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsLatestReading.EOF Then
        strLatestTime = rsLatestReading("LatestEntry")
    
        'JNM 04/24/03 - AT Fix - Do not store in int since this can exceed
        '   the maximum for int
        'intTimeLeft = DateDiff("n", strLatestTime, EndExposure)
        
        'Instead, compare the values immediately
        If DateDiff("n", strLatestTime, EndExposure) > 30 Then
            HasTimeLeft = True
        End If
    End If
    
    rsLatestReading.Close
    Set rsLatestReading = Nothing

    Exit Function
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Function
