Attribute VB_Name = "modExport"

Option Explicit
'----------------------------------------------------------------------------------
'*
'* Filename: modExport.mod
'*
'* DESCRIPTION
'* -----------
'* This module contains the ExportFile() function and supporting routines used to
'* dump information from the currently opened file to tab-delimited text file
'*
'* Copyright (c) 2002 Accenture.  All rights reserved.
'---------------------------------------------------------------------------------
'*
'* REVISION HISTORY
'*
'* Date         Author          Description
'* 3/31/03      Jett Gamboa     initial revision
'----------------------------------------------------------------------------------


'********************************************************************************
'   Sub ExportFile()
'   DESCRIPTION
'   This is the function that gets executed when the user does a File - Export
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  3/31/2003       Jett Gamboa       Initial Revision
'  4/28/2003       Jett Gamboa       Changed order of data input dump to match the
'                                    sort order in the tab
'  1/19/2006       Anabel Tomas      Removed the Dialog box to automatically
'                                    display the report after clicking export
'*********************************************************************************
Dim strExportFilename As String   ' name of the export file
Sub ExportFile()
    
    Const MODULE_NAME = "modExport.bas"
    Const FUNCTION_NAME = "ExportFile()"
    Dim strErrorDesc As String

    'BEL 01/25/2006 - Moved this variable outside the function
    'Dim strExportFilename As String   ' name of the export file
    
    Dim nFileNum As Integer         ' file number to use for the file
    
    Dim strSQL As String
    
    On Error GoTo Error_Handler
    'BEL 01/19/2006 - Removed the Dialog box
    '--START--
    'With frmMain.dlgCommonDialog
        
        ' set the properties of the common dialog control
    '    .FileName = ""
    '    .DialogTitle = "Export"
    '    .CancelError = True
    '    .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    '    .Filter = LoadResString(IDS_DUMPFILE_FILTER)
        
        ' turn off error handling so we can check if the user selected
        ' the cancel button
    '    On Error Resume Next
        
        ' display the save dialog control
    '    .ShowSave
        
        ' If the user did not select a file exit the subroutine
    '    If Err.Number = cdlCancel Or Len(.FileName) = 0 Then
    '        Exit Sub
    '    End If
        
    '    On Error GoTo Error_Handler
            
        ' set the filename to the name specified by the user
    '    strExportFilename = .FileName
            
    'End With
    
         strExportFilename = gsCurrentFile
         strExportFilename = Replace(strExportFilename, ".fum", ".txt")
    '--END--
    
    ' Retrieve the next free file number
    nFileNum = FreeFile()
    
    Open strExportFilename For Output As nFileNum
    
    ' --------------- Dump the Customer Information section ---------------------------
    DumpCustomerInfo nFileNum
'    strSQL = "SELECT Job.Company, Job.address1 AS [Address Line 1], Job.address2 AS [Address Line 2], Job.address3 AS [Address Line 3], " & _
'             "Job.City, Job.State As [State/Province],Job.zip AS [ZIP/Postal Code], Countries.name AS [Country], " & _
'             "Contact.name AS [Primary Contact Name], Contact.phone AS [Primary Contact Phone], Contact.email AS [Primary Contact Email], Contact_1.name AS [Secondary Contact Name], Contact_1.phone AS [Secondary Contact Phone], Contact_1.email AS [Secondary Contact Email], Job.Notes " & _
'             "FROM Contact, Contact AS Contact_1, Countries " & _
'             "INNER JOIN Job ON Countries.code = Job.country " & _
'             "WHERE (((Contact.Primary) = True) And ((Contact_1.Primary) <> True))"
'
'    DirectDump nFileNum, "CUSTOMER INFORMATION", strSQL
        
    ' --------------- Dump the Job Information section ---------------------------
    DumpJobInfo nFileNum
    
    ' --------------- Dump the Job Pest section ---------------------------
    DumpJobPest nFileNum
           
    ' --------------- Dump the Area Information section ---------------------------
    DumpAreaInfo nFileNum
'    strSQL = "SELECT JobArea.area_name AS [Area Name], JobArea.temperature AS [Temperature], JobArea.estimated_hlt AS [Estimated HLT], JobArea.volume As Volume, JobArea.planned_exposure As [Exposure Time], " & _
'             "JobArea.fumigant As [Fumigant Required], JobArea.concentration_time AS [Target CT], JobArea.initial_concentration AS [Initial Target CT], JobArea.calculation_warning As [Calculation Warning]" & _
'             "FROM JobArea " & _
'             "WHERE (((JobArea.deleted) = False)) " & _
'             "ORDER BY JobArea.order_entered"
'
'    DirectDump nFileNum, "AREA INFORMATION", strSQL


    ' --------------- Dump the Introduction Plan section ---------------------------
    
    strSQL = "SELECT JobArea.area_name As [Area Name], ShootingLine.name AS [Introduction Line], ShootingLine.humidity AS [Relative Humidity], ShootingLine.cylinder_temperature AS [Cylinder Temperature], ShootingLine.diameter AS [Inside Diameter], ShootingLine.length AS Length, ShootingLine.fan AS [Fan Capacity], " & _
             "ShootingLine.rate AS [Introduction Rate], ShootingLine.max_rate AS [Permitted Rate], ShootingLine.calculation_warning AS [Calculation Warning]" & _
             "FROM JobArea INNER JOIN ShootingLine ON JobArea.id = ShootingLine.jobarea_fkey " & _
             "WHERE (((ShootingLine.deleted)=False))"
             
    DirectDump nFileNum, "INTRODUCTION PLAN", strSQL

    ' --------------- Dump the Monitor Plan section ---------------------------

    'JNM 04/21/03 - AT Fixes - Get the undeleted records only
    strSQL = "SELECT JobArea.area_name AS [Area Name], MonitorPoint.name AS [Monitor Point Name], MonitorPoint.description AS [Location/Description] " & _
             "FROM JobArea INNER JOIN MonitorPoint ON JobArea.id = MonitorPoint.jobarea_fkey " & _
             "WHERE (((JobArea.deleted)=False)) and MonitorPoint.deleted = false"
             
    DirectDump nFileNum, "MONITOR PLAN", strSQL

    ' --------------- INTRODUCTION HISTORY ---------------------------
    
    strSQL = "SELECT ShootingLine.name AS [Introduction Line Name], IntroHistory.amt_introduced AS [Amount Introduced], IntroHistory.time_introduced AS [Time Introduced], IntroHistory.Note " & _
             "FROM ShootingLine INNER JOIN IntroHistory ON ShootingLine.id = IntroHistory.shootingline_fkey " & _
             "WHERE (((ShootingLine.deleted)=False) AND ((IntroHistory.deleted)=False))"

    DirectDump nFileNum, "INTRODUCTION HISTORY", strSQL
    
    ' --------------- MONITOR POINT INPUT ---------------------------
    
    strSQL = "SELECT JobArea.area_name AS [Area Name], MonitorPoint.name AS [Monitor Pt Name], MonitorPointValue.date_time AS [Date / Time], MonitorPointValue.fumiscope_conc AS Concentration, MonitorPointValue.hlt_pt2pt AS HLT, MonitorPointValue.general_Note AS Notes " & _
             "FROM (JobArea INNER JOIN MonitorPoint ON JobArea.id = MonitorPoint.jobarea_fkey) INNER JOIN MonitorPointValue ON MonitorPoint.id = MonitorPointValue.monitorpoint_fkey " & _
             "WHERE (((JobArea.deleted) = False) And ((MonitorPoint.deleted) = False) And ((MonitorPointValue.deleted) = False)) " & _
             "ORDER BY MonitorPoint.name, MonitorPointValue.date_time"
             
    DirectDump nFileNum, "MONITOR POINT INPUT", strSQL
    
    ' --------------- MONITOR POINT STATUS ---------------------------
    
    strSQL = "SELECT DISTINCT JobArea.area_name AS [Area Name], MonitorPoint.name AS [Monitor Point Name], MonitorPoint.last_entry AS [Last Entry], MonitorPoint.elapsed_time AS [Elapsed Time], MonitorPoint.hlt AS [HLT], MonitorPoint.start_time AS [Start Time of HLT], MonitorPoint.end_time AS [End Time of HLT], MonitorPoint.achieved_concentration_time AS [CT Achieved], MonitorPoint.projected_concentration_time AS [Projected CT] " & _
             "FROM (JobArea INNER JOIN MonitorPoint ON JobArea.id = MonitorPoint.jobarea_fkey) INNER JOIN MonitorPointValue ON MonitorPoint.id = MonitorPointValue.monitorpoint_fkey " & _
             "WHERE (((JobArea.deleted) = False) And ((MonitorPointValue.deleted) = False)) " & _
             "ORDER BY JobArea.area_name"
             
    DirectDump nFileNum, "MONITOR POINT STATUS", strSQL
    
    
    ' --------------- AREA STATUS ---------------------------
    
    DumpAreaStatus nFileNum


    Close nFileNum
    
    'BEL 01/19/2006 - Replaced it with calling the Fumiguide Report
    'MsgBox LoadResString(IDS_DATA_EXPORTED), vbInformation
    Call OpenFumiguideReport
       
    Exit Sub
    
Error_Handler:

    ' If the file cannot be opened (i.e. it is read only or there was error accessing the
    ' file) display a different message
    ' 75 - is file access error, 70 is permission denied
    If Err.Number = 75 Or Err.Number = 70 Then
    
        MsgBox LoadResString(IDS_ERR_OPENDUMPFILE), vbCritical
        Exit Sub
        
    ElseIf Err.Number = 5903 Then
        'If the template file has already data in it, this error is generated
        'the empty template file should be copied back to profume folder
        MsgBox "Fumiguide Report file is not empty, copy the original back to ProFume folder.", vbCritical
        Exit Sub
        
    Else
        ' untrapped error, call the critical error handler function
        strErrorDesc = Err.Number & " - " & Err.Description
        psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc
    End If
    
End Sub


'********************************************************************************
'   Sub DirectDump(nFileNumber As Integer, strSectionName As String, strSQL As String)
'   DESCRIPTION
'   This is performs a generic dump using the SQL statement specified in the
'   parameter. It uses the file number passed to determine the the file to write to.
'   It first writes the field names (column names in SELECT clause) followed by
'   the actual contents. Data is tab-delimited.
'
'   PARAMETERS
'   nFileNumber - integer specifying file number of open file
'   strSectionName - the name of the section
'   strSQL - the (SELECT) SQL used to fetch the data
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  3/31/2003       Jett Gamboa       Initial Revision
'*********************************************************************************
Sub DirectDump(nFileNumber As Integer, strSectionName As String, strSQL As String)

    Const MODULE_NAME = "modExport.bas"
    Const FUNCTION_NAME = "DirectDump()"
    Dim strErrorDesc As String

    Dim rs As ADODB.Recordset
    Dim strFieldNames As String
    
    Dim dbField As ADODB.Field
    
    On Error GoTo Error_Handler
    
    Set rs = New ADODB.Recordset
    
    rs.Open strSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    ' write the section name to the file
    Print #nFileNumber, strSectionName
    
    ' Retrieve the fields list
    For Each dbField In rs.Fields
        strFieldNames = strFieldNames & dbField.Name & vbTab
    Next
    
    'Write the field list to the file
    Print #nFileNumber, strFieldNames
    
    ' write the records to the file
    While Not rs.EOF
        Print #nFileNumber, rs.GetString(adClipString)
        
        If Not rs.EOF Then
            rs.MoveNext
        End If
    Wend
    
    Print #nFileNumber, ""
    
    ' cleanup
    rs.Close
    Set rs = Nothing


    Exit Sub
    
Error_Handler:

    Set rs = Nothing

    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Sub


'********************************************************************************
'   Sub DumpCustomerInfo(nFileNumber As Integer)
'   DESCRIPTION
'   This dumps the Customer Information to a text file specified by
'   nFileNumber. Data in this section is dumped using a special routine since
'   the Notes field needs to be cleaned-up of tabs and carriage returns.
'
'   PARAMETERS
'   nFileNumber - integer specifying file number of open file
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  May 7, 2003     Joseph Manyaga      Initial Revision
'*********************************************************************************
Sub DumpCustomerInfo(nFileNumber As Integer)

    Const MODULE_NAME = "modExport.bas"
    Const FUNCTION_NAME = "DumpCustomerInfo()"
    Dim strErrorDesc As String

    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim strDataLine As String
    
    Dim dbField As ADODB.Field
    Dim strFieldNames As String
    
    Dim vntFields As Variant
    Dim nCol As Integer
    Dim nRow As Integer
    
    Dim strTemp
    
    On Error GoTo Error_Handler
    
    Set rs = New ADODB.Recordset
    
    strSQL = "SELECT Job.Company, Job.address1 AS [Address Line 1], Job.address2 AS [Address Line 2], Job.address3 AS [Address Line 3], " & _
             "Job.City, Job.State As [State/Province],Job.zip AS [ZIP/Postal Code], Countries.name AS [Country], " & _
             "Contact.name AS [Primary Contact Name], Contact.phone AS [Primary Contact Phone], Contact.email AS [Primary Contact Email], Contact_1.name AS [Secondary Contact Name], Contact_1.phone AS [Secondary Contact Phone], Contact_1.email AS [Secondary Contact Email], Job.Notes " & _
             "FROM Contact, Contact AS Contact_1, Countries " & _
             "INNER JOIN Job ON Countries.code = Job.country " & _
             "WHERE (((Contact.Primary) = True) And ((Contact_1.Primary) <> True))"
             
    rs.Open strSQL, gdbConn

    ' write the section name to the file
    Print #nFileNumber, "CUSTOMER INFORMATION"

    ' Retrieve the fields list
    For Each dbField In rs.Fields
        strFieldNames = strFieldNames & dbField.Name & vbTab
    Next
    
    'Write the field list to the file
    Print #nFileNumber, strFieldNames
    
    vntFields = rs.GetRows
    nCol = 0
    strDataLine = ""
    
    'Stop before the Notes field
    While nCol < UBound(vntFields, 1)
        strDataLine = strDataLine & vntFields(nCol, 0) & vbTab
        nCol = nCol + 1
    Wend
    
    strDataLine = strDataLine & CleanString(CStr(vntFields(nCol, 0)))
    
    'Write the line to the file
    Print #nFileNumber, strDataLine
    
    ' write a blank line
    Print #nFileNumber, ""
    
    ' cleanup
    rs.Close
    Set rs = Nothing
    
    Exit Sub
    
Error_Handler:

    Set rs = Nothing

    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Sub


'********************************************************************************
'   Sub DumpAreaInfo(nFileNumber As Integer)
'   DESCRIPTION
'   This dumps the Area Information to a text file specified by
'   nFileNumber. Data in this section is dumped using a special routine since
'   the Warning field needs to be cleaned-up of tabs and carriage returns.
'
'   PARAMETERS
'   nFileNumber - integer specifying file number of open file
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE        NAME                    CHANGE
'  May 7, 2003      Joseph Manyaga          Initial Revision
'  April 11, 2005   Maria Cecilia Gamboa    Added the user_defined_ct field for export dump
'*********************************************************************************
Sub DumpAreaInfo(nFileNumber As Integer)

    Const MODULE_NAME = "modExport.bas"
    Const FUNCTION_NAME = "DumpAreaInfo()"
    Dim strErrorDesc As String

    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim strDataLine As String
    
    Dim dbField As ADODB.Field
    Dim strFieldNames As String
    
    Dim vntFields As Variant
    Dim nCol As Integer
    Dim nRow As Integer
    
    Dim sTemp As String
    
    On Error GoTo Error_Handler
    
    Set rs = New ADODB.Recordset
    
    strSQL = "SELECT JobArea.area_name AS [Area Name], JobArea.temperature AS [Temperature], JobArea.estimated_hlt AS [Estimated HLT], JobArea.volume As Volume, JobArea.planned_exposure As [Exposure Time], " & _
             "JobArea.fumigant As [Fumigant Required], JobArea.concentration_time AS [Target CT], JobArea.user_defined_ct as [User Defined CT], JobArea.initial_concentration AS [Initial Target CT], JobArea.calculation_warning As [Calculation Warning]" & _
             "FROM JobArea " & _
             "WHERE (((JobArea.deleted) = False)) " & _
             "ORDER BY JobArea.order_entered"
             
    rs.Open strSQL, gdbConn

    ' write the section name to the file
    Print #nFileNumber, "AREA INFORMATION"

    ' Retrieve the fields list
    For Each dbField In rs.Fields
        strFieldNames = strFieldNames & dbField.Name & vbTab
    Next
    
    'Write the field list to the file
    Print #nFileNumber, strFieldNames
    
    If Not rs.BOF And Not rs.EOF Then
        vntFields = rs.GetRows
        nCol = 0
        nRow = 0
        
        While nRow < (UBound(vntFields, 2) + 1)
            strDataLine = ""
            
            'Stop before the Warning field
            While nCol < UBound(vntFields, 1)
                strDataLine = strDataLine & vntFields(nCol, nRow) & vbTab
                nCol = nCol + 1
            Wend
                        
            sTemp = IIf(IsNull(vntFields(nCol, nRow)), "", vntFields(nCol, nRow))
                        
            strDataLine = strDataLine & CleanString(CStr(sTemp))
            
            'Write the line to the file
            Print #nFileNumber, strDataLine
            nRow = nRow + 1
            nCol = 0
        Wend
    End If
    
    ' write a blank line
    Print #nFileNumber, ""
    
    ' cleanup
    rs.Close
    Set rs = Nothing
    
    Exit Sub
    
Error_Handler:

    Set rs = Nothing

    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Sub

'********************************************************************************
'   Sub DumpJobInfo(nFileNumber As Integer)
'   DESCRIPTION
'   This is performs dumps the Job Information to a text file specified by
'   nFileNumber. Data in this section is dumped using a special routine because
'   several fields only contain numbers that have to be determined using various
'   methods from dpcalc
'
'   PARAMETERS
'   nFileNumber - integer specifying file number of open file
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  3/31/2003       Jett Gamboa       Initial Revision
'*********************************************************************************
Sub DumpJobInfo(nFileNumber As Integer)

    Const MODULE_NAME = "modExport.bas"
    Const FUNCTION_NAME = "DumpJobInfo()"
    Dim strErrorDesc As String

    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim strDataLine As String
    
    Dim dbField As ADODB.Field
    Dim strFieldNames As String
    
    Dim strTemp
    
    On Error GoTo Error_Handler
    
    Dim oCalculator As DPCALCLib.ProFumeCalculator
    Set rs = New ADODB.Recordset
    
    ' instantiate a calculator object. We will use it to get some string values
    Set oCalculator = New DPCALCLib.ProFumeCalculator
    
    strSQL = "SELECT Job.name AS [Job Name], Job.location_name AS [Site Name], Job.fumigation_date AS [Fumigation Date], Job.fumigator_name AS [Licensed Fumigator], Job.commodity_enum AS Commodity, Job.fumigation_type_enum AS [Fumigation Type], Job.lifestage_enum AS [Life Stage], Job.treatment_type_enum AS [Pressure Type], Job.load_factor AS [Load Factor], Job.fumigant AS [Amount of Fumigant], Job.cylinders AS Cylinders, Job.avg_concentration_time AS [Average Concentration-Time], Job.avg_initial_concentration AS [Average Initial Concentration], Job.avg_estimated_hlt AS [Average Estimated HLT], Job.calculation_warning AS [Warnings], " & _
             " Job.calc_version AS [DPCalc Version], Job.volume AS [Total Structure Volume], Job.startintro_datetime AS [Start Introduction], Job.startexp_datetime AS [Start Exposure], Job.endexp_datetime AS [End Exposure] " & _
             "FROM Job"
             
    rs.Open strSQL, gdbConn

    ' write the section name to the file
    Print #nFileNumber, "JOB INFORMATION"

    ' Retrieve the fields list
    For Each dbField In rs.Fields
        strFieldNames = strFieldNames & dbField.Name & vbTab
    Next
    
    'Write the field list to the file
    Print #nFileNumber, strFieldNames
    
    While Not rs.EOF
    
        strDataLine = ""
        
        strDataLine = strDataLine & rs("Job Name") & vbTab
        strDataLine = strDataLine & rs("Site Name") & vbTab
        strDataLine = strDataLine & rs("Fumigation Date") & vbTab
        strDataLine = strDataLine & rs("Licensed Fumigator") & vbTab
        
        ' use the calculator to retrieve the name of our commodity
        If rs("Commodity").Value <> -1 Then
            strTemp = oCalculator.GetCommodityString(rs("Commodity").Value)
            strDataLine = strDataLine & strTemp & vbTab
        End If
               
        ' use the calculator to retrieve the name of our fumigation type
        If rs("Fumigation Type").Value <> -1 Then
            strTemp = oCalculator.GetFumigationTypeString(rs("Fumigation Type").Value)
            strDataLine = strDataLine & strTemp & vbTab
        End If
        
        ' use the calculator to retrieve the name of our life stage
        If rs("Life Stage").Value <> -1 Then
            strTemp = oCalculator.GetLifeStageString(rs("Life Stage").Value)
            strDataLine = strDataLine & strTemp & vbTab
        End If
        
        ' use the calculator to retrieve the name of our pressure type
        If rs("Pressure Type").Value <> -1 Then
            strTemp = oCalculator.GetTreatmentTypeString(rs("Pressure Type").Value)
            strDataLine = strDataLine & strTemp & vbTab
        End If
        
        strDataLine = strDataLine & rs("Load Factor") & vbTab
        strDataLine = strDataLine & rs("Amount of Fumigant") & vbTab
        strDataLine = strDataLine & rs("Cylinders") & vbTab
        strDataLine = strDataLine & rs("Average Concentration-Time") & vbTab
        strDataLine = strDataLine & rs("Average Initial Concentration") & vbTab
        strDataLine = strDataLine & rs("Average Estimated HLT") & vbTab
        strDataLine = strDataLine & rs("Warnings") & vbTab
        strDataLine = strDataLine & rs("DPCALC Version") & vbTab
        strDataLine = strDataLine & rs("Total Structure Volume") & vbTab
        strDataLine = strDataLine & rs("Start Introduction") & vbTab
        strDataLine = strDataLine & rs("Start Exposure") & vbTab
        strDataLine = strDataLine & rs("End Exposure") & vbTab
        
        'Write the line to the file
        Print #nFileNumber, strDataLine
    
        rs.MoveNext
    
    Wend
    
    ' write a blank line
    Print #nFileNumber, ""
    
    ' cleanup
    rs.Close
    Set rs = Nothing
    
    Set oCalculator = Nothing


    Exit Sub
    
Error_Handler:

    Set rs = Nothing
    Set oCalculator = Nothing

    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Sub


'********************************************************************************
'   Sub DumpJobPest(nFileNumber As Integer)
'   DESCRIPTION
'   This is performs dumps the Pest Information to a text file specified by
'   nFileNumber. Data in this section is dumped using a special routine because
'   the pest names have to be determined using a method in dpcalc
'
'   PARAMETERS
'   nFileNumber - integer specifying file number of open file
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  3/31/2003       Jett Gamboa       Initial Revision
'*********************************************************************************
Sub DumpJobPest(nFileNumber As Integer)

    Const MODULE_NAME = "modExport.bas"
    Const FUNCTION_NAME = "DumpJobPest()"
    Dim strErrorDesc As String

    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim strDataLine As String
    
    Dim dbField As ADODB.Field
    Dim strFieldNames As String
    
    Dim strTemp
    
    Dim oCalculator As DPCALCLib.ProFumeCalculator
    
    On Error GoTo Error_Handler
    
    ' instantiate a calculator object. We will use it to get the pest names
    Set oCalculator = New DPCALCLib.ProFumeCalculator
    
    Set rs = New ADODB.Recordset
    
    'JNM 04/21/03 - AT Fix - Get the undeleted records only
    strSQL = "SELECT pest_enum As [Pest Name] FROM JobPest WHERE deleted = false"
    
    rs.Open strSQL, gdbConn

    ' write the section name to the file
    Print #nFileNumber, "JOB PEST"

    ' Retrieve the fields list
    For Each dbField In rs.Fields
        strFieldNames = strFieldNames & dbField.Name & vbTab
    Next
    
    'Write the field list to the file
    Print #nFileNumber, strFieldNames
    
    While Not rs.EOF
    
        strDataLine = ""
        
        strTemp = oCalculator.GetPestString(rs("Pest Name").Value)
        strDataLine = strDataLine & strTemp & vbTab
        
        'Write the pest name to the file
        Print #nFileNumber, strDataLine
                
        rs.MoveNext
        
    Wend
    
    ' write a blank line
    Print #nFileNumber, ""
    
    ' cleanup
    rs.Close
    Set rs = Nothing
    
    Set oCalculator = Nothing


    Exit Sub
       
Error_Handler:

    Set rs = Nothing
    Set oCalculator = Nothing

    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Sub


'********************************************************************************
'   Sub DumpAreaStatus(nFileNumber As Integer)
'   DESCRIPTION
'   This is performs dumps the Area Status. The actual status string is determined
'   by calling the GetJobAreaStatus subroutine in the modGetJobAreaStatus module
'
'   PARAMETERS
'   nFileNumber - integer specifying file number of open file
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  3/31/2003       Jett Gamboa       Initial Revision
'  2/9/2004        Pat San Miguel    SIMS 2118618: Change the SQL statement strSQL
'*********************************************************************************
Sub DumpAreaStatus(nFileNumber As Integer)

    Const MODULE_NAME = "modExport.bas"
    Const FUNCTION_NAME = "DumpAreaStatus()"
    Dim strErrorDesc As String

    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    
    Dim strSQL As String
    Dim strSQL2 As String
    Dim strDataLine As String
    
    Dim dbField As ADODB.Field
    Dim strFieldNames As String
    
    Dim x As Integer
    
    Dim strTemp
    
    On Error GoTo Error_Handler
       
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    
    'JNM 04/22/03 - GetJobAreaStatus also needs add_profume and mean_ntc
    'strSQL = "SELECT DISTINCT JobArea.area_name AS [Area Name], JobArea.hlt As HLT, JobArea.Achieved_concentration_time AS [Achieved CT], " & _
    '         "JobArea.projected_concentration_time AS [Projected CT], JobArea.status AS [Area Status], " & _
    '         "JobArea.status, JobArea.projected_time, JobArea.id, JobArea.fumigant " & _
    '         "FROM JobArea " & _
    '         "WHERE JobArea.deleted = False"
    'strSQL = "SELECT DISTINCT JobArea.area_name AS [Area Name], JobArea.hlt As HLT, JobArea.Achieved_concentration_time AS [Achieved CT], " & _
    '        "JobArea.projected_concentration_time AS [Projected CT], JobArea.status AS [Area Status], " & _
    '         "JobArea.status, JobArea.projected_time, JobArea.id, JobArea.fumigant, JobArea.add_profume, JobArea.mean_ntc " & _
    '         "FROM JobArea " & _
    '         "WHERE JobArea.deleted = False"
    'Pat 02/09/2004 SIMS Ticket 2118618: Changed the SQL statement to get the correct values of HLT and Dosage for each Area Name
    strSQL = "SELECT * FROM " & _
             "(SELECT DISTINCT JobArea.area_name AS [Area Name], JobArea.hlt As HLT, JobArea.Achieved_concentration_time AS [Achieved CT], " & _
             "JobArea.projected_concentration_time AS [Projected CT], JobArea.status AS [Area Status], " & _
             "JobArea.status , JobArea.projected_time, JobArea.id, JobArea.fumigant, JobArea.add_profume, JobArea.mean_ntc, order_entered " & _
             "From JobArea WHERE JobArea.deleted = False) ORDER BY order_entered"
             
    ' we need this second sql statement because the GetJobAreaStatus routine needs this.
    ' This is bad, we may want to find another way to do this.
    'JNM 04/22/03 - GetJobAreaStatus also needs the endexp_datetime field
    strSQL2 = "SELECT startexp_datetime, endexp_datetime FROM Job"
    
    rs.Open strSQL, gdbConn
    rs2.Open strSQL2, gdbConn

    ' write the section name to the file
    Print #nFileNumber, "AREA STATUS"

    ' initialize our counter
    x = 0

    ' Retrieve the fields list
    For Each dbField In rs.Fields
    
        ' We have to process area status fields differently. Our SQL statement includes
        ' fields that we don't really dump but we need to fetch because GetJobAreaStatus
        ' needs a recordset with those fields
        If x < 5 Then
            strFieldNames = strFieldNames & dbField.Name & vbTab
        End If
        
        x = x + 1
        
    Next
    
    'Write the field list to the file
    Print #nFileNumber, strFieldNames
    
    While Not rs.EOF
    
        strDataLine = ""
        
        strDataLine = strDataLine & rs("Area Name") & vbTab
        strDataLine = strDataLine & rs("HLT") & vbTab
        strDataLine = strDataLine & rs("Achieved CT") & vbTab
        strDataLine = strDataLine & rs("Projected CT") & vbTab
        
        ' Call GetJobAreaStatus to get the status for our area and append it
        ' to the line
        ' 5/22/03 Jett - "clean" the newlines in the status messages
        strTemp = CleanString(GetJobAreaStatus(rs, rs2))
        strDataLine = strDataLine & strTemp & vbTab
        
        ' Write the area status to the file
        Print #nFileNumber, strDataLine
    
        rs.MoveNext
        
    Wend
    
    ' write a blank line
    Print #nFileNumber, ""
    
    rs2.Close
    Set rs2 = Nothing
    
    ' cleanup
    rs.Close
    Set rs = Nothing
    
    Exit Sub
    
Error_Handler:

    Set rs = Nothing
    Set rs2 = Nothing

    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc
   
    
End Sub

'********************************************************************************
'   Function CleanString(sString as string)
'   DESCRIPTION
'   This removes tabs and carriage returns within a string.
'
'   PARAMETERS
'   sString = the string to be cleaned
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  5/7/2003       Joseph manyaga       Initial Revision
'  1/23/2004      Jett Gamboa          If it encounters a carriage return this will be replaced by &1310&
'                                      SIMS Ticket 2116769
'*********************************************************************************
Function CleanString(sString As String) As String
    Dim sTemp As String
    Dim sReplace As String
    Dim sReturn As String
    
    sReplace = "     "
    sReturn = "&1310&"
    
    sTemp = Replace(sString, vbTab, sReplace)
    sTemp = Replace(sTemp, vbCrLf, sReturn)
    sTemp = Replace(sTemp, vbCr, sReturn)
    sTemp = Replace(sTemp, vbLf, sReturn)
    sTemp = Replace(sTemp, vbNewLine, sReturn)
    
    CleanString = sTemp
    
End Function

'********************************************************************************
'   Function OpenFumiguideReport()
'   DESCRIPTION
'   This function open the fumiguide report document
'
'   PARAMETERS
'   sString = the string to be cleaned
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  1/18/2006       Anabel Tomas      Initial Revision
'  3/29/2007       David Smith       Changed to one click execute.
'*********************************************************************************
Function OpenFumiguideReport()
   Dim wrdApp As Object
   Dim wrdDoc As Object
   Dim strFileName As String
   strFileName = "C:\Program Files\DAS\Fumiguide\" & LoadResString(IDS_FUMIGUIDE_REPORT)
    
   
   
    Set wrdApp = New Word.Application
    With wrdApp
        .Visible = True
        .Documents.Open (strFileName)
        '.ActiveDocument.ShowMacro (strExportFilename)
        '.Run MacroName:="main(strExportFilename)"
        
        .ActiveDocument.Variables.Add Name:="ExportFilename", Value:=strExportFilename
        
        If gbIsMetric = True Then
            .ActiveDocument.Variables.Add Name:="IsMetric", Value:=True
        Else
            .ActiveDocument.Variables.Add Name:="IsMetric", Value:=False
        End If
        
        .Run MacroName:="main"
        
        
    End With
    Set wrdApp = Nothing
End Function


