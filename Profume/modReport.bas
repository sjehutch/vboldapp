Attribute VB_Name = "modReport"
'------------------------------------------------------
'*
'* Filename: modReport.bas
'*
'* DESCRIPTION
'* -----------
'* This module prints the ProFume Fumiguide Report.
'*
'* Copyright (c) 2002 Accenture.  All rights reserved.
'-----------------------------------------------------
'*
'* REVISION HISTORY
'* Revision:    0
'* Author:      Manyaga, Joseph
'* Date:        October 4, 2002
'*
'-----------------------------------------------------

Option Explicit

'Declare constants
Private Const MODULE_NAME = "modReport"
Private Const BOTTOM_MARGIN_TWIPS As Integer = 1440

'Declare variables
Dim nDecimal As Integer     'the number of decimal digits in the results
Dim sJobSQL As String

'Create objects
Dim rsJob As New ADODB.Recordset

'Declare module-specific variables. These variables would contain the
'maximum number of twips and characters that each item in the report
'would occupy.
Dim nMaxLabel As Integer
Dim nMaxSpace1 As Integer
Dim nMaxSpace2 As Integer
Dim nMaxValue As Integer
Dim nMaxArea As Integer
Dim nMaxHLT As Integer
Dim nMaxCT As Integer
Dim nMaxStatus As Integer
Dim nMaxLabelChar As Integer
Dim nMaxSpaceChar1 As Integer
Dim nMaxSpaceChar2 As Integer
Dim nMaxValueChar As Integer
Dim nMaxAreaChar As Integer
Dim nMaxHLTChar As Integer
Dim nMaxCTChar As Integer
Dim nMaxStatusChar As Integer

Public Sub PrintMain()
'**************************************************************
'   DESCRIPTION
'       This is the main routine of the Report module.
'       We are treating the output pages as tables. This means that we will
'       be printing the report items in separate columns. The first
'       page would have 3 columns (the label, a space, the actual value).
'       The Area Status page would have 7 columns (Area Name, HLT, CT,
'       Ending Status, 3 spaces).
'
'   PARAMETERS
'           NONE
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************

On Error GoTo Error_Handler

'Declare the constants that will be used in this subroutine
Const PRINTER_FONT As String = "Courier"
Const PRINTER_FONTSIZE As Integer = 12
'The constants below contain the width of the items to be printed in
'terms of the percentage of the page width that these items would occupy.
'first page
Const WIDTH_LABEL As Double = 0.45
Const WIDTH_SPACE1 As Double = 0.05
Const WIDTH_VALUE As Double = 0.5
'second page
Const WIDTH_AREA As Double = 0.35
Const WIDTH_SPACE2 As Double = 0.02
Const WIDTH_HLT As Double = 0.1
Const WIDTH_CT As Double = 0.1
Const WIDTH_STATUS As Double = 0.39
'Sample text to be used in determining the twips per character ratio.
Const SAMPLE_TEXT As String = "ABCDE"

Dim nTwipPerChar As Integer  'the number of twips per printed character
Dim dlgPrint As CommonDialog

Set dlgPrint = frmMain.dlgCommonDialog

'Check if the user really wants to print
If PrintDialog(dlgPrint) = False Then
    Set rsJob = Nothing
    Set dlgPrint = Nothing
    Exit Sub
End If

'Initialize the printer settings
With Printer
    .Orientation = dlgPrint.Orientation
    .FontName = PRINTER_FONT
    .FontSize = PRINTER_FONTSIZE
    .Font.Bold = False
    .Copies = dlgPrint.Copies
End With

'Get the maximum number of twips for each column. Round this value
'to the lower integer part.
nMaxLabel = Int(WIDTH_LABEL * (Printer.ScaleWidth))
nMaxSpace1 = Int(WIDTH_SPACE1 * (Printer.ScaleWidth))
nMaxValue = Int(WIDTH_VALUE * (Printer.ScaleWidth))
nMaxArea = Int(WIDTH_AREA * (Printer.ScaleWidth))
nMaxSpace2 = Int(WIDTH_SPACE2 * (Printer.ScaleWidth))
nMaxHLT = Int(WIDTH_HLT * (Printer.ScaleWidth))
nMaxCT = Int(WIDTH_CT * (Printer.ScaleWidth))
nMaxStatus = Int(WIDTH_STATUS * (Printer.ScaleWidth))

'Convert the maximum number of twips to the equivalent number of characters.
nTwipPerChar = Int(Printer.TextWidth(SAMPLE_TEXT) / Len(SAMPLE_TEXT))
nMaxLabelChar = Int(nMaxLabel / nTwipPerChar)
nMaxSpaceChar1 = Int(nMaxSpace1 / nTwipPerChar)
nMaxValueChar = Int(nMaxValue / nTwipPerChar)
nMaxAreaChar = Int(nMaxArea / nTwipPerChar)
nMaxSpaceChar2 = Int(nMaxSpace2 / nTwipPerChar)
nMaxHLTChar = Int(nMaxHLT / nTwipPerChar)
nMaxCTChar = Int(nMaxCT / nTwipPerChar)
nMaxStatusChar = Int(nMaxStatus / nTwipPerChar)

nDecimal = IIf(gbIsMetric, 2, 0)

'We are ready to print!!!
'Print the report heading in the first page.
Call PrintPageOne(LoadResString(IDS_PRINT_REPORT_TITLE), "")
Call PrintPageOne(" ", " ")

'Print the General Info section of the report.
Call PrintCompanyInfo
Call PrintCustomerInfo
Call PrintContactInfo
Call PrintJobInfo

Printer.NewPage

'Print the Area Status section of the report.
Call PrintAreas

'Go forth and print the report!
Printer.EndDoc

'Clean-up time!
rsJob.Close
Set rsJob = Nothing

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "PrintMain", Err.Description)
End If

End Sub


Private Function PrintDialog(ByRef PFPrintDialog As CommonDialog) As Boolean
'**************************************************************
'   DESCRIPTION
'       Displays the Printer Dialog box to the user.
'
'   PARAMETERS
'           Printer Dialog object - the Common Dialog control object
'               which is defined in the parent window module.
'
'   RETURNS - True if the user clicks on the Ok button
'             False if the user clicks on the Cancel button
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************
'We're using a different error handling style here. We need to catch
'the event that the user is clicking on the Cancel button.
On Error Resume Next

With PFPrintDialog
    .PrinterDefault = False
    .Flags = cdlPDDisablePrintToFile Or cdlPDNoPageNums Or cdlPDNoSelection
    .CancelError = True
    .ShowPrinter
    If Err = 0 Then
        PrintDialog = True
    Else
        PrintDialog = False
    End If
End With

End Function

Private Sub PrintCompanyInfo()
'**************************************************************
'   DESCRIPTION
'       Prints the Company Information section.
'       The Company Information is retrieved from the registry.
'       Public functions will be called in order to read the values from
'       the registry.
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

'Declare variables
Dim sCompany As String
Dim sAddress1 As String
Dim sAddress2 As String
Dim sAddress3 As String
Dim sCity As String
Dim sState As String
Dim sZip As String
Dim sCountry As String
Dim sName As String

'Get the values from the registry.
sCompany = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_COMPANY)
sAddress1 = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_ADDRESS1)
sAddress2 = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_ADDRESS2)
sAddress3 = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_ADDRESS3)
sCity = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_CITY)
sState = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_STATE)
sZip = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_ZIP)
sCountry = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_COUNTRY)
sName = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_NAME)

'Create the items to be printed.
Call PrintPageOne(LoadResString(IDS_PRINT_COMPANY), sCompany)
Call PrintPageOne(LoadResString(IDS_PRINT_ADDRESS), sAddress1)
If Trim(sAddress2) <> "" Then
    Call PrintPageOne("", sAddress2)
End If
If Trim(sAddress3) <> "" Then
    Call PrintPageOne("", sAddress3)
End If
Call PrintPageOne(LoadResString(IDS_PRINT_CITY), sCity)
Call PrintPageOne(LoadResString(IDS_PRINT_STATE), sState)
Call PrintPageOne(LoadResString(IDS_PRINT_ZIP), sZip)
Call PrintPageOne(LoadResString(IDS_PRINT_COUNTRY), sCountry)
Call PrintPageOne(LoadResString(IDS_PRINT_CONTACT), sName)

'Print a space before the next section
Call PrintPageOne(" ", " ")

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "PrintCompanyInfo", Err.Description)
End If

End Sub

Private Sub PrintCustomerInfo()
'**************************************************************
'   DESCRIPTION
'       Prints the Customer Information section.
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

'Declare the variables
Dim sCustomerSQL As String
Dim rsCustomer As New ADODB.Recordset
Dim sCompany As String
Dim sAddress1 As String
Dim sAddress2 As String
Dim sAddress3 As String
Dim sCity As String
Dim sState As String
Dim sZip As String
Dim sCountry As String

'Create the recordset
sCustomerSQL = "Select company, address1, address2, address3, city, state, zip, country " & _
    "from Job where deleted = False"
rsCustomer.Open sCustomerSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
rsCustomer.MoveFirst

'Translate "Null" to "" if value is "Null".
sCompany = IIf(IsNull(rsCustomer("company")), "", rsCustomer("company"))
sAddress1 = IIf(IsNull(rsCustomer("address1")), "", rsCustomer("address1"))
sAddress2 = IIf(IsNull(rsCustomer("address2")), "", rsCustomer("address2"))
sAddress3 = IIf(IsNull(rsCustomer("address3")), "", rsCustomer("address3"))
sCity = IIf(IsNull(rsCustomer("city")), "", rsCustomer("city"))
sState = IIf(IsNull(rsCustomer("state")), "", rsCustomer("state"))
sZip = IIf(IsNull(rsCustomer("zip")), "", rsCustomer("zip"))
sCountry = IIf(IsNull(rsCustomer("country")), "", rsCustomer("country"))

'Go forth and print
Call PrintPageOne(LoadResString(IDS_PRINT_CUSTOMER), sCompany)
Call PrintPageOne(LoadResString(IDS_PRINT_ADDRESS), sAddress1)
If Trim(sAddress2) <> "" Then
    Call PrintPageOne("", sAddress2)
End If
If Trim(sAddress3) <> "" Then
    Call PrintPageOne("", sAddress3)
End If
Call PrintPageOne(LoadResString(IDS_PRINT_CITY), sCity)
Call PrintPageOne(LoadResString(IDS_PRINT_STATE), sState)
Call PrintPageOne(LoadResString(IDS_PRINT_ZIP), sZip)
Call PrintPageOne(LoadResString(IDS_PRINT_COUNTRY), sCountry)

'And of course, clean-up your mess
rsCustomer.Close
Set rsCustomer = Nothing

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "PrintCustomerInfo", Err.Description)
End If

End Sub

Private Sub PrintContactInfo()
'**************************************************************
'   DESCRIPTION
'       Prints the Customer Contact Information section.
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

'Declare the variables
Dim sContactSQL As String
Dim sContact As String
Dim bOnlyOne As Boolean
Dim rsContact As New ADODB.Recordset
Dim sPhone As String
Dim sEmail As String

bOnlyOne = True

'Create the recordset
sContactSQL = "Select Contact.* from Contact where deleted = False"
rsContact.Open sContactSQL, gdbConn, adOpenForwardOnly, adLockReadOnly

'Check if the recordset is empty or not
If rsContact.BOF And rsContact.EOF Then
    'Print that no data is available
    Call PrintPageOne(LoadResString(IDS_PRINT_CUST_CONTACT), LoadResString(IDS_PRINT_NONE))
Else
    rsContact.MoveFirst
    While Not rsContact.EOF
        'Go forth and print
        sContact = IIf(IsNull(rsContact("name")), "", rsContact("name"))
        sPhone = IIf(IsNull(rsContact("phone")), "", rsContact("phone"))
        sEmail = IIf(IsNull(rsContact("email")), "", rsContact("email"))
        
        'Attach the phone and email if these have values
        If Len(sPhone) > 0 Or Len(sEmail) > 0 Then
            sContact = sContact & " ("
            sContact = sContact & sPhone
            If Len(sPhone) > 0 And Len(sEmail) > 0 Then
                sContact = sContact & ", "
            End If
            sContact = sContact & sEmail & ")"
        End If
        'Do not display the label if we are beyond the first record
        If bOnlyOne Then
            Call PrintPageOne(LoadResString(IDS_PRINT_CUST_CONTACT), sContact)
        Else
            Call PrintPageOne("", sContact)
        End If
        rsContact.MoveNext
        bOnlyOne = False
    Wend
End If

'Print a space before the next section
Call PrintPageOne(" ", " ")

'And of course, clean-up your mess
rsContact.Close
Set rsContact = Nothing

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "PrintContactInfo", Err.Description)
End If

End Sub

Private Sub PrintJobInfo()
'**************************************************************
'   DESCRIPTION
'       Prints the Job Information section.
'
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

'Declare the variables
Dim sJobSQL As String
Dim dVolume As Double
Dim sFumigationDate As String
Dim dConVolume As Double
Dim sUnitName As String
Dim sPest As String
Dim sFumigationType As String
Dim sLifeStage As String
Dim sCommodity As String
Dim sTreatmentType As String
Dim sLoadFactor As String
Dim sLocation As String
Dim sFumigatorName As String

Dim pCalc As New ProFumeCalculator
Dim rsJobPest As New ADODB.Recordset

'Create the recordset
sJobSQL = "Select Job.* from Job where deleted = False and id = '" & gsJobID & "'"
rsJob.Open sJobSQL, gdbConn, adOpenForwardOnly, adLockReadOnly

'It is assumed that there will always be a Job record in the database because
'the default database always has a record.
rsJob.MoveFirst

sLocation = IIf(IsNull(rsJob("location_name")), "", rsJob("location_name"))
sFumigationDate = IIf(IsNull(rsJob("fumigation_date")), "", rsJob("fumigation_date"))
sFumigatorName = IIf(IsNull(rsJob("fumigator_name")), "", rsJob("fumigator_name"))

Call PrintPageOne(LoadResString(IDS_PRINT_LOCATION), sLocation)
'For the structure total volume, we need to convert the units of the volume
'to the preference of the user.
If rsJob("volume") = "" Then
    Call PrintPageOne(LoadResString(IDS_PRINT_STRUCT_VOL), "")
Else
    dConVolume = 0#
    sUnitName = ""
    dVolume = IIf(rsJob("volume") < 0#, 0, rsJob("volume"))
    Call UnitConvert(modeDisplay, unitVOLUME, dVolume, dConVolume, sUnitName)
    dConVolume = Round(dConVolume, nDecimal)
    Call PrintPageOne(LoadResString(IDS_PRINT_STRUCT_VOL), FormatNumber(dConVolume, nDecimal) & " " & sUnitName)
End If
Call PrintPageOne(" ", " ")

Call PrintPageOne(LoadResString(IDS_PRINT_FUM_DATE), sFumigationDate)
Call PrintPageOne(" ", " ")

Call PrintPageOne(LoadResString(IDS_PRINT_LIC_FUM), sFumigatorName)
Call PrintPageOne(" ", " ")

'Create the Job Pest recordset
Call CreatePestRS(rsJobPest)
If rsJobPest.BOF And rsJobPest.EOF Then
    'Print that no data is available
    Call PrintPageOne(LoadResString(IDS_PRINT_TARGET_PESTS), LoadResString(IDS_PRINT_NONE))
Else
    rsJobPest.MoveFirst
    sPest = ""
    'Get the equivalent strings of the pest enums using the Profume Calculator
    While Not rsJobPest.EOF
        sPest = sPest & pCalc.GetPestString(rsJobPest("pest_enum"))
        sPest = sPest & ", "
        rsJobPest.MoveNext
    Wend
    'Remove the last comma from the string.
    sPest = Left(sPest, (InStrRev(sPest, ",") - 1))
    'Go forth and print
    Call PrintPageOne(LoadResString(IDS_PRINT_TARGET_PESTS), sPest)
End If

'Get the equivalent strings of the enum values. Use 0 if enum value is -1.
sFumigationType = pCalc.GetFumigationTypeString(IIf(rsJob("fumigation_type_enum") < 0, _
    0, rsJob("fumigation_type_enum")))
sLifeStage = pCalc.GetLifeStageString(IIf(rsJob("lifestage_enum") < 0, _
    0, rsJob("lifestage_enum")))
sCommodity = pCalc.GetCommodityString(IIf(rsJob("commodity_enum") < 0, _
    0, rsJob("commodity_enum")))
sTreatmentType = pCalc.GetTreatmentTypeString(IIf(rsJob("treatment_type_enum") < 0, _
    0, rsJob("treatment_type_enum")))
sLoadFactor = IIf(rsJob("load_factor") <> "", rsJob("load_factor") & "%", "")

'Print em all!!!
Call PrintPageOne(LoadResString(IDS_PRINT_FUMIGATION_TYPE), sFumigationType)
Call PrintPageOne(LoadResString(IDS_PRINT_LIFE_STAGES), sLifeStage)
Call PrintPageOne(LoadResString(IDS_PRINT_COMMODITY), sCommodity)
Call PrintPageOne(LoadResString(IDS_PRINT_TREATMENT_TYPE), sTreatmentType)
Call PrintPageOne(LoadResString(IDS_PRINT_STRUCT_LOAD_FACTOR), sLoadFactor)

'And of course, who would forget the space
Call PrintPageOne(" ", " ")

Call PrintTotals

'And of course, clean-up your mess
rsJobPest.Close

Set rsJobPest = Nothing
Set pCalc = Nothing

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "PrintJobInfo", Err.Description)
End If

End Sub
Private Sub CreatePestRS(ByRef rsJobPest As ADODB.Recordset)
'**************************************************************
'   DESCRIPTION
'       This subroutine creates the Job Pest recordset.
'
'   PARAMETERS
'           rsJobPest = the Job Pest recordset which is initially empty
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************

On Error GoTo Error_Handler

Dim sPestSQL As String

sPestSQL = "select pest_enum from JobPest where deleted = False"
rsJobPest.Open sPestSQL, gdbConn, adOpenForwardOnly, adLockReadOnly

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "CreatePestRS", Err.Description)
End If

End Sub

Private Sub PrintTotals()
'**************************************************************
'   DESCRIPTION
'       Prints the Totals section.
'
'
'   PARAMETERS
'
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************

On Error GoTo Error_Handler

Dim dFumigant As Double
Dim sFumigantUnit As String
Dim lTotalExposureTime As Long

dFumigant = 0#
sFumigantUnit = ""

Call UnitConvert(modeDisplay, unitWEIGHT, rsJob("fumigant"), _
                    dFumigant, sFumigantUnit)
Call PrintPageOne(LoadResString(IDS_PRINT_TOTAL_FUM), _
    FormatNumber(Round(dFumigant, 2), 2) & " " & sFumigantUnit)

Dim MonitorExpert As New clsMonitorExpert
If MonitorExpert.GetActualExposureTime(gsJobID, lTotalExposureTime) Then
    Call PrintPageOne(LoadResString(IDS_PRINT_TOTAL_EXP), _
        FormatNumber(Round(lTotalExposureTime / 3600, 2), 2) & " hours")
Else
    Call PrintPageOne(LoadResString(IDS_PRINT_TOTAL_EXP), LoadResString(IDS_PRINT_NONE))
End If

Set MonitorExpert = Nothing

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "PrintTotals", Err.Description)
End If

End Sub
Private Sub PrintAreas()
'**************************************************************
'   DESCRIPTION
'       This is the subroutine that generates the items that are
'       going to be printed on the Area Status page.
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

Dim rsJobArea As New ADODB.Recordset

'Create the JobArea recordset
Call CreateJobAreaRS(rsJobArea)

Printer.Font.Bold = True
Call PrintPageTwo(LoadResString(IDS_PRINT_AREA_NAME), LoadResString(IDS_PRINT_HLT), LoadResString(IDS_PRINT_ACHIEVED_CT), _
LoadResString(IDS_PRINT_ENDING_STATUS))

Printer.Font.Bold = False

'Do not print anything if there are no Job Area records
If rsJobArea.BOF And rsJobArea.EOF Then
    'do not do anything
Else
    rsJobArea.MoveFirst
    While Not rsJobArea.EOF
        Call PrintArea(rsJobArea)
        rsJobArea.MoveNext
    Wend
End If

'Clean-up time
rsJobArea.Close
Set rsJobArea = Nothing

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "PrintAreas", Err.Description)
End If

End Sub
Private Sub CreateJobAreaRS(ByRef rsJobArea As ADODB.Recordset)
'**************************************************************
'   DESCRIPTION
'       This subroutine creates the Job Area recordset.
'
'   PARAMETERS
'           rsJobArea = the Job Area recordset which is initially empty
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************

On Error GoTo Error_Handler

Dim sJobAreaSQL As String

sJobAreaSQL = "select * from JobArea where JobArea.deleted = false and " & _
    "job_fkey = '" & gsJobID & "' order by JobArea.order_entered"
rsJobArea.Open sJobAreaSQL, gdbConn, adOpenForwardOnly, adLockReadOnly

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "CreateJobAreaRS", Err.Description)
End If

End Sub


Private Sub PrintArea(ByRef rsJobArea As ADODB.Recordset)
'**************************************************************
'   DESCRIPTION
'       This subroutine creates the row to be printed in the Area
'       status page.
'
'   PARAMETERS
'           rsJobArea = the Job Area recordset
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'   12/12/02       Jett Gamboa       do not display achieved CT if we did not calculate
'**************************************************************

On Error GoTo Error_Handler

Dim bHasActiveMonitorPoints As Boolean
Dim sMonitorPointSQL As String
Dim rsMonitorPoint As New ADODB.Recordset
Dim sAreaName As String
Dim sHLT As String
Dim sCT As String
Dim sStatus As String

bHasActiveMonitorPoints = False
sAreaName = ""
sHLT = ""
sCT = ""
sStatus = ""

'Check if the Job Area has active monitor points.
sMonitorPointSQL = "select top 1 * from MonitorPoint where " & _
    "MonitorPoint.deleted = false and " & _
    "MonitorPoint.jobarea_fkey = '" & rsJobArea("id") & "'"
rsMonitorPoint.Open sMonitorPointSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
If Not rsMonitorPoint.BOF And Not rsMonitorPoint.EOF Then
    bHasActiveMonitorPoints = True
End If

'Set-up the values to be printed
'Don't worry about the area name, the app would ensure that this is not null.
sAreaName = rsJobArea("area_name")
If bHasActiveMonitorPoints Then
    sHLT = IIf(rsJobArea("hlt") >= 0#, Round(rsJobArea("hlt"), 2), " ")
    sCT = IIf(rsJobArea("achieved_concentration_time"), _
        Round(rsJobArea("achieved_concentration_time"), 2), " ")
    
    ' There may be instance were we don't calculate (like when elapsed times for
    ' monitoring points are not within each other, set it to blank so we do
    ' not display anything on the report
    If sCT = "-1" Then
        sCT = " "
    End If
    
    sStatus = GetJobAreaStatus(rsJobArea, rsJob)
Else
    sStatus = LoadResString(IDS_PRINT_NO_ACTIVE_MONITOR_POINTS)
End If

Call PrintPageTwo(sAreaName, sHLT, sCT, sStatus)

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "PrintArea", Err.Description)
End If

End Sub



Private Sub PrintPageOne(ByVal sLabel As String, ByVal sValue As String)
'**************************************************************
'   DESCRIPTION
'       This is the subroutine that prints the items in the first page.
'
'   PARAMETERS
'           sLabel = the label
'           sValue = the actual value to be printed
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************

On Error GoTo Error_Handler

'Declare variables.
Dim sValueFirst As String
Dim sValueLast As String
Dim nCurrPosX As Integer
Dim nCurrPosY As Integer

'Start printing on a new page if there is no more space in the current page.
If Printer.CurrentY >= (Printer.ScaleHeight - BOTTOM_MARGIN_TWIPS) Then
    Printer.NewPage
End If

'Initialize variables.
'We are only word-wrapping the value part because it is assumed that the labels
'fit the Label column.
sValueFirst = ""
sValueLast = ""

'Call the WordWrap subroutine in order to fit the text within the margin
Call WordWrap(sValue, nMaxValueChar, sValueFirst, sValueLast)

'Print the strings
nCurrPosX = Printer.CurrentX
nCurrPosY = Printer.CurrentY
Printer.CurrentX = nMaxLabel - Printer.TextWidth(sLabel)
Printer.Print sLabel
Printer.CurrentX = nCurrPosX + nMaxLabel + nMaxSpace1
Printer.CurrentY = nCurrPosY
Printer.Print sValueFirst

'Call this subroutine again if sValue was cut.
If sValueLast <> "" Then
   Call PrintPageOne("", sValueLast)
End If

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "PrintPageOne", Err.Description)
End If

End Sub
Private Sub PrintPageTwo(ByVal sArea As String, ByVal sHLT As String, _
ByVal sCT As String, ByVal sStatus As String)
'**************************************************************
'   DESCRIPTION
'       This is the subroutine that does the actual printing of the
'       items in the Area Status page.
'
'   PARAMETERS
'           sArea = the Area Name string to be printed
'           sHLT = the HLT string to be printed
'           sCT = the CT string to be printed
'           sSTatus = the status string to be printed
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************

On Error GoTo Error_Handler

'Declare variables.
Dim sAreaFirst As String
Dim sAreaLast As String
Dim sHLTFirst As String
Dim sHLTLast As String
Dim sCTFirst As String
Dim sCTLast As String
Dim sStatusFirst As String
Dim sStatusLast As String

Dim nCurrPosX As Integer
Dim nCurrPosY As Integer

'Check if there is still a space in the current page before printing.
If Printer.CurrentY >= (Printer.ScaleHeight - BOTTOM_MARGIN_TWIPS) Then
    Printer.NewPage
End If

'Initialize variables.
sAreaFirst = ""
sAreaLast = ""
sHLTFirst = ""
sHLTLast = ""
sCTFirst = ""
sCTLast = ""
sStatusFirst = ""
sStatusLast = ""

'Call the WordWrap subroutine in order to fit the text within the margin
Call WordWrap(sArea, nMaxAreaChar, sAreaFirst, sAreaLast)
Call WordWrap(sHLT, nMaxHLTChar, sHLTFirst, sHLTLast)
Call WordWrap(sCT, nMaxCTChar, sCTFirst, sCTLast)
Call WordWrap(sStatus, nMaxStatusChar, sStatusFirst, sStatusLast)


'Print the strings
nCurrPosX = Printer.CurrentX
nCurrPosY = Printer.CurrentY
Printer.CurrentX = nMaxArea - Printer.TextWidth(sAreaFirst)
Printer.Print sAreaFirst
Printer.CurrentX = nCurrPosX + nMaxArea + nMaxSpace2
Printer.CurrentY = nCurrPosY
Printer.Print sHLTFirst
Printer.CurrentX = nCurrPosX + nMaxArea + nMaxSpace2 + nMaxHLT + nMaxSpace2
Printer.CurrentY = nCurrPosY
Printer.Print sCTFirst
Printer.CurrentX = nCurrPosX + nMaxArea + nMaxSpace2 + nMaxHLT + nMaxSpace2 _
    + nMaxCT + nMaxSpace2
Printer.CurrentY = nCurrPosY
Printer.Print sStatusFirst

'Call this subroutine again if any string was cut by the WordWrap subroutine
If sAreaLast <> "" Or sHLTLast <> "" Or sCTLast <> "" Or sStatusLast <> "" Then
   Call PrintPageTwo(sAreaLast, sHLTLast, sCTLast, sStatusLast)
End If

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "PrintPageTwo", Err.Description)
End If

End Sub

Private Sub WordWrap(ByVal sWholeWord As String, ByVal nMaxWidth As Integer, _
ByRef sFirstPart As String, ByRef sLastPart As String)
'**************************************************************
'   DESCRIPTION
'       This is the subroutine that performs the word-wrapping procedure.
'
'   PARAMETERS
'           sWholeWord = the string to be word-wrapped
'           nMaxWidth = the maximum number of characters allowed for the column
'           sFirtPart = initially a NULL variable but would be updated with the
'                       first part of sWholeWord that would fit within nMaxWidth
'           sLastPart = initially a NULL variable but would be updated with the
'                       second part of sWholeWord that is beyond the last allowable
'                       character
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************

On Error GoTo Error_Handler

'Declare variables
Dim sTempFirstPart As String    'variable that would the First Part value temporarily
Dim nSpacePos As Integer        'variable that would hold the location of the first space from the last character

'Initialize the variables
sTempFirstPart = ""
nSpacePos = 0

'Cut the string if it exceeds nMaxWidth.
If Len(sWholeWord) > nMaxWidth Then
    sTempFirstPart = Left(sWholeWord, nMaxWidth)
    nSpacePos = InStrRev(sTempFirstPart, " ")
    If nSpacePos = 0 Then
        sFirstPart = sTempFirstPart
        sLastPart = Right(sWholeWord, Len(sWholeWord) - nMaxWidth)
    Else
        sFirstPart = Left(sTempFirstPart, nSpacePos)
        sLastPart = Right(sWholeWord, Len(sWholeWord) - nSpacePos)
    End If
Else
    sFirstPart = sWholeWord
    sLastPart = ""
End If

Error_Handler:
If Err.Number <> 0 Then
    Printer.KillDoc
    Call psAbort(MODULE_NAME, "WordWrap", Err.Description)
End If

End Sub


