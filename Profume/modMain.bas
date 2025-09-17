Attribute VB_Name = "modMain"
Option Explicit
'------------------------------------------------------
'
' Filename: modMain.bas
'
' DESCRIPTION
' -----------
' This module contains the Main function which will be
' executed when the program runs. It contains different
' functions which are executed from the main program
'
' Copyright (c) 2002 Accenture.  All rights reserved.
'-----------------------------------------------------
'
' REVISION HISTORY
' Revision:    0
' Author:      Gamboa, Terence S. <NAI0797>
' Date:        September 25, 2002
'
' Date         Author          Description
' 10/10/02     Jett Gamboa     Added HasJobArea function
' 10/14/02     Jett Gamboa     Modified LoadTab function
' 10/14/02     Jett Gamboa     Added HasIntroLines function
' 10/14/02     Jett Gamboa     Added HandleShortcutKeys subroutine
' 10/25/02     Joseph Manyaga  Added Cut menu item
'  4/18/03     Jett Gamboa     Move the event handlers from the form to this module
'  3/7/03      Jett Gamboa     Added handling for language
'  11/17/2008  Sebouh          hide/unhide Load Factor/Sorption
'-----------------------------------------------------

' the name of our temporary file
Const TEMPLATE_FILE As String = "Template.ft"

' variable to contain our help file
Public HELP_FILE As String

' used in error handler
Const MODULE_NAME = "modMain.bas"

' Define constants for the toolbar buttons
Const BUTTON_NEW As Integer = 1
Const BUTTON_OPEN As Integer = 2
Const BUTTON_SAVE As Integer = 3
'Const BUTTON_PRINT As Integer = 5
Const BUTTON_CUT As Integer = 5
Const BUTTON_COPY As Integer = 6
Const BUTTON_PASTE As Integer = 7

' Define constants for the tabs
Public Const TAB_CUSTOMER_INFO As Integer = 1
Public Const TAB_DOSAGE_PLAN As Integer = 2
Public Const TAB_INTRO_PLAN As Integer = 3
Public Const TAB_MONITOR_PLAN As Integer = 4
Public Const TAB_INTRO_HISTORY As Integer = 5
Public Const TAB_MONITOR_INPUT As Integer = 6
Public Const TAB_MONITOR_STATUS As Integer = 7
Public Const TAB_GRAPH As Integer = 8
'12-11-06 David Smith Added entry for Fumigation Management Plan tab
Public Const TAB_FUM_MGNT As Integer = 9


Public gbForcedClick As Boolean         ' used to track whether we revalidate after a tab click
Private nPreviousTab As Integer
Dim rsIntroLine As New ADODB.Recordset

'to hide Load Factor/Sorption
Public gbNewFile As Boolean


'**************************************************************
'   DESCRIPTION
'   This is the function which gets executed when the program runs
'   It calls other modules such as the CrypKey authentication module
'   and is responsible for calling the other functions to set up the
'   form on first use.
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'  4/01/2003       Jett Gamboa       removed Print option and added Export File and Graph
'  4/28/2003       Jett Gamboa       Set the application help file
'**************************************************************
Sub Main()

    Const FUNCTION_NAME = "Main()"

    Dim iCurrentLanguage As Integer

    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    ' set defaults
    gsCurrentFile = ""
    gsTempFile = ""
    gbHasChanged = False
    gbFileReadOnly = False
    gbCanChangeTabs = True
    gbCanExit = True
    gbIsExiting = False
    gbIsMetric = True
    gbIsTestingActive = False
    gnExposureTime = 0
    blnConvert = True
    gbNewFile = False
    
    'JNM 05/02/03 - PT fix - Initialize new global
    gbIsSaving = False
    
    ' Jett 12/23/2004 - Check if there was a call to remove the CrypKey license
    ' call the subroutine to remove the license if this is the case then
    ' terminiate immediately
    If LCase(Command) = "/removelicense" Then
        RemoveCrypKeyLicense
        End
    End If
        
    ' call check registry to make sure that all entries are there
    CheckRegistry

    ' JETT 2/11/2004 LOCALIZATION
    
    ' Check if the user forced us to use resources for a specific language
    If gForceLanguage <> eNEUTRAL Then
        ' We only force the thread language if this is an NT-based OS
        ' SetThreadLocale is only supported on NT-based OSes
        '
        ' This is not the ideal way to force a language but this functionality
        ' will only be available to DAS support personnel so it is not that risky
        If GetWinPlatform() = "NT" Then
            SetThreadLocale gForceLanguage
        End If
    
    End If
    
    ' call the GetThreadLocale API function to determine our language
    iCurrentLanguage = GetThreadLocale()
       
    ' Get only the value of lower byte
    iCurrentLanguage = iCurrentLanguage And &HFF&

    ' use the correct resources depending on the language
    Select Case iCurrentLanguage
       'Case eENGLISH
       Case &H9&              ' 0x9 - English
            HELP_FILE = "Fumiguide Help.chm"
            gTimeFormat = "hh:Nn AMPM"
            gFullTimeFormat = "dd-mmm-yy hh:mm AMPM"
            
       'Case eFrench
        Case &HC&              ' 0xC - French
            HELP_FILE = "Fumiguide Help FR.chm"
            frmMain.ChartFX1.Language "Cfx4Lang_FR.dll"
            
            frmMain.dtpDummyTime.CustomFormat = "HH:mm"
            frmMain.dtpStartIntroTime.CustomFormat = "HH:mm"
            frmMain.dtpStartIntroTime2.CustomFormat = "HH:mm"
            frmMain.dtpStartExposureTime.CustomFormat = "HH:mm"
            frmMain.dtpStartExposureTime2.CustomFormat = "HH:mm"
            frmMain.dtpEndExposureTime.CustomFormat = "HH:mm"
            frmMain.dtpEndExposureTime2.CustomFormat = "HH:mm"
            gTimeFormat = "HH:mm"
            gFullTimeFormat = "dd-mmm-yy HH:mm"
        
        'case eGERMAN
        Case &H7&              ' 0x7 - German
            HELP_FILE = "Fumiguide Help DE.chm"
            frmMain.ChartFX1.Language "Cfx4Lang_DE.dll"
            
            frmMain.dtpDummyTime.CustomFormat = "HH:mm"
            frmMain.dtpStartIntroTime.CustomFormat = "HH:mm"
            frmMain.dtpStartIntroTime2.CustomFormat = "HH:mm"
            frmMain.dtpStartExposureTime.CustomFormat = "HH:mm"
            frmMain.dtpStartExposureTime2.CustomFormat = "HH:mm"
            frmMain.dtpEndExposureTime.CustomFormat = "HH:mm"
            frmMain.dtpEndExposureTime2.CustomFormat = "HH:mm"
            gTimeFormat = "HH:mm"
            gFullTimeFormat = "dd-mmm-yy HH:mm"
        
        'Case eITALIAN
        Case &H10&             ' 0x10 - Italian
            HELP_FILE = "Fumiguide Help IT.chm"
            frmMain.ChartFX1.Language "Cfx4Lang_IT.dll"
            
            frmMain.dtpDummyTime.CustomFormat = "HH:mm"
            frmMain.dtpStartIntroTime.CustomFormat = "HH:mm"
            frmMain.dtpStartIntroTime2.CustomFormat = "HH:mm"
            frmMain.dtpStartExposureTime.CustomFormat = "HH:mm"
            frmMain.dtpStartExposureTime2.CustomFormat = "HH:mm"
            frmMain.dtpEndExposureTime.CustomFormat = "HH:mm"
            frmMain.dtpEndExposureTime2.CustomFormat = "HH:mm"
            gTimeFormat = "HH:mm"
            gFullTimeFormat = "dd-mmm-yy HH:mm"
        
        '12/11/06 David Smith Added code for Spanish
        'Case eSPANISH
        Case &HA              ' 0x10 - Spanish
            HELP_FILE = "Fumiguide Help.chm"
            gTimeFormat = "hh:Nn AMPM"
            gFullTimeFormat = "dd-mmm-yy hh:mm AMPM"
            
            frmMain.dtpDummyTime.CustomFormat = "HH:mm"
            frmMain.dtpStartIntroTime.CustomFormat = "HH:mm"
            frmMain.dtpStartIntroTime2.CustomFormat = "HH:mm"
            frmMain.dtpStartExposureTime.CustomFormat = "HH:mm"
            frmMain.dtpStartExposureTime2.CustomFormat = "HH:mm"
            frmMain.dtpEndExposureTime.CustomFormat = "HH:mm"
            frmMain.dtpEndExposureTime2.CustomFormat = "HH:mm"
            gTimeFormat = "HH:mm"
            gFullTimeFormat = "dd-mmm-yy HH:mm"
            
        Case Else
            HELP_FILE = "Fumiguide Help.chm"
            gTimeFormat = "hh:Nn AMPM"
            gFullTimeFormat = "dd-mmm-yy HH:mm"
    End Select
    


    ' do not execute authentication if ALWAYSAUTH is true
    #If ALWAYSAUTH <> 1 Then
        If CrypkeyAuthentication() = False Then
            Exit Sub
        End If
    #End If

    ' load the resource strings
    LoadStrFromResource
    
    ' Set the application help file
    App.HelpFile = App.Path & "\" & HELP_FILE
       
    With frmMain
    
        frmMain.Height = 11115
        frmMain.Move 0, 0
    
        ' Do not display the tab control yet
        frmMain.tabTabs.Visible = False
    
        ' disable Save, SaveAs, Print, Close, Cut, Copy, Paste menu items
        .mnuFileSave.Enabled = False
        .mnuFileSaveAs.Enabled = False
        '.mnuFilePrint.Enabled = False
        .mnuFileExport.Enabled = False
        .mnuFileExportGraph.Enabled = False
        .mnuFileClose.Enabled = False
        .mnuEditCut.Enabled = False     'JNM 10/25/02 - Added Cut item
        .mnuEditCopy.Enabled = False
        .mnuEditPaste.Enabled = False
        .mnuFileOpen = True
        .mnuFileExit = True
        
       'Disable Save, Print, copy paste buttons in the toolbar
        .tbToolBar.Buttons(BUTTON_SAVE).Enabled = False
        '.tbToolBar.Buttons(BUTTON_PRINT).Enabled = False
        .tbToolBar.Buttons(BUTTON_CUT).Enabled = False
        .tbToolBar.Buttons(BUTTON_COPY).Enabled = False
        .tbToolBar.Buttons(BUTTON_PASTE).Enabled = False
        
        .sbStatusBar.Panels(1).Enabled = True
        .sbStatusBar.Panels(1).Text = ""
        
    End With
    
    ' get the command-line argument and assign it as file to open
    If Command <> "" Then
              
        gsCurrentFile = Command
        
        OpenFile True
        
        frmMain.Caption = LoadResString(IDS_FUMIGUIDE) & " - " & gsCurrentFile
        
    End If
          
    ' display the main form
    frmMain.Show
             
    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
             
End Sub


'**************************************************************
'   DESCRIPTION
'   Loads all the resource strings needed for the main form.
'   It then sets the captions for all controls to the specified
'   values
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'  4/01/2003       Jett Gamboa       removed Print option and added Export File and Graph
'**************************************************************
Sub LoadStrFromResource()

    Const FUNCTION_NAME = "LoadStrFromResource()"

    Dim sErrorDesc As String
    On Error GoTo Error_Handler
    

    With frmMain
        ' load the application title from the resource file
        .Caption = LoadResString(IDS_FUMIGUIDE)
        
        ' Load the menu captions from the resource file
        .mnuFile.Caption = LoadResString(IDS_MNU_FILE)
        .mnuFileNew.Caption = LoadResString(IDS_MNU_NEW)
        .mnuFileOpen.Caption = LoadResString(IDS_MNU_OPEN)
        .mnuFileClose.Caption = LoadResString(IDS_MNU_CLOSE)
        .mnuFileSave.Caption = LoadResString(IDS_MNU_SAVE)
        .mnuFileSaveAs.Caption = LoadResString(IDS_MNU_SAVEAS)
        '.mnuFilePrint.Caption = LoadResString(IDS_MNU_PRINT)
        .mnuFileExport.Caption = LoadResString(IDS_MNU_EXPORT)
        .mnuFileExportGraph.Caption = LoadResString(IDS_MNU_EXPORTGRAPH)
        .mnuFileExit.Caption = LoadResString(IDS_MNU_EXIT)
        .mnuEdit.Caption = LoadResString(IDS_MNU_EDIT)
        .mnuEditCut.Caption = LoadResString(IDS_MNU_CUT)    'JNM 10/25/02 - Added Cut
        .mnuEditCopy.Caption = LoadResString(IDS_MNU_COPY)
        .mnuEditPaste.Caption = LoadResString(IDS_MNU_PASTE)
        .mnuView.Caption = LoadResString(IDS_MNU_VIEW)
        .mnuViewToolbar.Caption = LoadResString(IDS_MNU_TOOLBAR)
        .mnuViewStatusBar.Caption = LoadResString(IDS_MNU_STATUSBAR)
        .mnuViewOptions.Caption = LoadResString(IDS_MNU_OPTIONS)
        .mnuHelp.Caption = LoadResString(IDS_MNU_HELP)
        .mnuHelpContents.Caption = LoadResString(IDS_MNU_CONTENTS)
        .mnuHelpAbout.Caption = LoadResString(IDS_MNU_ABOUT)
        'CG 04/05/04 Added menu strings (to use resource file instead of the hardcoded menu values)
        .mnuMonitorInputCopy.Caption = LoadResString(IDS_TOOLTIP_COPY)
        .mnuMonitorInputDelete.Caption = LoadResString(IDS_DELETE)
        .mnuMonitorInputPaste.Caption = LoadResString(IDS_TOOLTIP_PASTE)
        .mnuDosagePlanDelete.Caption = LoadResString(IDS_DELETE)
        .mnuIntroHistoryDelete.Caption = LoadResString(IDS_DELETE)
        .mnuIntroPlanDelete.Caption = LoadResString(IDS_DELETE)
        .mnuMonitorPlanDelete.Caption = LoadResString(IDS_DELETE)
                 
        ' Load the tab captions from the resource file
        .tabTabs.Tabs(1).Caption = LoadResString(IDS_TAB_CUSTINFO)
        .tabTabs.Tabs(2).Caption = LoadResString(IDS_TAB_DOSAGE_PLAN)
        .tabTabs.Tabs(3).Caption = LoadResString(IDS_TAB_INTRO_PLAN)
        .tabTabs.Tabs(4).Caption = LoadResString(IDS_TAB_MONITORING_PLAN)
        .tabTabs.Tabs(5).Caption = LoadResString(IDS_TAB_INTRO_HIST)
        .tabTabs.Tabs(6).Caption = LoadResString(IDS_TAB_MDATA_INPUT)
        .tabTabs.Tabs(7).Caption = LoadResString(IDS_TAB_MONITOR_STATUS)
        .tabTabs.Tabs(8).Caption = LoadResString(IDS_TAB_GRAPH)
        
        .lblAboutProfume.Font.Bold = True
        .lblAboutProfume.Caption = LoadResString(IDS_HEADER)
        
        'Pat 3/4/04 - Load the buttons tooltip from the resource file
        
        .tbToolBar.Buttons("New").ToolTipText = LoadResString(IDS_TOOLTIP_NEW)
        .tbToolBar.Buttons("Open").ToolTipText = LoadResString(IDS_TOOLTIP_OPEN)
        .tbToolBar.Buttons("Save").ToolTipText = LoadResString(IDS_TOOLTIP_SAVE)
        .tbToolBar.Buttons("Cut").ToolTipText = LoadResString(IDS_TOOLTIP_CUT)
        .tbToolBar.Buttons("Copy").ToolTipText = LoadResString(IDS_TOOLTIP_COPY)
        .tbToolBar.Buttons("Paste").ToolTipText = LoadResString(IDS_TOOLTIP_PASTE)
        
                
        'JNM 10/25/2002 - Load the tooltip text
        .tabTabs.Tabs(1).ToolTipText = LoadResString(IDS_TIP_TAB_CUSTINFO)
        .tabTabs.Tabs(2).ToolTipText = LoadResString(IDS_TIP_TAB_DOSAGE_PLAN)
        .tabTabs.Tabs(3).ToolTipText = LoadResString(IDS_TIP_TAB_INTRO_PLAN)
        .tabTabs.Tabs(4).ToolTipText = LoadResString(IDS_TIP_TAB_MONITORING_PLAN)
        .tabTabs.Tabs(5).ToolTipText = LoadResString(IDS_TIP_TAB_INTRO_HIST)
        .tabTabs.Tabs(6).ToolTipText = LoadResString(IDS_TIP_TAB_MDATA_INPUT)
        .tabTabs.Tabs(7).ToolTipText = LoadResString(IDS_TIP_TAB_MONITOR_STATUS)
        
        'JNM 11/15 - Set text in textbox
        .lblGetHeading = LoadResString(IDS_PROFUME_INSTRUCT_HEAD)
        '.lblGetStarted.Caption = Space(50) & LoadResString(IDS_PROFUME_INSTRUCT_HEAD) & _
            vbCrLf & vbCrLf & LoadResString(IDS_PROFUME_INSTRUCT)
        'Pat 3/3/2004: remove one carriage return
        .lblGetStarted.Caption = vbCrLf & LoadResString(IDS_PROFUME_INSTRUCT)
        .lblTradeMark.Caption = LoadResString(IDS_TRADEMARK)
        
        ' CG 08/26 - Ticket # 2099731 Added the copyright text
        .lblCopyright.Caption = LoadResString(IDS_MAIN_COPYRIGHT)
                                             
    End With
    
    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'**************************************************************
'   DESCRIPTION
'   Loads the Job id for this single job in to the global variable
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'**************************************************************
Sub GetJobID()

    Const FUNCTION_NAME = "GetJobID()"

    Dim sErrorDesc As String
    Dim rsSQLJob As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo Error_Handler
    
    'Sets the SQL string
    sSQL = "SELECT id FROM Job"
    
    'Opens the recordset with
    rsSQLJob.Open sSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText



    ' if there were no records then this must be an invalid template
    If rsSQLJob.EOF Then
        sErrorDesc = LoadResString(IDS_ERR_INVALIDTEMPLATE)
        Set rsSQLJob = Nothing
        psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
        
    End If
    
    gsJobID = CStr(rsSQLJob("id").Value)
    
    rsSQLJob.Close
    Set rsSQLJob = Nothing

    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub

'**************************************************************
'   DESCRIPTION
'   Opens the access database file specified in the gsTempFile
'   global variable. It sets the global connection object to the
'   open ADO connection object
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
' 12/23/2004       Jett Gamboa       added call to UpgradeFile
'**************************************************************
Public Sub OpenDatabase()

    Const FUNCTION_NAME = "OpenDatabase()"

    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    ' instantiate a new ADO connection object
    Set gdbConn = New ADODB.Connection
    
    ' Sets the properties and opens the database connection.
    With gdbConn
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .Properties("Data Source") = gsTempFile
            .Open
    End With
   
   'MsgBox ("Here is the database  " & gsTempFile)
                
    ' Validates the template sets the global id
    GetJobID
    
    ' Call upgrade data file to see if any data file upgrades are needed
    UpgradeDataFile
               
    ' set the fileopen flag to true
    gbIsFileOpen = True
    
    ' reset the HasChanged flag
    gbHasChanged = False

    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'   Checks if the data file that was open is the latest version
'   of the data file. If not, then the data file is upgraded by
'   adding the necessary fields and tables to the database.
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  12/23/2004      Jett Gamboa       Initial Revision
'  02/10/2006      Pat San Miguel    Added new column in Job Table
'  04/17/2007      David Smith       Changed code to check each column and alter as needed not just
'                                    the version number as it was change but columns were missing.
'**************************************************************
Sub UpgradeDataFile()

    Const FUNCTION_NAME = "UpgradeDataFile()"

    Dim sErrorDesc As String
    Dim rsSQLSettings As New ADODB.Recordset
    Dim sSQL As String
    
    Dim sDataFileVersion As String
    
    On Error Resume Next
        
    'Checks to see if file being opened has a version number present
    sSQL = "SELECT Property_Value FROM Settings Where Property = ""Datafile_Version"""
    rsSQLSettings.Open sSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText
    ' if there were no records then this must be an invalid template
    If rsSQLSettings.EOF Then
       gdbConn.Execute "INSERT INTO Settings (Property, Property_Value) VALUES (""Datafile_Version"", ""2"")"
     Else
        ' retrieve the version number
        sDataFileVersion = CStr(rsSQLSettings("Property_Value").Value)
    End If
   rsSQLSettings.Close
   
   
    'Checks to see if file being opened has a user_defined_ct column in the  JobArea table
         sSQL = "SELECT user_defined_ct FROM JobArea "
         rsSQLSettings.Open sSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText
         ' if there were no records then this must be an invalid template
         If rsSQLSettings.EOF Then
             gdbConn.Execute "ALTER TABLE JobArea ADD COLUMN user_defined_ct Number"
         End If
         If rsSQLSettings.State = 1 Then
            rsSQLSettings.Close
         End If
       
    
         
        'Checks to see if file being opened has a one_fourth_length column in the Job ShootingLine
         sSQL = "SELECT one_fourth_length FROM ShootingLine "
         rsSQLSettings.Open sSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText
         ' if there were no records then this must be an invalid template
         ' Error number -2147217904 indicates that there was no column in that table of that name
         If Err.Number = -2147217904 Then
          'Pat 02/10/2006: Added new table for the value of the new field in Dosage Plan tab
              gdbConn.Execute "ALTER TABLE ShootingLine ADD COLUMN one_fourth_length Number"
         End If
         If rsSQLSettings.State = 1 Then
            rsSQLSettings.Close
        End If
        
        'Checks to see if file being opened has a one_eight_length column in the Job ShootingLine
         sSQL = "SELECT one_eight_length FROM ShootingLine "
         rsSQLSettings.Open sSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText
         ' if there were no records then this must be an invalid template
         ' Error number -2147217904 indicates that there was no column in that table of that name
         If Err.Number = -2147217904 Then
           'Pat 02/10/2006: Added new table for the value of the new field in Dosage Plan tab
           gdbConn.Execute "ALTER TABLE ShootingLine ADD COLUMN one_eight_length Number"
         End If
         If rsSQLSettings.State = 1 Then
            rsSQLSettings.Close
         End If
        
        
        
        
        'Checks to see if file being opened has a one_eight_length column in the Job ShootingLine
         sSQL = "SELECT one_eight_length FROM ShootingLine "
         rsSQLSettings.Open sSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText
         ' if there were no records then this must be an invalid template
         ' Error number -2147217904 indicates that there was no column in that table of that name
         If Err.Number = -2147217904 Then
           'Pat 02/10/2006: Added new table for the value of the new field in Dosage Plan tab
            gdbConn.Execute "ALTER TABLE ShootingLine ADD COLUMN numcylinder Number"
         End If
         If rsSQLSettings.State = 1 Then
            rsSQLSettings.Close
        End If
        
        
        'Checks to see if file being opened has a time_min_empty_cylinder column in the Job ShootingLine
         sSQL = "SELECT time_min_empty_cylinder FROM ShootingLine "
         rsSQLSettings.Open sSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText
         ' if there were no records then this must be an invalid template
         ' Error number -2147217904 indicates that there was no column in that table of that name
         If Err.Number = -2147217904 Then
           'Pat 02/10/2006: Added new table for the value of the new field in Dosage Plan tab
            gdbConn.Execute "ALTER TABLE ShootingLine ADD COLUMN time_min_empty_cylinder Number"
         End If
         If rsSQLSettings.State = 1 Then
            rsSQLSettings.Close
         End If
        
        
        
        
        
        'Checks to see if file being opened has a Pressure0_or_temperature1 column in the Job ShootingLine
         sSQL = "SELECT Pressure0_or_temperature1 FROM ShootingLine "
         rsSQLSettings.Open sSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText
         ' if there were no records then this must be an invalid template
         ' Error number -2147217904 indicates that there was no column in that table of that name
         If Err.Number = -2147217904 Then
           'Pat 02/10/2006: Added new table for the value of the new field in Dosage Plan tab
            gdbConn.Execute "ALTER TABLE ShootingLine ADD COLUMN Pressure0_or_temperature1 Number"
        End If
       If rsSQLSettings.State = 1 Then
        rsSQLSettings.Close
     End If
        
        
        
        
                'Checks to see if file being opened has a Pressure0_or_temperature1 column in the Job ShootingLine
         sSQL = "SELECT numcylinder FROM ShootingLine "
         rsSQLSettings.Open sSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText
         ' if there were no records then this must be an invalid template
         ' Error number -2147217904 indicates that there was no column in that table of that name
         If Err.Number = -2147217904 Then
           'Pat 02/10/2006: Added new table for the value of the new field in Dosage Plan tab
            gdbConn.Execute "ALTER TABLE ShootingLine ADD COLUMN numcylinder Number"
        End If
       If rsSQLSettings.State = 1 Then
        rsSQLSettings.Close
     End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
       'Checks to see if file being opened has a NoAdjSorp column in the Job table
         sSQL = "SELECT NoAdjSorp FROM Job "
         rsSQLSettings.Open sSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText
         ' if there were no records then this must be an invalid template
         ' Error number -2147217904 indicates that there was no column in that table of that name
         If Err.Number = -2147217904 Then
          'Pat 02/10/2006: Added new table for the value of the new field in Dosage Plan tab
             gdbConn.Execute "ALTER TABLE Job ADD COLUMN NoAdjSorp Number"
             Err.Clear
        End If
    
    'Final close on is needed after each open befor the next open
     If rsSQLSettings.State = 1 Then
        rsSQLSettings.Close
     End If
     
     Set rsSQLSettings = Nothing
    
    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub



'**************************************************************
'   DESCRIPTION
'   Closes the currently open database
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Sub CloseDatabase()

    Const FUNCTION_NAME = "CloseDatabase()"

    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    gdbConn.Close
    
    Set gdbConn = Nothing
    
    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc


End Sub


'**************************************************************
'   DESCRIPTION
'   This is the function which gets executed when the program runs
'   It calls other modules such as the CrypKey authentication module
'   and is responsible for calling the other functions to set up the
'   form on first use.
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Sub NewFile()

    Const FUNCTION_NAME = "NewFile()"

    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    ' If a file is currently open and we're trying to create a new one ...
    If gbIsFileOpen = True Then
    
        ' close the file that is currently opened
        CloseFile
        
        ' If the file is still open, it means that the user cancelled on a save
        ' dialog so we do not proceed with the creation of a new file
        If gbIsFileOpen = True Then
            Exit Sub
        End If
        
    End If
    
    ' temporarily set current file to the template file
    gsCurrentFile = App.Path & "\" & TEMPLATE_FILE
    ' call the Open file to "open" the temporary file
    OpenFile True
    
    ' 5/8/2003 Jett - disable Save As if this is a new file
    frmMain.mnuFileSaveAs.Enabled = False
    
    ' set the current file to blank so we are prompted to save the file
    gsCurrentFile = ""
    
    gbNewFile = True
       
    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'   Handles file opening routines. Is called when the user clicks on the
'   open file toolbar button or the file-open menu. Before opening a new
'   file it closes any open files.
'   When a file is opened we are not really opening the file, rather we
'   a creating a copy of it in the TEMP folder. All work is done on the
'   temporay file and not actual file. The opened file is only updated
'   when the user does a save
'
'   PARAMETERS
'       pvbFromTemplate - specifies whether the file being opened is the template
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'  4/01/2003       Jett Gamboa       removed Print option and added Export File and Graph
'**************************************************************
Public Sub OpenFile(Optional pvbFromTemplate As Boolean)

    Const FUNCTION_NAME = "OpenFile()"

    Dim sNewFile As String
    Dim sErrorDesc As String
    On Error GoTo Error_Handler


    ' prompt the user for a filename if this was not called from the NewFile procedure
    If Not pvbFromTemplate Then
    
        With frmMain.dlgCommonDialog
        
            ' set the properties of the common dialog control
            .DialogTitle = "Open"
            .CancelError = True
            .Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
            .Filter = LoadResString(IDS_FILE_FILTER)
            
            ' turn off error handling so we can catch if the user selected
            ' cancel from the dialog
            On Error Resume Next
            
            ' display the open dialog control
            .ShowOpen
            
            ' If the user did not select a file exit the subroutine
            If Err.Number = cdlCancel Or Len(.FileName) = 0 Then
                Exit Sub
            End If
            
            On Error GoTo Error_Handler
            
            ' TODO gsCurrentFile = .FileName
            sNewFile = .FileName
            
        End With
        
        ' if a file is currently open while we're trying to open a new one close it
        ' first
        If gbIsFileOpen Then
        
            CloseFile
                   
            ' if the file was not closed then it means that the user cancelled so we
            ' do not proceed with the file open
            If gbIsFileOpen Then
                Exit Sub
            End If
        End If
        
        gsCurrentFile = sNewFile
        
    End If
        
    gsTempFile = GetTemporaryFilename("~PRF")
        
    If Not CopyAFile(gsCurrentFile, gsTempFile) Then
        
        MsgBox LoadResString(IDS_ERR_OPENFILE), vbCritical, LoadResString(IDS_ERROR)
        
        ' Delete the temporary file
        Kill gsTempFile
        
        gsCurrentFile = ""
        
        ' exit from the subroutine
        Exit Sub
    End If
    
    ' call the OpenDatabase subroutine to open the file
    OpenDatabase
    
    ' 5/16/2003 Jett - force a recalculation when they open a file
    CalculateMonDataInput
        
    ' display the tab control
    frmMain.tabTabs.Visible = True
        
    'gbForcedClick = True
    
    ' "click" the first tab so that it is activated
    frmMain.tabTabs.Tabs(1).Selected = True
    
    ' load first frame
    'LoadTab frmMain.fraCustomerInfo
    
    'LoadTab frmMain.fraCustomerInfo
    'gbForcedClick = False
       
    ' retrieve the exposure time
    gnExposureTime = CheckExposureTimes()
    
    With frmMain
        ' enable Save, SaveAs, Print, Close, Cut, Copy, Paste menu items
        .mnuFileSave.Enabled = True
        .mnuFileSaveAs.Enabled = True
        '.mnuFilePrint.Enabled = True
        .mnuFileExport.Enabled = True
        .mnuFileExportGraph.Enabled = True
        .mnuFileClose.Enabled = True
        .mnuEditCut.Enabled = True      'JNM 10/25/02 - Added Cut item
        .mnuEditCopy.Enabled = True
        .mnuEditPaste.Enabled = True
        
       ' enable Save and Print buttons in the toolbar
        .tbToolBar.Buttons(BUTTON_SAVE).Enabled = True
        '.tbToolBar.Buttons(BUTTON_PRINT).Enabled = True
        .tbToolBar.Buttons(BUTTON_CUT).Enabled = True
        .tbToolBar.Buttons(BUTTON_COPY).Enabled = True
        .tbToolBar.Buttons(BUTTON_PASTE).Enabled = True
        
        If Not pvbFromTemplate Then
            .Caption = LoadResString(IDS_FUMIGUIDE) & " - " & gsCurrentFile
        Else
            .Caption = LoadResString(IDS_FUMIGUIDE) & " - New Document"
        End If

    End With
       
    ' we just opened the file so reset isExiting flag to false
    gbIsExiting = False
          
    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
                    
End Sub


'**************************************************************
'   DESCRIPTION
'   Deletes job-specific data when a file is Saved As  to a
'   different filename
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
' 10/21/2002       Jett Gamboa       Added additional code to "remove" specific records in some tables
'**************************************************************
Public Sub SaveAsProcessing()

    Const FUNCTION_NAME = "SaveAsProcessing()"

    Dim sErrorDesc As String
    Dim sSQL As String
    
    On Error GoTo Error_Handler

    gdbConn.BeginTrans
    gdbConn.Execute "DELETE FROM LogHistory"
    gdbConn.Execute "DELETE FROM MonitorPointValue"
    gdbConn.Execute "DELETE FROM IntroHistory"
        
    ' clean up the job summary
    sSQL = "UPDATE JobArea SET hlt = -1, achieved_concentration_time = 0, achieved_dose = 0, actual_hlt = -1, " & _
           "projected_concentration_time = -1,projected_time = -1, mean_ntc = -1, status = 0, add_profume = -1"
    
    gdbConn.Execute sSQL
    
    ' set to default the monitor point values specific to this job
    sSQL = "UPDATE MonitorPoint SET hlt = -1, last_entry = NULL, elapsed_time = -1, start_time = NULL, " & _
           "end_time = NULL, achieved_concentration_time = -1, achieved_dose = -1, projected_concentration_time = -1, " & _
           "projected_time = -1, future_target_concentration = -1, future_concentration = -1"
    
    gdbConn.Execute sSQL
    
    gdbConn.CommitTrans

    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'   This function takes care of doing a Save or Save As. It prompts
'   the user for a filename and copies the working temporary file
'   to the filename specified by the user.
'
'   PARAMETERS
'       pvbDoSaveAs - specifies whther this is an ordinary Save or
'                     a Save As. This parameter is optional
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Sub SaveFile(Optional pvbDoSaveAs As Boolean)
    
    Const FUNCTION_NAME = "SaveFile()"

    Dim sTempFilename As String

    Dim sErrorDesc As String
    On Error GoTo Error_Handler
    
    'JNM 05/02/03 - PT Fix - Save the last control. It's ok if the value is invalid since
    '   there is a requirement in the first release that the user should still
    '   be able to exit, except that the invalid value is not saved.
    gbIsSaving = True
    frmMain.ValidateControls
    gbIsSaving = False
        
    'JNM 05/02/03 - We'll have to force the temp file to be updated immediately
    gdbConn.BeginTrans
    gdbConn.CommitTrans
        
    ' if this is the first time we are saving this file or we are doing a
    ' save as, we need to prompt the user for a filename
    If gsCurrentFile = "" Or pvbDoSaveAs Then
                
        With frmMain.dlgCommonDialog
            
            ' set the properties of the common dialog control
            
            ' Put the correct dialog title depending on the save "mode"
            If pvbDoSaveAs Then
                .DialogTitle = "Save As"
            Else
                .DialogTitle = "Save"
            End If
            
            .CancelError = True
            .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
            .Filter = LoadResString(IDS_FILE_FILTER)
            
            ' turn off error handling so we can check if the user selected
            ' the cancel button
            On Error Resume Next
            
            ' display the save dialog control
            .ShowSave
            
            ' If the user did not select a file exit the subroutine
            If Err.Number = cdlCancel Or Len(.FileName) = 0 Then
                Exit Sub
            End If
            
            On Error GoTo Error_Handler
                
            ' set the filename to the name specified by the user
            gsCurrentFile = .FileName
                
        End With
        
    End If
    
    ' if doing a save as, ask the user if they want to remove the job-specific
    ' information before saving
    If pvbDoSaveAs Then
        If MsgBox(LoadResString(IDS_SAVEAS_PROMPT), vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            
            SaveAsProcessing
            
        End If
    End If

    ' copy the temporary file to the filename specified (if doing a save as)
    ' or the current file (if doing a save)
    If Not CopyAFile(gsTempFile, gsCurrentFile) Then
    
        MsgBox LoadResString(IDS_ERR_SAVEFILE), vbCritical, LoadResString(IDS_ERROR)
                
     Else
        
        ' we saved the file succesfully, set changed flag to false
        gbHasChanged = False
        
        ' 5/8/2003 Jett - enable save as after we save (actually, we only need this when
        ' saving a new file, but it does not really matter if we enable what is already
        ' enabled)
        frmMain.mnuFileSaveAs.Enabled = True
        
    End If
       
    ' If this is a save us, we need to save and reopen the new filename
    If pvbDoSaveAs Then
    
        ' save the new file name before closing the old file
        sTempFilename = gsCurrentFile
    
        ' close the current file
        CloseFile
        
        gsCurrentFile = sTempFilename
        
        ' reopen the file
        OpenFile True
    
    End If
    
    frmMain.Caption = LoadResString(IDS_FUMIGUIDE) & " - " & gsCurrentFile

    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'   Called when the user closes an open file or when the user exits
'   the application. This function checks if anychanges were made
'   since the file was last saved and the user is prompted accordingly
'
'   PARAMETERS
'       pvbExiting - specifies whether we called this when main form exits
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
' 10/16/2002       Jett Gamboa       added pvbExiting and additional code
'                                    to handle instances where current file
'                                    is being closed while a new one is opened
'  4/01/2003       Jett Gamboa       removed Print option and added Export File and Graph
'**************************************************************

Sub CloseFile(Optional pvbExiting As Boolean)

    Const FUNCTION_NAME = "CloseFile()"

    Dim sErrorDesc As String
    Dim sTempFile As String
    
    On Error GoTo Error_Handler

    sTempFile = ""

    ' if the file has changed since it was last saved, prompt the user if he
    ' would like to save before closing
    If gbHasChanged Then
    
        Select Case MsgBox(LoadResString(IDS_SAVECHANGE_PROMPT), vbYesNoCancel + vbExclamation)
        
            ' user selected Cancel, do not proceed with closing the file
            Case vbCancel
            
                'TODO: check this gActiveFrame.SetFocus
            
                ' prevent the application from exiting
                gbCanExit = False
                Exit Sub
                
            ' user selected Yes, save the file
            Case vbYes
            
                ' this handles the special case where the user closes the main program
                ' after changes were made. We need to FORCE unload the form first so it
                ' saves all records before we save and close the database
                
                
                If gsCurrentFile = "" Then
                    With frmMain.dlgCommonDialog
                        
                        ' set the properties of the common dialog control
                        
                        .DialogTitle = "Save As"
                        
                        .CancelError = True
                        .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
                        .Filter = LoadResString(IDS_FILE_FILTER)
                        
                        ' turn off error handling so we can check if the user selected
                        ' the cancel button
                        On Error Resume Next
                        
                        ' display the save dialog control
                        .ShowSave
                        
                        ' If the user did not select a file exit the subroutine
                        ' set gbCanExit to False so that it will not terminate
                        If Err.Number = cdlCancel Or Len(.FileName) = 0 Then
                            gbCanExit = False
                            Exit Sub
                        End If
                        
                        On Error GoTo Error_Handler
                            
                        ' keep the current filename just in case the user was opening another
                        ' file
                        sTempFile = gsCurrentFile
                            
                        ' set the filename to the name specified by the user
                        gsCurrentFile = .FileName
                            
                    End With
                End If
                
                ' if this was called from the Unload handler in the main form,
                ' force unload the current form so we can save the changes
                'TODO: check this
                If pvbExiting Then
                    If Not gfrmActiveForm Is Nothing Then

                        ' set gbIsExiting to true, forces all subforms to yield
                        'gbIsExiting = True
                        gbIsExiting = False
                        Unload gfrmActiveForm

                    End If
                End If

                ' Save the file
                SaveFile
                
                ' if there was no file name it means that the user cancelled saving a new file
                If gsCurrentFile = "" Then
                    Exit Sub
                End If
                
                ' if the tempfile was not empty it means the user was in the process of opening
                ' another file after he made changes to the current file
                If sTempFile <> "" Then
                    gsCurrentFile = sTempFile
                End If
                   
        End Select
    End If
      
    With frmMain
    
        ' Set Is Exiting flag to true, the child forms have no choice but to
        ' unload
        'gbIsExiting = True
        gbIsExiting = False
    
        ' call unload tab to unload the current tab
        ' TODO: Check this out Unload gfrmActiveForm
        UnloadTab gActiveFrame
        LoadTab .fraSplash
    
        ' hide the tab
        .tabTabs.Visible = False
    
        ' disable Save, SaveAs, Print, Close, Cut, Copy, Paste menu items
        .mnuFileSave.Enabled = False
        .mnuFileSaveAs.Enabled = False
        '.mnuFilePrint.Enabled = False
        .mnuFileExport.Enabled = False
        .mnuFileExportGraph.Enabled = False
        .mnuFileClose.Enabled = False
        .mnuEditCut.Enabled = False     'JNM 10/25/02 - Added Cut item
        .mnuEditCopy.Enabled = False
        .mnuEditPaste.Enabled = False
        
       ' disable Save and Print buttons in the toolbar
        .tbToolBar.Buttons(BUTTON_SAVE).Enabled = False
        '.tbToolBar.Buttons(BUTTON_PRINT).Enabled = False
        .tbToolBar.Buttons(BUTTON_CUT).Enabled = False
        .tbToolBar.Buttons(BUTTON_COPY).Enabled = False
        .tbToolBar.Buttons(BUTTON_PASTE).Enabled = False
        
        .Caption = LoadResString(IDS_FUMIGUIDE)
    
    End With
    
    ' clean up
    CloseDatabase
    Kill gsTempFile
    
    gsCurrentFile = ""
    gbIsFileOpen = False
    
    Set gfrmActiveForm = Nothing
    
    gbNewFile = False

    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'   Loads the form specified in the paramter inside the fraFrame
'   control using the SetParent API function
'
'   PARAMETERS
'       prfrmChild - the form to load inside the fraFrame control
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
' 10/14/2002       Jett Gamboa       removed certain tab checking, we are
'                                    checking it during the tab click event
'  4/17/2003       Jett Gamboa       Used frame instead of form
'**************************************************************
Public Sub LoadTab(prfrmChild As Frame)

    Const FUNCTION_NAME = "LoadTab()"

    Dim sErrorDesc As String
    On Error GoTo Error_Handler
    
'    If UnloadTab(gActiveFrame) Then
    
        If gActiveFrame.Name = "fraSplash" Then
            UnloadTab gActiveFrame
        End If
        
    
        ' set the active frame to our new frame
        Set gActiveFrame = prfrmChild

        ' make the tab control visible
        'frmMain.tabTabs.Visible = True
                      
        ' Call the forms load event
        Select Case prfrmChild.Name
            Case "fraSplash"
                frmMain.fraSplash.Visible = True
            Case "fraCustomerInfo"
                Load_CustomerInfo
            Case "fraDosagePlan"
                Load_DosagePlan
            Case "fraIntroductionPlan"
                Load_IntroductionPlan
            Case "fraMonitorPlan"
                Load_MonitorPlan
            Case "fraIntroHistory"
                Load_IntroHistory
            Case "fraMonitorDataInput"
                Load_MonitorInput
            Case "fraMonitorStatus"
                Load_MonitorStatus
            Case "fraGraph"
                Load_Graph
        End Select
        
        'JNM 04/21/03 - AT Fix - Transferred this line from before Select in
        '       order to prevent the flickers when changing tabs.
        ' show the frame
        gActiveFrame.Visible = True
        
        'JNM 04/22/03 - AT Fix - Display message box if we are in data input
        If prfrmChild.Name = "fraMonitorDataInput" Then
            Form_Activate_Handler
        End If
        
'    End If

    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description & " - " & Erl
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


' this will replace the form unload code
Public Function UnloadTab(prfrmChild As Frame) As Boolean
    
   Select Case prfrmChild.Name
        Case "fraSplash"
            frmMain.fraSplash.Visible = False
        Case "fraCustomerInfo"
            Unload_CustomerInfo
        Case "fraDosagePlan"
            Unload_DosagePlan
        Case "fraIntroductionPlan"
            Unload_IntroductionPlan
        Case "fraMonitorPlan"
            Unload_MonitorPlan
        Case "fraIntroHistory"
            Unload_IntroHistory
        Case "fraMonitorDataInput"
            Unload_MonitorInput
        Case "fraMonitorStatus"
            Unload_MonitorStatus
        Case "fraGraph"
            Unload_Graph
    End Select
    
    ' if we can change tabs, then hide the current frame
    If gbCanChangeTabs Then
        prfrmChild.Visible = False
    End If
    
    UnloadTab = True

End Function



'**************************************************************
'   DESCRIPTION
'   Called by the individual tabs when the user clicks on the next
'   tab button. This does not actually load the tab but rather it
'   it triggers a click event for the next tab. The On_Click and
'   Before_Click event handler for the tab controls handles the
'   rest.
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Sub NextTab()

    Const FUNCTION_NAME = "NextTab()"

    Dim sErrorDesc As String

    Dim nNextTabIndex As Integer
    
    ' set the next tab index to the next tab
    nNextTabIndex = frmMain.tabTabs.SelectedItem.Index
    nNextTabIndex = nNextTabIndex + 1

    ' trigger the next tab to be selected
    frmMain.tabTabs.Tabs(nNextTabIndex).Selected = True
   
   Exit Sub
   
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
   
End Sub


'**************************************************************
'   DESCRIPTION
'   Checks the HKEY_CURRENT_USER tree to see if the registry entries
'   required by the ProFume application is present. They are created
'   if they do not exist.
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'  12/11/2006      David Smith       Added code for spanish version
'**************************************************************
Public Sub CheckRegistry()

    Const FUNCTION_NAME = "CheckRegistry()"
    
    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    Dim vKeyValue As Variant

    ' check all of the registry keys we need under HKEY_CURRENT_USER exists
    ' all registry locations are hardcoded since changing this will also affect
    ' the Installshield project.
    If Not regDoes_Key_Exist(HKEY_CURRENT_USER, "Software\DAS") Then
        regCreate_A_Key HKEY_CURRENT_USER, "Software\DAS"
    End If
    
    If Not regDoes_Key_Exist(HKEY_CURRENT_USER, "Software\DAS\Profume") Then
        regCreate_A_Key HKEY_CURRENT_USER, "Software\DAS\Profume"
    End If
    
    If Not regDoes_Key_Exist(HKEY_CURRENT_USER, "Software\DAS\Profume\Options") Then
        regCreate_A_Key HKEY_CURRENT_USER, "Software\DAS\Profume\Options"
    End If
    
    If Not regDoes_Key_Exist(HKEY_CURRENT_USER, "Software\DAS\Profume\Options\Units") Then
        regCreate_A_Key HKEY_CURRENT_USER, "Software\DAS\Profume\Options\Units"
    End If
    
    ' check the metric key and set the global IsMetric variable
    vKeyValue = regQuery_A_Key(HKEY_CURRENT_USER, "Software\DAS\Profume\Options\Units", "Metric")
    If vKeyValue = "" Then
        regCreate_Key_Value HKEY_CURRENT_USER, "Software\DAS\Profume\Options\Units", "Metric", 1
        gbIsMetric = True
    Else
        If vKeyValue = 1 Then
            gbIsMetric = True
        Else
            gbIsMetric = False
        End If
    End If
    
    
    ' check if the secret testing key is there or not
    ' we do not create the key if it does not exist (it's supposed to be secret :) )
    vKeyValue = regQuery_A_Key(HKEY_CURRENT_USER, "Software\DAS\Profume\Options", "Testing")
    If vKeyValue <> "" Then
        If vKeyValue = 1 Then
            gbIsTestingActive = True
        Else
            gbIsTestingActive = False
        End If
    End If
      
    ' Jett 2/11/2004 LOCALIZATION
    ' added checking of a secret registry entry to force the UI language
    
    vKeyValue = regQuery_A_Key(HKEY_CURRENT_USER, "Software\DAS\Profume\Options", "ForceLanguage")
    If vKeyValue <> "" Then
        Select Case UCase(vKeyValue)
            Case "ENGLISH":
                gForceLanguage = eENGLISH
            Case "FRENCH":
                gForceLanguage = eFrench
            Case "GERMAN"
                gForceLanguage = eGERMAN
            Case "ITALIAN"
                gForceLanguage = eITALIAN
           '12/11/06 David Smith Added code for Spanish
            Case "SPANISH"
                gForceLanguage = eSPANISH
           Case Else  ' default to neutral if we do not have a valid language string
                gForceLanguage = eNEUTRAL
        End Select
    Else
        ' default to neutral language
        gForceLanguage = eNEUTRAL

    End If
  
    If Not regDoes_Key_Exist(HKEY_CURRENT_USER, "Software\DAS\Profume\Registration") Then
        regCreate_A_Key HKEY_CURRENT_USER, "Software\DAS\Profume\Registration"
    End If
    
    If Not regDoes_Key_Exist(HKEY_CURRENT_USER, "Software\DAS\Profume\Settings") Then
        regCreate_A_Key HKEY_CURRENT_USER, "Software\DAS\Profume\Settings"
    End If
    
    'regCreate_Key_Value HKEY_CURRENT_USER, sGridSettingsKey, arrGridColumns(x).sColumnName, arrGridColumns(x).nColumnWidth

   Exit Sub
   
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub

'**************************************************************
'   DESCRIPTION
'   Checks if there is a need for a calculation. We need to calculate if
'   there are no calc results and there are monitoring point data
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
' 10/11/2002       Jett Gamboa       modified to always check if there is monitoring
'                                    point value data
'**************************************************************
Public Function NeedsCalculation() As Boolean

    Const FUNCTION_NAME = "NeedsCalculation()"

    Dim sErrorDesc As String
    Dim rsRecord As New ADODB.Recordset
    Dim sSQL As String
    Dim nFumigantReq As Double
    
    nFumigantReq = 0
    
    On Error GoTo Error_Handler
    
    ' default to do not need to calculate
    NeedsCalculation = False
    
    ' check if the fumigant is -1 if it is, it means we need to calculate
    sSQL = "SELECT fumigant FROM Job"

    rsRecord.Open sSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText

    ' if there were no records then this must be an invalid template
    If rsRecord.EOF Then
    
        sErrorDesc = LoadResString(IDS_ERR_INVALIDTEMPLATE)
        Set rsRecord = Nothing
        psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
    Else
        ' get fumigant required
        nFumigantReq = CStr(rsRecord("fumigant").Value)
        
        ' close recordset
        rsRecord.Close
        
    End If
          
    sSQL = "SELECT * FROM MonitorPointValue WHERE Deleted = False"
    rsRecord.Open sSQL, gdbConn, adOpenStatic, adLockReadOnly, adCmdText
    
    ' if no records were found then do not calculate
    If nFumigantReq = -1 And Not rsRecord.EOF Then
    
        NeedsCalculation = True
        
    End If
       
    rsRecord.Close
    Set rsRecord = Nothing
    
    Exit Function
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function


'**************************************************************
'   DESCRIPTION
'   Checks if there were job areas entered in the dosage plan
'   tab. If there were none, we should prevent the user from
'   moving to the other tabs
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/10/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Function HasJobArea() As Boolean

    Const FUNCTION_NAME = "HasJobArea()"

    Dim sErrorDesc As String
    Dim rsJobArea As ADODB.Recordset
    Dim sJobAreaSQL As String
    
    On Error GoTo Error_Handler
    
    Set rsJobArea = New ADODB.Recordset

    'Check if there are JobArea records. If there are none, do not display the form.
    sJobAreaSQL = "select id, area_name from JobArea where deleted = false and " & _
        "job_fkey = '" & gsJobID & "' order by area_name"
    
    rsJobArea.Open sJobAreaSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rsJobArea.BOF And rsJobArea.EOF Then
        HasJobArea = False
    Else
        HasJobArea = True
    End If
    
    rsJobArea.Close
    Set rsJobArea = Nothing

    Exit Function
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function


'**************************************************************
'   DESCRIPTION
'   Checks if there are any introduction lines
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/10/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Function HasIntroLines() As Boolean

    Const FUNCTION_NAME = "HasIntroLines()"

    Dim sErrorDesc As String
    Dim rsIntroLines As ADODB.Recordset
    Dim sIntroLinesSQL As String
    
    On Error GoTo Error_Handler
    
    Set rsIntroLines = New ADODB.Recordset

    'Check if there are any introduction lines
    sIntroLinesSQL = "Select JobArea.id as area_id, JobArea.area_name, ShootingLine.name, ShootingLine.id as line_id from JobArea , ShootingLine where JobArea.deleted = false and ShootingLine.deleted = false and Shootingline.jobarea_fkey = JobArea.id"
    
    rsIntroLines.Open sIntroLinesSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rsIntroLines.BOF And rsIntroLines.EOF Then
        HasIntroLines = False
    Else
        HasIntroLines = True
    End If
    
    rsIntroLines.Close
    Set rsIntroLines = Nothing

    Exit Function
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function


'**************************************************************
'   DESCRIPTION
'   Checks if there are any monitor points
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/10/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Function HasMonitorPoints() As Boolean

    Const FUNCTION_NAME = "HasIntroLines()"

    Dim sErrorDesc As String
    Dim rsMonPts As ADODB.Recordset
    Dim sMonPtsSQL As String
    
    On Error GoTo Error_Handler
    
    Set rsMonPts = New ADODB.Recordset

    'Check if there are any introduction lines
    sMonPtsSQL = "Select id, name, jobarea_fkey from MonitorPoint where deleted = false"
    
    rsMonPts.Open sMonPtsSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rsMonPts.BOF And rsMonPts.EOF Then
        HasMonitorPoints = False
    Else
        HasMonitorPoints = True
    End If
    
    rsMonPts.Close
    Set rsMonPts = Nothing

    Exit Function
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function

'**************************************************************
'   DESCRIPTION
'   Checks if there are concentration readings in data input
'
'   PARAMETERS
'       none
'
'   RETURNS - True if there are
'             False if there are none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  04/27/2003      Joseph Manyaga    Initial Revision
'**************************************************************
Private Function HasConcReadings() As Boolean

    Const FUNCTION_NAME = "HasConcReadings()"
    
    Dim rsConc As ADODB.Recordset
    Dim sConcSQL As String
    Dim sErrorDesc As String
  
    On Error GoTo Error_Handler
    
    Set rsConc = New ADODB.Recordset

    sConcSQL = "select id from monitorpointvalue where deleted = false and " & _
        "fumiscope_conc >= 0"
    rsConc.Open sConcSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
      
    If rsConc.BOF And rsConc.EOF Then
        HasConcReadings = False
    Else
        HasConcReadings = True
    End If

    rsConc.Close
    Set rsConc = Nothing

    Exit Function
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function

'**************************************************************
'   DESCRIPTION
'   Checks if there are any monitor points
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/10/2002       Jett Gamboa       Initial Revision
'   5/27/2003      Jett Gamboa       Changed Function type to Single so we
'                                    get the decimal places
'**************************************************************
Public Function CheckExposureTimes() As Single

    Const FUNCTION_NAME = "CheckExposureTimes()"

    Dim sErrorDesc As String
    Dim rsAreas As ADODB.Recordset
    Dim sAreasSQL As String
    
    On Error GoTo Error_Handler
    
    Set rsAreas = New ADODB.Recordset

    ' sleect distinct planned exposure
    sAreasSQL = "SELECT DISTINCT planned_exposure AS exposure FROM JobArea WHERE deleted = No"
    
    rsAreas.Open sAreasSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rsAreas.BOF And rsAreas.EOF Then
        CheckExposureTimes = 0
    Else
        If rsAreas.RecordCount > 1 Then
            CheckExposureTimes = -1
        Else
            '--JG 6/03/2003, removed Rounding off
            'CheckExposureTimes = Round(rsAreas("exposure").Value, 1)
            CheckExposureTimes = rsAreas("exposure")
        End If
    End If
    
    rsAreas.Close
    Set rsAreas = Nothing

    Exit Function
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    Set rsAreas = Nothing
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function


'*************************************************************
'   DESCRIPTION
'   Hides/Displays the toolbar
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
' 11/12/2002       Jett Gamboa       fixed bug where Active form is moved even if it does not exist
'**************************************************************
Public Sub mnuViewToolbar_Click_Handler()

    ' TODO: fix the view toolbar shit

    Const FUNCTION_NAME = "mnuViewToolbar_Click_Handler()"
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler

    With frmMain
        ' toggle the visibility of the menu bar
        
        .mnuViewToolbar.Checked = Not .mnuViewToolbar.Checked
        .tbToolBar.Visible = .mnuViewToolbar.Checked
        
        ' move the tab control up if the toolbar is not visible
        If .mnuViewToolbar.Checked Then
            .tabTabs.Top = .tabTabs.Top + .tbToolBar.Height
            .tabTabs.Height = .tabTabs.Height - .tbToolBar.Height
            
            gActiveFrame.Top = gActiveFrame.Top + .tbToolBar.Height
            .Height = .Height + .tbToolBar.Height
        Else
            .tabTabs.Top = .tabTabs.Top - .tbToolBar.Height
            .tabTabs.Height = .tabTabs.Height + .tbToolBar.Height
            
            gActiveFrame.Top = gActiveFrame.Top - .tbToolBar.Height
            .Height = .Height - .tbToolBar.Height
        End If
    End With
    
    ' 11/12/02 TSG - only move the form if it is there.
'    If Not gfrmActiveForm Is Nothing Then
'        gfrmActiveForm.Move 0, 0
'    End If
    
    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc
    
End Sub



'*************************************************************
'   DESCRIPTION
'   Hides/Displays the Status bar
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Sub ViewStatusBar_Handler()
    
    Const FUNCTION_NAME = "ViewStatusBar_Handler()"
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler
    
    ' toggle display of the status bar
    frmMain.mnuViewStatusBar.Checked = Not frmMain.mnuViewStatusBar.Checked
    frmMain.sbStatusBar.Visible = frmMain.mnuViewStatusBar.Checked

    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'   Is called by the KeyDown event handler for the child form. Since
'   the tab forms are technically separate forms, they need to explicitly
'   trap the shortcut keys for the menu
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/17/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Sub HandleShortcutKeys(pvKeyCode As Integer, pvShift As Integer)

    DoEvents

    ' If the control key was pressed
    If pvShift = 2 Then
    
        Select Case pvKeyCode
            Case 78       ' N
                NewFile
            Case 83       ' S
                SaveFile
            Case 79       ' O
                OpenFile
        End Select
            
    End If
    
    If pvShift = 0 And pvKeyCode = 123 Then
        SaveFile True
    End If

    DoEvents

End Sub

' TODO: evaluate if we need this
Public Function ReturnCurrTab(nTab As Integer)
'**************************************************************
'   DESCRIPTION
'       This function returns the tab to the current tab when there
'       is an invalid value in the form controls.
'
'   PARAMETERS
'           nTab = the current tab number
'
'   RETURNS - NONE
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'**************************************************************

    gbForcedClick = True
    frmMain.tabTabs.Tabs(nTab).Selected = True
    gbForcedClick = False

End Function

'*******************************************************
' SUBROUTINE : FormatFumDate
'
' DESCRIPTION: Formats the date and time. Instead of using the
'              standard VB FormatDate function, this function
'              will be used instead because of special user
'              requirements.
'
'  INPUTS:     none
'  Output:     None
'
' REVISION HISTORY
' Revision:    0
' Author:      Manyaga, Joseph (NAI1227)
' Date:        October 18, 2002
'*******************************************************
Public Function FormatFumDate(dtValue As Date) As String

    Dim sDate As String
    Dim sDay As String
    Dim sHour As String
    Dim sMinute As String
    Dim sAMPM As String
    
    On Error GoTo Error_Handler
       
    sDay = Day(dtValue)
    sDay = IIf(Len(sDay) = 1, "0" & sDay, sDay)
    

    sDate = sDay & "-" & MonthName(Month(dtValue), True) & " "
    If Hour(dtValue) = 0 Then
        sHour = "12:"
    Else
        sHour = IIf(Hour(dtValue) > 12, Hour(dtValue) - 12, Hour(dtValue)) & ":"
    End If
    
    ' Jett 3/5/04 LOCALIZATION - used system time format
    
'    sMinute = Minute(dtValue) & " "
'    If Len(sMinute) = 2 Then
'        sMinute = "0" & sMinute
'    End If
'    If Hour(dtValue) > 11 Then
'        sAMPM = "PM"
'    Else
'        sAMPM = "AM"
'    End If
    
'    FormatFumDate = sDate & sHour & sMinute & sAMPM
     FormatFumDate = sDate & " " & FormatDateTime(dtValue, vbShortTime)
    
Error_Handler:
        If Err.Number <> 0 Then
            Call psAbort(MODULE_NAME, "FormatFumDate", Err.Description)
        End If

End Function



'*************************************************************
'   DESCRIPTION
'   Handles the Copy routine. There is a generic code to handle
'   Pasting for normal controls. If the function was called from
'   the Monitor Data Input grid, that grid's function is called
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'  4/17/2003       Jett Gamboa       converted to a handler
'**************************************************************
Public Sub mnuEditCopy_Click_Handler()
    
    Const FUNCTION_NAME = "mnuEditCopy_Click_Handler()"
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler

    ' make sure that a control was selected before doing anything
    If Not frmMain.ActiveControl Is Nothing Then
    
        ' if it's from the data input grid, call the special function
        If frmMain.ActiveControl.Name = "grdMonitorDataInput" Then
    
            'TODO: fix this--- frmMonitorDataInput.mnuGridCopy_Click
    
        Else
        
            ' don't bother copying if it's not a text box
            If TypeOf frmMain.ActiveControl Is TextBox Or _
               TypeOf frmMain.ActiveControl Is RichTextBox Then _
                        
                ' clear the clipboard and set the text
                Clipboard.SetText frmMain.ActiveControl.SelText
                
            End If
        
        End If
        
    End If
        
    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc
        
End Sub



'*************************************************************
'   DESCRIPTION
'   Handles the Paste routine. There is a generic code to handle
'   Pasting for normal controls. If the function was called from
'   the Monitor Data Input grid, that grid's function is called
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'  4/17/2003       Jett Gamboa       converted to a handler
'**************************************************************
Public Sub mnuEditPaste_Click_Handler()

    Const FUNCTION_NAME = "mnuEditPaste_Click_Handler()"
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler

    ' make sure that a control was selected before doing anything
    If Not frmMain.ActiveControl Is Nothing Then
    
        ' if it's from the data input grid, call the special function
        If frmMain.ActiveControl.Name = "grdMonitorDataInput" Then
    
            'TODO: Fix this---- frmMonitorDataInput.mnuGridPaste_Click
    
        Else
        
            ' don't bother pasting if it's not a text box
            If TypeOf frmMain.ActiveControl Is TextBox Or _
               TypeOf frmMain.ActiveControl Is RichTextBox Then
        
                ' set the text from the clipboard
                frmMain.ActiveControl.SelText = Clipboard.GetText()
                
            End If
        
        End If
        
    End If

    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc
    
End Sub



'*************************************************************
'   DESCRIPTION
'   Handles the Cut routine. There is a generic code to handle
'   Pasting for normal controls.
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
' 10/25/2002       Joseph Manyaga    Initial Revision
'  4/17/2003       Jett Gamboa       converted to a handler
'**************************************************************
Public Sub mnuEditCut_Click_Handler()

    Const FUNCTION_NAME = "mnuEditCut_Click_Handler()"
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler

    ' make sure that a control was selected before doing anything
    If Not frmMain.ActiveControl Is Nothing Then
    
        ' don't bother cutting if it's not a text box
        If TypeOf frmMain.ActiveControl Is TextBox Or _
            TypeOf frmMain.ActiveControl Is RichTextBox Then _
                        
            ' clear the clipboard and set the text
            Clipboard.SetText frmMain.ActiveControl.SelText
            ' delete the selected item from the text box
            frmMain.ActiveControl.SelText = ""
                    
        End If
        
    End If
        
    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc
End Sub



'**************************************************************
'   DESCRIPTION
'   Called when form is unloaded. Any open files are closed and
'   the active subform is also closed
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
' 10/11/2002       Jett Gamboa       removed forced validation of active control when
'                                    program is closed
' 10/16/2002       Jett Gamboa       Added special handling when closing the file
'  4/17/2003       Jett Gamboa       converted to a handler
'**************************************************************
Public Sub Form_QueryUnload_Handler(Cancel As Integer, UnloadMode As Integer)
    
    Const FUNCTION_NAME = "Form_QueryUnload_Handler()"
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler
    
    Dim I As Integer

    gbCanChangeTabs = True
    
    ' If a file is open, close it
    If gbIsFileOpen Then
    
        CloseFile
        
    End If

    ' check if we can exit (if the user cancelled on a save dialog,
    ' we should not exit
    If Not gbCanExit Then
        
        ' Cancel the unload event
        Cancel = True
        
        ' switch the flag back
        gbCanExit = True
               
        Exit Sub
        
    End If
         
    ' 11/15/02 TSG - make sure that we set the cursor to "normal" before closing
    Screen.MousePointer = vbDefault
                       
    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc
           
End Sub


'*************************************************************
'   DESCRIPTION
'   Handles the Help -> About menu item
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'  4/17/2003       Jett Gamboa       converted to a handler
'**************************************************************
Public Sub mnuHelpAbout_Click_Handler()
    
    Const FUNCTION_NAME = "mnuHelpAbout_Click_Handler()"
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler
    
    frmAbout.Show vbModal, frmMain

    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Sub


Public Sub Load_Main()

    frmMain.fraCustomerInfo.Visible = False
    frmMain.fraDosagePlan.Visible = False
    frmMain.fraIntroductionPlan.Visible = False
    frmMain.fraMonitorPlan.Visible = False
    frmMain.fraIntroHistory.Visible = False
    frmMain.fraMonitorDataInput.Visible = False
    frmMain.fraMonitorStatus.Visible = False
    frmMain.fraGraph.Visible = False
    frmMain.fraCustomerInfo.Visible = False
       
    ' tabs are invisible
    frmMain.tabTabs.Visible = False
    
    ' splash screen is visible
    frmMain.fraSplash.Visible = True
    
    Set gActiveFrame = frmMain.fraSplash

End Sub



Public Sub mnuViewOptions_Click_Handler()
    
    Const FUNCTION_NAME = "mnuViewOptions_Click_Handler()"

    Dim bOldUnitSetting As Boolean
    Dim nKeyValue As Integer
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler

    ' store the old unit of measure
    bOldUnitSetting = gbIsMetric

    ' display the options window
    frmOptions.Show vbModal, frmMain
    
    'If gfrmActiveForm Is Nothing Then
    If gActiveFrame.Name = "fraSplash" Then
        Exit Sub
    End If
    
    ' If the old unit of measure is not the same as the new one
    ' refresh the current tab
    If bOldUnitSetting <> gbIsMetric Then
        ' unload and reload the active tab
        UnloadTab gActiveFrame
               
        ' TSG 4/1/03 Reset the global UOM variable if the previous tab did not allow it to switch
        If Not gbCanChangeTabs Then
        
            nKeyValue = regQuery_A_Key(HKEY_CURRENT_USER, "Software\DAS\Profume\Options\Units", "Metric")
            regCreate_Key_Value HKEY_CURRENT_USER, "Software\DAS\ProFume\Options\Units", "Metric", IIf(nKeyValue = 0, 1, 0)
            
            gbIsMetric = Not gbIsMetric
        
        End If
        
        LoadTab gActiveFrame
    End If
    
    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc
   
    
End Sub


Public Sub tabTabs_BeforeClick_Handler(Cancel As Integer)

    If gbForcedClick Then
        Exit Sub
    End If
     
    ' keep tabs on the previous tab
    nPreviousTab = frmMain.tabTabs.SelectedItem.Index
    
    'JNM 05/10/03 - Do not allow the user to leave Monitor Input or
    '   Monitor Status if he has
    '   not clicked the Accept Date:Time button
    If nPreviousTab = TAB_MONITOR_INPUT Then
        Cancel = Not CanExitMonInput
    Else
        If nPreviousTab = TAB_MONITOR_STATUS Then
            Cancel = Not CanExitMonStatus
        End If
    End If
    
End Sub



'*************************************************************
'   DESCRIPTION
'   Loads the appropriate form after a tab is selected by the
'   user
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
' 10/14/2002       Jett Gamboa       Added validation checks before unloading
'                                    the current tab
' 12/20/2002       Jett Gamboa       Check if exposure times are equal before allowing
'                                    user to go to Input, Status, and Graph tab.
'  4/18/2003       Jett Gamboa       Move to module
'**************************************************************
Public Sub tabTabs_Click_Handler()

    Const FUNCTION_NAME = "tabTabs_Click_Handler()"
    
    Dim strErrorDesc As String
    
    On Error GoTo Error_Handler
    
    If gbNewFile Then
        frmMain.txtLoadFactor.Visible = False
        frmMain.lblLoadFactor.Visible = False
        frmMain.chkNoAdjSorp.Visible = False
    Else
        frmMain.txtLoadFactor.Visible = True
        frmMain.lblLoadFactor.Visible = True
        frmMain.chkNoAdjSorp.Visible = True
    End If
  
    If gbForcedClick Then
        Exit Sub
    End If
    
    ' if there was not previous tab, then set it to 1 (the first tab)
    If nPreviousTab = 0 Then
        nPreviousTab = 1
    End If
    
    ' if the current tab is clicked do not do anything
    If frmMain.tabTabs.SelectedItem.Index = nPreviousTab And gActiveFrame.Name <> "fraSplash" Then
        Exit Sub
    End If
       
    ' If the user clicked on a tab other than the first 2
    If frmMain.tabTabs.SelectedItem.Index > 2 Then

        ' If no areas were added,
        If Not HasJobArea() Then
            
            gbForcedClick = True
            frmMain.tabTabs.Tabs(nPreviousTab).Selected = True
            MsgBox LoadResString(IDS_INTRO_NO_AREA), vbExclamation
            gbForcedClick = False

            Exit Sub
        End If

        ' validate if a calculation needs to be done first
        If NeedsCalculation() Then
        
            gbForcedClick = True
            frmMain.tabTabs.Tabs(nPreviousTab).Selected = True
            MsgBox LoadResString(IDS_CALC_FUMIGANT_REQ), vbExclamation
            gbForcedClick = False
            
            Exit Sub
        End If

    End If
    
    ' 12/20/02 TSG   - added checking for exposure times
    ' check the exposure times if we are coming from the Fumigation Dosage Plan tab
    If nPreviousTab = TAB_DOSAGE_PLAN Then
        gnExposureTime = CheckExposureTimes()
         
        ' Display the warning message if user did NOT click on the Monitor Input
        ' or Status tabs (we do the same checking if checking was done on those tabs
        If gnExposureTime = -1 And frmMain.tabTabs.SelectedItem.Index < TAB_MONITOR_INPUT Then
            MsgBox LoadResString(IDS_EQUAL_EXPOSURE_TIMES), vbExclamation
        End If
        
    End If
    
    ' we cannot go to the Introduction History tab if lines have not yet been
    ' entered in the Introduction Plan page
    
    If frmMain.tabTabs.SelectedItem.Index = TAB_INTRO_HISTORY Then
    
        If Not HasIntroLines() Then
    
            gbForcedClick = True
            frmMain.tabTabs.Tabs(nPreviousTab).Selected = True
            MsgBox LoadResString(IDS_INTROHIST_MSG_NOJOBAREAS), vbExclamation
            gbForcedClick = False
            
            Exit Sub
                    
        End If

    End If
    
    ' We cannot go to the Monitor Data Input, Monitor Status, or Graph tabs
    ' if there are no monitoring points
    
    If frmMain.tabTabs.SelectedItem.Index = TAB_MONITOR_INPUT Or _
       frmMain.tabTabs.SelectedItem.Index = TAB_MONITOR_STATUS Or _
       frmMain.tabTabs.SelectedItem.Index = TAB_GRAPH Then
    
        ' If there are no Monitor points, do not allow the user to go to the
        ' Input and Status tabs
        If Not HasMonitorPoints() Then
            gbForcedClick = True
            frmMain.tabTabs.Tabs(nPreviousTab).Selected = True
            MsgBox LoadResString(IDS_MONDATA_MSG_SAVE_NO_MONITOR_POINTS), vbExclamation
            gbForcedClick = False
            
            Exit Sub
        End If
        
        ' 12/20/02 TSG   - added checking for exposure times. We do not allow the user
        ' go to the Monitoring Input or Status page if elapsed times are not equal
        If gnExposureTime = -1 Then
        
            gbForcedClick = True
            frmMain.tabTabs.Tabs(nPreviousTab).Selected = True
            MsgBox LoadResString(IDS_EQUAL_EXPOSURE_TIMES), vbExclamation
            gbForcedClick = False
            
            Exit Sub
            
        End If
    
    End If
    
    'JNM 04/27/03 - AT Fix - We cannot enter Graph tab if we haven't entered
    '   any data yet in Data input.
    If frmMain.tabTabs.SelectedItem.Index = TAB_GRAPH Then
        If Not HasConcReadings() Then
            gbForcedClick = True
            frmMain.tabTabs.Tabs(nPreviousTab).Selected = True
            MsgBox LoadResString(IDS_GRAPH_NO_CONC_YET), vbExclamation
            gbForcedClick = False
            Exit Sub
        End If
    End If
        
    ' default gbCanChangeTabs to true to clear out status for any previous errors
    gbCanChangeTabs = True
        
    ' We passed all previous validation so we can now unload the current form
    'Unload gfrmActiveForm
    
    UnloadTab gActiveFrame

    ' switch back to the previous form if the form did not allow us to change tabs
    If Not gbCanChangeTabs Then

        gbForcedClick = True
        frmMain.tabTabs.Tabs(nPreviousTab).Selected = True
        gbForcedClick = False
        
        Exit Sub

    End If
    
    
    'Pat 4/5/05: If the user tries to transfer tabs without perforeming a conversion, an error message shall prompt the user
    If frmMain.btnConvert.Enabled = True Then
        If blnConvert = False Then
            gbForcedClick = True
            frmMain.tabTabs.Tabs(nPreviousTab).Selected = True
            If gbForcedClick = True Then
                MsgBox LoadResString(IDS_CLICK_CONVERT), vbExclamation
            End If
            gbForcedClick = False
            blnChangeTab = True
        End If
    End If
       
    Select Case frmMain.tabTabs.SelectedItem.Index
        Case 1
            LoadTab frmMain.fraCustomerInfo
        Case 2
            LoadTab frmMain.fraDosagePlan
        Case 3
            LoadTab frmMain.fraIntroductionPlan
        Case 4
            LoadTab frmMain.fraMonitorPlan
       Case 5
            LoadTab frmMain.fraIntroHistory
        Case 6
            LoadTab frmMain.fraMonitorDataInput
        Case 7
            LoadTab frmMain.fraMonitorStatus
        Case 8
            LoadTab frmMain.fraGraph
    End Select

    gbForcedClick = False

    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc


End Sub


'*************************************************************
'   DESCRIPTION
'   Loads the appropriate form after a tab is selected by the
'   user
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
' 10/25/2002       Joseph Manyaga    Added the Cut item
'  3/25/2003       Jett Gamboa       commented out print toolbar code
'  4/18/2003       Jett Gamboa       moved to module. calling action functions directly.
'**************************************************************
Public Sub tbToolBar_ButtonClick_Handler(ByVal Button As MSComctlLib.Button)
    
    Const FUNCTION_NAME = "tbToolBar_ButtonClick_Handler()"
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler
    
    ' call the corresponding menu routine depending on the button clicked
    Select Case Button.Key
        Case "New"
            NewFile
        Case "Open"
            OpenFile
        Case "Save"
            SaveFile
        'Case "Print"               commented out print routine. it is no longer
        '    mnuFilePrint_Click     available from the toolbar
        Case "Cut"
            mnuEditCut_Click_Handler
        Case "Copy"
            mnuEditCopy_Click_Handler
        Case "Paste"
            mnuEditPaste_Click_Handler
    End Select
    
    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc
        
End Sub


'*************************************************************
'   DESCRIPTION
'   Handles the menu File - Exit event by unloading the Main form
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'  4/18/2003       Jett Gamboa       moved to module
'**************************************************************
Public Sub mnuFileExit_Click_Handler()
    
    Const FUNCTION_NAME = "mnuFileExit_Click_Handler()"
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler
    
    'unload the form
    Unload frmMain
    End

    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Sub



'**************************************************************
'   DESCRIPTION
'   Handles the Help -> Contents menu item
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'  4/18/2003       Jett Gamboa       moved to module
'**************************************************************
Public Sub mnuHelpContents_Click_Handler()

    Const FUNCTION_NAME = "mnuHelpContents_Click_Handler()"
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler

    HtmlHelp frmMain.hWnd, App.HelpFile, HH_DISPLAY_TOPIC, 0

    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Sub



'*************************************************************
'   DESCRIPTION
'   Hides/Displays the Status bar
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/25/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Sub mnuViewStatusBar_Click_Handler()
    
    Const FUNCTION_NAME = "mnuViewStatusBar_Click_Handler()"
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler
    
    ' toggle display of the status bar
    frmMain.mnuViewStatusBar.Checked = Not frmMain.mnuViewStatusBar.Checked
    frmMain.sbStatusBar.Visible = frmMain.mnuViewStatusBar.Checked
    
    If frmMain.sbStatusBar.Visible Then
        frmMain.Height = frmMain.Height + frmMain.sbStatusBar.Height
    Else
        frmMain.Height = frmMain.Height - frmMain.sbStatusBar.Height
    End If

    Exit Sub
           
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Sub


Public Sub mnuFileExport_Click_Handler()
    ExportFile
End Sub


Public Sub mnuFileExportGraph_Click_Handler()
    'JNM 04/27/03 - AT Fix - Tab to graph so that graph will be reset first.
    '   As an added benefit, Export Graph would therefore have the same
    '   validation as the Graph tab.
    frmMain.tabTabs.Tabs(TAB_GRAPH).Selected = True
    If frmMain.tabTabs.SelectedItem.Index = TAB_GRAPH Then
        ExportGraph
    End If
End Sub

'**************************************************************
'   DESCRIPTION
'   Validates the Start Intro, Start Exposure and End Exposure date and time.
'
'   Input Parameters: the dates to be validated
'
'   RETURNS - The name of the control that is invalid ("Start Exposure" or
'                   "Start Introduction")
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  May 9, 2003     Joseph Manyaga    Initial Revision
'**************************************************************
Public Function ValidateJobDates(dtStartIntroduction As Date, dtStartExposure As Date, _
dtEndExposure As Date) As String

    Const FUNCTION_NAME = "ValidateJobDates()"
    
    Dim sSQL As String
    Dim rsDateSeq As ADODB.Recordset
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler
    
    'Do not allow users to change Start Intro prior to the earliest concentration
    'reading
    sSQL = "Select fumiscope_conc from MonitorPointValue where deleted = false " & _
            "and fumiscope_conc > 0 and date_time < CDate('" & dtStartIntroduction & "')"
            
    Set rsDateSeq = New ADODB.Recordset
            
    rsDateSeq.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    If rsDateSeq.RecordCount <> 0 Then
        MsgBox LoadResString(IDS_MONDATA_MSG_INTRO_AFTER_CONCENTRATION), vbExclamation
                
        rsDateSeq.Close
        Set rsDateSeq = Nothing
        
        ValidateJobDates = "Start Introduction"
        
        Exit Function
    Else
        rsDateSeq.Close
        Set rsDateSeq = Nothing
    End If
    
    'Evaluate Start Introduction and Start Exposure
    If dtStartIntroduction <= dtStartExposure Then
        'Evaluate Start Exposure and End Exposure
        If dtStartExposure < dtEndExposure Then
            'Save the dates in the database
            SaveJobDates dtStartIntroduction, dtStartExposure, dtEndExposure
            ValidateJobDates = ""
        Else
            MsgBox LoadResString(IDS_MONDATA_MSG_END_INTRO_AFTER_EXPOSURE), vbExclamation
            ValidateJobDates = "Start Exposure"
        End If
    Else
        MsgBox LoadResString(IDS_MONDATA_MSG_INTRO_END_BEFORE_START), vbExclamation
        ValidateJobDates = "Start Introduction"
    End If
    
    Exit Function
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Function

'**************************************************************
'   DESCRIPTION
'   Saves the Start Intro, Start Exposure and End Exposure date and time.
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  May 9, 2003     Joseph Manyaga    Initial Revision
'**************************************************************
Private Sub SaveJobDates(dtStartIntroduction As Date, dtStartExposure As Date, _
dtEndExposure As Date)
    
    Dim sUpdateSQL As String
    Const FUNCTION_NAME = "SaveJobDates()"
    
    Dim strErrorDesc As String
    On Error GoTo Error_Handler
    
    sUpdateSQL = "UPDATE Job " & _
                "SET startintro_Datetime = '" & dtStartIntroduction & "', " & _
                "startexp_datetime = '" & dtStartExposure & "', " & _
                "endexp_datetime = '" & dtEndExposure & "'"
    gdbConn.Execute sUpdateSQL
    
    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    strErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, strErrorDesc

End Sub
