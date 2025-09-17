Attribute VB_Name = "modMonitoringPlan"
'*********************************************************************************************
'   DESCRIPTION
'       Functions and subroutines used in the MonitorPlan page
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
' 10/14/2002       LN Gosiengfiao    For AT defect correction,
'                                       inserted a DeleteEmptyRow function,
'                                       a PopulateMonPointGrid and a StartEdit
'                                       function.  Added a logic in the unload
'                                       function to check if row is empty and
'                                       then skip validation.  Modified some
'                                       settings to match the over all look
'                                       and feel of the other tabs
' 11/13/2002       Jett Gamboa      added default column widths for grid
' 11/14/2002       Joseph Manyaga    Removed all processing for Inactive column.
'************************************************************************************
Option Explicit

'Define constants used in Monitoring Plan page
Private Const MODULE_NAME = "MonitoringPlan"
Private Const GRIDSIZE = 101
Private Const GRID_COLOR = "&H80000008"
Private Const BACK_COLOR_ALTERNATE = "&H80000016"
Private Const READ_ONLY_COLOR = "&H8000000B"

Const AREANAME_COLSIZE = 3045
Const MONPTNAME_COLSIZE = 4230
Const LOCDESC_COLSIZE = 3960

'Global variables for the record set connections
Dim rsAreaName As New ADODB.Recordset
Dim rsMonitorPoint As New ADODB.Recordset
Dim rsDelete As New ADODB.Recordset
Dim rsStartIntro As New ADODB.Recordset
Dim rsCheckName As New ADODB.Recordset
Dim rsMPV As New ADODB.Recordset
Dim rsExposure As New ADODB.Recordset
Dim rsDefMonPt As New ADODB.Recordset

'Global variables used in Monitoring Plan page
Dim sDefArea As String
Dim sDefMonPt As String
Dim MyGrid As New clsGridWrapper


'**************************************************************
'   DESCRIPTION
'       Sub called upon load of frmMonitorPlan form
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Public Sub Load_MonitorPlan()
    Const FUNCTION_NAME = "Load_MonitorPlan"

    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler

    'set global variable to change tabs initially to true
    gbCanChangeTabs = True
   
    With frmMain
        'set the label caption and set font to bold
        .lblMonitorPlan.Caption = LoadResString(IDS_MONITOR_PLAN_LABEL)
        .lblMonitorPlan.FontBold = True
        
        'set the button text
        .cmdNextStep4.Caption = LoadResString(IDS_BTN_NEXTSTEP)
    End With
    
    'call the functions to create record sets for area names and monitor point
    'information, call also function to paint the grid
    CreateAreaRS
    CreateMonPointRS
    PaintGridMonPlan
      
Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub
'**************************************************************
'   DESCRIPTION
'       Sub to create the record set containing the area names
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************

Private Sub CreateAreaRS()
    Const FUNCTION_NAME = "CreateAreaRS"
    
    Dim sSQLAreaName As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler

    'assign the sql call to get area names from database and execute the call to
    'create the record set
    sSQLAreaName = "Select id, area_name from JobArea where deleted = False"
    rsAreaName.Open sSQLAreaName, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub

'**************************************************************
'   DESCRIPTION
'       Sub to create the record set containing the monitor
'       point data
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************

Private Sub CreateMonPointRS()
    Const FUNCTION_NAME = "CreateMonPointRS"
    
    Dim sSQLMonitorPoint As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler

    'assign the sql call to get monitor point data from database and execute the
    'call to create the record set
    sSQLMonitorPoint = "Select id, name, description, deleted, row, jobarea_fkey, job_fkey from MonitorPoint where deleted = False order by row"
    rsMonitorPoint.Open sSQLMonitorPoint, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub

'**************************************************************
'   DESCRIPTION
'       Sub which takes care of painting the grdMonPlan
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************

Private Sub PaintGridMonPlan()
    Const FUNCTION_NAME = "PaintGridMonPlan"
    
    Dim sErrorDesc As String
    Dim iCtr As Integer
    
    On Error GoTo ErrorHandler
    
    'Initialize grid making use of the grid wrapper class
    MyGrid.Initialize "MonitorPlan", frmMain.grdMonPlan
    MyGrid.AddColumn "txtCounter", "", colTypeText, flexAlignCenterTop
    MyGrid.AddColumn "cboAreaName", LoadResString(IDS_AREA_NAME), colTypeCombo, flexAlignGeneral
    MyGrid.AddColumn "txtMonPointName", LoadResString(IDS_MONITOR_POINT_NAME), colTypeText, flexAlignGeneral
    MyGrid.AddColumn "txtLocDesc", LoadResString(IDS_LOCATION_DESCRIPTION), colTypeText, flexAlignGeneral
    
    ' default column widths (will be overridden if previous settings were saved
    With frmMain.grdMonPlan
        .ColWidth(1) = AREANAME_COLSIZE
        .ColWidth(2) = MONPTNAME_COLSIZE
        .ColWidth(3) = LOCDESC_COLSIZE
    End With
    
    MyGrid.LoadGridSettings GRID_SETTINGS_PATH
    
    frmMain.grdMonPlan.FixedCols = 1
    
    frmMain.grdMonPlan.Rows = GRIDSIZE
    frmMain.grdMonPlan.GridColor = GRID_COLOR
    'grdMonPlan.BackColorAlternate = BACK_COLOR_ALTERNATE
    frmMain.grdMonPlan.ColWidth(0) = COL_INDEX_WIDTH
            
    'set read only color
    For iCtr = 1 To (GRIDSIZE - 1) Step 2
        frmMain.grdMonPlan.Cell(flexcpBackColor, iCtr, 0) = READ_ONLY_COLOR
    Next
        
    PopulateGridMonPlan
    
Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub

'**************************************************************
'   DESCRIPTION
'       Sub which takes care of populating the grdMonPlan with
'       data
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************

Private Sub PopulateGridMonPlan()
    Const FUNCTION_NAME = "PopulateGridMonPlan"
    
    Dim sAreaName As String
    Dim iAreaCnt As Integer, iCnt As Integer, iCtr As Integer, iCount As Integer, iPopulate As Integer
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    With frmMain
    
        'clear all grid values
        .grdMonPlan.Cell(flexcpText, 1, 0, GRIDSIZE - 1, 3) = ""
        .grdMonPlan.Cell(flexcpData, 1, 0, GRIDSIZE - 1, 3) = ""
        
        'assign the first column to contain the row numbers
        For iCtr = 1 To (GRIDSIZE - 1)
            .grdMonPlan.Cell(flexcpText, iCtr, 0) = iCtr
            .grdMonPlan.Cell(flexcpData, iCtr, 0) = -1
        Next
        
        'create the combo box list options for the area names
        iCnt = 0
    
        Do While Not rsAreaName.EOF
            If iCnt <> 0 Then
                sAreaName = sAreaName & "|"
            End If
        
            sAreaName = sAreaName & "#" & rsAreaName("id") & ";" & rsAreaName("area_name")
            rsAreaName.MoveNext
            iCnt = iCnt + 1
        Loop
    
        'assign the combo box to column 1
        .grdMonPlan.ColComboList(1) = sAreaName
    
        'loop through the monitor point data record set and display information in grid
        iCount = 1
        iPopulate = 0
    
        If Not rsMonitorPoint.EOF Then
            Do While Not rsMonitorPoint.EOF
                If Str(.grdMonPlan.Cell(flexcpText, iCount, 0)) = Str(rsMonitorPoint("row")) Then
                    iPopulate = iPopulate + 1
                    .grdMonPlan.Cell(flexcpData, iCount, 0) = iPopulate
                    .grdMonPlan.Cell(flexcpText, iCount, 1) = rsMonitorPoint("jobarea_fkey")
                    .grdMonPlan.Cell(flexcpData, iCount, 2) = rsMonitorPoint("id")
                    .grdMonPlan.Cell(flexcpText, iCount, 2) = rsMonitorPoint("name")
                    .grdMonPlan.Cell(flexcpText, iCount, 3) = rsMonitorPoint("description")
                    rsMonitorPoint.MoveNext
                End If
                If Not rsMonitorPoint.EOF Then
                    If rsMonitorPoint("row") = 0 Then
                        rsMonitorPoint.MoveNext
                    End If
                End If
                iCount = iCount + 1
            Loop
        End If
        
    End With
    Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'       Sub called upon unload of frmMonitorPlan form
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Public Sub Unload_MonitorPlan()
    Const FUNCTION_NAME = "Unload_MonitorPlan"
    
    Dim sSQLCheckName As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    MyGrid.SaveGridSettings GRID_SETTINGS_PATH
    
    'check if user is exiting, if true, then unload form without validation
    If gbIsExiting = True Then
        'Cancel = vbFalse
        gbCanChangeTabs = True
    Else
        If frmMain.grdMonPlan.Row = 0 Or frmMain.grdMonPlan.Col = 0 Then
            'TODO check this --- Cancel = vbFalse
            
        'otherwise perform validation before unload
        Else
            'check if row is empty
            If frmMain.grdMonPlan.Cell(flexcpData, frmMain.grdMonPlan.Row, 0) = -1 Then
                'TODO check thisCancel = vbFalse
            Else
                'check for duplication monitor point names
                sSQLCheckName = "Select name from MonitorPoint where name = '" & frmMain.grdMonPlan.Cell(flexcpText, frmMain.grdMonPlan.Row, 2) & "' and row <> " & frmMain.grdMonPlan.Cell(flexcpText, frmMain.grdMonPlan.Row, 0) & " and deleted <> True"
                rsCheckName.Open sSQLCheckName, gdbConn, adOpenForwardOnly, adLockReadOnly
        
                If rsCheckName.EOF Then
                    gbCanChangeTabs = True
                    'TODO check this Cancel = vbFalse
                Else
                    MsgBox (LoadResString(IDS_MONPLAN_MSG_DUPLICATE_NAMES)), vbExclamation
                    gbCanChangeTabs = False
                    'TODO check this--- Cancel = vbTrue
    
                    rsCheckName.Close
                    Set rsCheckName = Nothing
                    Exit Sub
                End If
    
                'check if there are empty monitor point names
                If frmMain.grdMonPlan.Cell(flexcpText, frmMain.grdMonPlan.Row, 2) = "" Then
                    MsgBox (LoadResString(IDS_MONITORPOINT_NAME_EMPTY)), vbExclamation
                    gbCanChangeTabs = False
                    'TODO check this--- Cancel = vbTrue
    
                    rsCheckName.Close
                    Set rsCheckName = Nothing
                    Exit Sub
                Else
                    gbCanChangeTabs = True
                    'TODO check this --- Cancel = vbFalse
                End If

                rsCheckName.Close
                Set rsCheckName = Nothing
            
            End If
        
        End If
        
    End If
        
    rsAreaName.Close
    Set rsAreaName = Nothing
    
    rsMonitorPoint.Close
    Set rsMonitorPoint = Nothing
    
    'TODO Check this--- Unload frmMonitorPlan
    
Exit Sub

ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub

Public Sub grdMonPlan_AfterEdit_Handler(ByVal Row As Long, ByVal Col As Long)
    Const FUNCTION_NAME = "grdMonPlan_AfterEdit_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    rsMonitorPoint.Close
    Set rsMonitorPoint = Nothing

    rsAreaName.Close
    Set rsAreaName = Nothing
    
    'rebuild the record sets for area name and monitor point data and
    'repaint the grid
    CreateAreaRS
    CreateMonPointRS
    PopulateGridMonPlan
    
Exit Sub

ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub

Public Sub grdMonPlan_AfterUserResize_Handler(ByVal Row As Long, ByVal Col As Long)
    Const FUNCTION_NAME = "grdMonPlan_AfterUserResize_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    MyGrid.SaveGridSettings GRID_SETTINGS_PATH
    
Exit Sub

ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


'**************************************************************
'   DESCRIPTION
'       Sub called before edit of grdMonPlan cells
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Public Sub grdMonPlan_BeforeEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Const FUNCTION_NAME = "grdMonPlan_BeforeEdit_Handler"

    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    With frmMain
        'check if user is trying to modify first column which contains the row numbers
        If .grdMonPlan.Col = 0 Then
            Cancel = True
        End If
        
        If .grdMonPlan.Col = 2 Then
            .grdMonPlan.EditMaxLength = 20
        End If
        
        If .grdMonPlan.Col = 3 Then
            .grdMonPlan.EditMaxLength = 50
        End If
    End With
    
Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub

'**************************************************************
'   DESCRIPTION
'       Sub called before user resizes columns in the grdMonPlan
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Public Sub grdMonPlan_BeforeUserResize_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Const FUNCTION_NAME = "grdMonPlan_BeforeUserResize_Handler"

    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'check if user is trying to modify first column which contains the row numbers
    If Col = 0 Then
        Cancel = True
    End If
    
Exit Sub
    
ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


'**************************************************************
'   DESCRIPTION
'       Function which finds the start and end exposure times
'
'   PARAMETERS
'           mode - to indicate whether finding the start or
'                   end exposure time
'           id - monitor point id
'
'   RETURNS - start or end exposure time
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
' TODO: where is this function used?
' TODO: rsexposure as well
Private Function fFindExposure(mode, id)
    Const FUNCTION_NAME = "fFindExposure"
    
    Dim sSQLExposure As String, sMode As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'check if mode is start or end exposure then assign specific mdm modes and
    'whether to order by the datetime field descending or ascending
    Select Case mode
    Case "start"
        sMode = MDM_EXPOSURE
        sSQLExposure = "Select mode, date_time from MonitorPointValue mpv, MonitorPoint mp where mp.id = '" & id & "' and mp.id = mpv.monitorpoint_fkey order by date_time"
    Case "end"
        sMode = MDM_PRE_AERATION
        sSQLExposure = "Select mode, date_time from MonitorPointValue mpv, MonitorPoint mp where mp.id = '" & id & "' and mp.id = mpv.monitorpoint_fkey order by date_time desc"
    End Select

    rsExposure.Open sSQLExposure, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    'loop through record set to find the desired exposure date/time
    Do While Not rsExposure.EOF
         If rsExposure("mode") = sMode Then
            fFindExposure = rsExposure("date_time")
            Exit Do
         End If
         rsExposure.MoveNext
    Loop
    
    rsExposure.Close
    Set rsExposure = Nothing
    
Exit Function
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function


'**************************************************************
'   DESCRIPTION
'       Sub called to save information into database
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Private Sub SaveMonitoringPlan(Row, Col)
    Const FUNCTION_NAME = "SaveMonitoringPlan"
    
    Dim sGUID As String, sRow As String, sCol As String, sRowChanged As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
        
    'set global variable to true because form fields have changed
    gbHasChanged = True

    With frmMain

        'check if data in first column is -1 which indicates that user is inserting
        'a new row
        If .grdMonPlan.Cell(flexcpData, Row, 0) = -1 Then
            'get a primary key for the new row
            sGUID = GetGUID()
            'add new row in record set
            rsMonitorPoint.AddNew
            'assign record set values from the grid
            rsMonitorPoint("id") = sGUID
            rsMonitorPoint("job_fkey") = gsJobID
            rsMonitorPoint("Row") = .grdMonPlan.Cell(flexcpText, Row, 0)
            rsMonitorPoint("deleted") = "False"
            
            'modified cell is area name
            If Col = 1 Then
                rsMonitorPoint("jobarea_fkey") = .grdMonPlan.ComboData
            Else
                rsMonitorPoint("jobarea_fkey") = fGetDefArea()
            End If
            
            'modified cell is monitor point name
            If Col = 2 Then
                rsMonitorPoint("name") = .grdMonPlan.EditText
            Else
                rsMonitorPoint("name") = fGetDefMonPt(Row)
            End If
            
            'modified cell is description/location
            If Col = 3 Then
                rsMonitorPoint("description") = .grdMonPlan.EditText
            Else
                rsMonitorPoint("description") = .grdMonPlan.Cell(flexcpText, Row, 3)
            End If
            
            'commit update to recordset
            rsMonitorPoint.Update
            
        Else
            'if user modified an existing row, get the row number from first column
            sRowChanged = .grdMonPlan.Cell(flexcpData, Row, 0)
            'move to the currently modified row in the record set and assign modified
            'values from the grid
            rsMonitorPoint.MoveFirst
            rsMonitorPoint.Move (sRowChanged - 1)
            
            'modified cell is area name
            If Col = 1 Then
                rsMonitorPoint("jobarea_fkey") = .grdMonPlan.ComboData
            Else
                rsMonitorPoint("jobarea_fkey") = .grdMonPlan.Cell(flexcpText, Row, 1)
            End If
            
            'modified cell is monitor point name
            If Col = 2 Then
                rsMonitorPoint("name") = .grdMonPlan.EditText
            Else
                rsMonitorPoint("name") = .grdMonPlan.Cell(flexcpText, Row, 2)
            End If
            
            'modified cell is description/location
            If Col = 3 Then
                rsMonitorPoint("description") = .grdMonPlan.EditText
            Else
                rsMonitorPoint("description") = .grdMonPlan.Cell(flexcpText, Row, 3)
            End If
            
            'commit changes to record set
            rsMonitorPoint.Update
        End If
        
    End With

    gbCanChangeTabs = True
    
Exit Sub

ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'       Function to validate entered fields
'
'   PARAMETERS
'           Row - row of cell which field is being validated
'           Col - column of cell which field is being validated
'
'   RETURNS - True if field is valid and False if not
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Private Function ValidateField(Row, Col)
    Const FUNCTION_NAME = "ValidateField"
    
    Dim sSQLCheckName As String, sCheckMonPt As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'check if monitor point name is empty
    If Col = 2 And frmMain.grdMonPlan.EditText = "" Then
        MsgBox (LoadResString(IDS_MONITORPOINT_NAME_EMPTY)), vbExclamation
        ValidateField = "False"
        gbCanChangeTabs = False
        Exit Function
    Else
        ValidateField = "True"
    End If
    
    'JNM 10/23 - initialize sCheckMonPt
    sCheckMonPt = ""
    If Col = 2 Then
        sCheckMonPt = frmMain.grdMonPlan.EditText
    End If
    
    sSQLCheckName = "Select name from MonitorPoint where name = '" & sCheckMonPt & "' and row <> " & frmMain.grdMonPlan.Cell(flexcpText, Row, 0) & " and deleted <> True"
    rsCheckName.Open sSQLCheckName, gdbConn, adOpenForwardOnly, adLockReadOnly

    'check if there are duplication monitor point names
    If rsCheckName.BOF And rsCheckName.EOF Then 'JNM 10/23 - it should be tested for both BOF and EOF
        ValidateField = "True"
    Else
        ValidateField = "False"
        MsgBox (LoadResString(IDS_MONPLAN_MSG_DUPLICATE_NAMES)), vbExclamation
        gbCanChangeTabs = False
    
        rsCheckName.Close
        Set rsCheckName = Nothing
        Exit Function
    End If
    
    rsCheckName.Close
    Set rsCheckName = Nothing
    
Exit Function
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function


'**************************************************************
'   DESCRIPTION
'       Function which checks if a Monitor Point is used in the
'       input data page
'
'   PARAMETERS
'          id - Monitor Point Name id
'
'   RETURNS - NO if monitor point name is not used and YES
'               if it is
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Private Function IsPresentMonitoringData(id)
    Const FUNCTION_NAME = "IsPresentMonitoringData"
    
    Dim sSQLMPV As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler

    'check in database if monitor point id exists in the monitor point value table
    sSQLMPV = "Select id from MonitorPointValue where monitorpoint_fkey = '" & id & "' and deleted <> True"
    rsMPV.Open sSQLMPV, gdbConn, adOpenForwardOnly, adLockReadOnly

    If rsMPV.EOF Then
        IsPresentMonitoringData = "NO"
    Else
        IsPresentMonitoringData = "YES"
    End If

    rsMPV.Close
    Set rsMPV = Nothing
    
Exit Function
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Function

'**************************************************************
'   DESCRIPTION
'       Sub called when hits on delete button
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/3/2002       LN Gosiengfiao    Initial Revision
'**************************************************************

Public Sub grdMonPlan_KeyDown_Handler(KeyCode As Integer, Shift As Integer)
    Const FUNCTION_NAME = "grdMonPlan_KeyDown_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'check to see if user hits on the delete button
    If KeyCode = 46 Then
        DeleteEntry
    End If

Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


'**************************************************************
'   DESCRIPTION
'       Sub called when user right clicks on a cell in the
'       grdMonPlan
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Public Sub grdMonPlan_MouseDown_Handler(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Const FUNCTION_NAME = "grdMonPlan_MouseDown_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
        
    With frmMain
        'JNM 11/4 - Added code that would select the cell that was clicked.
        'on right click show pop up menu
        If Button = vbRightButton And .grdMonPlan.MouseRow > 0 And .grdMonPlan.MouseCol > 0 Then
        
            .grdMonPlan.Select .grdMonPlan.MouseRow, .grdMonPlan.MouseCol
            .PopupMenu .mnuMonitorPlanContext
            
        End If
    End With
    
Exit Sub

ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'       Sub called when user moves the mouse over the
'       grdMonPlan
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Public Sub grdMonPlan_MouseMove_Handler(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Const FUNCTION_NAME = "grdMonPlan_MouseMove_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    Select Case frmMain.grdMonPlan.MouseCol
    Case 1
        frmMain.grdMonPlan.ToolTipText = LoadResString(IDS_TOOLTIP_AREA_NAME)
    Case 2
        frmMain.grdMonPlan.ToolTipText = LoadResString(IDS_TOOLTIP_MONITOR_POINT_NAME)
    Case 3
        frmMain.grdMonPlan.ToolTipText = LoadResString(IDS_TOOLTIP_DESCRIPTION)
    End Select
    
Exit Sub
    
ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'       Sub called when a user starts to edit the grdMonPlan
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Public Sub grdMonPlan_StartEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Const FUNCTION_NAME = "grdMonPlan_StartEdit_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    'check if column being edited is area name column
    If frmMain.grdMonPlan.Col = 1 Then
        'call function to get default area name
        sDefArea = fGetDefArea()
        
        'set the default value of the area name if empty
        If frmMain.grdMonPlan.Cell(flexcpText, frmMain.grdMonPlan.Row, frmMain.grdMonPlan.Col) = "" Then
            frmMain.grdMonPlan.Cell(flexcpText, frmMain.grdMonPlan.Row, frmMain.grdMonPlan.Col) = sDefArea
        End If
    
    End If
    
    'The line below would ensure that the tab does not change if the user enter enters
    'an invalid value and then clicks on another tab
    ReturnCurrTab TAB_MONITOR_PLAN

Exit Sub
    
ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


'**************************************************************
'   DESCRIPTION
'       Sub called to validate modifications entered by user
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************

Public Sub grdMonPlan_ValidateEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Const FUNCTION_NAME = "grdMonPlan_ValidateEdit_Handler"
    
    Dim sValid As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'when initial load of form do not validate switch of cells
    If Row = 0 And Col = 0 Then
        Exit Sub
    End If

    'set global variable to indicate form has changed
    gbHasChanged = True

    'call function to validate fields
    sValid = ValidateField(Row, Col)
    Select Case sValid
    Case "False"
        Cancel = True
    Case "True"
        SaveMonitoringPlan Row, Col
    End Select
    
    'JNM 10/23 - Focus should return to the grid if value is invalid
    If Cancel Then
        frmMain.grdMonPlan.SetFocus
    End If
    
Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'**************************************************************
'   DESCRIPTION
'       Sub called when user clicks on the popup menu
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Public Sub mnuMonitorPlanDelete_Click_Handler()

    Const FUNCTION_NAME = "mnuMonitorPlanDelete_Click_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    DeleteEntry
    
Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'**************************************************************
'   DESCRIPTION
'       Function that gets the first area name in the list and
'       sets it as default
'
'   RETURNS - the default area name
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Private Function fGetDefArea()
    Const FUNCTION_NAME = "fGetDefArea"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'get the first area name
    rsAreaName.MoveFirst
    fGetDefArea = rsAreaName("id")
    rsAreaName.MoveFirst
    
Exit Function
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Function

'**************************************************************
'   DESCRIPTION
'       Function that sets the default Monitor Point Name
'
'   PARAMETERS
'           Row - Row number of cell being edited

'   RETURNS - the default Monitor Point Name
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Private Function fGetDefMonPt(Row)
    Const FUNCTION_NAME = "fGetDefMonPt"
    
    Dim sMonPtName As String, sExist As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'set default monitor point name to 'New Mon. Pt. <row number>'
    sMonPtName = LoadResString(IDS_DEFAULT_MONITOR_POINT_NAME) & " " & Row
    'call function to check if default name set already exists
    sExist = fIfExists(sMonPtName)
    
    If sExist = "False" Then
        'set the default name
        fGetDefMonPt = sMonPtName
    Else
        'append row number again to the default name to create a unique name
        'and check again if existing already
        Do Until sExist = "False"
            sMonPtName = sMonPtName & Row
            sExist = fIfExists(sMonPtName)
        Loop
        'set the default name
        fGetDefMonPt = sMonPtName
    End If
    
Exit Function
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function


'**************************************************************
'   DESCRIPTION
'       Function which checks if a monitor point name exists
'
'   PARAMETERS
'           Name - monitor point name to be checked
'
'   RETURNS - True if name exists and False if not
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Private Function fIfExists(Name)
    Const FUNCTION_NAME = "fIfExist"
    
    Dim sSQLDefMonPt As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'check if monitor point name already exists
    sSQLDefMonPt = "Select name from MonitorPoint where name = '" & Name & "' and deleted <> True"
    rsDefMonPt.Open sSQLDefMonPt, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    If rsDefMonPt.EOF Then
        fIfExists = "False"
    Else
        fIfExists = "True"
    End If
    
    rsDefMonPt.Close
    Set rsDefMonPt = Nothing
    
Exit Function
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function


'**************************************************************
'   DESCRIPTION
'       Sub called to delete an entry in the monitoring
'       plan grid
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'**************************************************************
Private Sub DeleteEntry()
    Const FUNCTION_NAME = "DeleteEntry"
    
    Dim sDelete As String, sSQLDelete As String
    Dim iConfirm As Integer, sID As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    With frmMain
    
        If .grdMonPlan.Cell(flexcpData, .grdMonPlan.Row, 0) = -1 Then
            DeleteEmptyRow
            Exit Sub
        End If
        
        'JNM 11/04 - moved the code that asks user if he wants to continue the delete
        'to the part after the monitor point has been evaluated for monitoring data
        'issue confirm message if user really wants to delete the selection
        'iConfirm = MsgBox(LoadResString(IDS_MSG_DELETE_MONITOR_POINT), vbOKCancel + vbDefaultButton2)
        
        'If iConfirm = 1 Then
            'move the record set to the selected row number
            rsMonitorPoint.MoveFirst
            rsMonitorPoint.Move (.grdMonPlan.Cell(flexcpData, .grdMonPlan.Row, 0) - 1)
             
            'get the monitor point id from the data in grid
            sID = .grdMonPlan.Cell(flexcpData, .grdMonPlan.Row, 2)
            
            'check if monitor point name is used in input data page
            sDelete = IsPresentMonitoringData(rsMonitorPoint("id"))
            
            Select Case sDelete
            Case "NO"
                'update record set with deleted entry
                
                'JNM 11/04 - transferred the evaluation for iConfirm here
                iConfirm = MsgBox(LoadResString(IDS_MSG_DELETE_MONITOR_POINT), vbOKCancel + vbDefaultButton2)
        
                If iConfirm = 1 Then
                    rsMonitorPoint("deleted") = "TRUE"
                    rsMonitorPoint("row") = 0
                    rsMonitorPoint.Update
                
                    'reset the row numbers of other entries and commit change in database
                    sSQLDelete = "Update MonitorPoint set row = row - 1 where row > " & .grdMonPlan.Cell(flexcpText, .grdMonPlan.Row, 0)
                    rsDelete.Open sSQLDelete, gdbConn, adOpenForwardOnly, adLockReadOnly
                
                    rsAreaName.Close
                    Set rsAreaName = Nothing
                
                    rsMonitorPoint.Close
                    Set rsMonitorPoint = Nothing
                
                    'recreate record set
                    CreateAreaRS
                    CreateMonPointRS
            
                    'repaint grid
                    PopulateGridMonPlan
                End If
                
            Case "YES"
                'issue warning that entry can not be deleted because used in input data
                'page
                MsgBox (LoadResString(IDS_MSG_NODELETE_MONITORPOINT)), vbExclamation
            End Select
        'End If
    
    End With
    
   
Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


'**************************************************************
'   DESCRIPTION
'       Sub called to delete an entry in the monitoring
'       plan grid if the chosen entry is currently empty
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/12/2002       LN Gosiengfiao    Initial Revision
'  10/16/2002      JN Manyaga        Changed the part wherein rsMonitorPoint
'                                   is tested if it is empty or not. Instead of
'                                   testing whether the pointer is NOT at EOF, this
'                                   was changed to test at both EOF and BOF.
'**************************************************************
Private Sub DeleteEmptyRow()
    Const FUNCTION_NAME = "DeleteEmptyRow"
    
    Dim sDelete As String, sSQLDelete As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'reset the row numbers of other entries and commit change in database
    sSQLDelete = "Update MonitorPoint set row = row - 1 where row > " & frmMain.grdMonPlan.Cell(flexcpText, frmMain.grdMonPlan.Row, 0)
    rsDelete.Open sSQLDelete, gdbConn, adOpenForwardOnly, adLockReadOnly
            
    rsAreaName.Close
    Set rsAreaName = Nothing
            
    'recreate area name record set
    CreateAreaRS
        
    'repaint grid
    If rsMonitorPoint.EOF And rsMonitorPoint.BOF Then
        'Do nothing since recordset is empty
    Else
        rsMonitorPoint.MoveFirst
    End If
    PopulateGridMonPlan
    
Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub

