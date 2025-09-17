Attribute VB_Name = "modIntroHistory"
' JETT HAS THOUROUgHLY REVIEWED AND OPTIMIZED THIS MODULE
'*********************************************************************************************
'   DESCRIPTION
'       Functions and subroutines used in the IntroductionHistory page
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'  9/14/2002       Jett Gamboa       Removed checking for intro lines
'                                       we now do this in the main program
' 10/15/2002       LN Gosiengfiao    For AT defect correction,
'                                       inserted a BeforeEdit function.
'                                       Modified some settings to match
'                                       the over all look and feel of the
'                                       other tabs.  Inserted code to use a
'                                       default value for the Introduction
'                                       Line Name if empty
' 10/17/2002       LN Gosiengfiao    Added a form_keyup function and startedit
'                                       function.
'************************************************************************************
Option Explicit

'Constants used in Introduction History page
Private Const MODULE_NAME = "IntroductionHistory"
Private Const GRID_ROWS = 101
Private Const GRID_COLOR = "&H80000008"
Private Const BACK_COLOR_ALTERNATE = "&H80000016"
Private Const READ_ONLY_COLOR = "&H8000000B"

Const LINENAME_COLSIZE = 4260
Const AMTINTRO_COLSIZE = 1185
Const TIMEINTRO_COLSIZE = 1335
Const NOTE_COLSIZE = 4305

Private Const SUBTOTAL_AREA_COLSIZE = 3260
Private Const SUBTOTAL_POUNDS_INTRO_COLSIZE = 2450


'Global variables for the database and record set connections
Dim rsArea As ADODB.Recordset
Dim rsIntroHist As ADODB.Recordset

'Global variables used in Introduction History page
Dim MyGrid As New clsGridWrapper
Dim MyOtherGrid As New clsGridWrapper
Dim dGrandTotal As Double
Dim sILNAreaList As String


'*****************************************************************
'   DESCRIPTION
'       Subroutine run upon load of Introduction History form
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************

Public Sub Load_IntroHistory()
    Const FUNCTION_NAME = "Load_IntroHistory"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'set global variable to change tabs initially to true
    gbCanChangeTabs = True
        
    With frmMain
        'set the label caption and settings
        .lblFumIntro.Caption = LoadResString(IDS_FUMIGANT_INTRODUCED)
        .lblFumIntro.FontBold = True
        .lblSubTot.Caption = LoadResString(IDS_AREA_SUBTOTALS)
        .lblSubTot.FontBold = True
        .lblGrandTot.Caption = LoadResString(IDS_GRAND_TOTAL)
        .lblGrandTot.FontBold = True
        .cmdNextStep5.Caption = LoadResString(IDS_BTN_NEXTSTEP)
        .txtGrandTot.ToolTipText = LoadResString(IDS_TOOLTIP_GRAND_TOTAL)
    End With
    
    'initialize and set grid properties
    PaintGrid
    
    'check the database if there are Areas existing for the job
    CreateAreaRS
    
    'create the record set for the Intro History data
    CreateIntroHistRS
    DoStuffToPaintGrid

Exit Sub

ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine to create a record set containing the area
'       and shooting line data
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************

Private Sub CreateAreaRS()
    Const FUNCTION_NAME = "CreateAreaRS"
    
    Dim sSQLArea
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    Set rsArea = New ADODB.Recordset
    
    'assign the sql call to get area names from database and execute the call to
    'create the record set
    sSQLArea = "Select JobArea.id as area_id, JobArea.area_name, ShootingLine.name, ShootingLine.id as line_id from JobArea , ShootingLine where JobArea.deleted = false and ShootingLine.deleted = false and Shootingline.jobarea_fkey = JobArea.id"
    rsArea.Open sSQLArea, gdbConn, adOpenForwardOnly, adLockReadOnly
    
Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine to create the record set containing the
'       introduction history information
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Private Sub CreateIntroHistRS()
    Const FUNCTION_NAME = "CreateIntroHistRS"
    
    Dim sSQLIntroHist
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    Set rsIntroHist = New ADODB.Recordset
       
    'get introduction history information from the database and create a record set
    sSQLIntroHist = "SELECT id, shootingline_fkey, jobarea_fkey, amt_introduced, time_introduced, note, row, deleted from IntroHistory where deleted = false order by row"
    rsIntroHist.Open sSQLIntroHist, gdbConn, adOpenStatic, adLockOptimistic, adCmdText

Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'*****************************************************************
'   DESCRIPTION
'       Function to create the Introduction Line Name (Area)
'       combo list choices
'
'   RETURNS - pipe delimited string list of the Area and Shooting
'               Line names
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Private Function CreateAreaList()
    Const FUNCTION_NAME = "CreateAreaList"
    
    Dim sErrorDesc As String, sPipeList As String
    Dim iCnt As Integer
    
    On Error GoTo ErrorHandler
        
    iCnt = 0
    
    If Not (rsArea.EOF And rsArea.BOF) Then
        rsArea.MoveFirst
        iCnt = 0
        
        'work with the Area record set and create the pipe delimited
        'string for the combo list
        Do While Not rsArea.EOF
            If iCnt <> 0 Then
                sPipeList = sPipeList & "|"
            End If
    
            sPipeList = sPipeList & "#" & rsArea("area_id") & "/" & rsArea("line_id") & ";" & rsArea("name") & " (" & rsArea("area_name") & ")"
            rsArea.MoveNext
            iCnt = iCnt + 1
        Loop
        
    End If
    
    CreateAreaList = sPipeList
    
Exit Function

ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Function


'*****************************************************************
'   DESCRIPTION
'       Subroutine to list the introduction history data into
'       the grid
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Private Sub DisplayGridData()
    Const FUNCTION_NAME = "DisplayGridData"
    
    Dim sSQLIntroHist As String, sILNArea As String, sUnit As String
    Dim dConverted As Double
    Dim iCtr As Integer, iPopulate As Integer, iCnt As Integer
    Dim sErrorDesc As String
        
    On Error GoTo ErrorHandler
    
    With frmMain
    
        'clear out contents of grid
        .grdFumIntro.Cell(flexcpText, 1, 0, GRID_ROWS - 1, 3) = ""
        .grdFumIntro.Cell(flexcpData, 1, 0, GRID_ROWS - 1, 3) = ""
        
        'assign the hidden column to contain the row numbers
        For iCnt = 1 To (GRID_ROWS - 1)
            .grdFumIntro.Cell(flexcpText, iCnt, 4) = iCnt
            .grdFumIntro.Cell(flexcpData, iCnt, 4) = -1
        Next
        
        'initiate the grand total variable
        dGrandTotal = 0
        
        If Not (rsIntroHist.EOF And rsIntroHist.BOF) Then
            iCtr = 1
            iPopulate = 0
        
            rsIntroHist.MoveFirst
    
             'loop through rsIntroHist and move data into the grid
             Do While Not rsIntroHist.EOF
                If Str(frmMain.grdFumIntro.Cell(flexcpText, iCtr, 4)) = Str(rsIntroHist("row")) Then
                    iPopulate = iPopulate + 1
                    .grdFumIntro.Cell(flexcpData, iCtr, 4) = iPopulate
                    .grdFumIntro.Cell(flexcpText, iCtr, 0) = rsIntroHist("jobarea_fkey") & "/" & rsIntroHist("shootingline_fkey")
                    
                    If gbIsMetric = False Then
                        UnitConvert modeDisplay, unitWEIGHT, rsIntroHist("amt_introduced"), dConverted, sUnit
                        .grdFumIntro.Cell(flexcpData, iCtr, 1) = dConverted
                        .grdFumIntro.Cell(flexcpText, iCtr, 1) = Round(dConverted, 0)
                        dGrandTotal = dGrandTotal + dConverted
                    Else
                        .grdFumIntro.Cell(flexcpData, iCtr, 1) = rsIntroHist("amt_introduced")
                        .grdFumIntro.Cell(flexcpText, iCtr, 1) = Round(rsIntroHist("amt_introduced"), 0)
                        dGrandTotal = dGrandTotal + rsIntroHist("amt_introduced")
                    End If
                    
                    .grdFumIntro.Cell(flexcpText, iCtr, 2) = rsIntroHist("time_introduced")
                    .grdFumIntro.Cell(flexcpText, iCtr, 3) = rsIntroHist("note")
                    .grdFumIntro.Cell(flexcpData, iCtr, 3) = rsIntroHist("id")
                    rsIntroHist.MoveNext
                
                End If
                
                If Not rsIntroHist.EOF Then
                    If rsIntroHist("row") = 0 Then
                        rsIntroHist.MoveNext
                    End If
                End If
                
                iCtr = iCtr + 1
            Loop
        
        End If
    End With

Exit Sub

ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine to call the different functions needed to
'       paint the grid
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Private Sub DoStuffToPaintGrid()
    Const FUNCTION_NAME = "DoStuffToPaintGrid"

    Dim sErrorDesc As String
        
    On Error GoTo ErrorHandler

    'call the subroutine CreateAreaList, which handles the creation of the pipe-delimited string with the list of the available lines with Areas
    sILNAreaList = CreateAreaList()
        
    'assign the drop down list to column 0
    frmMain.grdFumIntro.ColComboList(0) = sILNAreaList
        
    'call the DisplayGridData subroutine to show history data
    DisplayGridData
        
    'call the SetTotals subroutine to show the grand and sub totals
    SetSubTotals
        
    'set the grand total textbox value
    If gbIsMetric Then
        frmMain.txtGrandTot.Text = Round(dGrandTotal, 0)
    Else
        frmMain.txtGrandTot.Text = Round(dGrandTotal, 0)
    End If
    
Exit Sub

ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine to initialize and set the grid properties
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
' 10/22/2002       JN Manyaga        Changed type of colTimeIntro to text.
'*****************************************************************
Private Sub PaintGrid()
    Const FUNCTION_NAME = "PaintGrid"

    Dim sErrorDesc As String
        
    On Error GoTo ErrorHandler
       
    'initialize the grid
    MyGrid.Initialize "FumigantIntroduced", frmMain.grdFumIntro
    MyOtherGrid.Initialize "AreaSubTotals", frmMain.grdAreaSubTot
    
    'add columns to the grid
    MyGrid.AddColumn "colIntroLineName", LoadResString(IDS_INTRODUCTION_LINE_NAME_AREA), colTypeCombo, flexAlignGeneral
    
    If gbIsMetric = True Then
        MyGrid.AddColumn "colAmtIntroduced", LoadResString(IDS_KG_INTRODUCED), colTypeText, flexAlignGeneral
    Else
        MyGrid.AddColumn "colAmtIntroduced", LoadResString(IDS_LBS_INTRODUCED), colTypeText, flexAlignGeneral
    End If
    
    'JNM 10/21/02 - Changed type from colTypeTime to colTypeText
    MyGrid.AddColumn "colTimeIntro", LoadResString(IDS_TIME_INTRODUCED), colTypeText, flexAlignGeneral
    MyGrid.AddColumn "colNote", LoadResString(IDS_INTROHIST_NOTE), colTypeText, flexAlignGeneral
    MyGrid.AddColumn "colRow", "", colTypeText, flexAlignGeneral
    
    MyOtherGrid.AddColumn "colArea", LoadResString(IDS_AREA), colTypeText, flexAlignGeneral
    
    If gbIsMetric = True Then
        MyOtherGrid.AddColumn "colPoundsIntro", LoadResString(IDS_KILOGRAMS_INTRODUCED), colTypeText, flexAlignGeneral
    Else
        MyOtherGrid.AddColumn "colPoundsIntro", LoadResString(IDS_POUNDS_INTRODUCED), colTypeText, flexAlignGeneral
    End If
        
    With frmMain.grdFumIntro
        .ColWidth(0) = LINENAME_COLSIZE
        .ColWidth(1) = AMTINTRO_COLSIZE
        .ColWidth(2) = TIMEINTRO_COLSIZE
        .ColWidth(3) = NOTE_COLSIZE
    End With
        
    frmMain.grdAreaSubTot.ColWidth(0) = SUBTOTAL_AREA_COLSIZE
    frmMain.grdAreaSubTot.ColWidth(1) = SUBTOTAL_POUNDS_INTRO_COLSIZE
    
    'load grid settings
    MyGrid.LoadGridSettings GRID_SETTINGS_PATH
    MyOtherGrid.LoadGridSettings GRID_SETTINGS_PATH
    
    frmMain.grdAreaSubTot.GridColor = GRID_COLOR
    frmMain.grdAreaSubTot.BackColor = READ_ONLY_COLOR
    frmMain.grdAreaSubTot.ToolTipText = LoadResString(IDS_TOOLTIP_SUB_TOTAL)
        
    frmMain.grdFumIntro.Rows = GRID_ROWS
    frmMain.grdFumIntro.GridColor = GRID_COLOR
    'grdFumIntro.BackColorAlternate = BACK_COLOR_ALTERNATE
    frmMain.grdFumIntro.ColHidden(4) = True
       
Exit Sub

ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine called when Introduction History form is
'       unloaded
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Public Sub Unload_IntroHistory()
    Const FUNCTION_NAME = "Unload_IntroHistory"
    
    Dim sCheck As Boolean
    Dim sErrorDesc As String
    Dim dAmt As Double
    Dim sAmt As String
    
    On Error GoTo ErrorHandler
    
    'JNM 10/22 - Removed the grid validation below since the ValidateEdit
    '   event of the grid validates the grid already.
    'check if user is exiting, if true, then unload form without validation
    'If gbIsExiting = True Then
    '    Cancel = vbFalse
    
    'Else
        
    '   If grdFumIntro.Row = 0 Then
    '        Cancel = vbFalse
        
        'otherwise perform validation before unload
    '    Else
            
            'initialize sCheck to true
    '        sCheck = True
            'JNM 10/22 - set-up sAmt
    '        sAmt = Trim(grdFumIntro.Cell(flexcpText, grdFumIntro.Row, 1))
            
    '        If grdFumIntro.Cell(flexcpText, grdFumIntro.Row, 1) <> "" Then
    '            sCheck = IsNumeric(sAmt)
    '        End If
            
    '        If sCheck = False Then
    '            MsgBox (LoadResString(IDS_MSG_AMT_INTRO_NUMERIC))
    '            gbCanChangeTabs = False
    '            Cancel = vbTrue
    '            Exit Sub
    '        Else
                'JNM 10/22 - If we are in English units, Amount Intro should
                'be a whole number
    '            dAmt = Round(CDbl(sAmt), 0)
    '            If Not gbIsMetric And (dAmt - CDbl(sAmt)) <> 0 Then
    '                MsgBox LoadResString(IDS_MSG_AMT_WHOLE_NUM)
    '                gbCanChangeTabs = False
    '                Cancel = vbTrue
    '                Exit Sub
    '            Else
    '                If grdFumIntro.Cell(flexcpText, grdFumIntro.Row, 1) < 0 Then
    '                    MsgBox (LoadResString(IDS_MSG_AMT_NEGATIVE))
    '                    gbCanChangeTabs = False
    '                    Cancel = vbTrue
    '                    Exit Sub
    '                Else
    '                    gbCanChangeTabs = True
    '                    Cancel = vbFalse
    '                End If
    '            End If
    '        End If
    '
    '    End If
    
    'End If
        
    'save settings to grid and unload form
    MyGrid.SaveGridSettings GRID_SETTINGS_PATH
    MyOtherGrid.SaveGridSettings GRID_SETTINGS_PATH
    
    'close recordsets
    rsArea.Close
    Set rsArea = Nothing
    
    rsIntroHist.Close
    Set rsIntroHist = Nothing
        
Exit Sub

ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine called after user resizes the grdFumIntro
'       grid.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/03/2002      LN Gosiengfiao    Initial Revision
'*****************************************************************
Public Sub grdAreaSubTot_AfterUserResize_Handler(ByVal Row As Long, ByVal Col As Long)
    Const FUNCTION_NAME = "grdAreaSubTot_AfterUserResize_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'call the grid wrapper function to save new settings to registry
    MyOtherGrid.SaveGridSettings GRID_SETTINGS_PATH
    
Exit Sub

ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine called before a user edits a cell in the
'       grdAreaSubTot grid.  It sets the grid to be read only.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Public Sub grdAreaSubTot_BeforeEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Const FUNCTION_NAME = "grdAreaSubTot_BeforeEdit_Handler"

    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'set grid to read only
    Cancel = True

Exit Sub
    
ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


'*****************************************************************
'   DESCRIPTION
'       Created this subroutine so that grid can be refreshed after
'       editing it.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/22/2002       JN Manyaga    Initial Revision
'*****************************************************************
Public Sub grdFumIntro_AfterEdit_Handler(ByVal Row As Long, ByVal Col As Long)
    Const FUNCTION_NAME = "grdFumIntro_AfterEdit_Handler"
  
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    rsIntroHist.Close
    Set rsIntroHist = Nothing

    rsArea.Close
    Set rsArea = Nothing
    
    'rebuild the record sets for area name and monitor point data and
    'repaint the grid
    CreateAreaRS
    CreateIntroHistRS
    DoStuffToPaintGrid
        
Exit Sub

ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine called after user resizes the grdFumIntro
'       grid.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Public Sub grdFumIntro_AfterUserResize_Handler(ByVal Row As Long, ByVal Col As Long)

    Const FUNCTION_NAME = "grdFumIntro_AfterUserResize_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'call the grid wrapper function to save new settings to registry
    MyGrid.SaveGridSettings GRID_SETTINGS_PATH
    
Exit Sub

ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


Public Sub grdFumIntro_BeforeEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Const FUNCTION_NAME = "grdFumIntro_BeforeEdit_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'assign maximum length to note column
    If frmMain.grdFumIntro.Col = 3 Then
        frmMain.grdFumIntro.EditMaxLength = 500
    End If
    
Exit Sub

ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine called when user hits the delete button
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/3/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Public Sub grdFumIntro_KeyDown_Handler(KeyCode As Integer, Shift As Integer)
    Const FUNCTION_NAME = "grdFumIntro_KeyDown_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'checks if user hits the delete button
    If KeyCode = 46 Then
        DeleteEntry
    End If
    
Exit Sub

ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine called when user right clicks on mouse
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Public Sub grdFumIntro_MouseDown_Handler(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNCTION_NAME = "grdFumIntro_MouseDown_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    With frmMain
        'JNM 11/4 - Added code that would select the cell that was clicked.
        'on right click show pop up menu
        If Button = vbRightButton And .grdFumIntro.MouseRow > 0 Then
            .grdFumIntro.Select .grdFumIntro.MouseRow, .grdFumIntro.MouseCol
            frmMain.PopupMenu frmMain.mnuIntroHistoryContext
        End If
    End With
    
Exit Sub

ErrorHandler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine which sets the tool tip texts when mouse is
'       moved around the grid
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Public Sub grdFumIntro_MouseMove_Handler(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNCTION_NAME = "grdFumIntro_MouseMove_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'show tool tip text for the different columns
    Select Case frmMain.grdFumIntro.MouseCol
    Case 0
        frmMain.grdFumIntro.ToolTipText = LoadResString(IDS_TOOLTIP_ILN_AREA)
    Case 1
        frmMain.grdFumIntro.ToolTipText = LoadResString(IDS_TOOLTIP_AMT_INTRODUCED)
    Case 2
        frmMain.grdFumIntro.ToolTipText = LoadResString(IDS_TOOLTIP_TIME_INTRODUCED)
    Case 3
        frmMain.grdFumIntro.ToolTipText = LoadResString(IDS_TOOLTIP_NOTE)
    End Select
    
Exit Sub
    
ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine called to set the data on the grdAreaSubTot
'       grid.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Private Sub SetSubTotals()
    Const FUNCTION_NAME = "SetTotals"
    
    Dim sSQLSubTotal As String, sUnit As String
    Dim dGrandTotal As Double, dConverted As Double
    Dim iRow As Integer, iCnt As Integer
    Dim sErrorDesc As String
    
    Dim rsSubTotal As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    Set rsSubTotal = New ADODB.Recordset

    'set record set containing the distint areas and subtotals
    '**** NOTE:  This SQL call may not work against other databases, applicable only to Access databases
    sSQLSubTotal = "SELECT JobArea.area_name, Sum(IntroHistory.amt_introduced) AS subtotal, IntroHistory.deleted FROM JobArea INNER JOIN IntroHistory ON JobArea.id = IntroHistory.jobarea_fkey GROUP BY JobArea.area_name, IntroHistory.deleted Having (((IntroHistory.deleted) = False))"
    rsSubTotal.Open sSQLSubTotal, gdbConn, adOpenStatic, adLockReadOnly

    iCnt = 1
    iRow = rsSubTotal.RecordCount + 1
    
    With frmMain
        'set the grdAreaSubTot grid number to the number of distinct areas
        .grdAreaSubTot.Rows = iRow
        
        'loop through record set and populate data for grid
        Do While Not rsSubTotal.EOF
            .grdAreaSubTot.Cell(flexcpText, iCnt, 0) = rsSubTotal("area_name")
            If gbIsMetric = False Then
                UnitConvert modeDisplay, unitWEIGHT, rsSubTotal("subtotal"), dConverted, sUnit
                .grdAreaSubTot.Cell(flexcpText, iCnt, 1) = Round(dConverted, 0)
            Else
                .grdAreaSubTot.Cell(flexcpText, iCnt, 1) = Round(rsSubTotal("subtotal"), 0)
            End If
            rsSubTotal.MoveNext
            iCnt = iCnt + 1
        Loop
    End With
    
    rsSubTotal.Close
    Set rsSubTotal = Nothing
    
Exit Sub
    
ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'*****************************************************************
'   DESCRIPTION
'       Run before edit to set default for Introduction Line Name
'
'   REVISION LOG
'       DATE       NAME              CHANGE
' 10/17/2002       LN Gosiengfiao    Initial Revision
' 10/23/2002       JN Manyaga        Set the maximum length for Time Intro column
'*****************************************************************
Public Sub grdFumIntro_StartEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Const FUNCTION_NAME = "grdFumIntro_StartEdit"
    
    Dim sDefault As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'check if column being edited is introduction line name column
    If Col = 0 Then
        
        'set the default value for name if empty
        If frmMain.grdFumIntro.Cell(flexcpText, Row, Col) = "" Then
            'get default value for Introduction Line Name, use first
            'entry in record set
            rsArea.MoveFirst
            sDefault = rsArea("area_id") & "/" & rsArea("line_id")
            rsArea.MoveFirst
            
            frmMain.grdFumIntro.Cell(flexcpText, Row, Col) = sDefault
        End If
    
    End If
    
    'Set the maximum number of characters
    If Col = 2 Then
        frmMain.grdFumIntro.EditMaxLength = 50
    Else
        frmMain.grdFumIntro.EditMaxLength = 0   'no limit
    End If
    
    'The line below would ensure that the tab does not change if the user enter enters
    'an invalid value and then clicks on another tab
    ReturnCurrTab TAB_INTRO_HISTORY
    
Exit Sub
    
ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine called after user edits a cell in the
'       grdFumIntro grid
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Public Sub grdFumIntro_ValidateEdit_Handler(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Const FUNCTION_NAME = "grdFumIntro_ValidateEdit_Handler"
    
    Dim sValid As Boolean
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
        SaveIntroHist Row, Col
    End Select
    
    'JNM 10/23 - Focus should return to the grid if value is invalid
    If Cancel Then
        frmMain.grdFumIntro.SetFocus
    End If
        
Exit Sub
    
ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine called when user clicks on delete choice in
'       popup menu
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Public Sub mnuIntroHistoryDelete_Click_Handler()
    Const FUNCTION_NAME = "mnuIntroHistoryDelete_Click_Handler"
    
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler

    DeleteEntry
    
Exit Sub
    
ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine called to validate data entries in grid
'       grdFumIntro
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
' 10/21/2002       JN Manyaga        Changed the validation process
'*****************************************************************
Private Function ValidateField(Row, Col)
    Const FUNCTION_NAME = "ValidateFields"
    
    Dim bCheck As Boolean, sDefILN As String, sAmt As String
    Dim sErrorDesc As String
    Dim dAmt As Double
        
    On Error GoTo ErrorHandler
    
    'initialize sCheck to true
    bCheck = True
    
    'JNM 10/22 - Removed the code below. The default should be set in the
    ' SaveIntroHist subroutine only.
    'If Col <> 0 Then
        'check if the Introduction Line Name column is empty
    '    If grdFumIntro.Cell(flexcpText, Row, 0) = "" Then
            'get default value for Introduction Line Name, use first
            'entry in record set
    '        rsArea.MoveFirst
    '        sDefILN = rsArea("area_id") & "/" & rsArea("line_id")
    '        rsArea.MoveFirst
        
    '        grdFumIntro.Cell(flexcpText, Row, 0) = sDefILN
    '    End If
    'End If
    
    'JNM 10/22 - Changed the validation code so that new records which
    'are created by entering a value for Amount Introduced are also validated.
    If Col = 1 Then
        sAmt = Trim(frmMain.grdFumIntro.EditText)
        If Not IsNumeric(sAmt) Then
            MsgBox (LoadResString(IDS_MSG_AMT_INTRO_NUMERIC)), vbExclamation
            gbCanChangeTabs = False
            bCheck = False
        Else
            'JNM 10/22 - If we are in English units, Amount Intro should
            'be a whole number
            dAmt = Round(CDbl(sAmt), 0)
            If Not gbIsMetric And (dAmt - CDbl(sAmt)) <> 0 Then
                MsgBox LoadResString(IDS_MSG_AMT_WHOLE_NUM), vbExclamation
                bCheck = False
                gbCanChangeTabs = False
            Else
                If sAmt < 0 Then
                    MsgBox (LoadResString(IDS_MSG_AMT_NEGATIVE)), vbExclamation
                    gbCanChangeTabs = False
                    bCheck = False
                Else
                    gbCanChangeTabs = True
                End If
            End If
        End If
    End If
        
    ValidateField = bCheck
    
Exit Function
    
ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Function


'*****************************************************************
'   DESCRIPTION
'       Subroutine called to save modifications on the Introduction
'       History data
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
' 10/22/2002       JN Manyaga        Added the code that would set the
'                                    default value of Intro Line (Area)
'                                    when this is blank.
'*****************************************************************
Private Sub SaveIntroHist(Row, Col)
    Const FUNCTION_NAME = "SaveIntroHist"
    
    Dim sGUID As String, sILN, sFkeys() As String, sUnit As String, sRowChanged As String
    Dim dAmount As Double, dConverted As Double
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    'set global variable to true because form fields have changed
    gbHasChanged = True

    'check if data in hidden column is -1 which indicates that user is inserting
    'a new row
    If frmMain.grdFumIntro.Cell(flexcpData, Row, 4) = -1 Then
        'get a primary key for the new row
        sGUID = GetGUID()
        
        'add new row in record set
        rsIntroHist.AddNew
        
        'assign record set values from the grid
        rsIntroHist("id") = sGUID
        
        rsIntroHist("row") = frmMain.grdFumIntro.Cell(flexcpText, Row, 4)
        rsIntroHist("deleted") = "False"
        
        'modified cell is area name
        If Col = 0 Then
            sILN = frmMain.grdFumIntro.ComboData
        Else
            'get the job area and shooting line id's from the column 0
            sILN = frmMain.grdFumIntro.Cell(flexcpText, Row, 0)
        End If
        
        If sILN = "" Then
            'JNM 10/22 - the fkeys should not be null
            rsArea.MoveFirst
            rsIntroHist("jobarea_fkey") = rsArea("area_id")
            rsIntroHist("shootingline_fkey") = rsArea("line_id")
        Else
            sFkeys = Split(sILN, "/")
            rsIntroHist("jobarea_fkey") = sFkeys(0)
            rsIntroHist("shootingline_fkey") = sFkeys(1)
        End If
        
        'modified cell is amount introduced
        If Col = 1 Then
            'if cell is empty set value to 0
            If frmMain.grdFumIntro.EditText = "" Then
                dAmount = 0#
            Else
                dAmount = frmMain.grdFumIntro.EditText
            End If
        Else
            If frmMain.grdFumIntro.Cell(flexcpText, Row, 1) = "" Then
                dAmount = 0#
            Else
                dAmount = frmMain.grdFumIntro.Cell(flexcpText, Row, 1)
            End If
        End If
        
        'call convert function if not in Metric mode
        If gbIsMetric = False Then
            UnitConvert modeSave, unitWEIGHT, dAmount, dConverted, sUnit
            rsIntroHist("amt_introduced") = dConverted
        Else
            rsIntroHist("amt_introduced") = dAmount
        End If
        
        'modified cell is time introduced
        If Col = 2 Then
            'set the value to Null if cell is empty
            If frmMain.grdFumIntro.EditText = "" Then
                rsIntroHist("time_introduced") = Null
            Else
                rsIntroHist("time_introduced") = frmMain.grdFumIntro.EditText
            End If
        Else
            If frmMain.grdFumIntro.Cell(flexcpText, Row, 2) = "" Then
                rsIntroHist("time_introduced") = Null
            Else
                rsIntroHist("time_introduced") = frmMain.grdFumIntro.Cell(flexcpText, Row, 2)
            End If
        End If
        
        'modified cell is note
        If Col = 3 Then
            rsIntroHist("note") = frmMain.grdFumIntro.EditText
        Else
            rsIntroHist("note") = frmMain.grdFumIntro.Cell(flexcpText, Row, 3)
        End If
        
        'commit update to recordset
        rsIntroHist.Update
        
    Else
        'if user modified an existing row, get the row number from hidden column
        sRowChanged = frmMain.grdFumIntro.Cell(flexcpData, Row, 4)
        
        'move to the currently modified row in the record set and assign modified
        'values from the grid
        rsIntroHist.MoveFirst
        rsIntroHist.Move (sRowChanged - 1)
        
        'modified cell is area name
        If Col = 0 Then
            sFkeys = Split(frmMain.grdFumIntro.ComboData, "/")
            rsIntroHist("jobarea_fkey") = sFkeys(0)
            rsIntroHist("shootingline_fkey") = sFkeys(1)
        Else
            'get the job area and shooting line id's from the column 0
            sFkeys = Split(frmMain.grdFumIntro.Cell(flexcpText, Row, 0), "/")
            rsIntroHist("jobarea_fkey") = sFkeys(0)
            rsIntroHist("shootingline_fkey") = sFkeys(1)
        End If
        
        'modified cell is amount introduced
        If Col = 1 Then
            'set value to 0 if cell is empty
            If frmMain.grdFumIntro.EditText = "" Then
                dAmount = 0#
            Else
                dAmount = frmMain.grdFumIntro.EditText
            End If
        Else
            If frmMain.grdFumIntro.Cell(flexcpText, Row, 1) = "" Then
                dAmount = 0#
            Else
                dAmount = frmMain.grdFumIntro.Cell(flexcpData, Row, 1)
            End If
        End If
                
        'convert amount introduced if mode is not metric
        If gbIsMetric = False Then
            UnitConvert modeSave, unitWEIGHT, dAmount, dConverted, sUnit
            rsIntroHist("amt_introduced") = dConverted
        Else
            rsIntroHist("amt_introduced") = dAmount
        End If
        
        'modified cell is time introduced
        If Col = 2 Then
            'set value to Null if cell is empty
            If frmMain.grdFumIntro.EditText = "" Then
                rsIntroHist("time_introduced") = Null
            Else
                rsIntroHist("time_introduced") = frmMain.grdFumIntro.EditText
            End If
        Else
            If frmMain.grdFumIntro.Cell(flexcpText, Row, 2) = "" Then
                rsIntroHist("time_introduced") = Null
            Else
                rsIntroHist("time_introduced") = frmMain.grdFumIntro.Cell(flexcpText, Row, 2)
            End If
        End If
        
        'modified cell is note
        If Col = 3 Then
            rsIntroHist("note") = frmMain.grdFumIntro.EditText
        Else
            rsIntroHist("note") = frmMain.grdFumIntro.Cell(flexcpText, Row, 3)
        End If
        
        'commit changes to record set
        rsIntroHist.Update
    End If

    'JNM - All the refresh procedures below are transferred to the AfterEdit event so
    'that the user can have the option of cancelling the edit.
    'rsIntroHist.Close
    'Set rsIntroHist = Nothing

    'rsArea.Close
    'Set rsArea = Nothing
    
    'rebuild the record sets for area name and intro history
    'CreateAreaRS
    'CreateIntroHistRS
    
    'call subroutine to repaint grid
    'DoStuffToPaintGrid

    gbCanChangeTabs = True
    
Exit Sub
    
ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


'*****************************************************************
'   DESCRIPTION
'       Subroutine called when user clicks on delete choice in
'       popup menu or when user hits the delete button to
'       delete an entry from the introduction history grid
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/30/2002       LN Gosiengfiao    Initial Revision
'*****************************************************************
Private Sub DeleteEntry()
    Const FUNCTION_NAME = "DeleteEntry"
    
    Dim rsDelete As ADODB.Recordset
    
    Dim sSQLDelete As String, iConfirm As String, sID As String
    Dim sErrorDesc As String
    
    On Error GoTo ErrorHandler
    
    Set rsDelete = New ADODB.Recordset

    If frmMain.grdFumIntro.Cell(flexcpData, frmMain.grdFumIntro.Row, 4) = -1 Then
        DeleteEmptyRow
        Set rsDelete = Nothing
        Exit Sub
    End If

    'issue confirm message if user really wants to delete the selection
    iConfirm = MsgBox(LoadResString(IDS_MSG_DELETE_INTRO_HIST), vbOKCancel + vbDefaultButton2)
    
    If iConfirm = 1 Then
        'move the record set to the selected row number
        rsIntroHist.MoveFirst
        rsIntroHist.Move (frmMain.grdFumIntro.Cell(flexcpData, frmMain.grdFumIntro.Row, 4) - 1)
         
        'get the introduction history id from the data in grid
        sID = frmMain.grdFumIntro.Cell(flexcpData, frmMain.grdFumIntro.Row, 3)
        
        'update record set with deleted entry
        rsIntroHist("deleted") = "TRUE"
        rsIntroHist("row") = 0
        rsIntroHist.Update
            
        'reset the row numbers of other entries and commit change in database
        'TODO: This should be call to exec not the open method
        sSQLDelete = "Update IntroHistory set row = row - 1 where row > " & frmMain.grdFumIntro.Cell(flexcpValue, frmMain.grdFumIntro.Row, 4)
        rsDelete.Open sSQLDelete, gdbConn, adOpenForwardOnly, adLockReadOnly
       
        rsArea.Close
        Set rsArea = Nothing
            
        rsIntroHist.Close
        Set rsIntroHist = Nothing
        
        'recreate record set
        CreateAreaRS
        CreateIntroHistRS
        
        'call subroutine to repaint grid
        DoStuffToPaintGrid
            
    End If

    Set rsDelete = Nothing

Exit Sub
    
ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub


Private Sub DeleteEmptyRow()
    Const FUNCTION_NAME = "DeleteEmptyRow"
    
    Dim sSQLDelete As String
    Dim sErrorDesc As String
    
    Dim rsDelete As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    Set rsDelete = New ADODB.Recordset

    'reset the row numbers of other entries and commit change in database
    sSQLDelete = "UPDATE IntroHistory SET row = row - 1 WHERE row > " & frmMain.grdFumIntro.Cell(flexcpValue, frmMain.grdFumIntro.Row, 4)
    rsDelete.Open sSQLDelete, gdbConn, adOpenForwardOnly, adLockReadOnly
    
    Set rsDelete = Nothing
            
    rsArea.Close
    Set rsArea = Nothing
        
    'recreate area name record set
    CreateAreaRS
    If Not rsIntroHist.EOF Then
        rsIntroHist.MoveFirst
    End If
            
    'call subroutine to repaint grid
    DoStuffToPaintGrid
            
Exit Sub
    
ErrorHandler:
    'untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
End Sub

