Attribute VB_Name = "modGraph"
'**************************************************************
'* Filename: frmGraph.frm
'*
'* DESCRIPTION
'* -----------
'* contains the form which will contain the controls (Chartfx)
'* listboxes, buttons etc. needed for the graph tab
'*
'* Copyright (c) 2002 Accenture.  All rights reserved.
'-----------------------------------------------------
'*
'* REVISION HISTORY
'* Revision:    0
'* Author:      Gamboa, Maria Cecilia <NAI1686>
'* Date:        March 25, 2003
'**************************************************************
Option Explicit

Private Const MODULE_NAME = "frmGraph"
Private sErrorDesc As String


Public Sub cmdOptions_Click_Handler()
    'show the properties dialog box
    frmMain.ChartFX1.ShowDialog CDIALOG_OPTIONS, 0
End Sub


'**************************************************************
'   DESCRIPTION
'   Handles the on click event of the Set Properties button
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'   April 24, 2003       Jett Gamboa      opened properties form as modal
'**************************************************************
Public Sub cmdSetProperties_Click_Handler()
    On Error GoTo Error_Handler
        
    'Save the changes made in the graph (through the Chartfx property box)
    SaveMonPtSpecificSettings
    Const FUNCTION_NAME = "cmdSetProperties_Click"
    frmGraphProperties.Show vbModal
    
    Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)

End Sub


'**************************************************************
'   DESCRIPTION
'   Handles the on load event of the form
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'   May 8, 2003          Amana Martinez   Fumiguide Enhancements 2003
'**************************************************************
Public Sub Load_Graph()

    On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "Form_Load"
    gbHasChanged = True
    
    'Load the control captions
    With frmMain
        .fraFilterBy.Caption = LoadResString(IDS_GRAPH_FILTER_BY)
        .cmdSetProperties.Caption = LoadResString(IDS_GRAPH_PROPERTIES_BUTTON)
        .cmdExportGraph.Caption = LoadResString(IDS_GRAPH_EXPORT_BUTTON)
        .lstDataFilter2.ToolTipText = LoadResString(IDS_TIP_DATA_FILTER)
        .cmdOptions.Caption = LoadResString(IDS_GRAPH_OPTIONS)
        .cmdOptions.ToolTipText = LoadResString(IDS_GRAPH_TOOLTIP_GRAPH_OPTIONS)
        .cmdSetProperties.ToolTipText = LoadResString(IDS_GRAPH_TOOLTIP_SET_PROPERTIES)
        
        ' Jett 4/18/03 clear filter box (since we are loading by frames)
        .lstDataFilter2.Clear
    End With

    'Set the Chart Properties that can not be changed by the user
    With frmMain.ChartFX1
        
         'Set the Axis Title colors
         .Axis(AXIS_Y).TitleColor = RGB(0, 0, 255)
         .Axis(AXIS_X).TitleColor = RGB(0, 0, 255)
         
         'Set the default gallery used
         .Gallery = LINES
         
         'Set the grid color
         '5/8/2003 - AM Commented the code below for Fumiguide Enhancements 2003 PT Defect 8711
         '.Axis(0).GridColor = &H8000000F
         '5/8/2003 - AM Commented the code above for Fumiguide Enhancements 2003 PT Defect 8711
         '5/8/2003 - AM Added the code below for Fumiguide Enhancements 2003 PT Defect 8711
         .Axis(AXIS_Y).GridColor = &H808080
         .Axis(AXIS_X).GridColor = &H808080
         '5/8/2003 - AM Added the code above for Fumiguide Enhancements 2003 PT Defect 8711
         
         
         'allows you to determine which property pages are shown at run time
         '(in other words, which property pages will be presented to the user
         'when it presses the Properties button in the Toolbar).
         .PropPageMask = 2
         .RigClk CHART_PROPERTIESCLK, 0
         
         .Style = Not CS_GALLERY And Not CS_EDITABLE
         .AutoIncrement = True

         .Axis(AXIS_X).LabelValue = 0
         .DataEditorObj.Visible = False
         
        ' UI Settings
         .AllowResize = True
         .AllowDrag = False
         .AllowEdit = False
         .ContextMenus = True
         .ShowTips = True
         ' 5/23/03 Jett - changed this to False so that everything fits in one
         ' screen
         .Scrollable = False
        
         ' Tool Windows Settings
         .LegendBoxObj.Moveable = False
         .LegendBoxObj.Sizeable = 0
         .SerLegBoxObj.Moveable = False
         .SerLegBoxObj.Sizeable = 0
         .DataEditorObj.Moveable = False
         .DataEditorObj.Sizeable = 0
         .ToolBarObj.Moveable = False
         
         ' Tool Windows Visibility
         .ToolBar = True
         .PaletteBar = False
         .PatternBar = False
         .DataEditor = False
        
         ' Chart Type Settings
         .Chart3D = False
         'Used to specify the point type used to paint markers in the chart
         .MarkerShape = MK_NONE
             
         ' Color Settings
         .Palette = "CHART_CSHATCH"
         .Border = False
         
    
    End With
    
    Call LoadGraphSettings
    Call ListValues
    
Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)

End Sub


'**************************************************************
'   DESCRIPTION
'   Handles the unload event of the form
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub Unload_Graph()

    On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "Form_Unload"
    'All objects have been destroyed in their respective subroutines/functions
    
    'Save the Monitor Point specific properties set by the user.
        Call SaveMonPtSpecificSettings
    
    Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)

End Sub


'**************************************************************
'   DESCRIPTION
'   Handles the on click event of the Export Graph button
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub cmdExportGraph_Click_Handler()
    On Error GoTo Error_Handler
    
        Const FUNCTION_NAME = "cmdExportGraph_Click_Handler"
        Call ExportGraph
        
    Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)
    
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
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub lstDataFilter2_Click_Handler()
    Dim sSelected As String
    Dim I As Long
    Dim sSQLFilter As String
    Dim sSeriesNo As String
    Dim sSeriesYes As String

On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "lstDataFilter2_Click"
    sSelected = ""
    
    For I = 0 To frmMain.lstDataFilter2.ListCount - 1
        If frmMain.lstDataFilter2.Selected(I) Then
            If sSelected = "" Then
                sSelected = "'" & frmMain.lstDataFilter2.List(I) & "'"
                If I = 0 Then 'this means all is selected
                    sSeriesYes = ""
                Else
                    If sSeriesYes = "" Then
                        sSeriesYes = I - 1
                    Else
                        sSeriesYes = sSeriesYes & " , " & I - 1
                    End If
                End If
            Else
                sSelected = sSelected & " , " & "'" & frmMain.lstDataFilter2.List(I) & "'"
                If I = 0 Then 'this means all is selected
                    sSeriesYes = ""
                Else
                    If sSeriesYes = "" Then
                        sSeriesYes = I - 1
                    Else
                        sSeriesYes = sSeriesYes & " , " & I - 1
                    End If
                End If
            End If
        Else
            If I = 0 Then 'this means all is selected
                sSeriesNo = ""
            Else
                If sSeriesNo = "" Then
                    sSeriesNo = I - 1
                Else
                    sSeriesNo = sSeriesNo & " , " & I - 1
                End If
            End If
        End If
    Next
    
    ' TODO: jett, what's the deal with this? why not store in a variable?
    frmMain.txtSelectedY.Text = sSeriesYes
    frmMain.txtSelectedN.Text = sSeriesNo
    
    ' 2/12/04 JETT LOCALIZATION
    ' Changed the way we check if the "ALL" item was selected. Now that
    ' we are localizing, we we should not check for the word ALL but instead
    ' check if the user selected the first item on the list.
    '
    'If sSelected = "'ALL'" Or sSelected = "" Then
    If frmMain.lstDataFilter2.Selected(0) Or sSelected = "" Then
        'display all graphing points
        SetPoints gnGraphUnits
    Else
         SetPoints gnGraphUnits, sSeriesNo, sSeriesYes, sSelected
    End If

Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)

End Sub


'**************************************************************
'   DESCRIPTION
'   Populates the listbox with the Monitor Point names
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************
Private Sub ListValues()

    Dim rsMonPoint As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "ListValues"
        
    frmMain.lstDataFilter2.AddItem LoadResString(IDS_MONPOINTS_FILTER_ALL)
    
    sSQL = "SELECT name FROM MonitorPoint WHERE deleted = false"
    'Store the records in a recordset
    rsMonPoint.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    If rsMonPoint.BOF And rsMonPoint.EOF Then
        'there are no monitor points that exists for this table.
        MsgBox LoadResString(IDS_GRAPH_NO_MON_POINTS), vbExclamation
    Else
        rsMonPoint.MoveFirst
        
        While Not (rsMonPoint.EOF)
            frmMain.lstDataFilter2.AddItem rsMonPoint("name")
            rsMonPoint.MoveNext
        Wend
    End If

    'Close and destroy all recordset objects used
    rsMonPoint.Close
    Set rsMonPoint = Nothing

Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)
End Sub


