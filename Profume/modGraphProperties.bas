Attribute VB_Name = "modGraphProperties"
'**************************************************************
'* Filename: modGraphProperties.bas
'*
'* DESCRIPTION
'* -----------
'* This module contains the functions/subroutines which will be
'* executed when the program runs. It contains different
'* functions which used by the graph and graph properties forms.
'*
'* Copyright (c) 2002 Accenture.  All rights reserved.
'-----------------------------------------------------
'*
'* REVISION HISTORY
'* Revision:    0
'* Author:      Gamboa, Maria Cecilia <NAI1686>
'* Date:        March 25, 2003
'*
'* DATE         AUTHOR                DESCRIPTION
'* 4/24/03      Jett Gamboa           Fixed inefficient use of ADO object in SetPoints
'*                                    changed approach to hiding data points
'* 3/7/04       Jett Gamboa           Used gFullTimeFormat instead of hardcoded date and
'*                                    formats
'**************************************************************

Option Explicit
Public clsExport As clsWordExport
Public Const GSTITLE = 1
Public Const GSSHOWLINES = 2
Public Const GSUNITS = 3
Public Const GSX_MIN = 4
Public Const GSX_MAX = 5
Public Const GSY_MIN = 6
Public Const GSY_MAX = 7
Public Const GSSHOWPOINTS = 8
Public Const GSSHOWXGRID = 9
Public Const GSSHOWYGRID = 10
Public Const TIME_UNIT = 1
Public Const ELAPSEDTIME_UNIT = 2
Public Const GSNOPOINTS = -10

Const CHART_HIDDEN = 1E+308   ' ChartFX Constant (value for hidden points)

Private Const MODULE_NAME = "modGraphProperties"

Private sErrorDesc As String
Public bNoSettings As Boolean

Dim sTitle As String
Dim sShowLines As String
Dim sUnits As String
Dim sX_Min As String
Dim sX_Max As String
Dim sY_Min As String
Dim sY_Max As String
Dim sShowPoints As String
Dim sShowXGrid As String
Dim sShowYGrid As String


'**************************************************************
'   DESCRIPTION
'   Shows the constant lines (Start Exposure and End Exposure)
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'   May 8, 2003          Amana Martinez   Fumiguide Enhancements 2003
'**************************************************************

Public Sub ShowLines()

    Dim rsDates As New ADODB.Recordset
    Dim sDatesSQL As String
  
    On Error GoTo Error_Handler

    Const FUNCTION_NAME = "ShowLines"
    gbShowLines = True
    
    frmMain.ChartFX1.OpenDataEx COD_CONSTANTS, 2, 0

    sDatesSQL = "Select startexp_datetime, endexp_datetime from Job"
    rsDates.Open sDatesSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
      
    'JNM 04/22/03 - AT Fix - Check if dates have values or not
    'If rsDates.BOF And rsDates.EOF Then
    If IsNull(rsDates("startexp_datetime")) Or IsNull(rsDates("endexp_datetime")) Then
        'there are no dates yet
        MsgBox LoadResString(IDS_GRAPH_NO_DATES_YET), vbExclamation
    Else
        'Markers
        If sUnits = TIME_UNIT Then
            frmMain.ChartFX1.ConstantLine(0).Value = CDate(rsDates("startexp_datetime"))
            frmMain.ChartFX1.ConstantLine(1).Value = CDate(rsDates("endexp_datetime"))
        Else
            frmMain.ChartFX1.ConstantLine(0).Value = 0
            frmMain.ChartFX1.ConstantLine(1).Value = DateDiff("h", CDate(rsDates("startexp_datetime")), CDate(rsDates("endexp_datetime")))
        End If
        
        '5/8/2003 - AM Added more below Labels for Fumiguide Enhancements 2003 PT Defect 8721
        frmMain.ChartFX1.ConstantLine.Item(0).Label = LoadResString(IDS_GRAPH_START_EXPOSURE) & " - " & Format(rsDates("startexp_datetime"), gFullTimeFormat)
        '5/8/2003 - AM Added more above Labels for Fumiguide Enhancements 2003 PT Defect 8721
        frmMain.ChartFX1.ConstantLine.Item(0).LineColor = &H0&
        frmMain.ChartFX1.ConstantLine.Item(0).LineWidth = 2
        frmMain.ChartFX1.ConstantLine(0).LineStyle = CHART_DOT
        
        '5/8/2003 - AM Added more below Labels for Fumiguide Enhancements 2003 PT Defect 8721
        frmMain.ChartFX1.ConstantLine.Item(1).Label = LoadResString(IDS_GRAPH_END_EXPOSURE) & " - " & Format(rsDates("endexp_datetime"), gFullTimeFormat)
        
        '5/8/2003 - AM Added more above Labels for Fumiguide Enhancements 2003 PT Defect 8721
        frmMain.ChartFX1.ConstantLine.Item(1).LineColor = &H0&
        frmMain.ChartFX1.ConstantLine.Item(1).LineWidth = 2
        frmMain.ChartFX1.ConstantLine(1).LineStyle = CHART_DOT
        
        frmMain.ChartFX1.ConstantLine(0).Style = (Not CC_HIDE) And (Not CC_HIDETEXT) And (Not CC_RIGHTALIGNED)
        frmMain.ChartFX1.ConstantLine(1).Style = (Not CC_HIDE) And (Not CC_HIDETEXT) And (Not CC_RIGHTALIGNED)
        
        frmMain.ChartFX1.ConstantLine(0).Axis = AXIS_X
        frmMain.ChartFX1.ConstantLine(1).Axis = AXIS_X
    End If
    
    frmMain.ChartFX1.CloseData COD_CONSTANTS
    
    'Close and destroy the recordset object
    rsDates.Close
    Set rsDates = Nothing

    Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)

End Sub

'**************************************************************
'   DESCRIPTION
'   Loads the graph settings that was saved in the database
'   (or default values for first time users)
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'   March 24, 2003       Jett Gamboa      Move setting of colors after series is loaded
'   March 3, 2004       Jett Gamboa       Set the x-axis time format based on the global time format
'**************************************************************

Public Sub LoadGraphSettings(Optional argSelectedY As String, Optional argSelectedN As String)

    Dim rsSettings As New ADODB.Recordset
    Dim rsMonPointGraph As New ADODB.Recordset
    Dim rsDates As New ADODB.Recordset
    Dim rsJobArea As New ADODB.Recordset
    Dim sSQLGraph As String
    Dim I As Integer
    Dim sDatesSQL As String
    Dim sSQL As String
    Dim nRecordCount As Integer
    
    On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "LoadGraphSettings"
    I = 0
    
    sDatesSQL = "SELECT startintro_datetime " _
              & ", startexp_datetime " _
              & ", endexp_datetime " _
              & "FROM Job"
              
    rsDates.Open sDatesSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    sSQL = "SELECT max(initial_concentration) as MaxIntConc " _
         & "FROM JobArea " _
         & "WHERE deleted = false"
         
    rsJobArea.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    ' Retrieve the graph settings set by the user (or the default values, if no updates were made) from the Settings table
    sSQLGraph = "SELECT Property_ID " _
                & ", Property " _
                & ", Property_Value " _
                & "FROM Settings"
                
    rsSettings.Open sSQLGraph, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rsSettings.BOF And rsSettings.EOF Then
        MsgBox LoadResString(IDS_GRAPH_INCOMPATIBLE_JOB_FILE), vbExclamation
        rsSettings.Close
        Set rsSettings = Nothing
        Exit Sub
    Else
        rsSettings.MoveFirst
        While Not rsSettings.EOF
    
            Select Case rsSettings("Property_ID")
            
                Case GSTITLE
                    'Title
                    sTitle = rsSettings("Property_Value")
                Case GSSHOWLINES
                    'ShowLines
                    sShowLines = rsSettings("Property_Value")
                    gbShowLines = sShowLines
                Case GSUNITS
                    'Units
                   sUnits = rsSettings("Property_Value")
                   gnGraphUnits = sUnits
                Case GSX_MIN
                    'X_Min
                    If IsNull(rsSettings("Property_Value")) Then
                        'JNM 04/22/03 - AT Fix - Check if the date is null or not
                        'If Not (rsDates.BOF And rsDates.EOF) Then
                        If Not IsNull(rsDates("startintro_datetime")) Then
                            sX_Min = rsDates("startintro_datetime")
                        Else
                            'JNM 04/22/03 - Set default value
                            sX_Min = "January 1, 1900"
                        End If
                    Else
                        sX_Min = rsSettings("Property_Value")
                    End If
                    
                Case GSX_MAX
                    'X_Max
                    If IsNull(rsSettings("Property_Value")) Then
                        'JNM 04/22/03 - AT Fix - Check if the date is null or not
                        'If Not (rsDates.BOF And rsDates.EOF) Then
                        If Not IsNull(rsDates("endexp_datetime")) Then
                            sX_Max = rsDates("endexp_datetime")
                        Else
                            'JNM 04/22/03 - Set default value
                            sX_Max = "January 2, 1900"
                        End If
                    Else
                        sX_Max = rsSettings("Property_Value")
                    End If
                  
                Case GSY_MIN
                    'Y_Min
                    If IsNull(rsSettings("Property_Value")) Then
                        sY_Min = 0
                    Else
                        sY_Min = rsSettings("Property_Value")
                    End If
                Case GSY_MAX
                    'Y_Max
                    If IsNull(rsSettings("Property_Value")) Then
                        If Not (rsJobArea.BOF And rsJobArea.EOF) Then
                            sY_Max = rsJobArea("MaxIntConc")
                        End If
                    Else
                        sY_Max = rsSettings("Property_Value")
                    End If
                Case GSSHOWPOINTS
                   'ShowPoints
                    sShowPoints = rsSettings("Property_Value")
                Case GSSHOWXGRID
                   'ShowXGrid
                    sShowXGrid = rsSettings("Property_Value")
                Case GSSHOWYGRID
                   'ShowYGrid
                    sShowYGrid = rsSettings("Property_Value")
            End Select
        rsSettings.MoveNext
    Wend
    End If
    
    'Set the user-specific graph settings
    With frmMain
        .ChartFX1.Title(CHART_TOPTIT) = sTitle
        .ChartFX1.Axis(AXIS_X).Grid = sShowXGrid
        .ChartFX1.Axis(AXIS_Y).Grid = sShowYGrid
        'JNM 05/05/03 - PT fix - changed CInt to CLng in order to accept large values
        .ChartFX1.Axis(AXIS_Y).Min = CLng(sY_Min)
        .ChartFX1.Axis(AXIS_Y).Max = CLng(sY_Max)
        .ChartFX1.Axis(AXIS_Y).Title = LoadResString(IDS_GRAPH_Y_DEFAULT)
        .ChartFX1.Axis(AXIS_X).TextColor = &H0&
        .ChartFX1.Axis(AXIS_Y).TextColor = &H0&
    End With

    If sShowLines Then
        'Show the lines for the Start Exposure and End Exposure
        Call ShowLines
    Else
        Call HideLines
    End If
        
    If sUnits = TIME_UNIT Then
        frmMain.ChartFX1.Axis(AXIS_X).Title = LoadResString(IDS_GRAPH_X_UNITS_TIME)
        frmMain.ChartFX1.Axis(AXIS_X).Min = CDate(sX_Min)
        frmMain.ChartFX1.Axis(AXIS_X).Max = CDate(sX_Max)
        ' Set the selected format
        
        ' Check if this is a European time format, use the 24 hour format if it is
        ' otherwise use the format with am/pm
        If gTimeFormat = "HH:mm" Then
            frmMain.ChartFX1.Axis(AXIS_X).Format = "tH:mm"
        Else
            frmMain.ChartFX1.Axis(AXIS_X).Format = "th:mm tt"
        End If
        
        ' Set the Label Angle (for Time entries, set the label at 75°
        frmMain.ChartFX1.Axis(AXIS_X).LabelAngle = 75
        If (argSelectedY = "" And argSelectedN = "") Then
            Call SetPoints(TIME_UNIT)
        Else
            'Load the points
            Call SetPoints(TIME_UNIT)
            'Call the Set points, but this time show only the selected points
            Call SetPoints(TIME_UNIT, argSelectedN, argSelectedY, argSelectedY)
        End If
    Else
        'This means sUnits = ELAPSEDTIME_UNIT
        frmMain.ChartFX1.Axis(AXIS_X).Title = LoadResString(IDS_GRAPH_X_UNITS_ELAPSED)
        frmMain.ChartFX1.Axis(AXIS_X).Min = CDbl(sX_Min)
        frmMain.ChartFX1.Axis(AXIS_X).Max = CDbl(sX_Max)
        frmMain.ChartFX1.Axis(AXIS_X).Format = AF_NONE
        frmMain.ChartFX1.Axis(AXIS_X).LabelAngle = 0
        If (argSelectedY = "" And argSelectedN = "") Then
            Call SetPoints(ELAPSEDTIME_UNIT)
        Else
            'Load the points
            Call SetPoints(ELAPSEDTIME_UNIT)
            'Call the Set points, but this time show only the selected points
            Call SetPoints(ELAPSEDTIME_UNIT, argSelectedN, argSelectedY, argSelectedY)
        End If
    End If
    
    ' 4/24/2003 Jett
    ' Move the setting of line colors here. We need to have the series defined first
    ' before we can get to this point
    
    sSQLGraph = "SELECT line_color " _
                & ", line_style" _
                & ", line_point" _
                & ", line_weight" _
                & ", line_marker " _
                & "FROM MonitorPoint WHERE deleted = false"
    rsMonPointGraph.Open sSQLGraph, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Not (rsMonPointGraph.BOF And rsMonPointGraph.EOF) Then
        I = 0
        nRecordCount = rsMonPointGraph.RecordCount
        rsMonPointGraph.MoveFirst
        
        While Not rsMonPointGraph.EOF
       
            'if the fields are not null, set the graph properties.
            'otherwise, do nothing and the graph will set the series' default properties
            If Not (IsNull(rsMonPointGraph("line_color"))) Then
                frmMain.ChartFX1.Series(I).Color = rsMonPointGraph("line_color")
            End If
            If Not (IsNull(rsMonPointGraph("line_style"))) Then
                frmMain.ChartFX1.Series(I).LineStyle = rsMonPointGraph("line_style")
            End If
            If Not (IsNull(rsMonPointGraph("line_point"))) Then
                frmMain.ChartFX1.Series(I).PointLabels = rsMonPointGraph("line_point")
            End If
            If Not (IsNull(rsMonPointGraph("line_weight"))) Then
                frmMain.ChartFX1.Series(I).LineWidth = rsMonPointGraph("line_weight")
            Else
                ' Jett 5/30/03 - default line width to 1 so that users using 98 can
                ' still change the settings
                frmMain.ChartFX1.Series(I).LineWidth = 1
            End If
             If Not (IsNull(rsMonPointGraph("line_marker"))) Then
                frmMain.ChartFX1.Series(I).MarkerShape = rsMonPointGraph("line_marker")
            End If
           I = I + 1
           rsMonPointGraph.MoveNext
           
        Wend
        
    End If
    
    'JNM 05/13/03 - PT fix - moved code here in order to have the
    '   initialized value of nRecordCount
    If sShowPoints Then
        For I = 0 To nRecordCount - 1
            frmMain.ChartFX1.Series(I).PointLabels = True
        Next
    Else
        For I = 0 To nRecordCount - 1
            frmMain.ChartFX1.Series(I).PointLabels = False
        Next
    End If
    
    'Close and destroy the recordset object
    rsSettings.Close
    rsMonPointGraph.Close
    rsDates.Close
    rsJobArea.Close
    Set rsSettings = Nothing
    Set rsMonPointGraph = Nothing
    Set rsDates = Nothing
    Set rsJobArea = Nothing
    
    Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    If Err.Number = -2147217865 Then
        MsgBox LoadResString(IDS_GRAPH_INCOMPATIBLE_JOB_FILE), vbExclamation
    Else
        Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)
    End If
End Sub

'**************************************************************
'   DESCRIPTION
'   Saves the graph settings that was set in Graph tab.
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'   May 5, 2003          Joseph Manyaga   PT fix - set y arguments to long
'                                         instead of integer in order to allow
'                                         big values
'**************************************************************

Public Sub SaveGraphSettings(argTitle As String, _
                                argShowLines As Boolean, _
                                argUnits As Integer, _
                                argX_Min As Variant, _
                                argX_Max As Variant, _
                                argY_Min As Long, _
                                argY_Max As Long, _
                                argShowPoints As Boolean, _
                                argShowXGrid As Boolean, _
                                argShowYGrid As Boolean, _
                                argSelectedY As String, _
                                argSelectedN As String)
    Dim sSQL As String
    Dim rsSaveSettings As New ADODB.Recordset
    
    On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "SaveGraphSettings"
    
    sSQL = "SELECT Property " _
        & ", Property_ID " _
        & ", Property_value FROM Settings"
        
    rsSaveSettings.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rsSaveSettings.BOF And rsSaveSettings.EOF Then
        MsgBox LoadResString(IDS_GRAPH_INCOMPATIBLE_JOB_FILE), vbExclamation
        rsSaveSettings.Close
        Set rsSaveSettings = Nothing
        Exit Sub
    Else
        rsSaveSettings.MoveFirst
        While Not rsSaveSettings.EOF
    
            Select Case rsSaveSettings("Property_ID")
            
                Case GSTITLE
                    'Title
                    rsSaveSettings("Property_value") = argTitle
                    
                Case GSSHOWLINES
                    'ShowLines
                    rsSaveSettings("Property_value") = argShowLines
                    
                Case GSUNITS
                    'Units
                    rsSaveSettings("Property_value") = argUnits
                    
                Case GSX_MIN
                    'X_Min
                     rsSaveSettings("Property_value") = argX_Min
                    
                Case GSX_MAX
                    'X_Max
                     rsSaveSettings("Property_value") = argX_Max
                    
                Case GSY_MIN
                    'Y_Min
                    rsSaveSettings("Property_value") = argY_Min
                    
                Case GSY_MAX
                    'Y_Max
                    rsSaveSettings("Property_value") = argY_Max
                    
                Case GSSHOWPOINTS
                   'ShowPoints
                    rsSaveSettings("Property_value") = argShowPoints
                    
                Case GSSHOWXGRID
                   'ShowXGrid
                    rsSaveSettings("Property_value") = argShowXGrid
                    
                Case GSSHOWYGRID
                   'ShowYGrid
                    rsSaveSettings("Property_value") = argShowYGrid
                   
                    
            End Select
             rsSaveSettings.Update
        rsSaveSettings.MoveNext
    Wend
    End If
    
    gbHasChanged = False
    'Refresh the graph
    Call LoadGraphSettings(argSelectedY, argSelectedN)
    
    'Close and destroy the recordset object
    rsSaveSettings.Close
    Set rsSaveSettings = Nothing
    
    Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)

End Sub

'**************************************************************
'   DESCRIPTION
'   Hides the constant lines (Start Exposure and End Exposure)
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub HideLines()

On Error GoTo Error_Handler

Const FUNCTION_NAME = "HideLines"

frmMain.ChartFX1.OpenDataEx COD_CONSTANTS, 2, 0

    ' Hide the constant lines and label
    frmMain.ChartFX1.ConstantLine(0).Style = CC_HIDE
    frmMain.ChartFX1.ConstantLine(1).Style = CC_HIDE
    gbShowLines = False
    
frmMain.ChartFX1.CloseData COD_CONSTANTS

Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)
    
End Sub

'**************************************************************
'   DESCRIPTION
'   Exports the graph
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************
Public Sub ExportGraph()

    Set clsExport = New clsWordExport
    
    On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "ExportGraph"
    
    'JNM 05/05/03 - PT fix - check if the user has MSWord
    If clsExport.HasMSWord = False Then
        Set clsExport = Nothing
        'Leave!!!!
        Exit Sub
    End If
    
    ' Clear the clipboard
    Clipboard.Clear
    
    ' Copy the graph into clipboard
    Clipboard.SetData frmMain.ChartFX1.GetPicture(CHART_BITMAP), vbCFBitmap

    With clsExport
        .OpenNewDoc
        
        ' Set the page setup to Landscape or Portait
        .PageSetupDocument = PageSetups.Portrait
    
        ' Insert a title for the Word Document
        .InsertText LoadResString(IDS_GRAPH_DOC_TITLE), 12, True, Alignment.Left
        ' Insert Spaces
        .InsertLinesInDoc = 1
        ' Insert the label for the date
        .InsertText LoadResString(IDS_GRAPH_EXPORT_DATE), 10, False, Alignment.Left
        ' Insert the Current Date
        .InsertCurrentDate
        ' Insert Spaces
        .InsertLinesInDoc = 3
    
        ' Insert the Graph
        .InsertGraph
            
    End With

    Set clsExport = Nothing

Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)
    
End Sub

'**************************************************************
'   DESCRIPTION
'   Retrieves the values (Concentration, Time or Elapsed Time)
'   from the database and sets these points in the graph.
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'   April 24, 2003       Jett Gamboa      Fixed inefficient use of recordset. Changed the
'                                         way data points are hidden from the graph
'**************************************************************
Public Sub SetPoints(intUnits As Integer, Optional sSeriesN As Variant, Optional sSeriesY As Variant, Optional sMonitorPoint As String)

    ' Declare and initialize the variables
    Dim arrSeries As Variant
    Dim rsFilterData As ADODB.Recordset
    Dim I As Integer
    Dim j As Integer
    Dim sSQLFilter As String
    Dim sSQLMonPt As String
    Dim nHideSeries As Integer
    Dim nShowSeries As Integer
    Dim sSeriesString As Variant
    Dim x As Integer
    Dim Y As Integer
    Dim nAllMonPts As Integer
    Dim sMonPtId As String
    Dim rsMonPoint As ADODB.Recordset
    Dim sSQL As String
    Dim vSeries As Variant
    Dim nRecCount As Integer
    Dim nMaxNumValues As Integer
    
    Dim rsPointCount As ADODB.Recordset
    
    On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "SetPoints"
    
    j = 0
       
    'retrieve all the Monitor Point names (in the order by which they were entered
    'in the monitoring plan
    

    ' Check if a value was passed to sMonitorPoint. If no monitor points were passed,
    ' it means that we need to reload the data points for all the Monitor points.
    ' Otherwise (in else statement), the monitor points not selected will just be hidden
    If sMonitorPoint = "" Then
    
        ' 4/24/03 Jett - added call to delete the data currently in the chart
        frmMain.ChartFX1.ClearData CD_VALUES
        frmMain.ChartFX1.ClearData CD_XVALUES
    
        ' We need to know the highest number of concentration readings for the monitoring
        ' points. It will be used as the number of values in the series
        sSQL = "SELECT MAX(PointCount) AS HighestCount FROM (SELECT Count(*) AS PointCount " & _
               "FROM (SELECT id, monitorpoint_fkey, fumiscope_conc FROM MonitorPointValue WHERE fumiscope_conc >= 0 and deleted = false) " & _
               "GROUP BY MonitorPointValue.monitorpoint_fkey)"
                   
        Set rsPointCount = New ADODB.Recordset
        
        rsPointCount.Open sSQL, gdbConn, adOpenForwardOnly, adLockReadOnly
        
        'JNM 04/27/03 - AT Fix - Set to 0 if there are no records
        If IsNull(rsPointCount("HighestCount")) Then
            nMaxNumValues = 0
        Else
            nMaxNumValues = rsPointCount("HighestCount")
        End If
               
        rsPointCount.Close
        Set rsPointCount = Nothing
                     
        sSQL = "Select id from MonitorPoint where deleted = false"
        
        Set rsMonPoint = New ADODB.Recordset
        
        'Store the records in a recordset
        rsMonPoint.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
        
        If rsMonPoint.BOF And rsMonPoint.EOF Then
            'there are no monitor points that exists for this table.
            'MsgBox LoadResString(IDS_GRAPH_NO_MON_POINTS), vbExclamation
            rsMonPoint.Close
            Set rsMonPoint = Nothing
            Exit Sub
        Else
            vSeries = rsMonPoint.GetRows
            
            ' Get the number of monitoring points
            nRecCount = rsMonPoint.RecordCount
            
            'We won't need the recordset so lets close it
            rsMonPoint.Close
            Set rsMonPoint = Nothing
            
            Set rsFilterData = New ADODB.Recordset
            
            frmMain.ChartFX1.ClearLegend CHART_SERLEG
            frmMain.ChartFX1.ClearData CD_SERLABELS
            
            ' 4/26/03 Jett - moved here so we only do this ones. The entire loops feeds
            ' all the data points.
            
            'Used to open a communications channel to pass data to the chart
            frmMain.ChartFX1.OpenDataEx COD_VALUES, nRecCount, nMaxNumValues
            frmMain.ChartFX1.OpenDataEx COD_XVALUES, nRecCount, nMaxNumValues
            
            ' Loop through each monitoring point
            For x = 0 To UBound(vSeries, 2)
            
                ' get the monitoring point id
                sMonPtId = vSeries(0, x)
                
                ' I holds the current data point we are updating. Initialize it to 0 because
                ' we are starting with the first point
                I = 0
                
                If intUnits = TIME_UNIT Then
                       sSQLFilter = "SELECT MonitorPoint.name" _
                                & ", MonitorPointValue.date_time" _
                                & ", MonitorPointValue.fumiscope_conc" _
                                & " FROM MonitorPointValue" _
                                & ", MonitorPoint " _
                                & "WHERE MonitorPointValue.monitorpoint_fkey = MonitorPoint.id " _
                                & "AND MonitorPoint.id = '" & sMonPtId & "' " _
                                & "AND MonitorPointValue.deleted = FALSE " _
                                & "AND MonitorPointValue.fumiscope_conc >= 0 " _
                                & "ORDER BY MonitorPointValue.date_time"
                Else
                       'JNM 04/27/03 - AT Fix - It's a Sunday, I'm hungry and
                       '    I found this bug. Anyway, change the SQL statement...
                       'Unit is elapsed_time
                       'sSQLFilter = "SELECT MonitorPoint.name" _
                                & ", MonitorPointValue.fumiscope_conc" _
                                & ", (MonitorPointValue.date_time - Job.startexp_datetime) AS elapsed_time" _
                                & " FROM MonitorPointValue" _
                                & ", MonitorPoint" _
                                & ", Job" _
                                & " WHERE MonitorPointValue.monitorpoint_fkey = MonitorPoint.id " _
                                & "AND MonitorPoint.id = '" & sMonPtId & "' " _
                                & "AND MonitorPointValue.deleted = FALSE " _
                                & "AND MonitorPointValue.fumiscope_conc >= 0 " _
                                & "ORDER BY elapsed_time"
                        sSQLFilter = "SELECT MonitorPoint.name" _
                                & ", MonitorPointValue.fumiscope_conc" _
                                & ", DateDiff('s', Job.startexp_datetime, MonitorPointValue.date_time)/3600 as elapsed_time" _
                                & " FROM MonitorPointValue" _
                                & ", MonitorPoint" _
                                & ", Job" _
                                & " WHERE MonitorPointValue.monitorpoint_fkey = MonitorPoint.id " _
                                & "AND MonitorPoint.id = '" & sMonPtId & "' " _
                                & "AND MonitorPointValue.deleted = FALSE " _
                                & "AND MonitorPointValue.fumiscope_conc >= 0 " _
                                & "ORDER BY DateDiff('s', Job.startexp_datetime, MonitorPointValue.date_time)/3600"
                End If
                                   
                'Store the records in a recordset
                rsFilterData.Open sSQLFilter, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
                                
                If Not (rsFilterData.EOF And rsFilterData.BOF) Then
                    'Clear the series legend and labels
                    
                    frmMain.ChartFX1.Series(x).Legend = rsFilterData("name")
                    frmMain.ChartFX1.Series(x).Visible = 1
                    rsFilterData.MoveFirst
                                       
                    While Not (rsFilterData.EOF)
                        ' set the graph
                         'First we set the Y coordinates using the YValue Property
        
                            If intUnits = TIME_UNIT Then
                                frmMain.ChartFX1.Series(x).Xvalue(I) = CDate(rsFilterData("date_time"))
                            Else
                                frmMain.ChartFX1.Series(x).Xvalue(I) = CDbl(rsFilterData("elapsed_time"))
                            End If
                            
                            frmMain.ChartFX1.Series(x).Yvalue(I) = rsFilterData("fumiscope_conc")
                            
                            ' move on to the next point
                            I = I + 1
                            rsFilterData.MoveNext
                    Wend
                                                              
' 4/24/03 Jett: removed the section below. We will hide "extra" points in the loop below
'
'                Else
                
'                    frmMain.ChartFX1.OpenDataEx COD_VALUES, nRecCount, COD_UNKNOWN
'                    frmMain.ChartFX1.OpenDataEx COD_XVALUES, nRecCount, COD_UNKNOWN
'                    frmMain.ChartFX1.Series(x).Xvalue(0) = GSNOPOINTS
'                    frmMain.ChartFX1.Series(x).Yvalue(0) = GSNOPOINTS
'                    frmMain.ChartFX1.Series(x).Visible = 0
'                    frmMain.ChartFX1.CloseData COD_VALUES
'                    frmMain.ChartFX1.CloseData COD_XVALUES
                    
                End If
                    
                'Close and destroy the recordset object
                rsFilterData.Close
                
                ' 4/24/03 Jett
                ' Loop through the remaining "points". They don't really exist for monitoring
                ' points which have less than the maximum number of points so we hide them
                'While I < nRecCount
                While I < nMaxNumValues
                    frmMain.ChartFX1.Series(x).Xvalue(I) = CHART_HIDDEN
                    frmMain.ChartFX1.Series(x).Yvalue(I) = CHART_HIDDEN
                    
                    I = I + 1
                Wend
                
                ' Close our data channel to the chart
                frmMain.ChartFX1.CloseData COD_VALUES
                frmMain.ChartFX1.CloseData COD_XVALUES

            Next
                        
            Set rsFilterData = Nothing
                    
        End If
            
    Else
                
        If sSeriesN <> "" Then
            sSeriesString = Split(sSeriesN, ",")
                For Y = 0 To UBound(sSeriesString)
                    nHideSeries = CInt(sSeriesString(Y))
                    frmMain.ChartFX1.Series(nHideSeries).Visible = 0
                Next
        End If
        
        If sSeriesY = "" Then
            frmMain.ChartFX1.Series.Visible = 1
        Else
            sSeriesString = Split(sSeriesY, ",")
                For Y = 0 To UBound(sSeriesString)
                    nShowSeries = CInt(sSeriesString(Y))
                    If (frmMain.ChartFX1.Series(nShowSeries).Xvalue(0) = GSNOPOINTS And frmMain.ChartFX1.Series(nShowSeries).Yvalue(0) = GSNOPOINTS) Then
                        frmMain.ChartFX1.Series(nShowSeries).Visible = 0
                    Else
                        frmMain.ChartFX1.Series(nShowSeries).Visible = 1
                    End If
                Next
        End If
                
    End If
    
    ' Set the Layout Settings
    frmMain.ChartFX1.SerLegBoxObj.Sizeable = 0
    frmMain.ChartFX1.SerLegBox = True
    frmMain.ChartFX1.SerLegBoxObj.Docked = 258

    Exit Sub
Error_Handler:
    MsgBox x
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)

End Sub

'**************************************************************
'   DESCRIPTION
'   Saves the Monitor Point-specific settings (like line color
'   for each series, point markers etc.)
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************

Public Sub SaveMonPtSpecificSettings()
    Dim I As Integer
    Dim rsMonPointGraph As New ADODB.Recordset
    Dim sSQLGraph As String
    
    On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "SaveMonPtSpecificSettings"
    
    I = 0
    sSQLGraph = "SELECT line_color" _
                & ", line_style" _
                & ", line_point" _
                & ", line_weight" _
                & ", line_marker " _
                & "FROM MonitorPoint " _
                & "WHERE deleted = false"
    
    rsMonPointGraph.Open sSQLGraph, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Not (rsMonPointGraph.BOF And rsMonPointGraph.EOF) Then
        While Not rsMonPointGraph.EOF
           rsMonPointGraph("line_color") = frmMain.ChartFX1.Series(I).Color
           rsMonPointGraph("line_style") = frmMain.ChartFX1.Series(I).LineStyle
           rsMonPointGraph("line_point") = frmMain.ChartFX1.Series(I).PointLabels
           rsMonPointGraph("line_weight") = frmMain.ChartFX1.Series(I).LineWidth
           rsMonPointGraph("line_marker") = frmMain.ChartFX1.Series(I).MarkerShape
           I = I + 1
           rsMonPointGraph.Update
           rsMonPointGraph.MoveNext
        Wend
    End If
    
    gbHasChanged = False
    
    'Close and destroy the recordset object
    rsMonPointGraph.Close
    Set rsMonPointGraph = Nothing
    
    Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)
End Sub

'**************************************************************
'   DESCRIPTION
'       This function verifies if the date format is correct.
'
'   PARAMETERS
'       sDate = the input date value
'
'   RETURNS - TRUE if the format is valid
'             FALSE if the format is invalid
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'       04/01/03  Cel Gamboa        Initial Revision.
'**************************************************************

Function CheckDate(objControl)
On Error GoTo Error_Handler

    Const FUNCTION_NAME = "CheckDate"
    If ValidateDate(Trim(objControl.Text)) = False Then
        MsgBox LoadResString(IDS_INVALID_DATE), vbExclamation
        objControl.SetFocus
        CheckDate = False
    Else
        CheckDate = True
    End If
    
Exit Function
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)
    
End Function
    
    
'**************************************************************
'   DESCRIPTION
'       This function verifies if the date format is correct.
'
'   PARAMETERS
'       sDate = the input date value
'
'   RETURNS - TRUE if the format is valid
'             FALSE if the format is invalid
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'       04/01/03  Cel Gamboa        Initial Revision.
'**************************************************************

Function ValidateDate(sDate)

    On Error GoTo Error_Handler
       
        Const FUNCTION_NAME = "ValidateDate"
        sDate = Trim(sDate)
        'Check if length of entered value is correct dd-mmm-yy hh:mm am/pm
        If Len(sDate) <> 18 Then
            ValidateDate = False
        Else
            'Check if there are slashes and a colon
            If Mid(sDate, 3, 1) <> "-" Or Mid(sDate, 7, 1) <> "-" Or Mid(sDate, 13, 1) <> ":" Then
                ValidateDate = False
            Else
                If IsDate(sDate) = True Then
                    ValidateDate = True
                Else
                    ValidateDate = False
                End If
            End If
        End If
        
    Exit Function
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)
    
End Function

'**************************************************************
'   DESCRIPTION
'   Checks if the  Start Intro, Start Exposure and End Exposure
'   were entered by the user in the previous tab.
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************

Function DatesExist()

    Dim bDefault As Boolean
    Dim rsDates As New ADODB.Recordset
    Dim sDatesSQL As String
    
    On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "DatesExist"
    
    bDefault = True
    sDatesSQL = "SELECT startintro_datetime " _
              & ", startexp_datetime " _
              & ", endexp_datetime " _
              & "FROM Job"
              
    rsDates.Open sDatesSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Not (rsDates.BOF And rsDates.EOF) Then
        'no dates has been set, there could be instances that some are null
        If IsNull(rsDates("startintro_datetime")) Then
            'MsgBox LoadResString(IDS_GRAPH_NO_DATES_YET), vbExclamation
            bDefault = False
        End If
        If IsNull(rsDates("startexp_datetime")) Then
            'MsgBox LoadResString(IDS_GRAPH_NO_DATES_YET), vbExclamation
            bDefault = False
            Exit Function
        End If
        If IsNull(rsDates("endexp_datetime")) Then
            'MsgBox LoadResString(IDS_GRAPH_NO_DATES_YET), vbExclamation
            bDefault = False
        End If
    Else
        'this means the recordset is empty
        bDefault = False
    End If
    
    rsDates.Close
    Set rsDates = Nothing
    
    DatesExist = bDefault
    
    Exit Function
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)

End Function
