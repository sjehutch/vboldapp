VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGraphProperties 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "frmGraphProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   6375
      Begin VB.CheckBox chkShowPoints 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   5535
      End
   End
   Begin VB.Frame frayrange 
      Height          =   615
      Left            =   2160
      TabIndex        =   16
      Top             =   1680
      Width           =   4335
      Begin VB.TextBox txtymin 
         Height          =   285
         Left            =   960
         MaxLength       =   9
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtymax 
         Height          =   285
         Left            =   2280
         MaxLength       =   9
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblto2 
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame fraxrange 
      Height          =   735
      Left            =   2160
      TabIndex        =   13
      Top             =   960
      Width           =   4335
      Begin MSComCtl2.DTPicker dtpMaxDate 
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yy hh:mm tt"
         Format          =   19726339
         CurrentDate     =   37739
      End
      Begin MSComCtl2.DTPicker dtpMinDate 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yy hh:mm tt"
         Format          =   19726339
         CurrentDate     =   37739
      End
      Begin VB.TextBox txtxmax 
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtxmin 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblto1 
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraGraphTitle 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   6375
      Begin VB.TextBox txtGraphTitle 
         Height          =   285
         Left            =   120
         MaxLength       =   200
         TabIndex        =   8
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.ComboBox cboUnits 
      Height          =   315
      ItemData        =   "frmGraphProperties.frx":0E42
      Left            =   240
      List            =   "frmGraphProperties.frx":0E44
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame fraUnits 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   6375
      Begin VB.CheckBox chkYes 
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   160
         Width           =   5775
      End
   End
End
Attribute VB_Name = "frmGraphProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'* Filename: frmGraphProperties.frm
'*
'* DESCRIPTION
'* -----------
'* contains the form which will contain the controls (Chartfx)
'* listboxes, buttons etc. needed to set the properties of the graph
'*
'* Copyright (c) 2002 Accenture.  All rights reserved.
'-----------------------------------------------------
'*
'* REVISION HISTORY
'* Revision:    0
'* Author:      Gamboa, Maria Cecilia <NAI1686>
'* Date:        March 25, 2003
'*
'* 3/7/2004  Jett Gamboa      Use gFullTimeFormat instead of hardcoded time formats
'**************************************************************
Private Const MODULE_NAME = "frmGraphProperties"
Private sModName As String
Private sErrorDesc As String

Dim sMinValue As String
Dim sMaxValue As String

Private Sub cboUnits_Click()
    'JNM AT Fix
    If cboUnits.Text = LoadResString(IDS_GRAPH_X_UNITS_TIME) Then
        txtxmin.Visible = False
        txtxmax.Visible = False
        dtpMinDate.Visible = True
        dtpMaxDate.Visible = True
    Else
        txtxmin.Visible = True
        txtxmax.Visible = True
        dtpMinDate.Visible = False
        dtpMaxDate.Visible = False
    End If
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
'   March 7, 2004        Jett Gamboa      Set date format for dropdowns depending on time format
'**************************************************************
Private Sub Form_Load()

    Dim rsSettings As New ADODB.Recordset
    Dim rsDates As New ADODB.Recordset
    Dim rsJobArea As New ADODB.Recordset
    Dim sDatesSQL As String
    Dim sSQL As String
    Dim sSQLGraph As String
    
    On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "Form_Load"
    
    If gTimeFormat = "HH:mm" Then
        dtpMinDate.CustomFormat = "dd-MMM-yy HH:mm"
        dtpMaxDate.CustomFormat = "dd-MMM-yy HH:mm"
    Else
        dtpMinDate.CustomFormat = "dd-MMM-yy hh:mm tt"
        dtpMaxDate.CustomFormat = "dd-MMM-yy hh:mm tt"
    End If
    
    frmGraphProperties.Caption = LoadResString(IDS_GRAPH_PROPERTIES_BUTTON)
    fraGraphTitle.Caption = LoadResString(IDS_GRAPH_TITLE)
    fraUnits.Caption = LoadResString(IDS_GRAPH_UNIT_FOR_AXIS)
    chkYes.Caption = LoadResString(IDS_GRAPH_DISPLAY_VER_LINE)
    cmdOk.Caption = LoadResString(IDS_GRAPH_OK)
    cmdCancel.Caption = LoadResString(IDS_GRAPH_CANCEL)
    fraxrange.Caption = LoadResString(IDS_GRAPH_X_AXIS_TITLE)
    frayrange.Caption = LoadResString(IDS_GRAPH_Y_AXIS_TITLE)
    lblto1.Caption = LoadResString(IDS_GRAPH_TO_TITLE)
    lblto2.Caption = LoadResString(IDS_GRAPH_TO_TITLE)
    cboUnits.AddItem LoadResString(IDS_GRAPH_X_UNITS_TIME)
    cboUnits.AddItem LoadResString(IDS_GRAPH_X_UNITS_ELAPSED)
    chkShowPoints.Caption = LoadResString(IDS_GRAPH_SHOW_POINTS)
    
    'tool tip text
    chkYes.ToolTipText = LoadResString(IDS_GRAPH_TOOLTIP_SHOWLINES)
    chkShowPoints.ToolTipText = LoadResString(IDS_GRAPH_TOOLTIP_SHOWPOINTS)
    cboUnits.ToolTipText = LoadResString(IDS_GRAPH_TOOLTIP_UNITS)
    txtxmin.ToolTipText = LoadResString(IDS_GRAPH_TOOLTIP_XMIN)
    txtxmax.ToolTipText = LoadResString(IDS_GRAPH_TOOLTIP_XMAX)
    'JNM 04/28/03 - AT Fix - Add datepicker
    dtpMinDate.ToolTipText = LoadResString(IDS_GRAPH_TOOLTIP_XMIN)
    dtpMaxDate.ToolTipText = LoadResString(IDS_GRAPH_TOOLTIP_XMAX)
    
    txtymin.ToolTipText = LoadResString(IDS_GRAPH_TOOLTIP_YMIN)
    txtymax.ToolTipText = LoadResString(IDS_GRAPH_TOOLTIP_YMAX)
    txtGraphTitle.ToolTipText = LoadResString(IDS_GRAPH_TOOLTIP_TITLE)
    
    ' Retrieve the graph settings set by the user (or the default values, if no updates were made) from the Settings table
    sSQLGraph = "SELECT Property " _
                & ", Property_ID " _
                & ", Property_value " _
                & "FROM Settings " _
                & "WHERE Property_id IN (1, 2, 3, 4, 5, 6, 7, 8)"
    rsSettings.Open sSQLGraph, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rsSettings.BOF And rsSettings.EOF Then
        MsgBox LoadResString(IDS_GRAPH_INCOMPATIBLE_JOB_FILE), vbExclamation
        rsSettings.Close
        Set rsSettings = Nothing
        Exit Sub
    Else
        
        sSQL = "SELECT max(initial_concentration) AS MaxIntConc " _
            & "FROM JobArea " _
            & "WHERE deleted = false"
            
        rsJobArea.Open sSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
        
        sDatesSQL = "SELECT startintro_datetime " _
                & ", endexp_datetime " _
                & "FROM Job"
        rsDates.Open sDatesSQL, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
        
        rsSettings.MoveFirst
        While Not rsSettings.EOF
    
            Select Case rsSettings("Property_ID")
            
                Case GSTITLE
                    'Title
                    txtGraphTitle.Text = rsSettings("Property_Value")
                Case GSSHOWLINES
                    'ShowLines
                    If rsSettings("Property_Value") Then
                        chkYes.Value = 1
                    Else
                        chkYes.Value = 0
                    End If
                Case GSUNITS
                    'Units
                   If rsSettings("Property_Value") = 1 Then
                        cboUnits.Text = LoadResString(IDS_GRAPH_X_UNITS_TIME)
                        'JNM 04/28/03 - AT Fix - Add datepicker
                        txtxmin.Visible = False
                        txtxmax.Visible = False
                        dtpMinDate.Visible = True
                        dtpMaxDate.Visible = True
                   Else
                        cboUnits.Text = LoadResString(IDS_GRAPH_X_UNITS_ELAPSED)
                        'JNM 04/28/03 - AT Fix - Add datepicker
                        txtxmin.Visible = True
                        txtxmax.Visible = True
                        dtpMinDate.Visible = False
                        dtpMaxDate.Visible = False
                   End If
                Case GSX_MIN
                    'X_Min
                    'JNM 04/28/03 - At Fix - Use datepickers
                    If cboUnits.Text = LoadResString(IDS_GRAPH_X_UNITS_TIME) Then
                        If IsNull(rsSettings("Property_Value")) Then
                            dtpMinDate.Value = Format(rsDates("startintro_datetime"), gFullTimeFormat)
                        Else
                            dtpMinDate.Value = Format(rsSettings("Property_Value"), gFullTimeFormat)
                        End If
                        txtxmin.Text = 0
                    Else
                        If IsNull(rsSettings("Property_Value")) Then
                            txtxmin.Text = 0
                        Else
                            txtxmin.Text = rsSettings("Property_Value")
                        End If
                        dtpMinDate.Value = Format(rsDates("startintro_datetime"), gFullTimeFormat)
                    End If
                    
                Case GSX_MAX
                    'X_Max
                    'JNM 04/28/03 - At Fix - Use datepickers
                    If cboUnits.Text = LoadResString(IDS_GRAPH_X_UNITS_TIME) Then
                        If IsNull(rsSettings("Property_Value")) Then
                            dtpMaxDate.Value = Format(rsDates("endexp_datetime"), gFullTimeFormat)
                        Else
                            dtpMaxDate.Value = Format(rsSettings("Property_Value"), gFullTimeFormat)
                        End If
                        txtxmax.Text = 24
                    Else
                        If IsNull(rsSettings("Property_Value")) Then
                            txtxmax.Text = 24
                        Else
                            txtxmax.Text = rsSettings("Property_Value")
                        End If
                        dtpMaxDate.Value = Format(rsDates("endexp_datetime"), gFullTimeFormat)
                    End If
                   
                Case GSY_MIN
                    'Y_Min
                    If IsNull(rsSettings("Property_Value")) Then
                        txtymin.Text = 0
                    Else
                        txtymin.Text = rsSettings("Property_Value")
                    End If
                    
                Case GSY_MAX
                    'Y_Max
                    If IsNull(rsSettings("Property_Value")) Then
                        If Not (rsJobArea.BOF And rsJobArea.EOF) Then
                            txtymax.Text = rsJobArea("MaxIntConc")
                        End If
                    Else
                        txtymax.Text = rsSettings("Property_Value")
                    End If
                    
                Case GSSHOWPOINTS
                    'ShowPoints
                    If rsSettings("Property_Value") Then
                        chkShowPoints.Value = 1
                    Else
                        chkShowPoints.Value = 0
                    End If
            End Select
        rsSettings.MoveNext
    Wend
    End If
    
    'close and destroy all recordset objects used
    rsSettings.Close
    Set rsSettings = Nothing
    rsJobArea.Close
    Set rsJobArea = Nothing
    rsDates.Close
    Set rsDates = Nothing
    
    Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)
End Sub

'**************************************************************
'   DESCRIPTION
'   Handles the on click event of the Cancel button
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************

Private Sub cmdCancel_Click()
    Unload Me
    Set frmGraphProperties = Nothing
End Sub

'**************************************************************
'   DESCRIPTION
'   Handles the on click event of the OK button
'
'   Input Parameters: None
'
'   RETURNS - None
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************

Private Sub cmdOK_Click()
    Dim bShowPoints As Boolean
    Dim bShowlines As Boolean
    Dim bXGrid As Boolean
    Dim bYGrid As Boolean
    Dim sNewTitle As String
    Dim sSQL As String
    Dim sPrevSelectedY As String
    Dim sPrevSelectedN As String
    Dim nUnitsSelected As Integer
    Dim nYMin As Long
    Dim nYMax As Long
    Dim vXMin As Variant
    Dim vXmax As Variant
    
    On Error GoTo Error_Handler
    
    Const FUNCTION_NAME = "cmdOK_Click"
    
    If chkYes.Value = False Then
        bShowlines = False
    Else
        bShowlines = True
    End If
    
    
    If chkShowPoints.Value = False Then
        bShowPoints = False
    Else
        bShowPoints = True
    End If
    
    If cboUnits.Text = LoadResString(IDS_GRAPH_X_UNITS_TIME) Then
        nUnitsSelected = 1
        If Trim(dtpMinDate.Value) = "" Then
            MsgBox LoadResString(IDS_GRAPH_SET_DEFAULT_X), vbExclamation
            txtxmin.SetFocus
            Exit Sub
        End If
        If Trim(dtpMaxDate.Value) = "" Then
            MsgBox LoadResString(IDS_GRAPH_SET_DEFAULT_X), vbExclamation
            txtxmax.SetFocus
            Exit Sub
        End If

        If (CDate(Format(dtpMinDate.Value, gFullTimeFormat)) >= CDate(Format(dtpMaxDate.Value, gFullTimeFormat))) Then
            MsgBox LoadResString(IDS_INVALID_MIN_MAX), vbExclamation
            dtpMinDate.SetFocus
            Exit Sub
        Else
            nXMin = CDate(Format(dtpMinDate.Value, gFullTimeFormat))
            nXmax = CDate(Format(dtpMaxDate.Value, gFullTimeFormat))
        End If

    Else
        nUnitsSelected = 2
        If Trim(txtxmin.Text) = "" Then
            MsgBox LoadResString(IDS_GRAPH_SET_DEFAULT_X), vbExclamation
            txtxmin.SetFocus
            Exit Sub
        End If
        If Trim(txtxmax.Text) = "" Then
            MsgBox LoadResString(IDS_GRAPH_SET_DEFAULT_X), vbExclamation
            txtxmax.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(Trim(txtxmin.Text)) Then
           MsgBox LoadResString(IDS_INVALID_ELAPSEDTIME), vbExclamation
           txtxmin.SetFocus
           Exit Sub
        End If
        If Not IsNumeric(Trim(txtxmax.Text)) Then
            MsgBox LoadResString(IDS_INVALID_ELAPSEDTIME), vbExclamation
            txtxmax.SetFocus
            Exit Sub
        End If
        If (CDbl(Trim(txtxmin.Text)) >= CDbl(Trim(txtxmax.Text))) Then
            MsgBox LoadResString(IDS_INVALID_MIN_MAX), vbExclamation
            txtxmin.SetFocus
            Exit Sub
        Else
            nXMin = CDbl(txtxmin.Text)
            nXmax = CDbl(txtxmax.Text)
        End If
    End If
    
    If Trim(txtGraphTitle.Text) = "" Then
        MsgBox LoadResString(IDS_NO_GRAPH_TITLE), vbExclamation
        txtGraphTitle.SetFocus
        Exit Sub
    End If
    
    sNewTitle = txtGraphTitle.Text
    
    If Not IsNumeric(Trim(txtymin.Text)) Then
        MsgBox LoadResString(IDS_INVALID_CONCENTRATION), vbExclamation
        txtymin.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(Trim(txtymax.Text)) Then
        MsgBox LoadResString(IDS_INVALID_CONCENTRATION), vbExclamation
        txtymax.SetFocus
        Exit Sub
    End If
    'JNM 05/05/03 - PT Fix - Changed CInt to CLong in order to accept large numbers
    If (CLng(txtymin.Text)) >= (CLng(txtymax.Text)) Then
       MsgBox LoadResString(IDS_INVALID_MIN_MAX), vbExclamation
       txtymin.SetFocus
       Exit Sub
    Else
       nYMin = CLng(txtymin.Text)
       nYMax = CLng(txtymax.Text)
    End If
    
    bXGrid = frmMain.ChartFX1.Axis(AXIS_X).Grid
    bYGrid = frmMain.ChartFX1.Axis(AXIS_Y).Grid
    sPrevSelectedY = frmMain.txtSelectedY.Text
    sPrevSelectedN = frmMain.txtSelectedN.Text
    
    'Save settings in the database
    Call SaveGraphSettings(sNewTitle, _
                           bShowlines, _
                           nUnitsSelected, _
                           nXMin, _
                           nXmax, _
                           nYMin, _
                           nYMax, _
                           bShowPoints, _
                           bXGrid, _
                           bYGrid, _
                           sPrevSelectedY, _
                           sPrevSelectedN)
    
    Unload Me
    Set frmGraphProperties = Nothing
    
    Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(MODULE_NAME, FUNCTION_NAME, sErrorDesc)

End Sub


