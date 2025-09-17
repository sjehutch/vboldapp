VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Caption         =   "btn&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "btn&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Frame fraUnitType 
      Caption         =   " fraSelect Unit Type "
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optMetric 
         Caption         =   "optMetric"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optEnglish 
         Caption         =   "optEnglish"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MODULE_NAME = "frmOptions.frm"

'**************************************************************
'   DESCRIPTION
'       Closes the window
'
'   PARAMETERS
'       None
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/18/2002       Sheryll Go        Initial Revision
'**************************************************************
Private Sub btnCancel_Click()
    Unload Me
    Set frmOptions = Nothing
End Sub

'**************************************************************
'   DESCRIPTION
'       Saves the selected user setting in the registry
'
'   PARAMETERS
'       None
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/18/2002       Sheryll Go        Initial Revision
'**************************************************************
Private Sub btnOK_Click()

    Const FUNCTION_NAME = "btnOK_Click()"

    On Error GoTo Error_Handler
    
    If frmOptions.optEnglish.Value = True Then
        gbIsMetric = False
        regCreate_Key_Value HKEY_CURRENT_USER, "Software\DAS\ProFume\Options\Units", "Metric", 0
    End If
    
    If frmOptions.optMetric.Value = True Then
        gbIsMetric = True
        regCreate_Key_Value HKEY_CURRENT_USER, "Software\DAS\ProFume\Options\Units", "Metric", 1
    End If
    
    'Err.Raise 101, "Profume", "Cannot save to registry."
    Unload Me
    Set frmOptions = Nothing

Exit Sub

Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving setting to the registry."
    
End Sub
'**************************************************************
'   DESCRIPTION
'       Checks for the global variable gbIsMetric
'
'   PARAMETERS
'       None
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/18/2002       Sheryll Go        Initial Revision
'**************************************************************
Private Sub Form_Load()
    
    Const FUNCTION_NAME = "Form_Load()"

    On Error GoTo Error_Handler
    
    'call sub to load resource strings
    OptionsLoadResString
    
    If gbIsMetric = True Then
        frmOptions.optMetric.Value = True
    Else
        frmOptions.optEnglish.Value = True
    End If
    
Exit Sub

Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in loading the form."
    
End Sub
'**************************************************************
'   DESCRIPTION
'       Procedure to load the resource strings in the form
'
'   PARAMETERS
'       None
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/18/2002       Sheryll Go        Initial Revision
'**************************************************************
Public Sub OptionsLoadResString()

    Const FUNCTION_NAME = "OptionsLoadResString()"

    On Error GoTo Error_Handler
    
    With frmOptions
        .Caption = LoadResString(IDS_FRM_OPTIONS)
        .fraUnitType.Caption = LoadResString(IDS_FRA_SELECTUNITTYPE)
        .optEnglish.Caption = LoadResString(IDS_OPT_ENGLISH)
        .optMetric.Caption = LoadResString(IDS_OPT_METRIC)
        .btnOK.Caption = LoadResString(IDS_BTN_OK)
        .btnCancel.Caption = LoadResString(IDS_BTN_CANCEL)
    End With
    
Exit Sub

Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in loading resource strings."
    
End Sub
