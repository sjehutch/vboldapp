VERSION 5.00
Begin VB.Form frmAppError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmAppError"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmAppError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   45
      TabIndex        =   0
      Top             =   38
      Width           =   6975
      Begin VB.TextBox txtErrText 
         Height          =   1815
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmAppError.frx":0E42
         Top             =   1200
         Width           =   6810
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "cmdOK"
         Height          =   375
         Left            =   5760
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblErrText 
         Caption         =   "lblErrText"
         Height          =   855
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   4815
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   120
         Picture         =   "frmAppError.frx":0E4F
         Top             =   120
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmAppError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SetForm As String
Public SetFunction As String
Public SetAddInfo As String

'**************************************************************
'   SUB NAME: cmdOK_Click
'
'   DESCRIPTION
'       Unloads the form
'
'   PARAMETERS
'       None
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/23/2002       Davis Grana        Initial Revision
'**************************************************************
Private Sub cmdOK_Click()

    Unload Me
    
End Sub

'**************************************************************
'   SUB NAME: Form_Load
'
'   DESCRIPTION
'       Loads resource strings and builds the error
'       message text
'
'   PARAMETERS
'       None
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/23/2002       Davis Grana        Initial Revision
'**************************************************************

Private Sub Form_Load()

    Dim strMessage As String
    
    '-- LOAD WIDGET LABELS
    frmAppError.Caption = LoadResString(IDS_ERR_TITLE)
    lblErrText.Caption = LoadResString(IDS_ERR_TEXT)
    cmdOK.Caption = LoadResString(IDS_BTN_OK)
    
    '--BUILD MESSAGE STRING
    strMessage = "[Source] " & Err.Source & vbCrLf & _
                 "[Module] " & SetForm & vbCrLf & _
                 "[Function] " & SetFunction & vbCrLf & _
                 "[Err Number] " & Err.Number & vbCrLf & _
                 "[Err Description] " & Err.Description & vbCrLf & _
                 "[Additional Info] " & SetAddInfo
                 
    txtErrText = strMessage
    
End Sub

'**************************************************************
'   SUB NAME: Form_UnLoad
'
'   DESCRIPTION
'       Unloads the form
'
'   PARAMETERS
'       None
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/23/2002       Davis Grana        Initial Revision
'**************************************************************

Private Sub Form_Unload(Cancel As Integer)
    
    '--TERMINATE APPLICATION
    Set frmAppError = Nothing
    End
    
End Sub
