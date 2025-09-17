VERSION 5.00
Begin VB.Form frmWarningMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmWarningMsg"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   Icon            =   "frmWarningMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   45
      TabIndex        =   0
      Top             =   38
      Width           =   8535
      Begin VB.TextBox txtErrorIntroLine 
         Height          =   1215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   8295
      End
      Begin VB.CheckBox chkEnableCmdBtn 
         Caption         =   "chkEnableCmdBtn"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   1560
         Width           =   4695
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   1920
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmWarningMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkEnableCmdBtn_Click()
    If chkEnableCmdBtn.Value = 1 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub cmdOK_Click()
    Unload frmWarningMsg
End Sub

