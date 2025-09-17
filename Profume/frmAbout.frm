VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1032"
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3960
      TabIndex        =   4
      Tag             =   "1034"
      Top             =   120
      Width           =   1230
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   240
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   510
   End
   Begin VB.Label lbldpcalc 
      Caption         =   "dpcalc Version"
      Height          =   225
      Left            =   1080
      TabIndex        =   5
      Tag             =   "1036"
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblDescription 
      Caption         =   "ProApp Description"
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1080
      TabIndex        =   3
      Tag             =   "1038"
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   2
      Tag             =   "1037"
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   1
      Tag             =   "1036"
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------
'*
'* Filename: frmAbout.frm
'*
'* DESCRIPTION
'* -----------
'* About form
'*
'* Copyright (c) 2002 Accenture.  All rights reserved.
'-----------------------------------------------------
'*
'* REVISION HISTORY
'* Revision:    0
'* Author:      Gamboa, Terence S. <NAI0797>
'* Date:        September 25, 2002
'*
'-----------------------------------------------------


'**************************************************************
'   DESCRIPTION
'       loads all of the constants from the resource file
'
'   PARAMETERS
'       none
'
'   RETURNS -
'       none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Private Sub Form_Load()

    Dim Calculator As DPCALCLib.ProFumeCalculator
    
    Set Calculator = New DPCALCLib.ProFumeCalculator
    
    frmAbout.Caption = LoadResString(IDS_ABOUT)
    lblTitle.Caption = LoadResString(IDS_DOW_AGRO)
    lblVersion.Caption = LoadResString(IDS_ABOUT_PROFUME) & " " & App.Major & "." & App.Minor & "." & App.Revision
    lblDescription.Caption = LoadResString(IDS_COPYRIGHT)
    picIcon.Picture = LoadResPicture(101, vbResIcon)
    
    lbldpcalc.Caption = LoadResString(IDS_ABOUT_DPCALC) & " " & Calculator.GetVersion()
    
    Set Calculator = Nothing
End Sub


'**************************************************************
'   DESCRIPTION
'       unloads the current form when the Ok button is clicked
'
'   PARAMETERS
'       none
'
'   RETURNS -
'       none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Private Sub cmdOK_Click()
        Unload Me
End Sub

