VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCrypKeyAuth 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ProFume* Gas Fumigant Fumiguide* - Please Enter Your Site Key"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   4680
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox rtfMessage 
      Height          =   1095
      Left            =   120
      TabIndex        =   23
      Top             =   3480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1931
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmCrypKeyAuth.frx":0000
   End
   Begin VB.CommandButton btnValidate 
      Caption         =   "Validate"
      Height          =   375
      Left            =   5880
      TabIndex        =   22
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtSiteKey 
      Height          =   285
      Left            =   1080
      TabIndex        =   21
      Top             =   3000
      Width           =   4455
   End
   Begin VB.TextBox txtSiteCode 
      BackColor       =   &H80000013&
      CausesValidation=   0   'False
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2640
      Width           =   4455
   End
   Begin VB.TextBox txtCountry 
      Height          =   285
      Left            =   6240
      TabIndex        =   17
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtZIP 
      Height          =   285
      Left            =   4440
      TabIndex        =   16
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   3360
      TabIndex        =   15
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   720
      TabIndex        =   14
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtAddress3 
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Top             =   1680
      Width           =   6615
   End
   Begin VB.TextBox txtAddress2 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   1320
      Width           =   6615
   End
   Begin VB.TextBox txtAddress1 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Top             =   960
      Width           =   6615
   End
   Begin VB.TextBox txtCompany 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   600
      Width           =   6615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label lblSiteKey 
      Caption         =   "Site Key"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblSiteCode 
      Caption         =   "Site Code"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblCountry 
      Caption         =   "Country"
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblZIP 
      Caption         =   "ZIP"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblState 
      Caption         =   "State"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCity 
      Caption         =   "City"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblAddress3 
      Caption         =   "Address 3"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblAddress2 
      Caption         =   "Address 2"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblAddress1 
      Caption         =   "Address 1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblCompany 
      Caption         =   "Company"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmCrypKeyAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------
'*
'* Filename: frmCrypKeyAuth.frm
'*
'* DESCRIPTION
'* -----------
'* Code for the form which requests user information
'*
'* Copyright (c) 2002 Accenture.  All rights reserved.
'-----------------------------------------------------
'*
'* REVISION HISTORY
'* Revision:    0
'* Author:      Gamboa, Terence S. <NAI0797>
'* Date:        Octoboer 7, 2002
'*
'-----------------------------------------------------

' stores the current machine's site code
Dim sSiteCode As String

' used in error handler
Const MODULE_NAME = "frmCrypKeyAuth.frm"


'**************************************************************
'   DESCRIPTION
'   Unloads the current form when the close button is clicked
'
'   PARAMETERS
'       none
'
'   RETURNS -
'       none
'
'   REVISION LOG
'       DATE      NAME                CHANGE
'  10/2/2002      Jett Gamboa         Initial Revision
'  03/4/2004      Patrick San Miguel  Modify the code so that it will terminate the application
'**************************************************************
Private Sub btnClose_Click()
       
    Const FUNCTION_NAME = "btnClose_Click()"

    Dim sErrorDesc As String
    On Error GoTo Error_Handler
       
    Unload frmCrypKeyAuth
    'Pat 3/4/04: Terminating the application
    End
    
    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Sub



'**************************************************************
'   DESCRIPTION
'   Validates teh CrypKey Site Key entered for the sitecode by
'   calling the CrypKey SDK functions
'
'   PARAMETERS
'       none
'
'   RETURNS -
'       none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/2/2002       Jett Gamboa       Initial Revision
'**************************************************************
Private Sub btnValidate_Click()

    Const FUNCTION_NAME = "btnValidate_Click()"

    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    Dim sSiteKey As String
    Dim nErrorCode As Integer
    
    sSiteKey = Trim(CStr(txtSiteKey.Text))

    ' Save the site key. If it is invalid return an error message
    nErrorCode = objCrypKey.SaveSiteKey(sSiteKey)

    If nErrorCode < 0 Then
    
        MsgBox LoadResString(IDS_SITEKEY_INVALID), vbExclamation
        
    Else
        
        nErrorCode = objCrypKey.GetAuthorization2(1)
        gbCrypKeyPassed = True
        
        ' unload the form
        Unload Me
    End If

    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'   initializes the CrypKey form which prompts the user for
'   information necessary to authenticate
'
'   PARAMETERS
'       none
'
'   RETURNS -
'       none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/2/2002       Jett Gamboa       Initial Revision
'**************************************************************
Private Sub Form_Load()

    Const FUNCTION_NAME = "Form_Load()"

    Dim sErrorDesc As String
    On Error GoTo Error_Handler
    
    
    With frmCrypKeyAuth
        .lblName = LoadResString(IDS_REG_NAME)
        .lblAddress1 = LoadResString(IDS_REG_ADDRESS1)
        .lblAddress2 = LoadResString(IDS_REG_ADDRESS2)
        .lblAddress3 = LoadResString(IDS_REG_ADDRESS3)
        .lblCity = LoadResString(IDS_REG_CITY)
        .lblCompany = LoadResString(IDS_REG_COMPANY)
        .lblCountry = LoadResString(IDS_REG_COUNTRY)
        .lblName = LoadResString(IDS_CK_NAME)
        .lblSiteCode = LoadResString(IDS_CK_SITECODE)
        .lblSiteKey = LoadResString(IDS_CK_SITEKEY)
        .lblState = LoadResString(IDS_CK_STATE)
        .lblZIP = LoadResString(IDS_CK_ZIP)
    End With

    ' retrieve the site code for this machine
    Call objCrypKey.GetSiteCode(sSiteCode)
      
    ' disable the validate button
    btnValidate.Enabled = False

    'Populate the fields from the registry. default to blank if there is nothing
    
    txtName.Text = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_NAME)
    txtCompany.Text = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_COMPANY)
    txtAddress1.Text = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_ADDRESS1)
    txtAddress2.Text = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_ADDRESS2)
    txtAddress3.Text = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_ADDRESS3)
    txtCity.Text = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_CITY)
    txtState.Text = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_STATE)
    txtZIP.Text = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_ZIP)
    txtCountry.Text = regQuery_A_Key(HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_COUNTRY)
       
    txtSiteCode.Text = sSiteCode

    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'   Unloads the current form and save data in the fields that the
'   user entered
'
'   PARAMETERS
'       none
'
'   RETURNS -
'       none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/2/2002       Jett Gamboa       Initial Revision
'**************************************************************
Private Sub Form_Unload(Cancel As Integer)

    Const FUNCTION_NAME = "Form_Unload()"

    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    ' Save values of the form to the registry
        
    regCreate_Key_Value HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_NAME, txtName.Text
    regCreate_Key_Value HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_COMPANY, txtCompany.Text
    regCreate_Key_Value HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_ADDRESS1, txtAddress1.Text
    regCreate_Key_Value HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_ADDRESS2, txtAddress2.Text
    regCreate_Key_Value HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_ADDRESS3, txtAddress3.Text
    regCreate_Key_Value HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_CITY, txtCity.Text
    regCreate_Key_Value HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_STATE, txtState.Text
    regCreate_Key_Value HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_ZIP, txtZIP.Text
    regCreate_Key_Value HKEY_CURRENT_USER, REG_KEY_PATH, SETTING_VAL_COUNTRY, txtCountry.Text

    ' If site key is empty or CrypKey authentication failed
    ' prompt the user.
    
    If Trim(txtSiteKey.Text) = "" Or Not gbCrypKeyPassed Then
    
        If MsgBox(LoadResString(IDS_SITEKEY_NOT_VALIDATED), vbExclamation + vbYesNo) = vbNo Then
                
            Cancel = True
                
        End If
        
    End If

    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'   enables the validate button if the site key field is not
'   empty
'
'   PARAMETERS
'       none
'
'   RETURNS -
'       none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/2/2002       Jett Gamboa       Initial Revision
'**************************************************************
Private Sub txtSiteKey_Change()

    Const FUNCTION_NAME = "txtSiteKey_Change()"

    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    ' disable the Validate button if the user did not enter anything in the site key
    If Trim(txtSiteKey.Text) = "" Then
        btnValidate.Enabled = False
    Else
        btnValidate.Enabled = True
    End If

    Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub
