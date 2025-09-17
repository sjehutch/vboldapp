Attribute VB_Name = "modCrypKey"
Option Explicit

'------------------------------------------------------
'*
'* Filename: modCrypKey.bas
'*
'* DESCRIPTION
'* -----------
'* Module that handles all the CrypKey authentication
'* related calls
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

' used in error handler
Const MODULE_NAME = "modCrypKey.bas"
'[1-14-08] David Smith  Upgraded to CrytKey 7.0 old name was CrypKeySDK
Public objCrypKey As New CrypKeySDK7

Public gbCrypKeyPassed As Boolean


'**************************************************************
'   DESCRIPTION
'   Calls CrypKey SDK routine to initialize the CrypKey engine
'
'   PARAMETERS
'       none
'
'   RETURNS - true if the application has been CrypKey authenticated,
'       false if otherwise.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/2/2002       Jett Gamboa       Initial Revision
'**************************************************************
Function Initialize_CrypKey()

    ' These are the keys that have been assigned to us by CrypKey
    ' this is tied to the executable that we are using
    '[1-14-08] David Smith changed for update to CrypKey 7.0
    'Const MASTER_KEY As String = "e75f7352b11d0337fa03bd3902eb6d445f82d062624a491ffdd8f37556cdd2bd6caf7e674d7bb0835ab99334f4cc021a4ca880af947690e7ec0864034bde5b135c8013008ba8fb33f1dcf5f70488ee2a6d207e6cada27e5ac1b8aeb64bc68dcd8dfc98c22d9e51113386381bc94769330429ad2dc41bb8dd26f8fba691661c7d"
    'Const USER_KEY As String = "C4EF1ECA89F1C6ED1968EE83E6"
    Const MASTER_KEY As String = "2f63af80b5fb8d9d06003e2096431422ccef9107043acd1e25fc38882bd73eec77d0f1304f184d61ae5782d29d190fd59636bfdc44e068f929017c6753b4705e9b526be22cb25deb6232133d20b84d6e8d0c6044fcec427ab9c88ba6a9a5e0279481080412d13f6a4a3ded8513ecc1a76b27a419a25f126040592d2b6d14f0ff"
    Const USER_KEY As String = "C4EF1ECA89F1C6ED1968EE83E6"
    
    
    Const FUNCTION_NAME = "Initialize_CrypKey"
    frmMain.txtOutput.text = frmMain.txtOutput.text & "Entering " & FUNCTION_NAME & vbCrLf
 
    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    Dim nResult As Integer
    Dim sAppPath As String

    ' default to CrypKey not passing
    gbCrypKeyPassed = False

    ' I am hardcoding the executable name because the name
    ' is directly tied to our CrypKey license
    frmMain.txtOutput.text = frmMain.txtOutput.text & App.Path & vbCrLf
    sAppPath = "C:\Program Files\DAS\Fumiguide\ProFume.exe"
    'sAppPath = App.Path & "\ProFume.exe"
    frmMain.txtOutput.text = frmMain.txtOutput.text & sAppPath & vbCrLf
    
    frmMain.txtOutput.text = frmMain.txtOutput.text & "Calling InitCrypKey" & vbCrLf
    nResult = objCrypKey.InitCrypKey(sAppPath, MASTER_KEY, USER_KEY, 0, 0)
    frmMain.txtOutput.text = frmMain.txtOutput.text & "InitCrypKey Result: " & nResult & vbCrLf
    
    If nResult <> CKInOk Then
        ' get the error message description
        Call objCrypKey.ExplainErr2(CKInitializeError, nResult, sErrorDesc)
        
        frmMain.txtOutput.text = frmMain.txtOutput.text & "Explain Error: " & nResult & " : " & sErrorDesc & vbCrLf
        
        'MsgBox nResult & " " & LoadResString(IDS_MSG_CRYPKEY_FAILED) & vbCrLf & vbCrLf & sErrorDesc, vbCritical
        Initialize_CrypKey = False
    Else
        Initialize_CrypKey = True
    End If

    frmMain.txtOutput.text = frmMain.txtOutput.text & "Exiting " & FUNCTION_NAME & vbCrLf

    Exit Function

Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function

'**************************************************************
'   DESCRIPTION
'   Determines whether the current application has been authenticated
'   by CrypKey, if not it displays a form where the user can enter
'   information include Site Code used to unlock the application
'
'   PARAMETERS
'       none
'
'   RETURNS - true if the application has been CrypKey authenticated,
'       false if otherwise.
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  10/2/2002       Jett Gamboa       Initial Revision
'  12/17/2004      Jett Gamboa       Modified so we initialize objCrypKey here
'                                    instead of globally
'**************************************************************
Function CrypkeyAuthentication()

    Const FUNCTION_NAME = "CrypkeyAuthentication"
    frmMain.txtOutput.text = frmMain.txtOutput.text & "Entering " & FUNCTION_NAME & vbCrLf
    
    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    Dim nErrorCode As Integer
    
    Dim nAllowed As Integer
    Dim nUsed As Integer
    Dim nResult As Integer
    Dim nDaysLeft As Integer
    
    Dim sExpireWarning As String

    ' create an instance of the CrypKey object
    Set objCrypKey = New CrypKeySDK7

    gbCrypKeyPassed = False

    ' if we were unable to initialize CrypKey, display a critical
    ' warning to the user (we should not be here)
    frmMain.txtOutput.text = frmMain.txtOutput.text & "Start Initialize_CrypKey" & vbCrLf
    If Initialize_CrypKey() = False Then
        CrypkeyAuthentication = False
        frmMain.txtOutput.text = frmMain.txtOutput.text & "CrypkeyAuthentication= " & CrypkeyAuthentication & vbCrLf

        Exit Function
    End If
    
    frmMain.txtOutput.text = frmMain.txtOutput.text & "Entering GetAuthorization2" & vbCrLf

    nErrorCode = objCrypKey.GetAuthorization2(1)
    
    frmMain.txtOutput.text = frmMain.txtOutput.text & "GetAuthorization2 nErrorCode: " & nErrorCode & vbCrLf

    ' display the form if request for Authorization failed
    If nErrorCode < 0 Then

        Load frmCrypKeyAuth

        ' Load the strings from the resource file to the form
        With frmCrypKeyAuth
            .lblName = LoadResString(IDS_CK_NAME)
            .lblCompany = LoadResString(IDS_CK_COMPANY)
            .lblAddress1 = LoadResString(IDS_CK_ADDRESS1)
            .lblAddress2 = LoadResString(IDS_CK_ADDRESS2)
            .lblAddress3 = LoadResString(IDS_CK_ADDRESS3)
            .lblCity = LoadResString(IDS_CK_CITY)
            .lblState = LoadResString(IDS_CK_STATE)
            .lblZIP = LoadResString(IDS_CK_ZIP)
            .lblSiteCode = LoadResString(IDS_CK_SITECODE)
            .lblSiteKey = LoadResString(IDS_CK_SITEKEY)
            .rtfMessage.text = LoadResString(IDS_SITEKEY_INSTRUCTIONS)
        End With
        
        frmCrypKeyAuth.Show vbModal

        ' end CrypKey
        objCrypKey.EndCrypKey
        
        ' if we passed the CrypKey validation from the form
        If gbCrypKeyPassed Then
            
            ' display message that we passed
            MsgBox LoadResString(IDS_SITEKEY_VALIDATED), vbInformation
            
            CrypkeyAuthentication = True
            
        End If

    Else
    
        ' Jett 12/17/2004 - Added this section so that a user is warned if the CrypKey
        ' license is about to expire in X number of days.
    
        ' Check what limitations are in place. We ignore unlimited runs or limits
        ' based on number of runs
        frmMain.txtOutput.text = frmMain.txtOutput.text & "Entering Get1RestInfo" & vbCrLf
        nResult = objCrypKey.Get1RestInfo(1)
        frmMain.txtOutput.text = frmMain.txtOutput.text & "Get1RestInfo nResult: " & nResult & vbCrLf
        
        Select Case nResult
            'Case 0
            '    MsgBox "Unlimited Runs"
            Case 1
                ' User was limited by number of days
                nAllowed = objCrypKey.Get1RestInfo(2)
                nUsed = objCrypKey.Get1RestInfo(3)
                 
                nDaysLeft = nAllowed - nUsed
                
                ' Check the number of days left, if it is less than or equal to our
                ' limit, display a messagebox to warn the user
                If nDaysLeft <= CRYPKEY_DAYS_TILL_EXPIRE Then
                
                    sExpireWarning = LoadResString(IDS_CK_EXPIRE)
                    MsgBox FormatSWParams(sExpireWarning, CStr(nDaysLeft)), vbExclamation
                    
                End If
                 
            'Case 2
                ' User was limited by number of days
                'nAllowed = objCrypKey.Get1RestInfo(2)
                'nUsed = objCrypKey.Get1RestInfo(3)
                
                'MsgBox "Runs Limit: " & nAllowed & " Used: " & nUsed
                
        End Select
    
        ' Application is authorized so terminate CrypKey and proceed
        objCrypKey.EndCrypKey
        CrypkeyAuthentication = True
    
    End If

    Set objCrypKey = Nothing
    
    frmMain.txtOutput.text = frmMain.txtOutput.text & "Exiting " & FUNCTION_NAME & vbCrLf

    Exit Function

Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function


'**************************************************************
'   DESCRIPTION
'   Removes the CrypKey license for this user
'   No error checking since it will not matter if this fails
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  12/23/2004      Jett Gamboa       Initial revision
'**************************************************************
Sub RemoveCrypKeyLicense()

    Dim nErrorCode As Integer
    Dim nReturnCode As Integer
    Dim sConfirmCode As String

    On Error Resume Next

    ' Instantiate a CrypKey object
    Set objCrypKey = New CrypKeySDK7

    ' Initialize CrypKey first
    If Initialize_CrypKey Then
    
        ' Check if we are authorized
        nErrorCode = objCrypKey.GetAuthorization2(1)
        
        ' Only remove the license if this copy is authorized. We cannot remove
        ' a license if we did not grant it in the first place
        If nErrorCode = CKAuthOk Then
            
            ' Call the KillLicense method to remove the license
            nReturnCode = objCrypKey.KillLicense(sConfirmCode)
            
            ' We ignore errors for now. If this failed, some CrypKey error has occurred,
            ' should not be a problem.
            
        End If

    End If
    
    Set objCrypKey = Nothing

End Sub
