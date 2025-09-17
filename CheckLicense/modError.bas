Attribute VB_Name = "modError"
'**************************************************************
'   SUB NAME: psAbort
'
'   DESCRIPTION
'       Loads the frmAppError
'
'   PARAMETERS
'       strForm - Form Name during error
'       strFunction - Function Name during error
'       strAddInfo - Additional Info regarding the error if any
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/23/2002       Davis Grana        Initial Revision
'**************************************************************
Public Sub psAbort(strForm As String, strFunction As String, strAddInfo As String)
    
    frmAppError.SetForm = strForm
    frmAppError.SetFunction = strFunction
    frmAppError.SetAddInfo = strAddInfo
    frmMain.txtOutput.text = frmMain.txtOutput.text & strForm & ":" & strFunction & ":" & strAddInfo
    frmAppError.Show vbModal

End Sub
