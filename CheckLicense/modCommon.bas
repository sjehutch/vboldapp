Attribute VB_Name = "modCommon"

Option Explicit

'**************************************************************
'   DESCRIPTION
'       Accepts a string with
'
'   PARAMETERS
'           pvsFormat - string with parameter codes (i.e. %1, %2)
'           pvsParams - string delimited by the "|" character. The value
'                       of each "field" in the delimited string is used to
'                       replace the string in the first parameter
'
'   RETURNS - formatted string with parameter
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Function FormatSWParams(pvsFormat As String, pvsParams As String)

    ' Declare variables
    Dim x As Integer
    Dim sTemp As String
    Dim sNewString As String
    Dim arrParameters As Variant
    
    ' Exit the function if no format string or parameters were passed
    If pvsFormat = "" Or pvsParams = "" Then
        FormatSWParams = ""
        Exit Function
    End If
    
    ' split the parameters string to an array
    arrParameters = Split(pvsParams, "|")
    
    ' Assign the original string to our temporary string
    sNewString = pvsFormat
    
    ' Loop through the parameters and replace the placeholders in the format string.
    ' We move through the parameters backward so that the substring is replaced
    ' correctly in instances where there are more than 9 parameters
    For x = UBound(arrParameters) To 0 Step -1
        sTemp = "%" & CStr(x + 1)
        sNewString = Replace(sNewString, sTemp, arrParameters(x))
    Next
    
    ' return the formatted string
    FormatSWParams = sNewString

End Function
Function RoundOff(vValue As Variant, nDecimal As Integer)
'**************************************************************
'   DESCRIPTION
'       Rounds-off dValue to nDecimal places
'       Created this since Word VB does not have Round function
'
'   PARAMETERS
'          NONE
'
'   RETURNS - None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'   May 2, 2003  Joseph Manyaga      Initial Revision
'   Feb 4, 2003  Pat San Miguel      Copied this function from the code of the Fumiguide .dot file
'                                    to have consistency in rounding off numbers in line with
'                                    the SIMS Ticket 2116768
'**************************************************************
    Dim dValue As Double
    Dim nFactor As Integer
    Dim x As Integer
    Dim sErrorString As String
            
    Const FUNCTION_NAME = "RoundOff()"
       
    On Error GoTo Error_Handler
    
    dValue = CDbl(vValue)
    nFactor = 1
    x = 1
    
    While x <= nDecimal
        nFactor = nFactor * 10
        x = x + 1
    Wend
    
    RoundOff = Int(dValue * nFactor + 0.5) / nFactor
    
    Exit Function

Error_Handler:
    sErrorString = FUNCTION_NAME & vbCrLf & vbCrLf & "Error: " & Err.Number & " - " & Err.Description
    MsgBox sErrorString, vbCritical, "Error"

End Function
