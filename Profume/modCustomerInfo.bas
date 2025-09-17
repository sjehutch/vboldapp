Attribute VB_Name = "modCustomerInfo"
' JETT HAS THOUROUgHLY REVIEWED AND OPTIMIZED THIS MODULE


'**************************************************************
'   DESCRIPTION
'       Customer Info Form
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/20/2002       Sheryll Go        Initial Revision
'  10/11/02        Cel Gamboa        Resized the controls to fit the Main Form
'                                    Set the MultiLine property of the Notes Textbox
'                                     to enable "word wrapping"
                                     
'**************************************************************

Option Explicit
Dim rsPriContact As ADODB.Recordset
Dim rsSecContact As ADODB.Recordset
Dim rsAddress As ADODB.Recordset

Dim sqlPriContact As String
Dim sqlSecContact As String
Dim sqlAddress As String
Dim bInvalid As Boolean

Const IDS_CUSTOMER_COMPANY = "The Customer Company may not be empty"
Const MODULE_NAME = "Customer Info"

Private Sub btnNextTab_Click()
    Const FUNCTION_NAME = "btnNextTab_Click()"
    
    On Error GoTo Error_Handler
    
   gbCanChangeTabs = True
    'Call function from the main program to switch to the next tab
    NextTab
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in Next Tab"
    
End Sub


Public Sub cboCountry_Click_Handler()
    Dim sValue As String
    Dim sCountry As String
    
    gbHasChanged = True
    sValue = frmMain.cboCountry.Text
    sCountry = fgetCountryCode(sValue)
    If sCountry <> "USA" Then
        frmMain.cboState.Clear
        frmMain.cboState.Enabled = False
    Else
    'PSM 1-23-2004: Modify the code so that whenever we are changing tabs
    '               we can still see the value of the state.
        If frmMain.cboState.ListCount = 0 Then
            PopulateState
            frmMain.cboState.Enabled = True
        End If
    End If
End Sub

Public Sub cboCountry_Validate_Handler(Cancel As Boolean)
    Dim sValue As String
    Dim sCountry As String
    Const FUNCTION_NAME = "cboCountry_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = frmMain.cboCountry.Text
    sCountry = fgetCountryCode(sValue)
    rsAddress!Country = sCountry
    rsAddress.Update
    
    gbCanChangeTabs = True
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Country."
    
End Sub


Public Sub cboState_Click_Handler()
    gbHasChanged = True
End Sub

Public Sub cboState_Validate_Handler(Cancel As Boolean)
    Dim sValue As String
    Dim sState As String
    Const FUNCTION_NAME = "cboState_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = frmMain.cboState.Text
    sState = fgetStateCode(sValue)
    If sState = "" Then
        rsAddress!State = ""
    Else
        rsAddress!State = sState
    End If
    rsAddress.Update
    
   gbCanChangeTabs = True
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving State."
    
End Sub


'**************************************************************
'   DESCRIPTION
'       When form loads, connect to the database, populate State
'       and Country drop-down boxes and load existing customer info
'
'   PARAMETERS
'       None
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/20/2002       Sheryll Go        Initial Revision
'  10/16/2002      Joseph Manyaga    Added a code that would reset the value of
'                                    gbHasChanged
'**************************************************************
'Private Sub Form_Load()
Public Sub Load_CustomerInfo()
    Dim bPrevHasChanged As Boolean
    
    Const FUNCTION_NAME = "CustomerInfo_Load()"

    On Error GoTo Error_Handler
    
    Set rsPriContact = New ADODB.Recordset
    Set rsSecContact = New ADODB.Recordset
    Set rsAddress = New ADODB.Recordset

    'We need to store the original value of gbHasChanged so that we would
    'be able to protect this from the changes due to the form initialization.
    bPrevHasChanged = gbHasChanged
    
    'Call sub CustInfoLoadResString to load string captions in all the controls in this form
    CustInfoLoadResString
    
    PopulateState   'populate state combo box
    PopulateCountry     'populate country combo box
    sLoadCustomerInfo   'load existing customer info (if there's any)
    
    'Reset the global variable gbHasChanged
    gbHasChanged = bPrevHasChanged
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in loading the form."
    
End Sub
'**************************************************************
'   DESCRIPTION
'       Load resource strings in the form
'
'   PARAMETERS
'       None
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/20/2002       Sheryll Go        Initial Revision
'**************************************************************
Private Sub CustInfoLoadResString()
    Const FUNCTION_NAME = "CustInfoLoadResString()"
    
    On Error GoTo Error_Handler
    
    With frmMain
        
        .fraAddInfo.Caption = LoadResString(IDS_FRA_ADDRESSINFO)
        .fraContactInfo.Caption = LoadResString(IDS_FRA_CONTACTINFO)
        .fraNotes.Caption = LoadResString(IDS_FRA_NOTES)
        .fraCustomerInfo.Caption = LoadResString(IDS_FRM_CUSTINFO)
        
    
        .lblCompany.Caption = LoadResString(IDS_COMPANY)
        .lblAddress1.Caption = LoadResString(IDS_ADDRESS1)
        .lblAddress2.Caption = LoadResString(IDS_ADDRESS2)
        .lblAddress3.Caption = LoadResString(IDS_ADDRESS3)
        .lblCity.Caption = LoadResString(IDS_PRINT_CITY)
        .lblState.Caption = LoadResString(IDS_STATE_PROVINCE)
        .lblZIP.Caption = LoadResString(IDS_ZIP_POSTAL)
        .lblCountry.Caption = LoadResString(IDS_PRINT_COUNTRY)
        .lblPrimary.Caption = LoadResString(IDS_PRI_CONTACTINFO)
        .lblPrimary.Font.Underline = True
        .lblSecondary.Caption = LoadResString(IDS_SEC_CONTACTINFO)
        .lblSecondary.Font.Underline = True
        .lblName.Caption = LoadResString(IDS_NAME)
        .lblPhone.Caption = LoadResString(IDS_PHONE)
        .lblEmail.Caption = LoadResString(IDS_EMAIL)
        .lblNotes.Caption = LoadResString(IDS_NOTES)
        .cmdNextStep1.Caption = LoadResString(IDS_BTN_NEXTSTEP)
        
    End With
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in loading resource strings."
    
End Sub
'**************************************************************
'   DESCRIPTION
'       Populate State drop-down box
'
'   PARAMETERS
'       None
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/20/2002       Sheryll Go        Initial Revision
'**************************************************************
Private Sub PopulateState()
    Dim rsState As ADODB.Recordset
    Dim sqlState As String
    Const FUNCTION_NAME = "PopulateState()"
    
    On Error GoTo Error_Handler
    
    Set rsState = New ADODB.Recordset
    
    sqlState = "Select Code, Name from States"
    rsState.Open sqlState, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    'JNM 05/02/03 - PT fix
    frmMain.cboState.Clear
    
    While rsState.EOF <> True
        frmMain.cboState.AddItem rsState!Name
        rsState.MoveNext
    Wend
    
    rsState.Close
    Set rsState = Nothing
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in populating State."
    
End Sub
'**************************************************************
'   DESCRIPTION
'       Populate Country drop-down box
'
'   PARAMETERS
'       None
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/20/2002       Sheryll Go        Initial Revision
'**************************************************************
Private Sub PopulateCountry()
    Dim rsCountry As ADODB.Recordset
    Dim sqlCountry As String
    Const FUNCTION_NAME = "PopulateCountry()"
    
    On Error GoTo Error_Handler
    
    Set rsCountry = New ADODB.Recordset
    
    sqlCountry = "Select Code, Name from Countries order by Name"
    rsCountry.Open sqlCountry, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    'JNM 05/02/03 - PT fix - Reinit the control
    frmMain.cboCountry.Clear
    
    While rsCountry.EOF <> True
        frmMain.cboCountry.AddItem rsCountry!Name
        rsCountry.MoveNext
    Wend
    
    rsCountry.Close
    Set rsCountry = Nothing
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in populating Country."
    
End Sub
'**************************************************************
'   DESCRIPTION
'       Loads existing customer info
'
'   PARAMETERS
'       None
'
'   RETURNS
'   None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/20/2002       Sheryll Go        Initial Revision
'**************************************************************
Private Sub sLoadCustomerInfo()
    Dim sStateCode As String
    Dim sState As String
    Dim sCountryCode As String
    Dim sCountry As String
    Const FUNCTION_NAME = "sLoadCustomerInfo()"
    
    On Error GoTo Error_Handler
 
    gbCanChangeTabs = True

    sqlPriContact = "Select * from Contact where primary = True"
    sqlSecContact = "Select * from Contact where primary = False"
    sqlAddress = "Select id, company, address1, address2, address3, city, state, zip, country, notes from Job"
    rsPriContact.Open sqlPriContact, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    rsSecContact.Open sqlSecContact, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    rsAddress.Open sqlAddress, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    With frmMain
    
        'Loading Address Info
        'recordset should ALWAYS contain records
        If rsAddress.EOF <> True Then
            If rsAddress!company <> "" Then
                .txtCompany.Text = rsAddress!company
            Else
                .txtCompany.Text = ""
            End If
            
            If rsAddress!address1 <> "" Then
                .txtAddress1.Text = rsAddress!address1
            Else
                .txtAddress1.Text = ""
            End If
            
            If rsAddress!address2 <> "" Then
                .txtAddress2.Text = rsAddress!address2
            Else
                .txtAddress2.Text = ""
            End If
            
            If rsAddress!address3 <> "" Then
                .txtAddress3.Text = rsAddress!address3
            Else
                .txtAddress3.Text = ""
            End If
            
            If rsAddress!City <> "" Then
                .txtCity.Text = rsAddress!City
            Else
                .txtCity.Text = ""
            End If
            
            If rsAddress!Zip <> "" Then
                .txtZIP.Text = rsAddress!Zip
            Else
                .txtZIP.Text = ""
            End If
            
            If rsAddress!State <> "" Then
                'get name for State
                sStateCode = rsAddress!State
                sState = fgetStateName(sStateCode)
                .cboState.Text = sState
            End If
            
            'get name for Country
            sCountryCode = rsAddress!Country
            sCountry = fgetCountryName(sCountryCode)
            .cboCountry.Text = sCountry
            
            If rsAddress!Notes <> "" Then
                .txtNotes.Text = rsAddress!Notes
            Else
                .txtNotes.Text = ""
            End If
        Else
            'call error handler
            psAbort MODULE_NAME, FUNCTION_NAME, "The recordset rsAddress is empty."
        End If
        
        'Loading Contact Info
        'for primary contact
        If rsPriContact.EOF <> True Then
            If rsPriContact!Name <> "" Then
                .txtContactName1.Text = rsPriContact!Name
            Else
                .txtContactName1.Text = ""
            End If
                
            If rsPriContact!Phone <> "" Then
                .txtContactPhone1.Text = rsPriContact!Phone
            Else
                .txtContactPhone1.Text = ""
            End If
                
            If rsPriContact!Email <> "" Then
                .txtContactEmail1.Text = rsPriContact!Email
            Else
                .txtContactEmail1.Text = ""
            End If
        Else
            'call error handler
            psAbort MODULE_NAME, FUNCTION_NAME, "The recordset rsPriContact is empty."
        End If
            
        'for secondary contact
        If rsSecContact.EOF <> True Then
            If rsSecContact!Name <> "" Then
                .txtContactName2.Text = rsSecContact!Name
            Else
                .txtContactName2.Text = ""
            End If
                
            If rsSecContact!Phone <> "" Then
                .txtContactPhone2.Text = rsSecContact!Phone
            Else
                .txtContactPhone2.Text = ""
            End If
                
            If rsSecContact!Email <> "" Then
                .txtContactEmail2.Text = rsSecContact!Email
            Else
                .txtContactEmail2.Text = ""
            End If
        Else
            'call error handler
            psAbort MODULE_NAME, FUNCTION_NAME, "The recordset rsSecContact is empty."
        End If
    
    End With
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in loading customer information."
 
End Sub
'**************************************************************
'   DESCRIPTION
'       Converts state code given state name for display purposes
'
'   PARAMETERS
'       strStateCode
'
'   RETURNS
'       State Name
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/20/2002       Sheryll Go        Initial Revision
'**************************************************************
Private Function fgetStateName(sStateCode) As String
    Dim rsStateName As ADODB.Recordset
    Dim sqlStateName As String
    Const FUNCTION_NAME = "fgetStateName()"
    
    On Error GoTo Error_Handler
    
    Set rsStateName = New ADODB.Recordset
    
    sqlStateName = "Select Name from States where Code = '" & sStateCode & "'"
    rsStateName.Open sqlStateName, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    fgetStateName = rsStateName!Name
    
    'close recordset
    rsStateName.Close
    Set rsStateName = Nothing
    
Exit Function
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in getting the State Name."
    
End Function
'**************************************************************
'   DESCRIPTION
'       Converts country code given country name for display purposes
'
'   PARAMETERS
'       strCountryCode
'
'   RETURNS
'       Country Name
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/20/2002       Sheryll Go        Initial Revision
'**************************************************************
Private Function fgetCountryName(sCountryCode) As String
    Dim rsCountryName As ADODB.Recordset
    Dim sqlCountryName As String
    Const FUNCTION_NAME = "fgetCountryName()"
    
    On Error GoTo Error_Handler
    
    Set rsCountryName = New ADODB.Recordset
    
    sqlCountryName = "Select Name from Countries where Code = '" & sCountryCode & "'"
    rsCountryName.Open sqlCountryName, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    fgetCountryName = rsCountryName!Name
    
    'close recordset
    rsCountryName.Close
    Set rsCountryName = Nothing
    
Exit Function
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in getting the Country Name."
    
End Function
'**************************************************************
'   DESCRIPTION
'       When form unloads, validations are done for company field.
'       Saves data to the database
'
'   PARAMETERS
'       None
'
'   RETURNS
'       None
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/20/2002       Sheryll Go        Initial Revision
'  4/24/2003       Amana Martinez    Fumiguide Enhancements 2003 AT
'**************************************************************
Public Sub Unload_CustomerInfo()
    Dim sCompany As String
    
    '4/24/2003 - AM Removed the variables below as they will not be needed anymore for AT Defect 8496'
    'Dim sAddress1 As String
    'Dim sAddress2 As String
    'Dim sAddress3 As String
    'Dim sCity As String
    'Dim sState As String
    'Dim sZip As String
    'Dim sCountry As String
    'Dim sName1 As String
    'Dim sName2 As String
    'Dim sPhone1 As String
    'Dim sPhone2 As String
    'Dim sEmail1 As String
    'Dim sEmail2 As String
    'Dim sNotes As String
    '4/24/2003 - AM Removed the variables above as they will not be needed anymore for AT Defect 8496'
    
    Const FUNCTION_NAME = "CustomerInfo_Unload()"
    
    On Error GoTo Error_Handler
   
    With frmMain
        'get values for fields
        sCompany = Trim$(.txtCompany.Text)
        
        '4/24/2003 - AM Commented below for AT Defect 8496 Fumiguide Enhancement 2003
        'sAddress1 = Trim$(.txtAddress1.Text)
        'sAddress2 = Trim$(.txtAddress2.Text)
        'sAddress3 = Trim$(.txtAddress3.Text)
        'sCity = Trim$(.txtCity.Text)
        'sState = .cboState.Text
        'sZip = Trim$(.txtZip.Text)
        'sCountry = .cboCountry.Text
        'sName1 = Trim$(.txtContactName1.Text)
        'sName2 = Trim$(.txtContactName2.Text)
        'sPhone1 = Trim$(.txtContactPhone1.Text)
        'sPhone2 = Trim$(.txtContactPhone2.Text)
        'sEmail1 = Trim$(.txtContactEmail1.Text)
        'sEmail2 = Trim$(.txtContactEmail2.Text)
        'sNotes = Trim$(.txtNotes.Text)
        '4/24/2003 - AM Commented above for AT Defect 8496 Fumiguide Enhancement 2003
        
        If sCompany = "" Then
            MsgBox IDS_CUSTOMER_COMPANY, vbExclamation
            .txtCompany.SetFocus
            gbCanChangeTabs = False
        Else
        
            '4/24/2003 - AM Commented below for AT Defect 8496 Fumiguide Enhancement 2003
            ''save address info
            'rsAddress!company = sCompany
            'rsAddress!address1 = sAddress1
            'rsAddress!address2 = sAddress2
            'rsAddress!address3 = sAddress3
            'rsAddress!City = sCity
            'rsAddress!Zip = sZip
            'rsAddress!Notes = sNotes
            'sState = fgetStateCode(sState)
            
            'If sState = "" Then
            '    rsAddress!State = ""
            'Else
            '    rsAddress!State = sState
            'End If
            
            'sCountry = fgetCountryCode(sCountry)
            'rsAddress!Country = sCountry
            'rsAddress.Update
             
            'save contact info
            'rsPriContact!Name = sName1
            'rsSecContact!Name = sName2
            'rsPriContact!Phone = sPhone1
            'rsSecContact!Phone = sPhone2
            'rsPriContact!Email = sEmail1
            'rsSecContact!Email = sEmail2
            'rsPriContact.Update
            'rsSecContact.Update
            '4/24/2003 - AM Commented above for AT Defect 8496 Fumiguide Enhancement 2003
            gbCanChangeTabs = True

            '4/24/2003 - AM Commented above for AT Defect 8496 Fumiguide Enhancement 2003
           
            '' close recordsets only if we can change tabs (and do not need them)
            'Set rsPriContact = Nothing
            'Set rsSecContact = Nothing
            'Set rsAddress = Nothing
            '4/24/2003 - AM Commented above for AT Defect 8496 Fumiguide Enhancement 2003
           
        End If
    End With
   
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in unloading form."
       
End Sub

Public Sub txtAddress1_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtAddress1_Validate_Handler(Cancel As Boolean)
    Dim sValue As String
    Const FUNCTION_NAME = "txtAddress1_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtAddress1.Text)
    rsAddress!address1 = sValue
    rsAddress.Update
    
    gbCanChangeTabs = True
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Address 1."
    
End Sub

Public Sub txtAddress2_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtAddress2_Validate_Handler(Cancel As Boolean)
    Dim sValue As String
    Const FUNCTION_NAME = "txtAddress2_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtAddress2.Text)
    rsAddress!address2 = sValue
    rsAddress.Update
    
    gbCanChangeTabs = True

Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Address 2."
    
End Sub

Public Sub txtAddress3_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtAddress3_Validate_Handler(Cancel As Boolean)
    Dim sValue As String
    Const FUNCTION_NAME = "txtAddress3_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtAddress3.Text)
    rsAddress!address3 = sValue
    rsAddress.Update
    
    gbCanChangeTabs = True
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Address 3."
    
End Sub

Public Sub txtCity_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtCity_Validate_Handler(Cancel As Boolean)

    Dim sValue As String
    Const FUNCTION_NAME = "txtCity_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtCity.Text)
    rsAddress!City = sValue
    rsAddress.Update
    
    gbCanChangeTabs = True
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving City."
    
End Sub

Public Sub txtCompany_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtCompany_Validate_Handler(Cancel As Boolean)
    
    If Not bInvalid Then
        ValidateCompany
    Else
        frmMain.txtCompany.SetFocus
    End If
    
End Sub
'**************************************************************
'   DESCRIPTION
'       Converts state name to state code for saving purposes
'
'   PARAMETERS
'       strStateName
'
'   RETURNS
'       State Code
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/20/2002       Sheryll Go        Initial Revision
'**************************************************************
Private Function fgetStateCode(sStateName) As String

    Dim rsStateCode As ADODB.Recordset
    Dim sqlStateCode As String
    Const FUNCTION_NAME = "fgetStateCode()"
    
    On Error GoTo Error_Handler
    
    Set rsStateCode = New ADODB.Recordset
    
    sqlStateCode = "Select Code from States where Name = '" & sStateName & "'"
    rsStateCode.Open sqlStateCode, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    If frmMain.cboState.Text = "" Then
        fgetStateCode = ""
    Else
        fgetStateCode = rsStateCode!code
    End If
    
    'close recordset
    rsStateCode.Close
    Set rsStateCode = Nothing
    
Exit Function
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in getting the State Code."
    
End Function


'**************************************************************
'   DESCRIPTION
'       Converts country name to country code for saving purposes
'
'   PARAMETERS
'       strCountryName
'
'   RETURNS
'       Country Code
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/20/2002       Sheryll Go        Initial Revision
'**************************************************************
Private Function fgetCountryCode(sCountryName) As String

    Dim rsCountryCode As ADODB.Recordset
    Dim sqlCountryCode As String
    Const FUNCTION_NAME = "fgetCountryCode()"
    
    On Error GoTo Error_Handler
    
    Set rsCountryCode = New ADODB.Recordset
    
    sqlCountryCode = "Select Code from Countries where Name = '" & sCountryName & "'"
    rsCountryCode.Open sqlCountryCode, gdbConn, adOpenStatic, adLockOptimistic, adCmdText
    
    fgetCountryCode = rsCountryCode!code
    
    'close recordset
    rsCountryCode.Close
    Set rsCountryCode = Nothing
    
Exit Function
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in getting the Country Code."
    
End Function

Public Sub txtCompany_LostFocus_Handler()
    bInvalid = False
    ValidateCompany
End Sub

Public Sub txtContactEmail1_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtContactEmail1_Validate_Handler(Cancel As Boolean)
    Dim sValue As String
    Const FUNCTION_NAME = "txtContactEmail1_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtContactEmail1.Text)
    rsPriContact!Email = sValue
    rsPriContact.Update
    
    gbCanChangeTabs = True
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Primary Contact Email."
    
End Sub

Public Sub txtContactEmail2_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtContactEmail2_Validate_Handler(Cancel As Boolean)
    Dim sValue As String
    Const FUNCTION_NAME = "txtContactEmail2_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtContactEmail2.Text)
    rsSecContact!Email = sValue
    rsSecContact.Update
    
    gbCanChangeTabs = True
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Secondary Contact Email."
    
End Sub

Public Sub txtContactName1_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtContactName1_Validate_Handler(Cancel As Boolean)

    Dim sValue As String
    Const FUNCTION_NAME = "txtContactName1_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtContactName1.Text)
    rsPriContact!Name = sValue
    rsPriContact.Update
    
    gbCanChangeTabs = True
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Primary Contact Name."
    
End Sub

Public Sub txtContactName2_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtContactName2_Validate_Handler(Cancel As Boolean)

    Dim sValue As String
    Const FUNCTION_NAME = "txtContactName2_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtContactName2.Text)
    rsSecContact!Name = sValue
    rsSecContact.Update
    
    gbCanChangeTabs = True

Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Secondary Contact Name."
    
End Sub

Public Sub txtContactPhone1_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtContactPhone1_Validate_Handler(Cancel As Boolean)
    Dim sValue As String
    Const FUNCTION_NAME = "txtContactPhone1_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtContactPhone1.Text)
    rsPriContact!Phone = sValue
    rsPriContact.Update
   gbCanChangeTabs = True
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Primary Contact Phone."
    
End Sub

Public Sub txtContactPhone2_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtContactPhone2_Validate_Handler(Cancel As Boolean)

    Dim sValue As String
    Const FUNCTION_NAME = "txtContactPhone2_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtContactPhone2.Text)
    rsSecContact!Phone = sValue
    rsSecContact.Update
    
    gbCanChangeTabs = True
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Secondary Contact Phone."
    
End Sub

Public Sub txtNotes_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtNotes_Validate_Handler(Cancel As Boolean)

    Dim sValue As String
    Const FUNCTION_NAME = "txtNotes_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtNotes.Text)
    rsAddress!Notes = sValue
    rsAddress.Update
    
    gbCanChangeTabs = True
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Notes."
    
End Sub

Public Sub txtZip_Change_Handler()
    gbHasChanged = True
End Sub

Public Sub txtZip_Validate_Handler(Cancel As Boolean)

    Dim sValue As String
    Const FUNCTION_NAME = "txtZip_Validate_Handler(Cancel as Boolean)"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtZIP.Text)
    rsAddress!Zip = sValue
    rsAddress.Update
    
    gbCanChangeTabs = True
    
Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Zip."
    
End Sub

Public Sub ValidateCompany()

    Dim sValue As String
    Const FUNCTION_NAME = "txtCompany_LostFocus()"
    
    On Error GoTo Error_Handler
    
    sValue = Trim$(frmMain.txtCompany.Text)
    
    If sValue = "" Then
        If Not bInvalid Then
            MsgBox IDS_CUSTOMER_COMPANY, vbExclamation
        End If
        'TODO: fix this bug right here
        ReturnCurrTab TAB_CUSTOMER_INFO
        frmMain.txtCompany.SetFocus
        gbCanChangeTabs = False
        bInvalid = True
    Else
        rsAddress!company = sValue
        rsAddress.Update
        
        gbCanChangeTabs = True
        bInvalid = False
    End If

Exit Sub
    
Error_Handler:
    ' untrapped error, call the critical error handler function
    psAbort MODULE_NAME, FUNCTION_NAME, "Error in saving Company."

End Sub

