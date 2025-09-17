Attribute VB_Name = "modAPI"
Option Explicit

'------------------------------------------------------
'*
'* Filename: modAPI.mod
'*
'* DESCRIPTION
'* -----------
'* This module contains function declares for various
'* Win32 API routines used by the ProFume application.
'* It also includes certain wrappers for certain
'* routines.
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

' used in error handler
Const MODULE_NAME = "modAPI.bas"

' Type Declarations for GUID API
Private Type GUID
    lData1 As Long
    lData2 As Integer
    lData3 As Integer
    lData4(7) As Byte
End Type

' Registry Root key constants as defined in the API
Public Enum regRootKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum


' registry key types
Private Const REG_SZ As Long = 1                    ' null terminated string
Private Const REG_DWORD As Long = 4                 ' 32-bit number


' Registry API return constants
Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_NO_MORE_ITEMS As Long = 259

' Enumeration of Languages
Enum clLanguages                                  ' NOTE: Not all languages are included!
    eNEUTRAL = &H0        ' Language Neutral
    eAFRIKAANS = &H436    '"0436"="af;Afrikaans"
    eALBANIAN = &H41C     '"041C"="sq;Albanian"
    eARABIC = &H1         '"0001"="ar;Arabic"
    eBASQUE = &H42D       '"042D"="eu;Basque"
    eBULGARIAN = &H402    '"0402"="bg;Bulgarian"
    eBELARUSIAN = &H423   '"0423"="be;Belarusian"
    eCATALAN = &H403      '"0403"="ca;Catalan"
    eCHINEZE = &H4        '"0004"="zh;Chinese"
    eCROATIAN = &H41A     '"041A"="hr;Croatian"
    eCZECH = &H405        '"0405"="cs;Czech"
    eDANISH = &H406       '"0406"="da;Danish"
    eDUTCH = &H413        '"0413"="nl;Dutch (Standard)"
    eENGLISH = &H409      '"0409"="en-us;English (United States)"
    eESTONIAN = &H425     '"0425"="et;Estonian"
    eFAEROESE = &H438     '"0438"="fo;Faeroese"
    eFARSI = &H429        '"0429"="fa;Farsi"
    eFINNISH = &H40B      '"040B"="fi;Finnish"
    eFrench = &H40C       '"040C"="fr;French (Standard)"
    eGAELIC = &H43C       '"043C"="gd;Gaelic"
    eGERMAN = &H407       '"0407"="de;German (Standard)"
    eGREEK = &H408        '"0408"="el;Greek"
    eHEBREW = &H40D       '"040D"="he;Hebrew"
    eHINDI = &H439        '"0439"="hi;Hindi"
    eHUNGARIAN = &H40E    '"040E"="hu;Hungarian
    eICELANDIC = &H40F    '"040F"="is;Icelandic"
    eINDONESIAN = &H421   '"0421"="in;Indonesian"
    eITALIAN = &H410      '"0410"="it;Italian (Standard)"
    eJAPANESE = &H411     '"0411"="ja;Japanese"
    eKOREAN = &H412       '"0412"="ko;Korean"
    eLATVIAN = &H426      '"0426"="lv;Latvian"
    eLITHUANIAN = &H427   '"0427"="lt;Lithuanian"
    eMACEDONIAN = &H42F   '"042F"="mk;Macedonian"
    eMALAYSIAN = &H43E    '"043E"="ms;Malaysian"
    eMALTESE = &H43A      '"043A"="mt;Maltese"
    eNORWEGIAN = &H414    '"0414"="no;Norwegian (bokmal)"
    ePOLISH = &H415       '"0415"="pl;Polish"
    ePORTUGESE = &H816    '"0816"="pt;Portuguese (Standard)"
    eROMAINAN = &H418     '"0418"="ro;Romanian"
    eRUSSIAN = &H419      '"0419"="ru;Russian"
    eSERBIAN = &H81A      '"081A"="sr;Serbian (Latin)"
    eSERBIANC = &HC1A     '"0C1A"="sr;Serbian (Cyrillic)"
    eSLOVAK = &H41B       '"041B"="sk;Slovak"
    eSLOVENIAN = &H424    '"0424"="sl;Slovenian"
    eSPANISH = &H40A      '"040A"="es;Spanish"
    eSWEDISH = &H41D      '"041D"="sv;Swedish"
    eTHAI = &H41E         '"041E"="th;Thai"
    eTURKISH = &H41F      '"041F"="tr;Turkish"
    eUKRAINIAN = &H422    ''"0422"="uk;Ukrainian"
End Enum

' API type for determining OS platform type
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

'windows-defined type OSVERSIONINFO
Public Type OSVERSIONINFO
  OSVSize         As Long         'size, in bytes, of this data structure
  dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
  dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
  dwBuildNumber   As Long         'NT: build number of the OS
                                  'Win9x: build number of the OS in low-order word.
                                  '       High-order word contains major & minor ver nos.
  PlatformID      As Long         'Identifies the operating system platform.
  szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"
                                  'Win9x: 'arbitrary additional information'
End Type


' Constants for Html Help API
Public Const HH_DISPLAY_TOPIC = &H0


' variable to hold our return value for API Calls
Private lRetVal As Long
       
' API function declaration for GetGUID
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long


' API function declarations used by Registry functions
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngRootKey As Long) As Long
  
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String) As Long
  
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String) As Long
  
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
             lpType As Long, lpData As Any, lpcbData As Long) As Long
  
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
             ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long


' API functions to get temporary filename and path
Public Declare Function GetTempFilename Lib "kernel32" Alias _
            "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString _
            As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Public Declare Function GetTempPath Lib "kernel32" Alias _
"GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer _
As String) As Long


' API function for copying a file
Private Declare Function CopyFile Lib "kernel32" Alias _
            "CopyFileA" (ByVal lpExistingFileName As String, ByVal _
            lpNewFileName As String, ByVal bFailIfExists As Long) As Long

' API function to set the window handle of a control
Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

' API function to set locale of the current thread (works on NT-based OSes only)
Public Declare Function SetThreadLocale Lib "kernel32" (ByVal Locale As clLanguages) As Long

' API function to get locale of the current thread
Public Declare Function GetThreadLocale Lib "kernel32" () As Integer

' API function to retrieve Windows version
Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'**************************************************************
'   DESCRIPTION
'       Opens the HTML help file specified in the pszFile parameter
'       (Note: naming standards follow API conventions because this
'        is really an wrapper to an API function and NOT a user function)
'
'   PARAMETERS
'       hwndCaller - window handle of the calling window
'       pszFile - filename of the help file to load
'       uCommand - always set to HH_DISPLAY_TOPIC. If we get a decent
'                  help file in the future, this may change
'
'   RETURNS - the window handle of the help window. 0 if the call
'             was not succesful
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
' API function declarations for Html Help
Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
    (ByVal hwndCaller As Long, ByVal pszFile As String, _
     ByVal uCommand As Long, ByVal dwData As Long) As Long
        
     
'**************************************************************
'   DESCRIPTION
'       Calls the CoCreateGUID API function to get a unique GUID
'
'   PARAMETERS
'       none
'
'   RETURNS - a unique GUID string, blank string if an error occured
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Function GetGUID() As String

    Const FUNCTION_NAME = "GetGUID()"
    
    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    ' declare our GUID variable as the GUID user-defined data type
    Dim udtGUID As GUID
    
    ' call the API function to retrieve our GUID
    lRetVal = CoCreateGuid(udtGUID)

    If lRetVal = 0 Then
    
        ' if API function call was succesful, return the unique GUID
        GetGUID = _
        "{" & _
        String(8 - Len(Hex$(udtGUID.lData1)), "0") & Hex$(udtGUID.lData1) & "-" & _
        String(4 - Len(Hex$(udtGUID.lData2)), "0") & Hex$(udtGUID.lData2) & "-" & _
        String(4 - Len(Hex$(udtGUID.lData3)), "0") & Hex$(udtGUID.lData3) & "-" & _
        IIf((udtGUID.lData4(0) < &H10), "0", "") & Hex$(udtGUID.lData4(0)) & _
        IIf((udtGUID.lData4(1) < &H10), "0", "") & Hex$(udtGUID.lData4(1)) & _
        IIf((udtGUID.lData4(2) < &H10), "0", "") & Hex$(udtGUID.lData4(2)) & _
        IIf((udtGUID.lData4(3) < &H10), "0", "") & Hex$(udtGUID.lData4(3)) & _
        IIf((udtGUID.lData4(4) < &H10), "0", "") & Hex$(udtGUID.lData4(4)) & _
        IIf((udtGUID.lData4(5) < &H10), "0", "") & Hex$(udtGUID.lData4(5)) & _
        IIf((udtGUID.lData4(6) < &H10), "0", "") & Hex$(udtGUID.lData4(6)) & _
        IIf((udtGUID.lData4(7) < &H10), "0", "") & Hex$(udtGUID.lData4(7)) & _
        "}"
    Else
        ' otherwise, return an empty string
        GetGUID = ""
    End If

    Exit Function

Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function


'**************************************************************
'   DESCRIPTION
'       Deletes a subkey from the registry
'
'   PARAMETERS
'       pvlRootKey - see valid values in Enum regRootKey. Corresponds
'                    to root keys in the Windows registry
'       pvsRegKeyPath - the path under the root key to check
'       pvsRegSubKey - the name of subkey to delete
'
'   RETURNS - True if the key exists, False if otherwise
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Sub regDelete_Sub_Key(ByVal pvlRootKey As regRootKey, _
                      ByVal pvsRegKeyPath As String, _
                      ByVal pvsRegSubKey As String)

    Const FUNCTION_NAME = "regDelete_Sub_Key()"
    
    Dim sErrorDesc As String
    Dim lKeyHandle As Long
    
    On Error GoTo Error_Handler
    
    If regDoes_Key_Exist(pvlRootKey, pvsRegKeyPath) Then
  
        ' Get the key handle
        lRetVal = RegOpenKey(pvlRootKey, pvsRegKeyPath, lKeyHandle)
      
        ' Delete the sub key.  If it does not exist, then ignore it.
        lRetVal = RegDeleteValue(lKeyHandle, pvsRegSubKey)
  
        ' close the handle in the registry.
        lRetVal = RegCloseKey(lKeyHandle)
    End If

    Exit Sub

Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub

'**************************************************************
'   DESCRIPTION
'       Checks if a registry key exists
'
'   PARAMETERS
'       pvlRootKey - see valid values in Enum regRootKey. Corresponds
'                    to root keys in the Windows registry
'       pvsRegKeyPath - the path under the root key to check
'
'   RETURNS - True if the key exists, False if otherwise
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Function regDoes_Key_Exist(ByVal pvlRootKey As regRootKey, _
                                  ByVal pvsRegKeyPath As String) As Boolean

    Const FUNCTION_NAME = "regDoes_Key_Exist()"
    
    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    ' Define and initialize a handle to our registry key
    Dim lKeyHandle As Long
    lKeyHandle = 0
    
    ' call the RegOpenKey API function to open the key
    lRetVal = RegOpenKey(pvlRootKey, pvsRegKeyPath, lKeyHandle)

    ' If we got a key handle, it means the key exists
    If lKeyHandle = 0 Then
        regDoes_Key_Exist = False
    Else
        regDoes_Key_Exist = True
    End If

    ' Close the handle to the key
    lRetVal = RegCloseKey(lKeyHandle)
    
    Exit Function

Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc
    
End Function


'**************************************************************
'   DESCRIPTION
'       Checks if a registry key exists
'
'   PARAMETERS
'       pvlRootKey - see valid values in Enum regRootKey. Corresponds
'                    to root keys in the Windows registry
'       pvsRegKeyPath - the path under the root key to check
'       pvsRegSubKey - the name of the value key to check
'
'   RETURNS - True if the key exists, False if otherwise
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Function regQuery_A_Key(ByVal pvlRootKey As regRootKey, _
                               ByVal pvsRegKeyPath As String, _
                               ByVal pvsRegSubKey As String) As Variant
                               
    Const FUNCTION_NAME = "regQuery_A_Key()"
    
    Dim sErrorDesc As String
    On Error GoTo Error_Handler
                               
    ' Define and initialize our variables
    Dim lKeyHandle As Long
    Dim lBufferSize  As Long
    Dim nPosition As Integer
    Dim lDataType As Long
    Dim lBuffer As Long
    Dim sBuffer As String
  
    lKeyHandle = 0
    lBufferSize = 0
    lKeyHandle = 0
    
    ' call the RegOpenKey API function to open the key
    lRetVal = RegOpenKey(pvlRootKey, pvsRegKeyPath, lKeyHandle)
                               
    ' Exit the function if subkey does not exist
    If lKeyHandle = 0 Then
        regQuery_A_Key = ""
        lRetVal = RegCloseKey(lKeyHandle)
        Exit Function
    End If
                                          
    ' query the value of the key. retrieve key data type
    ' and size of buffer needed to hold the value
    lRetVal = RegQueryValueEx(lKeyHandle, pvsRegSubKey, 0&, _
                              lDataType, ByVal 0&, lBufferSize)
    
    ' if it was not able to query the value, it means the value key does not
    ' exist so return an empty string.
    If lRetVal <> 0 Then
        regQuery_A_Key = ""
        lRetVal = RegCloseKey(lKeyHandle)
        Exit Function
    End If

    Select Case lDataType
        ' String data
        Case REG_SZ:
        
            ' fill our buffer with spaces
            sBuffer = Space(lBufferSize)
            lRetVal = RegQueryValueEx(lKeyHandle, pvsRegSubKey, 0&, 0&, ByVal sBuffer, lBufferSize)
              
            If lRetVal <> ERROR_SUCCESS Then
                  regQuery_A_Key = ""
            Else
                  ' if we found it, get the the string data
                  nPosition = InStr(1, sBuffer, Chr(0))
                  If nPosition > 0 Then
                      ' if we found the null character, return everything before the NULL
                      regQuery_A_Key = VBA.Left(sBuffer, nPosition - 1)
                  Else
                      ' did not find one.  return the entire string
                      regQuery_A_Key = sBuffer
                  End If
            End If
            
        ' Numeric data (Integer)
         Case REG_DWORD:
            ' fetch the value into our buffer. the last parameter specifies
            ' a 4 byte work (a long integer)
            lRetVal = RegQueryValueEx(lKeyHandle, pvsRegSubKey, 0&, lDataType, _
                                      lBuffer, 4&)
              
            ' If we are not succesful, return blank
            If lRetVal <> ERROR_SUCCESS Then
                regQuery_A_Key = -1
            Else
                ' return the value from the buffer
                regQuery_A_Key = lBuffer
            End If
         
         ' if we don't know what type
         Case Else:
              regQuery_A_Key = ""
  End Select
  
  ' close the handle
  lRetVal = RegCloseKey(lKeyHandle)

    Exit Function

Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function


'**************************************************************
'   DESCRIPTION
'       Creates a registry key
'
'   PARAMETERS
'       pvlRootKey - see valid values in Enum regRootKey. Corresponds
'                    to root keys in the Windows registry
'       pvsRegKeyPath - the path under the root key to check
'
'   RETURNS - none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Sub regCreate_A_Key(ByVal pvlRootKey As regRootKey, ByVal pvsRegKeyPath As String)

    Const FUNCTION_NAME = "regCreate_A_Key()"
    
    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    Dim lKeyHandle As Long
  
    ' create the key (if it already exists, this API call will be ignored
    lRetVal = RegCreateKey(pvlRootKey, pvsRegKeyPath, lKeyHandle)

    lRetVal = RegCloseKey(lKeyHandle)

    Exit Sub

Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'       Create or update a registry key value
'
'   PARAMETERS
'       pvlRootKey - see valid values in Enum regRootKey. Corresponds
'                    to root keys in the Windows registry
'       pvsRegKeyPath - the path under the root key to check
'       pvsRegSubKey - the name of the SubKey to create
'       pvvRegData - the value of the regkey
'
'   RETURNS - True if the key exists, False if otherwise
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Sub regCreate_Key_Value(ByVal pvlRootKey As regRootKey, _
                        ByVal pvsRegKeyPath As String, _
                        ByVal pvsRegSubKey As String, _
                        PvvRegData As Variant)

    Const FUNCTION_NAME = "regCreate_Key_Value()"
    
    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    Dim lKeyHandle As Long
    Dim lDataType As Long
    Dim lKeyValue As Long
    Dim sKeyValue As String
  
    ' determine the data type based on the value of the registry data
    ' i.e. if it is a number, then store it as a DWORD, otherwise as string
    If IsNumeric(PvvRegData) Then
        lDataType = REG_DWORD
    Else
        lDataType = REG_SZ
    End If
  
    lRetVal = RegCreateKey(pvlRootKey, pvsRegKeyPath, lKeyHandle)
    
    Select Case lDataType
        Case REG_SZ:
            ' if it is a string data, null-terminate it and store
            sKeyValue = Trim(PvvRegData) & Chr(0)     ' null terminated
            lRetVal = RegSetValueEx(lKeyHandle, pvsRegSubKey, 0&, lDataType, _
                                        ByVal sKeyValue, Len(sKeyValue))
                                     
        Case REG_DWORD:
            ' if it is numeric, convert it to a long and store
            ' last parameter in call to RegSetValueEx specifies that we use 4 bytes to represent a WORD
            lKeyValue = CLng(PvvRegData)
            lRetVal = RegSetValueEx(lKeyHandle, pvsRegSubKey, 0&, lDataType, _
                                        lKeyValue, 4&)
    End Select
  
  lRetVal = RegCloseKey(lKeyHandle)

    Exit Sub

Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Sub


'**************************************************************
'   DESCRIPTION
'       Deletes a registry key
'
'   PARAMETERS
'       pvlRootKey - see valid values in Enum regRootKey. Corresponds
'                    to root keys in the Windows registry
'       pvsRegKeyPath - the path under the root key to check
'       pvsRegKeyName - the name of the key to delete
'
'   RETURNS - True if the key was deleted, False if otherwise
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Function regDelete_A_Key(ByVal pvlRootKey As regRootKey, _
                         ByVal pvsRegKeyPath As String, _
                         ByVal pvsRegKeyName As String) As Boolean
    
    Const FUNCTION_NAME = "regDelete_A_Key()"
    
    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    Dim lKeyHandle As Long
  
    regDelete_A_Key = False
  
    ' check if the key exists
    If regDoes_Key_Exist(pvlRootKey, pvsRegKeyPath) Then
  
        ' Get the key and delete it
        lRetVal = RegOpenKey(pvlRootKey, pvsRegKeyPath, lKeyHandle)
      
        lRetVal = RegDeleteKey(lKeyHandle, pvsRegKeyName)
      
        ' If the value returned is zero then the delete was a success
        If lRetVal = 0 Then regDelete_A_Key = True
      
        ' close the handle
        lRetVal = RegCloseKey(lKeyHandle)
    End If

    Exit Function

Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function

'**************************************************************
'   DESCRIPTION
'       creates a unique temporary file in the default temp path and returns
'       the filename to the calling function
'
'   PARAMETERS
'       pvPrefix - file prefix (optional)
'
'   RETURNS - True if the key exists, False if otherwise
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/18/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Function GetTemporaryFilename(Optional pvPrefix As String = "") As String
   
    Const FUNCTION_NAME = "GetTemporaryFilename()"
    
    Dim sErrorDesc As String
    On Error GoTo Error_Handler
   
    Dim sTempPath As String * 255
    Dim sTempFilename As String * 255
    
    lRetVal = GetTempPath(254, sTempPath)
    lRetVal = GetTempFilename(sTempPath & "\", pvPrefix, 0, sTempFilename)
    
    GetTemporaryFilename = Trim(sTempFilename)

    Exit Function

Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function

'**************************************************************
'   DESCRIPTION
'       Copies a file to the specified destination. Overwrites file
'       if it already exists
'
'   PARAMETERS
'       pvsSource - source file to copy
'       pvsTarget - target filename
'
'   RETURNS - True if the key exists, False if otherwise
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/18/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Function CopyAFile(pvsSource As String, pvsTarget As String)

    Const FUNCTION_NAME = "CheckRegistry()"
    
    Dim sErrorDesc As String
    On Error GoTo Error_Handler

    CopyAFile = False

    ' call the copy file API function (do not fail if the file exists
    ' because we know it is a tempfile)
    lRetVal = CopyFile(Trim(pvsSource), Trim(pvsTarget), False)

    If lRetVal Then
        CopyAFile = True
    End If

    Exit Function

Error_Handler:
    ' untrapped error, call the critical error handler function
    sErrorDesc = Err.Number & " - " & Err.Description
    psAbort MODULE_NAME, FUNCTION_NAME, sErrorDesc

End Function



'**************************************************************
'   DESCRIPTION
'       Returns the windows
'
'   PARAMETERS
'       none
'
'   RETURNS - "9x" if the program is running on a 9x-based OS (95, 98, ME)
'             "NT" if the program is running on a NT-based OS (NT, 2K, XP)
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  2/11/2004       Jett Gamboa       Initial Revision
'**************************************************************
Public Function GetWinPlatform() As String

    Dim OSV As OSVERSIONINFO
     
    OSV.OSVSize = Len(OSV)
    
    ' Call the GetVersionEx API function
    If GetVersionEx(OSV) = 1 Then
     
        ' Determine the platform ID
        Select Case OSV.PlatformID
        
            Case VER_PLATFORM_WIN32_NT:
                GetWinPlatform = "NT"
                
            Case VER_PLATFORM_WIN32_WINDOWS:
                GetWinPlatform = "9x"
                
        End Select
    Else
        ' there was an error retrieving the platform information so
        ' we just return the safest type - Win9x
    
        GetWinPlatform = "9x"
     
    End If

End Function



