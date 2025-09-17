VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGridWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------
'*
'* Filename: clsGridWrapper.cls
'*
'* DESCRIPTION
'* -----------
'* This is a class designed to wrap around the VSFlexGrid
'* control. It's primary goal is to provide an easy way
'* to call initialize the grid controls as well as to
'* facilitate saving grid column widths.
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


Dim NewColumn As ColumnInfo

'JNM 04/09/03 - Modified maximum number of columns
Const MAX_GRID_COLUMNS = 12

' private variable to store pointer to the "real" grid control
Private ctlUserGrid As VSFlexGrid

' user defined types for the grid columns
Public Enum gridColumnType
    colTypeText = 1
    colTypeDate = 2
    colTypeTime = 3
    colTypeCombo = 4
    colTypeCheckbox = 5
End Enum

' user defined type for column information
Private Type ColumnInfo
    nColumnNumber As Integer
    sColumnName As String
    nColumnWidth As Integer
End Type

Private gsGridName As String
Private nNumberOfColumns As Integer
Dim arrGridColumns(MAX_GRID_COLUMNS) As ColumnInfo

'**************************************************************
'   DESCRIPTION
'       Stores a reference to the grid control.
'
'   PARAMETERS
'       pvsTheGridName - name used in registry when saving grid settings
'       prctlMyGrid - the control on the form that this instance will
'            be associated to
'
'   RETURNS -
'       none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Sub Initialize(ByRef pvsTheGridName As String, prctlMyGrid As VSFlexGrid)

    Set ctlUserGrid = prctlMyGrid
    gsGridName = pvsTheGridName

    nNumberOfColumns = 0

    ' set the defaults for the grid
    With ctlUserGrid
        .Rows = 1
        .Cols = 0
        .Editable = flexEDKbdMouse
        .AutoResize = False
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
    End With
    
End Sub


'**************************************************************
'   DESCRIPTION
'       Adds a new column to the grid
'
'   PARAMETERS
'       pvsColumnName - name used in registry when saving grid settings
'       pvsColumnHeading - the heading for this column. Is displayed
'           in row 0 of the grid control
'       pvColumnType - type of data this column will hold
'       pvColAlign - alignment of text in this column
'
'   RETURNS -
'       none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
' 10/13/2002       Jett Gamboa       set color for column headings
'**************************************************************
Public Sub AddColumn(pvsColumnName As String, pvsColumnHeading As String, pvColumnType As gridColumnType, pvColAlign As AlignmentSettings)

    If gsGridName = "" Then
        MsgBox "Grid has not been initialized", vbCritical
        End
    End If

    nNumberOfColumns = nNumberOfColumns + 1
     
    ctlUserGrid.Cols = nNumberOfColumns
    ctlUserGrid.ColAlignment(nNumberOfColumns - 1) = pvColAlign
    
    With NewColumn
        .nColumnNumber = nNumberOfColumns
        .sColumnName = pvsColumnName
        .nColumnWidth = ctlUserGrid.ColWidth(nNumberOfColumns - 1)
        
        ' set the column heading to the specified column name
        ctlUserGrid.Cell(flexcpText, 0, nNumberOfColumns - 1) = pvsColumnHeading
        
        ' Set the column heading color
        'TSG 11/26 - removed this
        'ctlUserGrid.Cell(flexcpBackColor, 0, nNumberOfColumns - 1, 0, nNumberOfColumns - 1) = COLUMN_HEADING_COLOR
    End With
    
    'collGridColumns.Add NewColumn, NewColumn.sColumnName
    arrGridColumns(nNumberOfColumns - 1) = NewColumn

    Select Case pvColumnType
        Case colTypeText
            ctlUserGrid.ColDataType(nNumberOfColumns - 1) = flexDTString
        Case colTypeDate
            ctlUserGrid.ColDataType(nNumberOfColumns - 1) = flexDTString
            ctlUserGrid.ColFormat(nNumberOfColumns - 1) = "dd-mmm-yy"
        Case colTypeTime
            ctlUserGrid.ColDataType(nNumberOfColumns - 1) = flexDTString
            ' JETT 3/5/04 LOCALIZATION - used global time format variable
            ctlUserGrid.ColFormat(nNumberOfColumns - 1) = gTimeFormat '"H:nn" '"Medium Time" '"h:Nn AMPM"
        Case colTypeCombo
            ctlUserGrid.ColDataType(nNumberOfColumns - 1) = flexDTString
        Case colTypeCheckbox
            ctlUserGrid.ColDataType(nNumberOfColumns - 1) = flexDTBoolean
    End Select

End Sub


'**************************************************************
'   DESCRIPTION
'       Saves the settings for all the columns in the associated grid
'
'   PARAMETERS
'       pvsGridSettingsRegPath - registry path to save settings to
'
'   RETURNS -
'       none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Sub SaveGridSettings(pvsGridSettingsRegPath As String)

    Dim GridColumn As ColumnInfo
    Dim sGridSettingsKey As String
    Dim sGridColKey As String
    
    Dim x As Integer

    ' If the Initialize method was not called, warn the user and terminte the application
    If gsGridName = "" Then
        MsgBox "Grid has not been initialized", vbCritical
        End
    End If
    
    sGridSettingsKey = pvsGridSettingsRegPath & "\" & gsGridName
    
    ' check if the key already exists. We always check the HKCU node
    ' because settings are saved for each user. If it does not exist
    ' create the key
    If Not regDoes_Key_Exist(HKEY_CURRENT_USER, sGridSettingsKey) Then
        
        regCreate_A_Key HKEY_CURRENT_USER, sGridSettingsKey
    
    End If
    
    'For Each GridColumn In collGridColumns
    For x = 0 To nNumberOfColumns - 1
        sGridColKey = sGridSettingsKey & "\" & arrGridColumns(x).sColumnName
               
        ' save the new width
        arrGridColumns(x).nColumnWidth = ctlUserGrid.ColWidth(x)
               
        ' create the key for his column, if it does not exist it will be created
        regCreate_Key_Value HKEY_CURRENT_USER, sGridSettingsKey, arrGridColumns(x).sColumnName, arrGridColumns(x).nColumnWidth
    Next

End Sub


'**************************************************************
'   DESCRIPTION
'       Saves the settings for all the columns in the associated grid
'
'   PARAMETERS
'       pvsGridSettingsRegPath - registry path to load settings from
'
'   RETURNS -
'       none
'
'   REVISION LOG
'       DATE       NAME              CHANGE
'  9/11/2002       Jett Gamboa       Initial Revision
'**************************************************************
Public Sub LoadGridSettings(pvsGridSettingsRegPath As String)

    Dim x As Integer

    Dim GridColumn As ColumnInfo
    Dim sGridSettingsKey As String
    Dim sGridColKey As String
    Dim vKeyValue As Variant
    Dim nWidth As Integer
    
    ' If the Initialize method was not called, warn the user and terminte the application
    If gsGridName = "" Then
        MsgBox "Grid has not been initialized", vbCritical
        End
    End If

    sGridSettingsKey = pvsGridSettingsRegPath & "\" & gsGridName

    ' if the key for this grid does not exist, create it
    If Not regDoes_Key_Exist(HKEY_CURRENT_USER, sGridSettingsKey) Then
        ' SaveGridSettings pvsGridSettingsRegPath
        regCreate_A_Key HKEY_CURRENT_USER, sGridSettingsKey
    End If
    
    For x = 0 To nNumberOfColumns - 1

        sGridColKey = sGridSettingsKey & "\" & arrGridColumns(x).sColumnName
                      
        vKeyValue = regQuery_A_Key(HKEY_CURRENT_USER, sGridSettingsKey, arrGridColumns(x).sColumnName)
                      
        If vKeyValue <> "" Then
            nWidth = CInt(vKeyValue)
        
            arrGridColumns(x).nColumnWidth = nWidth
            ctlUserGrid.ColWidth(x) = nWidth
            
            ' refresh the grid
            ctlUserGrid.Refresh
        Else
            ' if it does not exist, use whatever is on the grid
            arrGridColumns(x).nColumnWidth = ctlUserGrid.ColWidth(x)
            
            ' save it to the registry
            regCreate_Key_Value HKEY_CURRENT_USER, sGridSettingsKey, arrGridColumns(x).sColumnName, arrGridColumns(x).nColumnWidth
        End If
    Next

End Sub


'**************************************************************
'   DESCRIPTION
'       common routine that can be called which copies contents
'       of selected cell(s) to the Windows clipboard
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
Public Sub GridSelectionCopy()
    Clipboard.Clear
    Clipboard.SetText ctlUserGrid.Clip
End Sub


'**************************************************************
'   DESCRIPTION
'       removes all of the rows in the grid
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
Public Sub RemoveAllRows()

    ctlUserGrid.Rows = 1

End Sub

