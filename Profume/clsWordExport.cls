VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWordExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
'**************************************************************
'* Filename: clsWordExport.cls
'*
'* DESCRIPTION
'* -----------
'* This code will contain the functions and subroutines
'* to be referenced by the modGrapgProperties and
'* components of the Fumiguide (ProFume) Application in
'* Exporting the graph to WORD.
'*
'* Copyright (c) 2002 Accenture.  All rights reserved.
'-----------------------------------------------------
'*
'* REVISION HISTORY
'* Revision:    0
'* Author:      Gamboa, Maria Cecilia <NAI1686>
'* Date:        March 25, 2003
'**************************************************************

Option Explicit

Public Enum PageSetups
   Landscape = wdOrientLandscape
   Portrait = wdOrientPortrait
End Enum

Public Enum Alignment
    Center = WdParagraphAlignment.wdAlignParagraphCenter
    Left = WdParagraphAlignment.wdAlignParagraphLeft
    Right = WdParagraphAlignment.wdAlignParagraphRight
    Justify = WdParagraphAlignment.wdAlignParagraphJustify
End Enum

Private Const CLSNAME = "clsWordExport.cls"
Private sProcedureName As String    'Name of current procedure: for error handling
Private nNumOfLines As Integer      'Number of blank lines to insert
Private sStrToAdd As String         'Text to add to Word doc
Private nFontSize As Integer        'Font size of StrToAdd
Private bFontBold As Boolean        'Is strToAdd bold or Not
Private nParaAlign As Integer       'StrToAdd alignment SEE ENUM ALIGNMENT
Private nPageSetup As Integer       'Page setup of Word Doc SEE ENUM PAGESETUPS
Private I As Integer                'Used in for next loop
Private wrdApp As Word.Application  'MS Word object
Private wrdDoc As Word.Document     'MS Word Document
Private wrdSelection As Word.Selection 'MS Word Selection
Private sDocName As String          'MS Word document name
Private sModName As String
Private sErrorDesc As String

'JNM 05/05/03 - PT fix - added this boolean in order to flag the app if the machine
'   has MSWord or not
Private bHasMSWord As Boolean

'**************************************************************
'   DESCRIPTION
'   Initializes the class.
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************

Private Sub Class_Initialize()

On Error GoTo Error_Handler

  sModName = "Class_Initialize"
  Set wrdApp = New Word.Application
  
  'JNM 05/05/03 - Tell the app that MSWord exists
  bHasMSWord = True
    
  'Set to false if you don't want to see the word doc
  wrdApp.Visible = True

Exit Sub
Error_Handler:
'JNM 05/05/03 - PT fix - Transferred Error handler here.
  If Err.Number <> 0 Then
     'this means the computer being used does not have WORD installed in it
     MsgBox LoadResString(IDS_GRAPH_NO_WORD), vbCritical
     bHasMSWord = False
     Exit Sub
  End If

End Sub

'**************************************************************
'   DESCRIPTION
'   Terminates the class.
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************

Private Sub Class_Terminate()

On Error GoTo Error_Handler

  sModName = "Class_Terminate"

  'wrdApp.Quit
  Set wrdSelection = Nothing
  Set wrdDoc = Nothing
  Set wrdApp = Nothing

Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(CLSNAME, sModName, sErrorDesc)

End Sub

'**************************************************************
'   DESCRIPTION
'   Opens a new instance of Word.
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************

Public Sub OpenNewDoc()

On Error GoTo Error_Handler

    sModName = "OpenNewDoc"
    Set wrdDoc = wrdApp.Documents.Add
    wrdDoc.Select
    
    Set wrdSelection = wrdApp.Selection
    
Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(CLSNAME, sModName, sErrorDesc)
End Sub

Public Property Let PageSetupDocument(IntPageSetup As Integer)

    nPageSetup = IntPageSetup
    wrdDoc.PageSetup.Orientation = nPageSetup
    
End Property

Public Property Let InsertLinesInDoc(numOfLines As Integer)

    nNumOfLines = numOfLines
    InsertLines nNumOfLines
    
End Property

'**************************************************************
'   DESCRIPTION
'   Inserts a text/string in the document
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************

Public Sub InsertText(strToAdd As String, IntFontSize As Integer, _
    blBold As Boolean, intParagraphAlign As Integer)

On Error GoTo Error_Handler

    sModName = "InsertText"
    sStrToAdd = strToAdd
    bFontBold = blBold
    nFontSize = IntFontSize
    nParaAlign = intParagraphAlign
    
    InsertTextIntoDoc
    
Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(CLSNAME, sModName, sErrorDesc)
    
End Sub

'**************************************************************
'   DESCRIPTION
'   Handles inserting the graph to the word document
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************

Public Sub InsertGraph()

On Error GoTo Error_Handler

    sModName = "InsertGraph"
    InsertGraphIntoDoc

Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(CLSNAME, sModName, sErrorDesc)
End Sub

'**************************************************************
'   DESCRIPTION
'   Inserts the graph into the word document from the clipboard
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************

Private Sub InsertGraphIntoDoc()

On Error GoTo Error_Handler

    sModName = "InsertGraphIntoDoc"
    wrdSelection.Paste
    
Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(CLSNAME, sModName, sErrorDesc)
    
End Sub

'**************************************************************
'   DESCRIPTION
'   This routine inserts text into the word document and sets
'   the font and alignment
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************
Private Sub InsertTextIntoDoc()


On Error GoTo Error_Handler

    sModName = "InsertTextIntoDoc"
    wrdSelection.ParagraphFormat.Alignment = nParaAlign
    wrdSelection.Font.Size = nFontSize
    wrdSelection.Font.Bold = bFontBold
    wrdSelection.TypeText sStrToAdd

Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(CLSNAME, sModName, sErrorDesc)
    
End Sub

'**************************************************************
'   DESCRIPTION
'   Inserts blank lines in Word document
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************

Private Sub InsertLines(LineNum As Integer)

Dim nCount As Integer
On Error GoTo Error_Handler

    sModName = "InsertLines"
    'Insert blank lines in Word document
    For nCount = 1 To LineNum
        wrdApp.Selection.TypeParagraph
    Next nCount
    
Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(CLSNAME, sModName, sErrorDesc)
    
End Sub
   
'**************************************************************
'   DESCRIPTION
'   Inserts the current date with the default font
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   March 25, 2003       Cel Gamboa       Initial Revision
'**************************************************************

Public Sub InsertCurrentDate()
On Error GoTo Error_Handler
    
    sModName = "InsertCurrentDate"

    wrdSelection.InsertDateTime _
    DateTimeFormat:="dddd, MMMM dd, yyyy hh:mm AM/PM", InsertAsField:=False
    
Exit Sub
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(CLSNAME, sModName, sErrorDesc)
    
End Sub

'**************************************************************
'   DESCRIPTION
'   Returns the value of the private variable bHasMSWord
'
'   PARAMETERS
'       none
'
'   RETURNS - none
'
'   REVISION LOG
'   DATE                 NAME             CHANGE
'   May 5, 2003       Joseph Manyaga      Initial Revision
'**************************************************************

Public Function HasMSWord()
On Error GoTo Error_Handler
    
    sModName = "HasMSWord"

    HasMSWord = bHasMSWord
    
Exit Function
Error_Handler:
    sErrorDesc = Err.Number & " - " & Err.Description
    Call psAbort(CLSNAME, sModName, sErrorDesc)
    
End Function

