VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Class Name       : clsEntry
'Description      : The class that is an object to represent a Entry item
'                   Types of Entry items are Menu input, Text input or Combo
'                   lookup. Used in conjunction with the frmDCEntry form
'                   (If VB5 - then this would have been an Activex Document).
'Properties       : Add, AddScanItem, Count, CountLookupItems, CountScanItems,
'                   EntryType, Escaped, getCode, getCodeDesc, getDesc,
'                   getLookupItem, getScanItem, IsLookup, LuTable, MainEsc,
'                   Response, ResponseValue, Title
'Called From      : frmLabor.BasEntry, frmInpEmplNo.BasEntry,
'                   frmDirLabor.InpLabJobSeq, frmDirLabor.DirLaborReEnter,
'                   frmDCMain.BasEntry, frmNotepad.NotepadProcess,
'                   frmLaborChk.LaborChkProcess, frmJobPrint.inpJobNum,
'                   Inputs.inpEndOfShift, Inputs.AnotherRec, Inputs.inpJobAndSeq,
'                   Inputs.inpWcAns, Inputs.inpWCID, Inputs.inpAddonHrs,
'                   Inputs.inpIndirHrs, Inputs.inpProdQty, Inputs.callCompl,
'                   Inputs.ReEnter, Inputs.inpNumbMachs, Inputs.inpMchHours,
'                   Inputs.inpSetHours, Inputs.inpOperNumb
'Dependencies     : clsCodeTable, clsCodeItem
'==============================================================================

' The Private Properties of an Entry Item
Private strTitle            As String           'Title of The menu Option / User Prompt
Private colEntryItems       As New Collection   'Menu Entries as a collection, each menu item is a collection member
Private colScanItems        As New Collection   'Keeps track of Scan Itemes, This prevents scans as being acted upon as a menu entry
Private colLookupItems      As clsCodeTable     'Collection of Items that hold the Combo box Items.
Private strResponse         As Variant          'Contains the Memnu / Scan / Combo Box Selection was.
Private iEntryType          As Integer          '0 - Any Char (Default), 1 Numeric Only
Private bMainEsc            As Boolean          'Flag that indicated the uses presses CTRL ALT X, then code would need to call shutdown
Private bEscaped            As Boolean          'If the User presses the ESC Key this flag is set.
Private bIsLookup           As Boolean          'Indicates that we are using a combo box
Private bESCDisabled        As Boolean
Public Sub Add(strCode As String, strDesc As String, _
        Optional vValueIn As Variant)
    'Add a Menu Item to the Menu Collection.  The CodeDesc field used in the
    '   clsCodeItem class is created by combining the code and desc.  The
    '   Value field is an optional argument, that is used when a value
    '   different from the code should be returned from the menu selection
    '   process.
    Dim newEntryItem As New clsCodeItem
    newEntryItem.Code = UCase(strCode)
    newEntryItem.Desc = strDesc
    newEntryItem.CodeDesc = UCase(strCode) & " " & strDesc
    newEntryItem.Value = vValueIn
    colEntryItems.Add Item:=newEntryItem, Key:=UCase(strCode)
End Sub

Public Sub AddScanItem(strCode As String)
    'Add a Scan Item to the Scan Item Collection.  This is used when the
    '   scanner is allowed to pass in a string while the menu is shown.
    'The Code field is the only field used in the clsCodeItem class.  This
    '   is used for specifying the first character of multiple character
    '   scannings in a menu selection.
    Dim newEntryItem As New clsCodeItem
    newEntryItem.Code = UCase(strCode)
    newEntryItem.Desc = ""
    colScanItems.Add Item:=newEntryItem, Key:=UCase(strCode)
End Sub


Property Get Count() As Long
    'Get a count of items in the Menu Collection.
    Count = colEntryItems.Count
End Property

Property Get CountScanItems() As Long
    'Get a count of items in the Scan Items Collection.
    CountScanItems = colScanItems.Count
End Property
Property Let ESCDisabled(ByVal blEDisabled As Boolean)
    bESCDisabled = blEDisabled
End Property
Property Get ESCDisabled() As Boolean
    'What is the title or prompt for the entry.
    ESCDisabled = bESCDisabled
End Property

Property Let Title(ByVal strNewTitle As String)
    'Set the Title/Prompt for the entry
    strTitle = strNewTitle
End Property
Property Get Title() As String
    'What is the title or prompt for the entry.
    Title = strTitle
End Property



Public Function getDesc(anItem As Variant) As String
    'Lookup a Menu Desc from the Menu Item Collection.
    On Error GoTo getDescError
    
    getDesc = colEntryItems.Item(anItem).Desc
    Exit Function
    
getDescError:
    getDesc = LoadResString(gLanguageOffset + 4)
End Function

Public Function getCodeDesc(anItem As Variant) As String
    'Lookup a Menu Code & Desc from the Menu Item Collection.
    
    On Error GoTo getCodeDescError
    getCodeDesc = colEntryItems.Item(anItem).CodeDesc
    Exit Function
    
getCodeDescError:
    getCodeDesc = LoadResString(gLanguageOffset + 4)
End Function


Public Function getCode(anItem As Variant) As String
    'Lookup a Menu Code from the Menu Item Collection.
    
    On Error GoTo getCodeError
    getCode = colEntryItems.Item(anItem).Code
    Exit Function
    
getCodeError:
    getCode = LoadResString(gLanguageOffset + 4)
End Function

Public Function getScanItem(anItem As Variant) As String
    'Lookup a Code from the Scan Items Collection.
    
    On Error GoTo getScanItemError:
    getScanItem = colScanItems.Item(anItem).Code
    Exit Function
    
getScanItemError:
    getScanItem = LoadResString(gLanguageOffset + 4)
End Function


Public Function getLookupItem(anItem As Variant) As String
    'Lookup a Code & Desc from the Code Lookup Collection.
    Dim vResponse As Variant
    
    If bIsLookup Then
        vResponse = colLookupItems.getCodeDesc(anItem)
    Else
        vResponse = ""
    End If
    If vResponse = Null Then
        getLookupItem = ""
    Else
        getLookupItem = vResponse
    End If
End Function


Public Property Get Response() As String
    'Get the response from the entry system.  Has the code of the Menu
    '   selection, the text entered, or the scanned input.
    Response = strResponse
End Property

Public Property Get ResponseValue() As Variant
    'Get the response from the entry system.  This is used only with the
    '   Menu system, and passes back the value corresponding to the Menu
    '   item selected.  This routine takes the response back from the Entry
    '   system and looks up the corresponding value, to pass back.
    Dim vValue As Variant
    On Error GoTo ResponseValueError
    If colEntryItems.Count > 0 Then     'Menu input
        vValue = colEntryItems.Item(strResponse).Value
        ResponseValue = vValue
    Else
        ResponseValue = strResponse
    End If
    Exit Property
ResponseValueError:
    ResponseValue = strResponse
End Property


Public Property Let Response(strNewResponse As String)
    'Used by the frmDCEntry to save/pass the response from the entry system.
    '   Has the code of the Menu selection, the text entered, or the
    '   scanned input.
    strResponse = strNewResponse
End Property

Private Sub Class_Initialize()
    strTitle = ""
    strResponse = ""
    bMainEsc = False
    iEntryType = 0
    bIsLookup = False
    bESCDisabled = False
End Sub


Public Property Get EntryType() As Integer
    'Expose the EntryType property of clsEntry.  This is used to select any
    '   printable characters, or just numbers.
    EntryType = iEntryType
End Property

Public Property Let EntryType(iEntryTypeNew As Integer)
    'Expose the EntryType property of clsEntry.  This is used to select any
    '   printable characters, or just numbers.
    If iEntryTypeNew >= 0 _
            And iEntryTypeNew <= 1 Then
            '0 is All Characters
            '1 is Numeric Only
        iEntryType = iEntryTypeNew
    End If
End Property

Public Property Get MainEsc() As Boolean
    'Property to see if Ctrl-Alt-X is allowed to shut down App
    MainEsc = bMainEsc
End Property

Public Property Let MainEsc(bNewValue As Boolean)
    'Property to allow or not to allow Ctrl-Alt-X to shut down App
    bMainEsc = bNewValue
End Property

Public Property Get Escaped() As Boolean
    'Expose the Escaped property of clsEntry.  This is used to determine if
    '   user just escaped out of the entry system.
    Escaped = bEscaped
End Property

Public Property Let Escaped(bNewValue As Boolean)
    'Expose the Escaped property of clsEntry.  This is used by frmDCEntry
    '   to indicate if the user just escaped out of the entry system.
    bEscaped = bNewValue
End Property


Public Property Get CountLookupItems() As Integer
    'Get a count of items in the Lookup Items Collection.
    If bIsLookup Then
        CountLookupItems = colLookupItems.Count
    Else
        CountLookupItems = 0
    End If
End Property



Public Property Let LuTable(aCol As clsCodeTable)
    'Property to pass in a lookup collection in a clsCodeTable object
    'Sets the property for the type of entry as lookup (combo box)
    Set colLookupItems = aCol
    bIsLookup = True
End Property



Public Property Get IsLookup() As Boolean
    'Property to see if entry type is a lookup combo box
    IsLookup = bIsLookup
End Property


Private Sub Class_Terminate()
  'MsgBox "Entry class being destroyed....", vbInformation
End Sub


