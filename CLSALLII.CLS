VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAllIITable"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Class Name       : clsAllIITable
'Description      : Class to contain all of the Indirect Incentive codes for
'                   the division.  This collection Class is used to build the
'                   Indirect Incentives Collection Class by Cost Center when
'                   it is required.
'Properties       : AddIIItem, Count, GetIIItem
'Called From      : DBCommon.LoadAllBatchIncnCols, Startup.Declarations
'Dependencies     : clsIIItem
'==============================================================================

' The Indirect Incentive Items Collection for all Cost Centers
Private colTable As New Collection

Public Sub AddIIItem(strCCID As String, iType As Integer, _
     iNumb As Integer, cWrkStd As Currency, strDesc As String)
    'Add a Indirect Incentive Code to the All Indirect Incentive Code Collection.
    'Method of clsAllIITable object to create a new clsIIItem object with the
    '   properties passed in, and put it into the All Indirect Incentive Items
    '   Collection.
    Dim newEntryItem As New clsIIItem
    newEntryItem.Code = Left(strDesc & String(30, " "), 30)
    newEntryItem.Desc = Format(cWrkStd, "00.0")
    newEntryItem.CodeDesc = strDesc & " " & newEntryItem.Code
    newEntryItem.CCID = UCase(Mid(strCCID, 1, 4))
    newEntryItem.IncenType = iType
    newEntryItem.MachNumb = iNumb
    newEntryItem.WrkStd = cWrkStd
    colTable.Add Item:=newEntryItem, _
        Key:=UCase(strCCID) & newEntryItem.Code
End Sub

Property Get Count() As Long
    'Get a count of items in the All Indirect Incentive Code Collection.
    'Property of claAllBITable object expose the Count property of the
    '   All Indirect Incentive Items Collection (colTable).
    Count = colTable.Count
End Property

Public Sub getIIItem(anIndex As Variant, ByRef theItem As clsIIItem)
    'Lookup a Indirect Incentive Code from the All Indirect Incentive Code
    '   Collection.
    'Method of clsAllIITable object. Populates the properties of the passed
    '   clsIIItem item with the properties of the Item (in the All Indirect
    '   Incentive Items Collection) that matches the anIndex index passed in.
    
    On Error GoTo getIIItemError
    theItem.Code = colTable.Item(anIndex).Code
    theItem.Desc = colTable.Item(anIndex).Desc
    theItem.CodeDesc = colTable.Item(anIndex).CodeDesc
    theItem.CCID = colTable.Item(anIndex).CCID
    theItem.MachNumb = colTable.Item(anIndex).MachNumb
    theItem.IncenType = colTable.Item(anIndex).IncenType
    theItem.WrkStd = colTable.Item(anIndex).WrkStd
    Exit Sub
    
getIIItemError:
    theItem.Code = ""
    theItem.Desc = LoadResString(gLanguageOffset + 4)
    theItem.CodeDesc = LoadResString(gLanguageOffset + 4)
    theItem.CCID = ""
    theItem.IncenType = 0
    theItem.MachNumb = 0
    theItem.WrkStd = 0
End Sub

