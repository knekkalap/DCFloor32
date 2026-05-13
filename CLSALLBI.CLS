VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAllBITable"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Class Name       : clsAllBITable
'Description      : Class to contain all of the Batch Incentive codes for
'                   the division.  This collection Class is used to build the
'                   Batch Incentives Collection Class by Work Center when
'                   it is required.
'Properties       : AddBIItem, Count, GetBIItem
'Called From      : DBCommon.LoadAllBatchIncnCols, Startup.Declarations
'Dependencies     : clsBIItem
'==============================================================================

' The Batch Incentive Items Collection
Private colTable As New Collection


Public Sub AddBIItem(strWCID As String, siMethod As Byte, _
        strDesc As String, siIncMin As Byte)
    'Add a Batch Incentive Code to the All Batch Incentive Code Collection.
    'Method of clsAllBITable object to create a new clsBIItem object with the
    '   properties passed in, and put it into the All Batch Incentive Items
    '   Collection.
    Dim newEntryItem As New clsBIItem
    newEntryItem.Code = Format(siMethod, "000")
    newEntryItem.Desc = strDesc
    newEntryItem.CodeDesc = newEntryItem.Code & " " & strDesc
    newEntryItem.WCID = UCase(strWCID)
    newEntryItem.Method = siMethod
    newEntryItem.IncMin = siIncMin
    colTable.Add Item:=newEntryItem, _
        Key:=UCase(strWCID) & newEntryItem.Code
End Sub


Property Get Count() As Long
    'Get a count of items in the All Batch Incentive Code Collection.
    'Property of claAllBITable object expose the Count property of the
    '   All Batch Incentive Items Collection (colTable).
    Count = colTable.Count
End Property


Public Sub getBIItem(anIndex As Variant, ByRef theItem As clsBIItem)
    'Lookup a Batch Incentive Code from the All Batch Incentive Code
    '   Collection.
    'Method of claAllBITable object. Populates the properties of the passed
    '   clsBIItem item with the properties of the Item (in the All Batch
    '   Incentive Items Collection) that matches the anIndex index passed in.
    On Error GoTo getBIItemError
    theItem.Code = colTable.Item(anIndex).Code
    theItem.Desc = colTable.Item(anIndex).Desc
    theItem.CodeDesc = colTable.Item(anIndex).CodeDesc
    theItem.WCID = colTable.Item(anIndex).WCID
    theItem.Method = colTable.Item(anIndex).Method
    theItem.IncMin = colTable.Item(anIndex).IncMin
    gBatchIncen.IncenMin = colTable.Item(anIndex).IncMin
    Exit Sub
    
getBIItemError:
    theItem.Code = ""
    theItem.Desc = LoadResString(gLanguageOffset + 4)
    theItem.CodeDesc = LoadResString(gLanguageOffset + 4)
    theItem.WCID = ""
    theItem.Method = 0
    theItem.IncMin = 0
    gBatchIncen.IncenMin = 0
   
End Sub




