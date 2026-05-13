VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWcBITable"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Class Name       : clsWcBITable
'Description      : Class to contains the collection of Batch Incentives
'                   for just one Work Center.  This table is refreshed and
'                   rebuilt for each transaction required.  Used to create
'                   global object gcolWcBICodes.
'Properties       : AddBIItem, Count, GetBIItem, Refresh
'Called From      : startup.declarations, DBCommon.LdBatchIncn
'Dependencies     : clsBIItem
'==============================================================================

' The Batch Incentive Items Collection for one Work Center
Private colTable As New Collection
Public Sub AddBIItem(strWCID As String, siMethod As Byte, _
        strDesc As String, siIncMin As Byte)
    'Add a Batch Incentive Code to the Batch Incentive Code Collection
    '   for one Work Center.
    Dim newEntryItem As New clsBIItem
    newEntryItem.Code = Format(siMethod, "000")
    newEntryItem.Desc = strDesc
    newEntryItem.CodeDesc = newEntryItem.Code & " " & strDesc
    newEntryItem.WCID = strWCID
    newEntryItem.Method = siMethod
    newEntryItem.IncMin = siIncMin
    colTable.Add Item:=newEntryItem, Key:=newEntryItem.Code
End Sub



Property Get Count() As Long
    'Get a count of items in the Batch Incentive Code Collection
    '   for one Work Center.
    Count = colTable.Count
End Property


Public Sub getBIItem(anIndex As Variant, anItem As clsBIItem)
    'Lookup a Batch Incentive Code from the Batch Incentive Code
    '   Collection for one Work Center.
    On Error GoTo getBIItemError
    Dim theItem As New clsBIItem
    theItem.Code = colTable.Item(anIndex).Code
    theItem.Desc = colTable.Item(anIndex).Desc
    theItem.CodeDesc = colTable.Item(anIndex).CodeDesc
    theItem.WCID = colTable.Item(anIndex).WCID
    theItem.Method = colTable.Item(anIndex).Method
    theItem.IncMin = colTable.Item(anIndex).IncMin
    gBatchIncen.IncenMin = colTable.Item(anIndex).IncMin
    anItem = theItem
    Exit Sub
    
getBIItemError:
    theItem.Code = ""
    theItem.Desc = LoadResString(gLanguageOffset + 4)
    theItem.CodeDesc = LoadResString(gLanguageOffset + 4)
    theItem.WCID = ""
    theItem.Method = 0
    theItem.IncMin = 0
    gBatchIncen.IncenMin = 0
    anItem = theItem
End Sub

Public Sub Refresh()
    'Remove all Batch Incentive Codes from the Batch Incentive Code
    '   Collection for one Work Center.
    Dim i As Variant
    
    If colTable.Count > 0 Then
        For i = colTable.Count To 1 Step -1
            colTable.Remove i
        Next i
    End If
End Sub


