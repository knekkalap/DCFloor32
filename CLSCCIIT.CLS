VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCCIITable"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Class Name       : clsCCIITable
'Description      : Class to contains the collection of Indirect Incentives
'                   for one Cost Center.  This table is refreshed and
'                   rebuilt for each transaction required.  Used to create
'                   global object gcolCCIICodes.
'Properties       : AddIIItem, Count, GetIIItem, Refresh
'Called From      : startup.declarations, DBCommon.LdIndirIncn,
'                   Inputs.GetIndirIncn
'Dependencies     : clsIIItem
'==============================================================================

' The Indirect Incentive Items Collection for one Cost Center
Private colTable As New Collection

Public Sub AddIIItem(strCCID As String, iType As Integer, _
     iNumb As Integer, cWrkStd As Currency, strDesc As String)
    'Add a Indirect Incentive Code to the Indirect Incentive Code Collection
    '   for one Cost Center.
    Dim newEntryItem As New clsIIItem
    newEntryItem.Code = Left(strDesc & String(30, " "), 30)
    newEntryItem.Desc = Format(cWrkStd, "00.0")
    newEntryItem.CodeDesc = newEntryItem.Code & " " & strDesc
    newEntryItem.CCID = UCase(strCCID)
    newEntryItem.IncenType = iType
    newEntryItem.MachNumb = iNumb
    newEntryItem.WrkStd = cWrkStd
    colTable.Add Item:=newEntryItem, _
        Key:=newEntryItem.Code
End Sub
Property Get Count() As Long
    'Get a count of items in the Indirect Incentive Code Collection
    '   for one Cost Center.
    Count = colTable.Count
End Property

Public Sub getIIItem(anIndex As Variant, ByRef theItem As clsIIItem)
    'Lookup a Indirect Incentive Code from the Indirect Incentive Code
    '   Collection for one Cost Center.
    
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

Public Sub Refresh()
    'Remove all Indirect Incentive Codes from the Indirect Incentive Code
    '   Collection for one Cost Center.
    Dim i As Variant
    
    If colTable.Count > 0 Then
        For i = colTable.Count To 1 Step -1
            colTable.Remove i
        Next i
    End If
End Sub

