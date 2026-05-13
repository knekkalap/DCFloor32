VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Class Name       : clsCodeTable
'Description      : The class that contains a collection of clsCodeItems used
'                   in the Menu Entry or Combo Box Lookup sections of the
'                   frmDCEntry form. Global instances of clsCodeTable include
'                   gcolIndirCodes, gcolScrapCodes, gcolUnplnCodes,
'                   gcolTlmntCodes, gcolOperCodes, gcolIICodes, gcolBICodes
'Properties       : Add, Count, GetCode, GetCodeDesc, GetDesc, Refresh
'Called From      : clsCodeTable.Add, clsEntry.LuTable, startup.LoadApp,
'                   startup.LoadLookupCols, startup.LoadOperCols,
'                   Inputs.inpIndirCd, Inputs.inpUnplnReas, Inputs.inpTlmntCd,
'                   Inputs.inpOperNumbScrapWind.txtScrapCd_LostFocus,
'                   DBCommon.LdIndirIncn, Inputs.getIndirIncn,
'                   DBCommon.ldBatchIncn, Inputs.getBatchIncn
'Dependencies     : clsCodeItem
'==============================================================================

' The Codes Collection
Private colTable As New Collection

Public Sub Refresh()
    'Remove all Code Lookup Items from the Code Lookup Table.
    Dim i As Integer
    
    If colTable.Count > 0 Then
        For i = colTable.Count To 1 Step -1
            colTable.Remove i
        Next i
    End If
End Sub




Public Sub Add(strCode As String, _
        strDesc As String)
    'Add a Code to the Code Lookup Collection.
    Dim newEntryItem As New clsCodeItem
    newEntryItem.Code = UCase(strCode)
    newEntryItem.Desc = strDesc
    newEntryItem.CodeDesc = UCase(strCode) & " " & strDesc
    colTable.Add Item:=newEntryItem, Key:=UCase(strCode)
End Sub


Property Get Count() As Long
    'Get a count of items in the Code Lookup Collection.
    Count = colTable.Count
End Property



Public Function getDesc(anItem As Variant) As String
    'Lookup a Description from the Code Lookup Collection.
    Dim i As Integer
    On Error GoTo getDescError:
    For i = 1 To colTable.Count
      If Trim(anItem) = Trim(colTable.Item(i).Code) Then
          getDesc = colTable.Item(i).Desc
          Exit Function
      End If
    Next i
  '  Exit Function
    
getDescError:
    getDesc = LoadResString(gLanguageOffset + 4)
End Function

Public Function getCodeDesc(anItem As Variant) As String
    'Lookup a Code & Description from the Code Lookup Collection.
    Dim i As Integer
    On Error GoTo getCodeDescError:
'    For i = 1 To colTable.Count
'      If anItem = colTable.Item(i).Code Then
     getCodeDesc = colTable.Item(anItem).CodeDesc
 '         Exit For
 '     End If
 '   Next i
    
    Exit Function
    
getCodeDescError:
    getCodeDesc = LoadResString(gLanguageOffset + 4)
End Function


Public Function getCode(anItem As Variant) As String
    'Lookup a Code from the Code Lookup Collection.
    
    On Error GoTo getCodeError
    getCode = colTable.Item(anItem).Code
    Exit Function
    
getCodeError:
    getCode = LoadResString(gLanguageOffset + 4)
End Function
