Attribute VB_Name = "DBCommon"
Option Explicit
Public Function GetBICode(ByVal biMethod As Integer, ByVal biWCID As String) As Boolean
    Dim BiItem As New clsBIItem
    
    On Error GoTo GetBICodeError
    
    Call gcolAllBICodes.getBIItem(biWCID & Format(biMethod, "000"), _
        BiItem)
    If BiItem.Code = "" Or BiItem.Code <> biMethod Then
        GetBICode = False
    Else
        GetBICode = True
    End If
    Exit Function
GetBICodeError:
    GetBICode = False
End Function
Public Sub initJobOpers()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : DBCommon.initJobOpers
'Description      : Clear all items in the JobOpers type structure.
'Parameters       : None
'Procedures Used  : None
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    gJobOpers.companyCd = 0
    gJobOpers.OrderNo = ""
    gJobOpers.OperSeqNo = 0
    gJobOpers.WCID = ""
    gJobOpers.Oper = ""
    gJobOpers.PlnQty = 0
    gJobOpers.PlnHrs = 0
    gJobOpers.SetStd = 0
    gJobOpers.MchStd = 0
    gJobOpers.EorSStd = ""
    gJobOpers.TimeBasisCd = ""
    gJobOpers.ActQty = 0
    gJobOpers.ActHrs = 0
    gJobOpers.scrapQty = 0
    gJobOpers.OperStatus = ""
End Sub
Public Function LdBatchIncn(strWCID As String)
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : DBCommon.LdBatchIncn
'Description      : Populate the two WCID Batch Incentive code collections
'                   with the entries that match the Work Center passed in
'                   (strWCID).  The gcolBICodes collection is used to pass
'                   into gEntry.LuTable for entry of the code by combo box.
'Parameters       : Work Center Code
'Procedures Used  : None
'Objects Used     : gcolAllBICodes(clsAllBITable),
'                   gcolWCBICodes(clsWcBITable), gcolBICodes(clsCodeTable)
'Forms Loaded     : None
'Return Parameters: Count of entries matching the Work Center
'Tables Updates   : None
'==============================================================================

    Dim I As Integer
    Dim iCount As Integer
    Dim theItem As New clsBIItem
    
    gcolWcBICodes.Refresh
    gcolBICodes.Refresh
    iCount = 0
    For I = 1 To gcolAllBICodes.Count
        Call gcolAllBICodes.getBIItem(I, theItem)
        If theItem.WCID = strWCID Then
            gcolWcBICodes.AddBIItem theItem.WCID, theItem.Method, _
                theItem.Desc, theItem.IncMin
            gcolBICodes.Add CStr(CInt(theItem.Code)), theItem.Desc
            iCount = iCount + 1
        End If
    Next I
    LdBatchIncn = iCount
End Function
Public Function LdIndirIncn(strCCID As String)
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : DBCommon.LdIndirIncn
'Description      : Populate the two CCID Indirece Incentive code collections
'                   with the entries that match the Cost Center passed in
'                   (strCCID).  The gcolIICodes collection is passed into
'                   the gEntry.LuTable for entry of the code by combo box.
'Parameters       : Cost Center Code
'Procedures Used  : None
'Objects Used     : gcolAllIICodes(clsAllIITable),
'                   gcolWCIICodes(clsWcIITable), gcolIICodes(clsCodeTable)
'Forms Loaded     : None
'Return Parameters: Count of entries matching the Cost Center
'Tables Updates   : None
'==============================================================================
    Dim I As Integer
    Dim iCount As Integer
    Dim theItem As New clsIIItem
    
    gcolCcIICodes.Refresh
    gcolIICodes.Refresh
    iCount = 0
    For I = 1 To gcolAllIICodes.Count
        Call gcolAllIICodes.getIIItem(I, theItem)
        If theItem.CCID = strCCID Then
            gcolCcIICodes.AddIIItem theItem.CCID, theItem.IncenType, _
                theItem.MachNumb, theItem.WrkStd, theItem.Code
            gcolIICodes.Add theItem.Code, theItem.Desc
            iCount = iCount + 1
        End If
    Next I
    LdIndirIncn = iCount
End Function

Public Function ErrorChk(ByVal aryNbr As Integer)
Dim rndHrs As String
Dim ernHrs As Currency
Dim pctEff As Currency
Dim lScrapTot As Long
Dim j As Integer
 
 If gLaborIn(aryNbr).RecType = "S" Then
    Exit Function
 End If
 
 If gLaborIn(aryNbr).RecType = "R" Then
    Call RecordError("R")
 End If
 
 If gLaborIn(aryNbr).SeqNo = 0 Or gLaborIn(aryNbr).SeqNo = 990 Then
    Call RecordError("U")
    Exit Function
 End If
 
 If gJobMaster.OrderStat = "C" Then
    Call RecordError("J")
 End If
 
 If gCompanyCd <> corry Then
    If gJobOpers.OperStatus = "C" Then
       Call RecordError("O")
    End If
    
    If Trim$(gLaborIn(aryNbr).WCID) <> Trim$(gJobOpers.WCID) Then
       Call RecordError("W")
    End If
 
   If gCompanyCd <> 77 Then
       If gLaborIn(aryNbr).setupHours = 0 And gLaborIn(aryNbr).machHours = 0 Then
            Call RecordError("H")
      End If
   End If
   
   If gCompanyCd = 77 Then
         lScrapTot = 0
        If gScrapColl Then
                For j = 1 To 8
                    lScrapTot = lScrapTot _
                        + gLaborIn(aryNbr).Scrap(j).Qty
                Next
        End If
        If lScrapTot > gLaborIn(aryNbr).ActQty Then
             Call RecordError("S")
       End If
  End If
  
    'If gLaborIn(aryNbr).MchStd > 0 And gLaborIn(aryNbr).machHours > 0 Then
    '   If gJobOpers.TimeBasisCd = 4 Then
    '     ernHrs = gLaborIn(aryNbr).ActQty / gLaborIn(aryNbr).MchStd
    '     rndHrs = Format(ernHrs, "000.00")
    '     ernHrs = CCur(rndHrs)
    '     pctEff = (ernHrs / gLaborIn(aryNbr).machHours) * 100#
    '   ElseIf gJobOpers.TimeBasisCd = 1 Then            'time per piece
    '     ernHrs = gLaborIn(aryNbr).ActQty * gLaborIn(aryNbr).MchStd
    '     rndHrs = Format(ernHrs, "000.00")
    '     ernHrs = CCur(rndHrs)
    '     pctEff = (ernHrs / gLaborIn(aryNbr).machHours) * 100#
    '   End If
    '   If pctEff > 200 Then
    '      Call RecordError("E")
    '   End If
    'End If
 
 End If      'if division not = corry
 
 If gLaborIn(aryNbr).addonHours > 0 Then
    Call RecordError("A")
 End If
End Function

'Public Function CheckForSQLServer() As Boolean

'Dim objWMIService As SWbemServices
'Dim colItems As WbemScripting.SWbemObjectSet
'Dim objItem As WbemScripting.SWbemObject
'Dim strMachineName As String
'
'strMachineName = gSQLServer
'
'On Error Resume Next
'Set objWMIService = GetObject("winmgmts:\\" & Trim(strMachineName) & "\root\cimv2")
'If Err.Number Then
'   MsgBox Err.Description
'End If
'
'Dim strQuery As String
'strQuery = "SELECT * FROM Win32_Service WHERE Name = 'MSSQLServer'"
'Set colItems = objWMIService.ExecQuery(strQuery, , wbemFlagForwardOnly Or wbemFlagReturnImmediately)
'If Err.Number Then
'   MsgBox Err.Description
'Else
'    If colItems.Count Then
'        For Each objItem In colItems
'            Debug.Print "Status: " & objItem.Status
'            Debug.Print "State: " & objItem.State
'        Next
'    Else
'        ' MSSQLServer service not running on strMachineName
'    End If
'End If
'End Function
