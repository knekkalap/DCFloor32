Attribute VB_Name = "DB"
Option Explicit
Private Const module_name = "DB.BAS"
Public Function InsertNote() As Boolean
Dim strfunctionname As String
Dim sql As String
Dim iNumRecs As Integer
Dim x As Integer
Dim iretnval As Integer
    
On Error GoTo ErrorHandler
strfunctionname = "InsertNote"
InsertNote = False
  
sql = "declare @status int Execute @status = SF_InsertNotepad "
sql = sql & gNotepad.companyCd & ", "
sql = sql & gNotepad.EmplNo & ", "
sql = sql & "'" & spDateParam(gNotepad.Actdate, 1) & "', "
sql = sql & "'" & Trim(gNotepad.MsgText) & "', "
sql = sql & "'" & gNotepad.PrintFlag & "' Select @status"
    
iretnval = DB.ExecuteQuery(sql)

If iretnval = 0 Then
    InsertNote = True
Else
    InsertNote = False
End If


Exit Function
    
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear
    
End Function
Public Function GetPaEmplyData() As Integer
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim strfunctionname As String
Dim sql As String
Dim iNumRecs As Integer
Dim x As Integer
    
On Error GoTo ErrorHandler
strfunctionname = "GetPAEmplyData"
    
GetPaEmplyData = 0
    
If gEmployee.ClassWrkr = 4 And div = stroms Then
        gEmployee.Actdate = gSysDate
End If

sql = "Exec SF_GetLaborForShopFloor " & gCompanyCd & ", "
sql = sql & gEmployee.EmplNo & ", '" & spDateParam(gEmployee.Actdate, 0) & "'"
    
iNumRecs = 0
Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)

'Set rs.ActiveConnection = Nothing
If rs.BOF And rs.EOF Then
Else
'        rs.MoveLast
'        rs.MoveFirst
        While Not rs.EOF
'            iNumRecs = iNumRecs + 1
            gPaEmplyData(iNumRecs).EmplNo = rs!Empl_No
            gPaEmplyData(iNumRecs).ShiftCd = rs!Shift_Cd
            gPaEmplyData(iNumRecs).IndirCd = rs!Indir_Cd
            gPaEmplyData(iNumRecs).OrderNo = rs!Order_No
            gPaEmplyData(iNumRecs).SeqNo = rs!seq_numb
            gPaEmplyData(iNumRecs).CCID = rs!CC_ID
            gPaEmplyData(iNumRecs).WCID = rs!WC_ID
            gPaEmplyData(iNumRecs).ActQty = rs!act_qty
            gPaEmplyData(iNumRecs).ActHrs = rs!act_hrs
            gPaEmplyData(iNumRecs).MachHrs = rs!Mach_hrs
            gPaEmplyData(iNumRecs).EmplEff = rs!empl_eff
            gPaEmplyData(iNumRecs).RecType = rs!rec_type
            gPaEmplyData(iNumRecs).scrapQty = rs!scrap_qty
            gPaEmplyData(iNumRecs).IncenCd = rs!incen_cd
            gPaEmplyData(iNumRecs).transdate = IIf(IsDate(rs!trans_date), Format(rs!trans_date, gDateformat & " " & gTimeFormat), " ")
            gPaEmplyData(iNumRecs).Oper_Cd = rs!Oper_Cd
            iNumRecs = iNumRecs + 1
            rs.MoveNext
        Wend
  End If
  If div <> stroms Then
        GetPaEmplyData = iNumRecs
        Exit Function
 End If
   
' next search open labor table for activity  - but only for stromsholmen

sql = "Exec SF_GetOpenLaborForShopFloor " & gCompanyCd & ", "
sql = sql & gEmployee.EmplNo
    
Set rs2 = New ADODB.Recordset
Set rs2 = DB.GetRecordsetData(sql)

If rs2.BOF And rs2.EOF Then
Else
'        rs2.MoveLast
'        rs2.MoveFirst
        While Not rs2.EOF
                gPaEmplyData(iNumRecs).EmplNo = rs2!Empl_No
                gPaEmplyData(iNumRecs).ShiftCd = rs2!Shift_Cd
                gPaEmplyData(iNumRecs).IndirCd = rs2!Indir_Cd
                gPaEmplyData(iNumRecs).OrderNo = rs2!Order_No
                gPaEmplyData(iNumRecs).SeqNo = rs2!seq_numb
                gPaEmplyData(iNumRecs).CCID = rs2!CC_ID
                gPaEmplyData(iNumRecs).WCID = rs2!WC_ID
                gPaEmplyData(iNumRecs).transdate = IIf(IsDate(rs2!start_datetime), Format(rs2!start_datetime, gDateformat & " " & gTimeFormat), " ")
                gPaEmplyData(iNumRecs).ActQty = 0
                gPaEmplyData(iNumRecs).ActHrs = 0
                gPaEmplyData(iNumRecs).MachHrs = 0
                gPaEmplyData(iNumRecs).EmplEff = 0
                gPaEmplyData(iNumRecs).RecType = "O"
                gPaEmplyData(iNumRecs).scrapQty = 0
                gPaEmplyData(iNumRecs).IncenCd = 0
                gPaEmplyData(iNumRecs).Oper_Cd = ""
                iNumRecs = iNumRecs + 1
                rs2.MoveNext
           Wend
End If
' set function to false only if nothing is returned
 If iNumRecs = 0 Then
        GetPaEmplyData = False
 Else
       GetPaEmplyData = iNumRecs
 End If

Cleanup:
    Set rs = Nothing
    Set rs2 = Nothing
    Exit Function
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    
    gError.ErrorRoutine
    gError.Clear
    
  End Function
Public Function GetMsgsForShopFloor(ByVal displaytime As String) As Integer
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim iResp As Integer
    Dim strfunctionname As String
    Dim iNumRecs As Integer
    Dim x As Integer
    Dim CurrentDate As Date
    GetMsgsForShopFloor = 0
  
    On Error GoTo ErrorHandler
    strfunctionname = "GetMsgsForShopFloor"
        
    sql = "Exec SF_GetMsgForShop "
    sql = sql & gEmployee.EmplNo & ", '"
    sql = sql & spDateParam(gSysDateTime, 1) & "', '"   'use sysdatetime instead of employee.actdate
    sql = sql & Trim(displaytime) & "' "

   
    iNumRecs = 0
    
    Set rs = New ADODB.Recordset
    Set rs = DB.GetRecordsetData(sql)
    If rs.EOF And rs.BOF Then
        GetMsgsForShopFloor = 0
    Else
'only retrieve one message - the one returned is prioritized for this employee
       iNumRecs = iNumRecs + 1
'       ReDim Preserve gOTOffering(iNumRecs)
       gMessage.msg_type = rs!message_type
       gMessage.message = rs!message
       gMessage.message_id = rs!message_id
       GetMsgsForShopFloor = 1
    End If
Cleanup:
    Set rs = Nothing
    Exit Function
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear
End Function
Public Function GetOTOfferingData() As Integer
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim iResp As Integer
    Dim strfunctionname As String
    Dim iNumRecs As Integer
    Dim x As Integer
    Dim CurrentDate As Date
    GetOTOfferingData = 0
  
    On Error GoTo ErrorHandler
    strfunctionname = "getOTOfferingdata"
    
    gEmployee.Actdate = gSysDate
        
    sql = "Exec SF_GetOTOffering "
    sql = sql & gEmployee.EmplNo & ", '"
    sql = sql & spDateParam(gEmployee.Actdate, 0) & "'"
   
    iNumRecs = 0
    
    Set rs = New ADODB.Recordset
    Set rs = DB.GetRecordsetData(sql)
    If rs.EOF And rs.BOF Then
        GetOTOfferingData = 0
    Else
'        rs.MoveLast
'        rs.MoveFirst
        While Not rs.EOF
               iNumRecs = iNumRecs + 1
               ReDim Preserve gOTOffering(iNumRecs)
               gOTOffering(iNumRecs).EmplNo = rs!Empl_No
               gOTOffering(iNumRecs).Actdate = Format(rs!Act_date, "MM/DD/YYYY")
               gOTOffering(iNumRecs).OTCCID = rs!OT_Costcenter
               gOTOffering(iNumRecs).OfferedHrs = rs!hrs_offered
               gOTOffering(iNumRecs).AcceptedHrs = rs!hrs_accepted
               rs.MoveNext
          Wend
          GetOTOfferingData = iNumRecs
    End If
Cleanup:
    Set rs = Nothing
    Exit Function
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    
    gError.ErrorRoutine
    gError.Clear

End Function


Public Function LoadShiftCodes(Info() As ShiftRecType) As Boolean
Dim rs As ADODB.Recordset
Dim sql As String
Dim iResp As Integer
Dim iCompanyCd, iShftNo, iShftGrp As Integer
Dim strfunctionname As String

On Error GoTo ErrorHandler
strfunctionname = "LoadshiftCodes"

LoadShiftCodes = True
    
sql = "exec SF_GetAllShiftRecs"

 Set rs = New ADODB.Recordset
 Set rs = DB.GetRecordsetData(sql)
 If rs.EOF And rs.BOF Then
     LoadShiftCodes = False
 Else
'       rs.MoveLast
'       rs.MoveFirst
       While Not rs.EOF
            iCompanyCd = CInt(rs(0))
            iShftGrp = CInt(rs(1))
            iShftNo = CInt(rs(2))
            gShiftData(iShftGrp, iShftNo).startTime = rs(3)
            gShiftData(iShftGrp, iShftNo).LunchOut = rs(4)
            gShiftData(iShftGrp, iShftNo).LunchIn = rs(5)
            gShiftData(iShftGrp, iShftNo).endTime = rs(6)
            gShiftData(iShftGrp, iShftNo).StParam = rs(7)
            gShiftData(iShftGrp, iShftNo).EndParam = rs(8)
            gShiftData(iShftGrp, iShftNo).ACT = rs(9)
            gShiftData(iShftGrp, iShftNo).lunchGrace = rs(10)
            gShiftData(iShftGrp, iShftNo).EOSStart = rs(11)
            gShiftData(iShftGrp, iShftNo).EOSEnd = rs(12)
            rs.MoveNext
        Wend
        LoadShiftCodes = True
 End If
Cleanup:
    Set rs = Nothing
    Exit Function
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear
End Function
Public Function GetEmplyShiftGrp(iShiftGrp As Integer)
Dim rs As ADODB.Recordset
Dim sql As String
Dim iResp As Integer
Dim iCompanyCd, iShftNo As Integer
Dim strfunctionname As String

On Error GoTo ErrorHandler
strfunctionname = "GetEmplyShiftGrp"

'LoadShiftCodes = True
    
sql = "exec SF_GetShiftRecData " & iShiftGrp

 Set rs = New ADODB.Recordset
 Set rs = DB.GetRecordsetData(sql)
 If rs.EOF And rs.BOF Then
'     LoadShiftCodes = False
 Else
      Erase gShiftRecData()
       While Not rs.EOF
            iCompanyCd = CInt(rs(0))
'            iShftGrp = CInt(rs(1))
            iShftNo = CInt(rs(2))
            gShiftRecData(iShftNo).startTime = rs(3)
            gShiftRecData(iShftNo).LunchOut = rs(4)
            gShiftRecData(iShftNo).LunchIn = rs(5)
            gShiftRecData(iShftNo).endTime = rs(6)
            gShiftRecData(iShftNo).StParam = rs(7)
            gShiftRecData(iShftNo).EndParam = rs(8)
            gShiftRecData(iShftNo).ACT = rs(9)
            gShiftRecData(iShftNo).lunchGrace = rs(10)
            gShiftRecData(iShftNo).EOSStart = rs(11)
            gShiftRecData(iShftNo).EOSEnd = rs(12)
            rs.MoveNext
        Wend
        'LoadShiftCodes = True
 End If
Cleanup:
    Set rs = Nothing
    Exit Function
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear
End Function

Public Function GetWorkCenter(ByVal sWCID As String) As Boolean
Dim rs As ADODB.Recordset
Dim sql As String
Dim strfunctionname As String

Dim iResp As Integer
Dim iNumRecs As Integer
Dim x As Integer
    
On Error GoTo ErrorHandler
strfunctionname = "getworkcenter"
GetWorkCenter = False
    
sql = "Exec SF_SelectWorkcenter "
sql = sql & gCompanyCd & ", '" & sWCID & "'"

iNumRecs = 0
Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)

If rs.BOF And rs.EOF Then
Else
'      rs.MoveLast
'      rs.MoveFirst
      iNumRecs = iNumRecs + 1
       GetWorkCenter = True
       gWorkCenter.companyCd = rs!company_cd
       gWorkCenter.WCID = rs!WC_ID
       gWorkCenter.WcGroup = rs!wc_group
       gWorkCenter.Location = rs!Location
       gWorkCenter.Descr = rs!Descr
       gWorkCenter.CCID = rs!CC_ID
       gWorkCenter.PaidLunch = rs!pd_lunch
 End If
Cleanup:
    Set rs = Nothing
    Exit Function
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear

End Function
Public Function GetCostCenter(ByVal sCcid As String) As Boolean
Dim rs As ADODB.Recordset
Dim sql As String
Dim iResp As Integer
Dim iNumRecs As Integer
Dim x As Integer
Dim strfunctionname As String

On Error GoTo ErrorHandler
strfunctionname = "GetCostCenter"

GetCostCenter = False
    
sql = "Exec SF_SelectCostcenter "
sql = sql & gCompanyCd & ", '" & sCcid & "'"

iNumRecs = 0
Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)
If rs.BOF And rs.EOF Then
Else
'      rs.MoveLast
'      rs.MoveFirst
      iNumRecs = iNumRecs + 1
      GetCostCenter = True
      gCostCenter.companyCd = rs!company_cd
      gCostCenter.CCID = rs!CC_ID
      gCostCenter.Descr = rs!Descr
      gCostCenter.SetupRate = rs!setup_rate
      gCostCenter.OvrhdRate = rs!ovrhd_rate
      gCostCenter.MachRate = rs!machn_rate
      gCostCenter.PaidLunch = rs!pd_lunch
 End If
Cleanup:
    Set rs = Nothing
    Exit Function
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear

End Function

Public Function GetJobMaster(ByVal sOrderNo As String) As Boolean
Dim rs As ADODB.Recordset
Dim strfunctionname As String
Dim sql As String
Dim iNumRecs As Integer
    
  On Error GoTo ErrorHandler
  strfunctionname = "GetJobMaster"
  
GetJobMaster = False
sql = "Exec SF_Selectjobmaster '" & sOrderNo & "'"
iNumRecs = 0

Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)

If rs.EOF And rs.BOF Then
Else
'     rs.MoveLast
'     rs.MoveFirst
     iNumRecs = iNumRecs + 1
     GetJobMaster = True
     gJobMaster.companyCd = rs!company_cd
     gJobMaster.OrderNo = rs!Order_No
     gJobMaster.OrderType = rs!Order_Type
     gJobMaster.OrderQty = rs!Order_Qty
     gJobMaster.PartNo = IIf(rs!part_number = "", "               ", rs!part_number)
     gJobMaster.PartRev = IIf(rs!Part_Rev = "", " ", rs!Part_Rev)
     gJobMaster.CustCode = IIf(rs!cust_code = "", " ", rs!cust_code)
     gJobMaster.CustName = IIf(rs!cust_name = "", " ", rs!cust_name)
     gJobMaster.Actdate = ConvSqlDate(rs!Act_date)
     gJobMaster.OrderStat = rs!order_status
      gJobMaster.MatlWgt = IIf(rs!matl_wgt = "", 0, rs!matl_wgt)
 End If

Cleanup:
    Set rs = Nothing
    Exit Function
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear

End Function
Public Function CheckWorkOrder(ByVal sOrderNo As String) As Boolean
Dim rs As ADODB.Recordset
Dim strfunctionname As String
Dim sql             As String
Dim iNumRecs        As Integer
Dim strWorkOrder    As String
    
On Error GoTo ErrorHandler
strfunctionname = "checkworkorder"

CheckWorkOrder = False
        
sql = "Exec SF_CheckWorkOrder '" & sOrderNo & "'"
Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)
    
    iNumRecs = 0
If rs.EOF And rs.BOF Then
Else
    iNumRecs = iNumRecs + 1
    CheckWorkOrder = True
    strWorkOrder = rs!Order_No
End If
Cleanup:
    Set rs = Nothing
    Exit Function
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear

End Function
Public Function GetJobOper(ByVal sOrderNo As String, ByVal iSeqNo As Integer) As Boolean
Dim rs As ADODB.Recordset
Dim strfunctionname As String
Dim sql As String
Dim iNumRecs As Integer
    
On Error GoTo ErrorHandler
strfunctionname = "getjoboper"
GetJobOper = False
    
sql = "exec SF_selectJobOper '" & sOrderNo
sql = sql & "', " & iSeqNo

Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)

iNumRecs = 0
If rs.EOF And rs.BOF Then
Else
'      rs.MoveLast
'      rs.MoveFirst
      iNumRecs = iNumRecs + 1
       GetJobOper = True
       gJobOpers.companyCd = rs!company_cd
       gJobOpers.OrderNo = rs!Order_No
       gJobOpers.OperSeqNo = rs!seq_numb
       gJobOpers.WCID = rs!WC_ID
       gJobOpers.Oper = rs!Oper_Cd
       gJobOpers.PlnQty = rs!pln_qty
       gJobOpers.PlnHrs = rs!pln_hrs
       gJobOpers.SetStd = rs!set_std
       gJobOpers.MchStd = rs!mch_std
       gJobOpers.EorSStd = rs!e_or_s_std
       gJobOpers.TimeBasisCd = rs!time_basis_cd
       gJobOpers.ActQty = rs!act_qty
       gJobOpers.ActHrs = rs!act_hrs
       gJobOpers.scrapQty = rs!scrap_qty
       gJobOpers.OperStatus = rs!oper_status
 End If
Cleanup:
    Set rs = Nothing
    Exit Function
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear

End Function
Public Function GetEmployee(ByVal lemplno As Long) As Boolean
Dim rs As ADODB.Recordset
Dim strfunctionname As String
Dim sql As String
Dim iNumRecs As Integer
    
On Error GoTo ErrorHandler
strfunctionname = "GetEmployee"

sql = "Exec SF_SelectEmployee " & lemplno

GetEmployee = False
iNumRecs = 0

Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)

If rs.EOF And rs.BOF Then
Else
  '    rs.MoveLast
 '     rs.MoveFirst
     GetEmployee = True
     iNumRecs = iNumRecs + 1
     gEmployee.companyCd = rs!company_cd
     gEmployee.EmplNo = rs!Empl_No
     gEmployee.shiftGrp = rs!Shift_Grp
     gEmployee.EmplName = rs!empl_name
     gEmployee.HomeCCID = rs!home_cc_id
     gEmployee.ClassWrkr = rs!class_worker
     gEmployee.TermCd = rs!term_cd
     gEmployee.PdLunch = rs!pd_lunch
     gEmployee.HourlyRate = IIf(rs!hourly_rate = "", 0, rs!hourly_rate)
     gEmployee.IncenRate = IIf(rs!incen_rate = "", 0, rs!incen_rate)
     gEmployee.GMRate = IIf(rs!gm_rate = "", 0, rs!gm_rate)
     gEmployee.AERRate = IIf(rs!aer_rate = "", 0, rs!aer_rate)
     gEmployee.SpeclRate = IIf(rs!specl_rate = "", 0, rs!specl_rate)
     gEmployee.Actdate = ConvSqlDate(rs!Act_date)
     gEmployee.AsgnShift = rs!asgn_shift
     gEmployee.GenShift = IIf(rs!gen_shift = "", 0, rs!gen_shift)
     gEmployee.IndirCCID = IIf(rs!indir_cc_id = "", "     ", rs!indir_cc_id)
     gEmployee.IndirCd = rs!Indir_Cd
     gEmployee.AttdStartDate = ConvSqlDate(rs!attd_start)
     gEmployee.AttdEndDate = ConvSqlDate(rs!attd_end)
'     gEmployee.AttdStartDate = IIf(IsNull(rs!attd_start), dBlankDate, ConvSqlDate(rs!attd_start))
'     gEmployee.AttdEndDate = IIf(IsNull(rs!attd_end), dBlankDate, ConvSqlDate(rs!attd_end))
     gEmployee.AttdStatus = IIf(rs!attd_status = "", 2, rs!attd_status)
'     gEmployee.LaborStartDate = IIf(IsNull(rs!labor_start), dBlankDate, ConvSqlDate(rs!labor_start))
'     gEmployee.ShiftStartDate = IIf(IsNull(rs!shift_start), dBlankDate, ConvSqlDate(rs!shift_start))
     gEmployee.LaborStartDate = ConvSqlDate(rs!labor_start)
      'gEmployee.ShiftStartDate = rs!shift_start
     gEmployee.ShiftStartDate = ConvSqlDate(rs!shift_start)
     gEmployee.NewHrlyRate = rs!new_hourly_rate
     gEmployee.NewRateEff = ConvSqlDate(rs!new_rate_eff)
     gEmployee.LeadHand = rs!lead_hand
     
     If gEmployee.ClassWrkr = 4 And div = stroms Then
         'always clocked in
         gEmployee.Actdate = gSysDate
         gEmployee.GenShift = 1
     End If
End If
Cleanup:
    Set rs = Nothing
    Exit Function
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear

End Function
Public Function GetDivParam(ByVal sParamKey As String) As Boolean
Dim rs As ADODB.Recordset
Dim sql As String
Dim strfunctionname As String
    
On Error GoTo ErrorHandler
strfunctionname = "GetDivParam"
GetDivParam = False

Dim iNumRecs As Integer
Dim x As Integer
   
sql = "Exec SF_GetDivParam '" & sParamKey & "'"

iNumRecs = 0

Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)
    
If rs.BOF And rs.EOF Then
Else
'       rs.MoveLast
'       rs.MoveFirst
       iNumRecs = iNumRecs + 1
       gDivParam.ParamKey = rs!param_key
       gDivParam.ParamData = rs!param_data
       GetDivParam = True
End If
     
Cleanup:
    Set rs = Nothing
    Exit Function
    
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    
    gError.ErrorRoutine
    gError.Clear

End Function

Public Function Ck_for_Attd_Rec(ByVal lemplno As Long, ByVal dActDate As Date) As Boolean
Dim rs As ADODB.Recordset
Dim strfunctionname As String
Dim sql As String
Dim iNumRecs As Integer
    
On Error GoTo ErrorHandler
strfunctionname = "ck_for_attd_rec"
Ck_for_Attd_Rec = False

sql = "Exec SF_SelectAttendance " & lemplno & ", "
sql = sql & "'" & spDateParam(dActDate, 0) & "'"
iNumRecs = 0

Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)
If rs.EOF And rs.BOF Then
Else
'            rs.MoveLast
'            rs.MoveFirst
            Ck_for_Attd_Rec = True
            iNumRecs = iNumRecs + 1
            gAttendance.companyCd = rs!company_cd
            gAttendance.EmplNo = rs!Empl_No
            gAttendance.Actdate = ConvSqlDate(rs!Act_date)
            gAttendance.ShiftCd = rs!Shift_Cd
            gAttendance.CCID = rs!CC_ID
            gAttendance.PdLunch = rs!pd_lunch
            gAttendance.HrsUnrnd = IIf(rs!hrs_unrnd = "", 0, rs!hrs_unrnd)
            gAttendance.ActHrs = IIf(rs!act_hrs = "", 0, rs!act_hrs)
            gAttendance.ErrorFlag = IIf(rs!Error_Flag = "", " ", rs!Error_Flag)
            gAttendance.ErrorLst = IIf(rs!Error_Lst = "", " ", rs!Error_Lst)
            gAttendance.ClockIn = ConvSqlDate(rs!CLOCK_IN)
            gAttendance.ClockOut = ConvSqlDate(rs!CLOCK_OUT)
            gAttendance.LunchStart = ConvSqlDate(rs!lunch_start)
            gAttendance.LunchEnd = ConvSqlDate(rs!lunch_end)
            gAttendance.SuspLbr1 = ConvSqlDate(rs!susp_lbr_1)
            gAttendance.ResmLbr1 = ConvSqlDate(rs!resm_lbr_1)
            gAttendance.SuspLbr2 = ConvSqlDate(rs!susp_lbr_2)
            gAttendance.ResmLbr2 = ConvSqlDate(rs!resm_lbr_2)
            gAttendance.CIStart = ConvSqlDate(rs!CI_start)
            gAttendance.CIEnd = ConvSqlDate(rs!CI_End)
'Note ConvSQLVals function needs to be tested in Sweden environment
            gAttendance.ClockInRnd = IIf(IsNull(rs!Clock_in_rnd), INIT_RND_TM, rs!Clock_in_rnd)
            gAttendance.ClockOutRnd = IIf(IsNull(rs!Clock_out_rnd), INIT_RND_TM, rs!Clock_out_rnd)
            gAttendance.LunchOutRnd = IIf(IsNull(rs!Lunch_Out_Rnd), INIT_RND_TM, rs!Lunch_Out_Rnd)
            gAttendance.LunchInRnd = IIf(IsNull(rs!Lunch_in_Rnd), INIT_RND_TM, rs!Lunch_in_Rnd)
            gAttendance.SuspLbr1Rnd = IIf(IsNull(rs!susp_lb1_rnd), INIT_RND_TM, rs!susp_lb1_rnd)
            gAttendance.ResmLbr1Rnd = IIf(IsNull(rs!resm_lb1_rnd), INIT_RND_TM, rs!resm_lb1_rnd)
            gAttendance.SuspLbr2Rnd = IIf(IsNull(rs!susp_lb2_rnd), INIT_RND_TM, rs!susp_lb2_rnd)
            gAttendance.ResmLbr2Rnd = IIf(IsNull(rs!resm_lb2_rnd), INIT_RND_TM, rs!resm_lb2_rnd)
            gAttendance.CIStartRnd = IIf(IsNull(rs!ci_start_rnd), INIT_RND_TM, rs!ci_start_rnd)
            gAttendance.CIEndRnd = IIf(IsNull(rs!ci_end_rnd), INIT_RND_TM, rs!ci_end_rnd)
            gAttendance.CallInHrs = rs!call_in_hrs
            gAttendance.Earnings = rs!Earnings
            gAttendance.PayRate = rs!pay_rate
 End If

Cleanup:
    Set rs = Nothing
     Exit Function
    
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear
  
End Function

Public Function Valid_Empl_badge(ByVal sEmplNo As String) As Boolean
Dim lemplno As Long
Valid_Empl_badge = False
sEmplNo = Trim(sEmplNo)
If ValidNumber(sEmplNo) = False _
    Or Len(sEmplNo) > 6 Or _
       Len(sEmplNo) < 1 Then
    ErrBox LoadResString(gLanguageOffset + 75), 4
    Exit Function
Else
    lemplno = CLng(sEmplNo)
    If GetEmployee(lemplno) = False Then
       ErrBox LoadResString(gLanguageOffset + 44), 4
       Exit Function
    ElseIf gEmployee.TermCd = "Y" Then
            ErrBox LoadResString(gLanguageOffset + 44), 4
            Exit Function
    Else
      Valid_Empl_badge = True
    End If
 End If
 GetEmplyShiftGrp (gEmployee.shiftGrp)
End Function

Public Sub Upd_Attd_rec()
Dim sql As String
Dim iResp As Integer
    
'Because the stored procedute Update_attendance will
'automatically create an indirect labor record if the
'indirect code on the employee master is non-blank, the
'code in the variable must be blanked out here.
If AttdIndirLbr = False Then
     gEmployee.IndirCd = "  "
End If

sql = "Declare @status int Exec @status = SF_UpdateAttendance "
sql = sql & gCompanyCd & ", "
sql = sql & uEmployee.EmplNo & ", "
If uEmployee.Actdate = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(uEmployee.Actdate, 0) & "', "
End If
sql = sql & uEmployee.GenShift & ", "
If uEmployee.AttdStartDate = dBlankDate Then
    sql = sql & "NULL, "
Else
   sql = sql & "'" & spDateParam(uEmployee.AttdStartDate, 1) & "', "
End If
If uEmployee.AttdEndDate = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(uEmployee.AttdEndDate, 1) & "', "
End If
sql = sql & uEmployee.AttdStatus & ", "
If uEmployee.LaborStartDate = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(uEmployee.LaborStartDate, 1) & "', "
End If
If uEmployee.ShiftStartDate = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(uEmployee.ShiftStartDate, 1) & "', "
End If
sql = sql & "'" & gAttendance.ErrorLst & "', "
sql = sql & "'" & gAttendance.ErrorFlag & "', "
sql = sql & "'" & gAttendance.PdLunch & "', "
sql = sql & "'" & gAttendance.CCID & "', "
sql = sql & Str(gAttendance.ActHrs) & ", "
sql = sql & Str(gAttendance.HrsUnrnd) & ", "
If gAttendance.ClockIn = dBlankDate Then
    sql = sql & "NULL, "
Else
   sql = sql & "'" & spDateParam(gAttendance.ClockIn, 1) & "', "
End If
If gAttendance.ClockOut = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.ClockOut, 1) & "', "
End If
If gAttendance.LunchStart = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.LunchStart, 1) & "', "
End If
If gAttendance.LunchEnd = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.LunchEnd, 1) & "', "
End If
If gAttendance.SuspLbr1 = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.SuspLbr1, 1) & "', "
End If
If gAttendance.ResmLbr1 = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.ResmLbr1, 1) & "', "
End If
If gAttendance.SuspLbr2 = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.SuspLbr2, 1) & "', "
End If
If gAttendance.ResmLbr2 = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.ResmLbr2, 1) & "', "
End If
If gAttendance.CIStart = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.CIStart, 1) & "', "
End If
If gAttendance.CIEnd = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.CIEnd, 1) & "', "
End If
If gAttendance.ClockInRnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.ClockInRnd) & ", "
End If
If gAttendance.ClockOutRnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.ClockOutRnd) & ", "
End If
If gAttendance.LunchOutRnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.LunchOutRnd) & ", "
End If
If gAttendance.LunchInRnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.LunchInRnd) & ", "
End If
If gAttendance.SuspLbr1Rnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.SuspLbr1Rnd) & ", "
End If
If gAttendance.ResmLbr1Rnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.ResmLbr1Rnd) & ", "
End If
If gAttendance.SuspLbr2Rnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.SuspLbr2Rnd) & ", "
End If
If gAttendance.ResmLbr2Rnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.ResmLbr2Rnd) & ", "
End If
If gAttendance.CIStartRnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.CIStartRnd) & ", "
End If
If gAttendance.CIEndRnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.CIEndRnd) & ", "
End If
sql = sql & Str(gAttendance.CallInHrs) & ", "
sql = sql & Str(gAttendance.Earnings) & ", "
sql = sql & Str(gAttendance.PayRate) & ", "
sql = sql & "'" & gEmployee.IndirCd & "', "
sql = sql & "'" & spDateParam(gSysDateTime, 1) & "', "
sql = sql & Str(gLabor.Earnings) & ", "
sql = sql & Str(gLabor.PayRate)
sql = sql & " Select @status"

iResp = DB.ExecuteQuery(sql)

End Sub
Public Sub Ins_Attd_rec()
Dim sql As String
Dim iResp As Integer

'Because the stored procedute Insert_attendance will
'automatically create an indirect labor record if the
'indirect code on the employee master is non-blank, the
'code in the variable must be blanked out here.
If AttdIndirLbr = False Then
     gEmployee.IndirCd = "  "
End If

sql = "Declare @status int Exec @status = SF_InsertAttendance "
sql = sql & gCompanyCd & ", "
sql = sql & uEmployee.EmplNo & ", "
If uEmployee.Actdate = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(uEmployee.Actdate, 0) & "', "
End If
sql = sql & uEmployee.GenShift & ", "
If uEmployee.AttdStartDate = dBlankDate Then
    sql = sql & "NULL, "
Else
   sql = sql & "'" & spDateParam(uEmployee.AttdStartDate, 1) & "', "
End If
If uEmployee.AttdEndDate = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(uEmployee.AttdEndDate, 1) & "', "
End If
sql = sql & uEmployee.AttdStatus & ", "
If uEmployee.LaborStartDate = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(uEmployee.LaborStartDate, 1) & "', "
End If
If uEmployee.ShiftStartDate = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(uEmployee.ShiftStartDate, 1) & "', "
End If
sql = sql & gAttendance.ShiftCd & ", "
sql = sql & "'" & gAttendance.ErrorLst & "', "
sql = sql & "'" & gAttendance.ErrorFlag & "', "
sql = sql & "'" & gAttendance.PdLunch & "', "
sql = sql & "'" & gAttendance.CCID & "', "
sql = sql & Str(gAttendance.ActHrs) & ", "
sql = sql & Str(gAttendance.HrsUnrnd) & ", "
If gAttendance.ClockIn = dBlankDate Then
    sql = sql & "NULL, "
Else
   sql = sql & "'" & spDateParam(gAttendance.ClockIn, 1) & "', "
End If
If gAttendance.ClockOut = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.ClockOut, 1) & "', "
End If
If gAttendance.LunchStart = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.LunchStart, 1) & "', "
End If
If gAttendance.LunchEnd = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.LunchEnd, 1) & "', "
End If
If gAttendance.SuspLbr1 = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.SuspLbr1, 1) & "', "
End If
If gAttendance.ResmLbr1 = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.ResmLbr1, 1) & "', "
End If
If gAttendance.SuspLbr2 = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.SuspLbr2, 1) & "', "
End If
If gAttendance.ResmLbr2 = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.ResmLbr2, 1) & "', "
End If
If gAttendance.CIStart = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.CIStart, 1) & "', "
End If
If gAttendance.CIEnd = dBlankDate Then
    sql = sql & "NULL, "
Else
    sql = sql & "'" & spDateParam(gAttendance.CIEnd, 1) & "', "
End If
If gAttendance.ClockInRnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.ClockInRnd) & ", "
End If
If gAttendance.ClockOutRnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.ClockOutRnd) & ", "
End If
If gAttendance.LunchOutRnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.LunchOutRnd) & ", "
End If
If gAttendance.LunchInRnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.LunchInRnd) & ", "
End If
If gAttendance.SuspLbr1Rnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.SuspLbr1Rnd) & ", "
End If
If gAttendance.ResmLbr1Rnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.ResmLbr1Rnd) & ", "
End If
If gAttendance.SuspLbr2Rnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.SuspLbr2Rnd) & ", "
End If
If gAttendance.ResmLbr2Rnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.ResmLbr2Rnd) & ", "
End If
If gAttendance.CIStartRnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.CIStartRnd) & ", "
End If
If gAttendance.CIEndRnd = INIT_RND_TM Then
    sql = sql & "NULL, "
Else
    sql = sql & Str(gAttendance.CIEndRnd) & ", "
End If
sql = sql & Str(gAttendance.CallInHrs) & ", "
sql = sql & Str(gAttendance.Earnings) & ", "
sql = sql & Str(gAttendance.PayRate) & ", "
sql = sql & "'" & gEmployee.IndirCd & "', "
sql = sql & "'" & spDateParam(gSysDateTime, 1) & "', "
sql = sql & Str(gLabor.Earnings) & ", "
sql = sql & Str(gLabor.PayRate)
sql = sql & " Select @status"

iResp = DB.ExecuteQuery(sql)

End Sub
Public Function First_Labor_Rec(ByVal EmplNo As Long, ByVal Actdate As Date) As Boolean
Dim strfunctionname As String
Dim rs As ADODB.Recordset
Dim sql As String
Dim iResp As Integer

On Error GoTo ErrorHandler
strfunctionname = "first_labor_rec"

First_Labor_Rec = True

If gEmployee.LaborStartDate = gEmployee.AttdStartDate Then
     First_Labor_Rec = True
     Exit Function
End If

sql = "Exec SF_FirstLaborRec " & EmplNo & ", '" & spDateParam(Actdate, 0) & "'"

Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)

If rs.BOF And rs.EOF Then
Else
'      rs.MoveLast
'      rs.MoveFirst
      iResp = rs(0)
      If iResp > 0 Then
          First_Labor_Rec = False
     End If
End If

Cleanup:
    Set rs = Nothing
    Exit Function
   
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear
End Function

Public Function OpenIndirectLabor(ByVal EmplNo As Long, ByVal Actdate As Date) As Boolean
Dim rs As ADODB.Recordset
Dim strfunctionname As String
Dim sql As String
Dim iResp As Integer

On Error GoTo ErrorHandler
strfunctionname = "OpenIndirectLabor"

OpenIndirectLabor = False

sql = "Exec SF_OpenIndirectLabor " & EmplNo & ", '" & spDateParam(Actdate, 0) & "' "

Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)

If rs.EOF And rs.BOF Then
Else
'      rs.MoveLast
'      rs.MoveFirst
      gOpenLabor.IndirCd = rs!Indir_Cd
      gOpenLabor.CCID = rs!CC_ID
      OpenIndirectLabor = True
End If

Cleanup:
    Set rs = Nothing
    Exit Function
   
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear

End Function

Public Function OpenDirectLabor(ByVal EmplNo As Long, ByVal OrderNo As String, ByVal SeqNo As Integer, ByVal Actdate As Date) As Boolean
Dim rs As ADODB.Recordset
Dim strfunctionname As String

Dim sql As String
Dim iResp As Integer
On Error GoTo ErrorHandler
strfunctionname = "OpenDirectLabor"

OpenDirectLabor = False

sql = "Exec SF_OpenDirectLabor " & EmplNo & ", '" & OrderNo & "', "
sql = sql & SeqNo & ", '" & spDateParam(Actdate, 0) & "' "

Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)
If rs.BOF And rs.EOF Then
Else
'     rs.MoveLast
'     rs.MoveFirst
     OpenDirectLabor = True
     gOpenLabor.companyCd = rs!company_cd
     gOpenLabor.EmplNo = rs!Empl_No
     gOpenLabor.OrderNo = rs!Order_No
     gOpenLabor.SeqNumb = rs!seq_numb
     gOpenLabor.OperCd = rs!Oper_Cd
     gOpenLabor.WCID = rs!WC_ID
     gOpenLabor.CCID = rs!CC_ID
     gOpenLabor.StartDateTime = ConvSqlDate(rs!start_datetime)
     gOpenLabor.UnplnReas = rs!unpln_reas
     gOpenLabor.ElapsedHrs = rs!Elapsed_hrs
End If

Cleanup:
    Set rs = Nothing
    Exit Function
   
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear

End Function

Public Function OkayToStartLabor(ByVal EmplNo As Long, ByVal sOrderNo As String, ByVal iSeqNo As Integer, ByVal Actdate As Date) As Boolean
Dim rs As ADODB.Recordset
Dim strfunctionname As String
Dim sql As String
Dim iResp As Integer
Dim lngActQty, lngScrapQty As Long
Dim dActHrs  As Currency
Dim iPrevSeqNo As Integer
OkayToStartLabor = False

On Error GoTo ErrorHandler
strfunctionname = "OkayToStartLabor"
OkayToStartLabor = False

If gCompanyCd <> 77 Or (iSeqNo = 0 Or iSeqNo = 990) Then
        OkayToStartLabor = True
        Exit Function
End If

'Return from this sp is actual quantity, actual hours, and scrap quantity of operation
'that precedes the operation being started. If nothing reported on previous operation,
'assumption is that, at best, the previous operation is in process, or is not being worked
'on at all. Therefore, the operation cannot be started.
sql = "exec SF_LaborAlreadyReported " & gCompanyCd
sql = sql & ", '" & sOrderNo & "', " & iSeqNo

Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)
If rs.EOF And rs.BOF Then
Else
'            rs.MoveLast
'            rs.MoveFirst
            lngActQty = CLng(rs(0))
            dActHrs = Val(rs(1))
            lngScrapQty = CLng(rs(2))
            iPrevSeqNo = CInt(rs(3))
            If lngActQty > 0 Or dActHrs > 0 Or lngScrapQty > 0 Or iPrevSeqNo = -1 Then
                 OkayToStartLabor = True
            End If
End If
Cleanup:
    Set rs = Nothing
    Exit Function
   
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear

End Function

Public Function CloseIndirect(ByVal EmplNo As Long, ByVal Actdate As Date) As Boolean
Dim sql As String
Dim iResp As Integer

sql = "Declare @status int Exec @status = SF_CloseIndirect " & EmplNo & ", '" & spDateParam(Actdate, 0) & _
           "', '" & spDateParam(gSysDateTime, 1) & "' Select @status"

iResp = DB.ExecuteQuery(sql)

End Function
Public Sub Ins_OpenLabor_Rec()
Dim sql As String
Dim iResp As Integer

sql = "Declare @status int Exec @status =  SF_InsertOpenLabor "
sql = sql & gCompanyCd & ", "
sql = sql & gOpenLabor.EmplNo & ", "
sql = sql & "'" & spDateParam(gOpenLabor.Actdate, 0) & "', "
sql = sql & "'" & gOpenLabor.OrderNo & "', "
sql = sql & gOpenLabor.SeqNumb & ", "
sql = sql & "'" & gOpenLabor.IndirCd & "', "
sql = sql & "'" & gOpenLabor.WCID & "', "
sql = sql & "'" & gOpenLabor.OperCd & "', "
sql = sql & "'" & gOpenLabor.CCID & "', "
sql = sql & "'" & gOpenLabor.MachRec & "', "
sql = sql & gOpenLabor.ShiftCd & ", "
sql = sql & gLabor.IncenCd & ", "
sql = sql & gLabor.MMachCd & ", "
sql = sql & "'" & gLabor.UnplnReas & "', "
sql = sql & "'" & spDateParam(gSysDateTime, 1) & "'"
sql = sql & " Select @status"
iResp = DB.ExecuteQuery(sql)

End Sub
Public Sub Ins_Labor_Rec()
Dim sql As String
Dim iResp As Integer
Dim x As Integer
 
' Insert labor record, conditionally insert carryover record,
' conditionally update job oper, update employee master.
'
sql = "declare @status int Exec @status = SF_InsertLabor "
sql = sql & gCompanyCd & ", "
sql = sql & gLabor.EmplNo & ", "
sql = sql & "'" & spDateParam(gLabor.Actdate, 0) & "', "
sql = sql & "'" & gLabor.OrderNo & "', "
sql = sql & gLabor.SeqNumb & ", "
sql = sql & "'" & gLabor.IndirCd & "', "
sql = sql & "'" & gLabor.WCID & "', "
sql = sql & "'" & gLabor.Oper & "', "
sql = sql & Str(gLabor.SetStd) & ", "
sql = sql & Str(gLabor.MchStd) & ", "
sql = sql & "'" & gLabor.EorSStd & "', "
sql = sql & gLabor.ActQty & ", "
sql = sql & Str(gLabor.ActHrs) & ", "
sql = sql & gLabor.ScrapQty1 & ", "
sql = sql & "'" & gLabor.ScrapCd1 & "', "
sql = sql & gLabor.ScrapQty2 & ", "
sql = sql & "'" & gLabor.ScrapCd2 & "', "
sql = sql & gLabor.ScrapQty3 & ", "
sql = sql & "'" & gLabor.ScrapCd3 & "', "
sql = sql & gLabor.ScrapQty4 & ", "
sql = sql & "'" & gLabor.ScrapCd4 & "', "
sql = sql & gLabor.ScrapQty5 & ", "
sql = sql & "'" & gLabor.ScrapCd5 & "', "
sql = sql & gLabor.ScrapQty6 & ", "
sql = sql & "'" & gLabor.ScrapCd6 & "', "
sql = sql & gLabor.ScrapQty7 & ", "
sql = sql & "'" & gLabor.ScrapCd7 & "', "
sql = sql & gLabor.ScrapQty8 & ", "
sql = sql & "'" & gLabor.ScrapCd8 & "', "
sql = sql & "'" & gLabor.CCID & "', "
sql = sql & "'" & gLabor.MachRec & "', "
sql = sql & Str(gLabor.PayRate) & ", "
sql = sql & Str(gLabor.Earnings) & ", "
sql = sql & Str(gLabor.EmplEff) & ", "
sql = sql & Str(gLabor.EarnHrs) & ", "
sql = sql & "'" & gLabor.EarnType & "', "
sql = sql & gLabor.ShiftNo & ", "
sql = sql & "'" & gLabor.OperStat & "', "
sql = sql & gLabor.IncenCd & ", "
sql = sql & "'" & gLabor.RecType & "', "
sql = sql & gLabor.MMachCd & ", "
sql = sql & Str(gLabor.ActHrsUnrnd) & ", "
sql = sql & Str(gLabor.MachHrs) & ", "
sql = sql & "'" & gLabor.ErrorLst & "', "
If Len(RTrim(gLabor.ErrorLst)) > 0 Then
     sql = sql & "'Y', "     'error flag
Else
    sql = sql & "' ', "
End If
sql = sql & "'" & gLabor.UnplnReas & "', "
sql = sql & "'" & spDateParam(gSysDateTime, 1) & "', "
sql = sql & Str(gLabor.AddonHrs) & ", "
sql = sql & "'  ', "     'Addon reason code is unused
sql = sql & gLabor.NumbLoads & ", "
sql = sql & gLabor.CycleTime & ", "
sql = sql & gLabor.batchMeth
If div = saline Then
    sql = sql & ", " & gLabor.box_nbr_1
    sql = sql & ", " & gLabor.box_nbr_2
    sql = sql & ", " & gLabor.box_nbr_3
    sql = sql & ", " & gLabor.box_nbr_4
    sql = sql & ", " & gLabor.box_nbr_5
    sql = sql & ", " & gLabor.box_nbr_6
    sql = sql & ", " & gLabor.box_nbr_7
    sql = sql & ", " & gLabor.box_nbr_8
End If
sql = sql & " select @status"
iResp = DB.ExecuteQuery(sql)

End Sub
Public Sub Close_Direct_Labor()
Dim AttdUpdFlag As Boolean
Dim sql As String
Dim iResp As Integer
Dim x As Integer
Dim aryNumb As Integer
If First_Labor_Rec(gEmployee.EmplNo, gSysDate) = True Then
   AttdUpdFlag = True
End If
For aryNumb = 1 To LoopCnt
   Init_labor_rec
   gLabor.companyCd = gCompanyCd
   gLabor.EmplNo = gEmployee.EmplNo
   gLabor.Actdate = gEmployee.Actdate
   gLabor.CCID = gLaborIn(aryNumb).CCID
   gLabor.TranDate = gSysDateTime
   gLabor.ShiftNo = gEmployee.GenShift
   gLabor.ErrorLst = gLaborIn(aryNumb).ErrorLst
   gLabor.RecType = gLaborIn(aryNumb).RecType
   If AttdUpdFlag Then
       gAttendance.CCID = gLabor.CCID
   End If
   gLabor.WCID = gLaborIn(aryNumb).WCID
   gLabor.OrderNo = gLaborIn(aryNumb).OrderNo
   gLabor.SeqNumb = gLaborIn(aryNumb).SeqNo
   gLabor.SetStd = gLaborIn(aryNumb).SetStd
   gLabor.MchStd = gLaborIn(aryNumb).MchStd
   gLabor.EorSStd = gLaborIn(aryNumb).EOrS
   gLabor.Oper = gLaborIn(aryNumb).Oper
   If gLaborIn(aryNumb).Scrap(1).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
      gLabor.ScrapQty1 = ConvScrapType(gLaborIn(aryNumb).Scrap(1).Qty, _
                          gLaborIn(aryNumb).MatlWgt)
   Else
      gLabor.ScrapQty1 = gLaborIn(aryNumb).Scrap(1).Qty
   End If
       gLabor.ScrapCd1 = gLaborIn(aryNumb).Scrap(1).Code
       If gLaborIn(aryNumb).Scrap(2).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty2 = ConvScrapType(gLaborIn(aryNumb).Scrap(2).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty2 = gLaborIn(aryNumb).Scrap(2).Qty
       End If
       gLabor.ScrapCd2 = gLaborIn(aryNumb).Scrap(2).Code
       If gLaborIn(aryNumb).Scrap(3).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty3 = ConvScrapType(gLaborIn(aryNumb).Scrap(3).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty3 = gLaborIn(aryNumb).Scrap(3).Qty
       End If
       gLabor.ScrapCd3 = gLaborIn(aryNumb).Scrap(3).Code
       If gLaborIn(aryNumb).Scrap(4).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty4 = ConvScrapType(gLaborIn(aryNumb).Scrap(4).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty4 = gLaborIn(aryNumb).Scrap(4).Qty
       End If
       gLabor.ScrapCd4 = gLaborIn(aryNumb).Scrap(4).Code
       If gLaborIn(aryNumb).Scrap(5).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty5 = ConvScrapType(gLaborIn(aryNumb).Scrap(5).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty5 = gLaborIn(aryNumb).Scrap(5).Qty
       End If
       gLabor.ScrapCd5 = gLaborIn(aryNumb).Scrap(5).Code
       If gLaborIn(aryNumb).Scrap(6).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty6 = ConvScrapType(gLaborIn(aryNumb).Scrap(6).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty6 = gLaborIn(aryNumb).Scrap(6).Qty
       End If
       gLabor.ScrapCd6 = gLaborIn(aryNumb).Scrap(6).Code
       If gLaborIn(aryNumb).Scrap(7).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty7 = ConvScrapType(gLaborIn(aryNumb).Scrap(7).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty7 = gLaborIn(aryNumb).Scrap(7).Qty
       End If
       gLabor.ScrapCd7 = gLaborIn(aryNumb).Scrap(7).Code
       If gLaborIn(aryNumb).Scrap(8).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty8 = ConvScrapType(gLaborIn(aryNumb).Scrap(8).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty8 = gLaborIn(aryNumb).Scrap(8).Qty
       End If
       gLabor.ScrapCd8 = gLaborIn(aryNumb).Scrap(8).Code
       gLabor.ActQty = gLaborIn(aryNumb).ActQty
       gLabor.OperStat = gLaborIn(aryNumb).OperStatus
       gLabor.ActHrsUnrnd = 0
       gLabor.ActHrs = gLaborIn(aryNumb).setupHours
       gLabor.MachHrs = gLaborIn(aryNumb).machHours
       gLabor.UnplnReas = gLaborIn(aryNumb).UnplnReas
       gLabor.MachRec = gLaborIn(aryNumb).MachRec
       gLabor.MMachCd = gLaborIn(aryNumb).NumbMachs
       gLabor.IncenCd = gLaborIn(aryNumb).incenRec
       gLabor.AddonHrs = gLaborIn(aryNumb).addonHours
       gLabor.NumbLoads = gLaborIn(aryNumb).NoOfLoads
       gLabor.CycleTime = gLaborIn(aryNumb).CycleTime
       gLabor.batchMeth = gLaborIn(aryNumb).biMethod

 
' Insert labor record, conditionally insert carryover record,
' conditionally update job oper, update employee master.
'
sql = "declare @status int Exec @status = SF_CloseDirect "
sql = sql & gCompanyCd & ", "
sql = sql & gLabor.EmplNo & ", "
sql = sql & "'" & spDateParam(gLabor.Actdate, 0) & "', "
sql = sql & "'" & gLabor.OrderNo & "', "
sql = sql & gLabor.SeqNumb & ", "
sql = sql & "'" & gLabor.IndirCd & "', "
sql = sql & "'" & gLabor.WCID & "', "
sql = sql & "'" & gLabor.Oper & "', "
sql = sql & Str(gLabor.SetStd) & ", "
sql = sql & Str(gLabor.MchStd) & ", "
sql = sql & "'" & gLabor.EorSStd & "', "
sql = sql & gLabor.ActQty & ", "
sql = sql & gLabor.ScrapQty1 & ", "
sql = sql & "'" & gLabor.ScrapCd1 & "', "
sql = sql & gLabor.ScrapQty2 & ", "
sql = sql & "'" & gLabor.ScrapCd2 & "', "
sql = sql & gLabor.ScrapQty3 & ", "
sql = sql & "'" & gLabor.ScrapCd3 & "', "
sql = sql & gLabor.ScrapQty4 & ", "
sql = sql & "'" & gLabor.ScrapCd4 & "', "
sql = sql & gLabor.ScrapQty5 & ", "
sql = sql & "'" & gLabor.ScrapCd5 & "', "
sql = sql & gLabor.ScrapQty6 & ", "
sql = sql & "'" & gLabor.ScrapCd6 & "', "
sql = sql & gLabor.ScrapQty7 & ", "
sql = sql & "'" & gLabor.ScrapCd7 & "', "
sql = sql & gLabor.ScrapQty8 & ", "
sql = sql & "'" & gLabor.ScrapCd8 & "', "
sql = sql & "'" & gLabor.CCID & "', "
sql = sql & "'" & gLabor.MachRec & "', "
sql = sql & gLabor.ShiftNo & ", "
sql = sql & "'" & gLabor.OperStat & "', "
sql = sql & gLabor.IncenCd & ", "
sql = sql & "'" & gLabor.RecType & "', "
sql = sql & gLabor.MMachCd & ", "
sql = sql & "'" & gLabor.ErrorLst & "', "
If Len(RTrim(gLabor.ErrorLst)) > 0 Then
     sql = sql & "'Y', "     'error flag
Else
    sql = sql & "' ', "
End If
sql = sql & "'" & gLabor.UnplnReas & "', "
sql = sql & "'" & spDateParam(gSysDateTime, 1) & "', "
sql = sql & Str(gLabor.AddonHrs) & ", "
sql = sql & "'  ', "     'Addon reason code is unused
sql = sql & gLabor.NumbLoads & ", "
sql = sql & gLabor.CycleTime & ", "
sql = sql & gLabor.batchMeth & ", "
sql = sql & Str(gLabor.ActHrs)
sql = sql & " select @Status"


 iResp = DB.ExecuteQuery(sql)
Next aryNumb

End Sub
Public Sub Ins_OSSendOut()
Dim sql As String
Dim iResp As Integer

sql = "declare @status int Exec @status = SF_InsOSSendOut "
sql = sql & gCompanyCd & ", "
sql = sql & gLabor.EmplNo & ", "
'sql = sql & "'" & spDateparam(gLabor.ActDate,0) & "', "
sql = sql & "'" & gLabor.OrderNo & "', "
sql = sql & gLabor.SeqNumb & ", "
sql = sql & gLabor.ActQty
sql = sql & " Select @status"

iResp = DB.ExecuteQuery(sql)

End Sub

Public Sub Ins_Assembly_Labor()
Dim sql As String
Dim iResp As Integer

sql = "Declare @status int Exec @Status = SF_InsAssemblyLabor "
sql = sql & gCompanyCd & ", "
sql = sql & gLabor.EmplNo & ", "
'sql = sql & "'" & spDateparam(gLabor.ActDate,0) & "', "
sql = sql & "'" & gLabor.OrderNo & "', "
sql = sql & gLabor.SeqNumb & ", "
sql = sql & gLabor.ActQty
sql = sql & " Select @status"

 iResp = DB.ExecuteQuery(sql)

End Sub
Public Sub Ins_Std_Labor()
Dim sql As String
Dim iResp As Integer

sql = "declare @status int Exec @Status = SF_InsStdLabor "
sql = sql & gCompanyCd & ", "
sql = sql & gLabor.EmplNo & ", "
sql = sql & "'" & spDateParam(gLabor.Actdate, 0) & "', "
sql = sql & "'" & gLabor.OrderNo & "', "
sql = sql & gLabor.SeqNumb & ", "
sql = sql & gLabor.ActQty & ", "
sql = sql & "'" & gLabor.OperStat & "', "
sql = sql & gLabor.ScrapQty1 & ", "
sql = sql & "'" & gLabor.ScrapCd1 & "', "
sql = sql & gLabor.ScrapQty2 & ", "
sql = sql & "'" & gLabor.ScrapCd2 & "', "
sql = sql & gLabor.ScrapQty3 & ", "
sql = sql & "'" & gLabor.ScrapCd3 & "', "
sql = sql & gLabor.ScrapQty4 & ", "
sql = sql & "'" & gLabor.ScrapCd4 & "', "
sql = sql & gLabor.ScrapQty5 & ", "
sql = sql & "'" & gLabor.ScrapCd5 & "', "
sql = sql & gLabor.ScrapQty6 & ", "
sql = sql & "'" & gLabor.ScrapCd6 & "', "
sql = sql & gLabor.ScrapQty7 & ", "
sql = sql & "'" & gLabor.ScrapCd7 & "', "
sql = sql & gLabor.ScrapQty8 & ", "
sql = sql & "'" & gLabor.ScrapCd8 & "' "
sql = sql & " select @status"

iResp = DB.ExecuteQuery(sql)

End Sub
Public Function ValidIndirCd(ByVal IndirCd As String) As Boolean
    Dim sParamKey As String
    
    On Error GoTo ValidIndirCdErr
    
    IndirCd = RTrim(IndirCd)
    If Len(IndirCd) > 2 Or Len(IndirCd) = 0 Then
        ValidIndirCd = False
    Else
        'sParamKey = "SCRAP_" & scrapCd
        'If GetDivParam(sParamKey) = True Then
        '    ValidScrapCd = True
        'Else
        '    ValidScrapCd = False
        'End If
        If Left(gcolIndirCodes.getDesc(IndirCd), 3) <> "***" Then
            ValidIndirCd = True
        Else
            ValidIndirCd = False
        End If
    End If
    Exit Function
ValidIndirCdErr:
    ValidIndirCd = False
End Function


Public Function ValidScrapCd(ByVal scrapCd As String) As Boolean
    Dim sParamKey As String
    
    On Error GoTo ValidScrapCdErr
    
    scrapCd = RTrim(scrapCd)
    If Len(scrapCd) > 2 Or Len(scrapCd) = 0 Then
        ValidScrapCd = False
    Else
  '    sParamKey = "SCRAP_" & scrapCd
  '      If GetDivParam(sParamKey) = True Then
  '          ValidScrapCd = True
  '      Else
  '          ValidScrapCd = False
  '      End If
        If Left(gcolScrapCodes.getDesc(scrapCd), 3) <> "***" Then
            ValidScrapCd = True
        Else
            ValidScrapCd = False
        End If
    End If
    Exit Function
ValidScrapCdErr:
    ValidScrapCd = False
End Function

Function spDateParam(ByVal dtmDateIn, ByVal Dateonly As Integer) As String
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : spDateParam
'Description      : Routine to properly format the entered date into an
'                   acceptable format for SQL Server Stored Procedure
'                   parameters (ccyy-mm-dd hh:mm:ss.ddd).
'Parameters       : dtmDateIn | date | date to be formatted
'                              dateonly  | integer | when this parameter = 0, specifies that dtmdatein
'                                                                  contains only date, and when = 1, dtmdatein is a
'                                                                  date and time
'
'Called From      : each module that passes a date to SQL Server
'Procedures Used  :
'Return Parameters: spDateParam | String | desc
'Tables Updates   : none
'
'==============================================================================
    spDateParam = _
       IIf(Len(CStr(Month(dtmDateIn))) = 1, "0", "") _
      & CStr(Month(dtmDateIn)) & "/" & _
       IIf(Len(CStr(Day(dtmDateIn))) = 1, "0", "") _
      & CStr(Day(dtmDateIn)) & "/" _
      & CStr(Year(dtmDateIn))
If Dateonly = 0 Then Exit Function
    spDateParam = spDateParam & " " & _
        IIf(Len(CStr(Hour(dtmDateIn))) = 1, "0", "") _
        & CStr(Hour(dtmDateIn)) & ":" & _
        IIf(Len(CStr(Minute(dtmDateIn))) = 1, "0", "") _
        & CStr(Minute(dtmDateIn)) & ":" & _
        IIf(Len(CStr(Second(dtmDateIn))) = 1, "0", "") _
        & CStr(Second(dtmDateIn))
      
End Function
Public Sub Ins_Std_chg()
Dim sql As String
Dim iResp As Integer

sql = "declare @status int Exec @Status = SF_InsertStdChgs "
sql = sql & gCompanyCd & ", "
sql = sql & gStdChg.EmplNo & ", "
sql = sql & "'" & gStdChg.OrderNo & "', "
sql = sql & gStdChg.SeqNumb & ", "
sql = sql & Str(gStdChg.OrigFixStd) & ", "
sql = sql & Str(gStdChg.RevFixedStd) & ", "
sql = sql & Str(gStdChg.origMchStd) & ", "
sql = sql & Str(gStdChg.RevMchStd) & ", "
sql = sql & "'" & QuoteSet(gStdChg.ReasonText) & "', "
sql = sql & gStdChg.ShiftCd & ", "
sql = sql & "'" & spDateParam(gStdChg.Actdate, 0) & "', "
sql = sql & "'" & spDateParam(gStdChg.transdate, 1) & "', "
sql = sql & "' ', null"    'last two fields are process flag and process date
sql = sql & " select @status"
iResp = DB.ExecuteQuery(sql)

End Sub
Public Sub InsEmplCommAction()
Dim sql As String
Dim SQLAnswer As String
Dim iResp As Integer
Select Case gMessage.msg_answer
   Case "A"
     If gMessage.msg_type = 1 Then 'message to an individual
         SQLAnswer = "A" 'answer is accept
     ElseIf gMessage.msg_type = 2 Then 'message to a home Dept
         SQLAnswer = "R" 'answer is acknowledge
     Else  'default to all employees
         SQLAnswer = "R" 'answer is acknowledge
     End If
  Case "D"
     If gMessage.msg_type = 1 Then 'message to an individual
         SQLAnswer = "D" 'answer is dispute
     End If
End Select
sql = "declare @status int Exec @Status = SF_InsertEmplCommAction "
sql = sql & gCompanyCd & ", " & gEmployee.EmplNo & ", "
sql = sql & gMessage.message_id & ", " & SQLAnswer
sql = sql & " select @status"
iResp = DB.ExecuteQuery(sql)

End Sub

Public Function UpdateAcceptedOT(ByVal lemplno As Long, ByVal lactdate As Date, ByVal lAccHrs As Currency)
Dim sql As String
Dim iResp As Integer

sql = "Declare @status int Exec @status = SF_updateAcceptedOT "
sql = sql & lemplno & ", '" & spDateParam(lactdate, 0) & "', " & Str(lAccHrs)
sql = sql & " Select @status"
iResp = DB.ExecuteQuery(sql)

End Function

Public Function GetServerTime() As String
    
'    Shell gSrvrTimeStr
    GetServerTime = ""

End Function

Public Function GetRecordsetData(strsql As String) As ADODB.Recordset

Dim Cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim strfunctionname As String

On Error GoTo ErrorHandler
strfunctionname = "GetRecordsetData"
    
' create command object
Set Cmd = New ADODB.Command
    
' set command parameters
With Cmd
    .ActiveConnection = gSQLConnection
    .CommandType = adCmdText
    .CommandText = strsql
End With
    
' create recordset object
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
    
' get recordset
Set rs = Cmd.Execute

Set GetRecordsetData = rs
Set rs.ActiveConnection = Nothing
    
Cleanup:
'    If gSQLConnection.State = adStateOpen Then
'        gSQLConnection.Close
'    End If
'    If Not (gSQLConnection Is Nothing) Then
'        Set gSQLConnection = Nothing
'    End If
    Set rs = Nothing
    Set Cmd = Nothing
    Exit Function
   
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
'    gError.ShowMsg (True)
    Resume Next

End Function

Public Function ExecuteQuery(strsql As String) As Integer
Dim Cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim Param As ADODB.Parameter
Dim strfunctionname As String
    
On Error GoTo ErrorHandler
strfunctionname = "ExecuteQuery"
ExecuteQuery = -1
Set Cmd = New ADODB.Command
    
With Cmd
  .ActiveConnection = gSQLConnection
  .CommandType = adCmdText
  .CommandText = strsql
End With
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
Set rs = Cmd.Execute
ExecuteQuery = rs(0)

Cleanup:
'    If gSQLConnection.State = adStateOpen Then
'        gSQLConnection.Close
'    End If
'
'    If Not (gSQLConnection Is Nothing) Then
'        Set gSQLConnection = Nothing
'    End If
    Set rs = Nothing
    Set Cmd = Nothing
    Exit Function
   
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
'    gError.ShowMsg (True)
    Resume Next
End Function

Public Function LoadLookupCols(ByVal sParamKey As String, ByVal iCodeSize As Integer, _
        ByRef aCol As clsCodeTable) As Boolean
'==============================================================================
'Author           : David A. Taylor
'Date             : June, 1998
'Procedure Name   : Startup.LoadLookupCols
'Description      : Populate the clsCodeTable Collection passed in with the
'                   appropriate Div Param table Data
'Parameters       : None
'Objects Used     : frmDCMain Error Labels
'Forms Loaded     : None
'Return Parameters: String with User ID retrieved, or Null string
'Tables Updates   : None
'==============================================================================
Dim rs As ADODB.Recordset
Dim sql As String
Dim strfunctionname As String
Dim iResp As Integer
Dim retnStr As String

On Error GoTo ErrorHandler
strfunctionname = "LoadLookupCols"
    
LoadLookupCols = False
    
sql = "exec SF_GetDivParam '" & sParamKey & "'"

Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)

If rs.BOF And rs.EOF Then
Else
  '    rs.MoveLast
  '    rs.MoveFirst
      While Not rs.EOF
          retnStr = rs(1)
         aCol.Add Left(retnStr, iCodeSize), Mid(retnStr, iCodeSize + 1, 20)
         rs.MoveNext
      Wend
      LoadLookupCols = True
End If
Cleanup:
    Set rs = Nothing
    Exit Function
   
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear
    Resume Next
End Function
Public Function LoadOperCol(aCol As clsCodeTable)
Dim rs As ADODB.Recordset
Dim strfunctionname As String
Dim sql As String
Dim RetnOper As Integer
Dim RetnDesc As String
'Dim CodeEntry As New clsCodeItem

On Error GoTo ErrorHandler
strfunctionname = "LoadOperCol"

LoadOperCol = False
    
sql = "Exec SF_GetOperCodes"

 Set rs = New ADODB.Recordset
 Set rs = DB.GetRecordsetData(sql)
  
  If rs.BOF And rs.EOF Then
  Else
'       rs.MoveLast
'       rs.MoveFirst
       While Not rs.EOF
              RetnOper = CInt(rs(0))
              RetnDesc = rs(1)
              aCol.Add Format(RetnOper, "000"), RetnDesc
              rs.MoveNext
       Wend
       LoadOperCol = True
  End If
Cleanup:
    Set rs = Nothing
    Exit Function
   
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear
    Resume Next
 
End Function

Public Function CheckValidQty(ByVal sOrderNo As String, ByVal iSeqNo As Integer, ByVal iInpQty As Long, ByVal sChkType As String) As Boolean
Dim rs As ADODB.Recordset
Dim sql As String
Dim iResp As Integer
Dim iNumRecs As Integer
Dim iRetnSeq As Integer
Dim iRetnQty As Long
Dim iOvrnQty As Long
Dim x As Integer
Dim strfunctionname As String

On Error GoTo ErrorHandler
strfunctionname = "CheckValidQty"
CheckValidQty = False

If gCompanyCd <> 77 Or _
   (iSeqNo = 0 Or iSeqNo = 990) Then
        CheckValidQty = True
        Exit Function
End If

    
sql = "exec SF_GetLastOperWorked " & gCompanyCd
sql = sql & ", '" & sOrderNo & "', " & iSeqNo

Set rs = New ADODB.Recordset
Set rs = DB.GetRecordsetData(sql)
If rs.EOF And rs.BOF Then
    Exit Function
Else
'          rs.MoveLast
'          rs.MoveFirst
          iRetnSeq = CInt(rs(0))
          iRetnQty = CLng(rs(1))
          iOvrnQty = CLng(rs(2))
End If
If sChkType = "S" Then     'operation being started has a net available quantity (prior oper actual minus scrap)
                                              ' minus current oper actual qty. If there is no net available qty, then
                                              ' operation has been completed
    If iRetnQty > 0 Then
        CheckValidQty = True
    End If
End If

If sChkType = "E" Then  'reported quantity for labor being finished must be less than or equal to net available qty
'The following code was disabled on 1/24/2001 after conversation with Sweden. Quantity restriction
' on first operation is too hard to manage. However, they asked if I could keep it available for future use.
' This change also affects sp_GetLastOperWorked stored procedure
'  If iRetnSeq = -1 Then 'first sequence reporting
'       If (iInpQty < (iRetnQty - iOvrnQty)) Then
'            CheckValidQty = True
'            gOKToComplete = False
'    ElseIf (iInpQty <= iRetnQty) Then
'          gOKToComplete = True
'          CheckValidQty = True
'    Else
'          gOKToComplete = False
'          CheckValidQty = False
'    End If
' ******************disabled code end  replacement follows below********************
  If iRetnSeq = -1 Then       'first sequence reporting
     If iInpQty < iRetnQty Then
            CheckValidQty = True
            gOKToComplete = False
    Else
          gOKToComplete = True
          CheckValidQty = True
    End If     'end of replacement code as noted above
  Else
      If (iInpQty < iRetnQty) Then
            CheckValidQty = True
            gOKToComplete = False
    ElseIf (iInpQty = iRetnQty) Then
          gOKToComplete = True
          CheckValidQty = True
    Else
          gOKToComplete = False
          CheckValidQty = False
    End If
  End If
End If

Cleanup:
    Set rs = Nothing
    Exit Function
   
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear
    Resume Next
End Function

Public Function LoadAllBatchIncnCols(ByRef aCol As clsAllBITable) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : DBCommon.LoadAllBatchIncnCols
'Description      : Populate the Full Batch Incentive code collection
'                   with the entries for all Work Centers.
'Parameters       : Collection Class clsAllBITable to hold all Batch Incentives
'                   (gcolAllBICodes).
'Procedures Used  : VBSQL
'Objects Used     : gcolAllBICodes(clsAllBITable)
'Forms Loaded     : None
'Return Parameters: Boolean True, unless error.
'Tables Updates   : None
'==============================================================================
    Dim rs As ADODB.Recordset
    Dim strfunctionname As String
    Dim sql As String
    Dim iResp As Integer
    Dim strWCID As String
    Dim siMethod As Byte
    Dim strDesc As String
    Dim siIncMin As Byte
    
    On Error GoTo ErrorHandler
    strfunctionname = "LoadAllBatchIncnCols"
    LoadAllBatchIncnCols = False
        
    sql = "Exec SF_GetbatchIncen " & gCompanyCd
    
    Set rs = New ADODB.Recordset
    Set rs = DB.GetRecordsetData(sql)
    If rs.EOF And rs.BOF Then
    Else
'            rs.MoveLast
'            rs.MoveFirst
            While Not rs.EOF
                  strWCID = rs!WC_ID
                  siMethod = rs!batch_method
                  strDesc = rs!method_desc
                  siIncMin = rs!inc_min
                  aCol.AddBIItem strWCID, siMethod, strDesc, siIncMin
                  rs.MoveNext
            Wend
            LoadAllBatchIncnCols = True
    End If

Cleanup:
    Set rs = Nothing
    Exit Function
   
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear
    Resume Next

End Function
Public Function LoadAllIndirIncnCols(ByRef aCol As clsAllIITable) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : DBCommon.LoadAllIndirIncnCols
'Description      : Populate the Full Indirect Incentive code collection
'                   with the entries for all Cost Centers.
'Parameters       : Collection Class clsAllIITable to hold all Indirect
'                   Incentives (gcolAllIICodes).
'Objects Used     : gcolAllIICodes(clsAllIITable)
'Forms Loaded     : None
'Return Parameters: Boolean True, unless error.
'Tables Updates   : None
'==============================================================================
    Dim rs As ADODB.Recordset
    Dim strfunctionname As String
    Dim sql As String
    Dim iResp As Integer
    Dim strCCID As String
    Dim iMachNumb As Integer
    Dim iIncenType As Integer
    Dim cWrkStd As Currency
    Dim strDesc As String
    
    On Error GoTo ErrorHandler
    strfunctionname = "LoadAllIndirIncnCols"
    LoadAllIndirIncnCols = False
    
    sql = "Exec SF_GetIndirIncen " & gCompanyCd
    
    Set rs = New ADODB.Recordset
    Set rs = DB.GetRecordsetData(sql)
    If rs.EOF And rs.BOF Then
    Else
'            rs.MoveLast
'            rs.MoveFirst
            While Not rs.EOF
                strCCID = rs!CC_ID
                iIncenType = rs!incen_type
                iMachNumb = rs!mach_numb
                cWrkStd = rs!wrk_std
                strDesc = rs!Descr
                aCol.AddIIItem strCCID, iIncenType, iMachNumb, cWrkStd, strDesc
                rs.MoveNext
            Wend
           LoadAllIndirIncnCols = True
    End If

Cleanup:
    Set rs = Nothing
    Exit Function
   
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = strfunctionname
    gError.ErrorRoutine
    gError.Clear
    Resume Next

End Function

Public Function ConnectToDB() As Long
    Dim strconnect As String
    
 On Error GoTo ErrorHandler
   ClearMainError
   
   ConnectToDB = vbError
       strconnect = "PROVIDER=SQLOleDB.1;Persist Security Info=False;" & _
                                "User ID=" & SQL_LOGIN_ID & ";Password=" & SQL_LOGIN_PSWD & _
                              ";Initial Catalog=" & gSQLDatabase & ";Data Source=" & gSQLServer
                              
      
     Set gSQLConnection = New ADODB.Connection
     With gSQLConnection
         .ConnectionString = strconnect
         .CursorLocation = adUseClient
         .Open
     End With
     ConnectToDB = vbOK
     Exit Function
ErrorHandler:
     ConnectToDB = vbError
'    get error info
'    gError.ErrNum = Err.Number
'    gError.ErrDesc = Err.Description
'    gError.ModuleName = MODULE_NAME
'    gError.FunctionName = "ConnectToDB"
'    gError.ErrorRoutine
'    gError.Clear
'
End Function

Public Sub DisconnectDB()
If Not (gSQLConnection Is Nothing) Then
    If gSQLConnection.State = adStateOpen Then
          gSQLConnection.Close
    End If
    Set gSQLConnection = Nothing
    gbConnected = False
End If

End Sub
