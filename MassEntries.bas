Attribute VB_Name = "MassEntries"
Option Explicit

Public Sub ClearCollection(entries As Collection)
    Set entries = Nothing
    Set entries = New Collection
End Sub

Public Sub BuildCollection(records As ADODB.Recordset, ByRef entries As Collection)
On Error Resume Next
    
    Dim i As Integer
    
    With records '//Use the recordset
        '//Loop through the recordset and build a collection of entries
        For i = 0 To .RecordCount - 1
            
            '//Instantiate a new labor record to be added to the collection
            Dim LbrRec As New MassEntry
            
            '//Give the object data from the grid row
            LbrRec.Actdate = IIf(IsDate(Format$(.Fields("Date").Value, gDateformat)), CDate(.Fields("Date").Value), Now)
            LbrRec.EmplNo = .Fields("Emp No").Value
            LbrRec.OrderNo = .Fields("Job No").Value
            LbrRec.SeqNumb = .Fields("Seq No").Value
            LbrRec.MachHrs = .Fields("Machine Hrs").Value
            LbrRec.ActQty = .Fields("Quantity").Value
            LbrRec.UnplnReas = IIf(IsNull(.Fields("Unplan Reas").Value), "", .Fields("Unplan Reas").Value)
            LbrRec.ScrapQty = IIf(IsNull(.Fields("Scrap Quantity").Value), "", .Fields("Scrap Quantity").Value)
            LbrRec.ScrapCd = IIf(IsNull(.Fields("Scrap Code").Value), "", .Fields("Scrap Code").Value)
            
            '//Add the labor record to the collection
            entries.Add LbrRec
            
            '//Destroy the labor record object to release memory
            Set LbrRec = Nothing
            .MoveNext
        Next i
    End With
End Sub

Public Function ValidateEntry(ByRef Entry As MassEntry) As Boolean
On Error GoTo Cleanup
    
    Dim cmd As New ADODB.Command
    
    If gbConnected Then
        
        If gSQLConnection.State = adStateOpen Then
            '//Build SQL query
            With cmd
                .ActiveConnection = gSQLConnection
                .CommandType = adCmdStoredProc
                .CommandText = "ValidateMassEntryRecord"
                .Parameters.Append .CreateParameter("@EmplNo", adInteger, adParamInput, , Entry.EmplNo)
                .Parameters.Append .CreateParameter("@OrderNo", adVarChar, adParamInput, 8, Entry.OrderNo)
                .Parameters.Append .CreateParameter("@SeqNumb", adInteger, adParamInput, , Entry.SeqNumb)
                .Parameters.Append .CreateParameter("@EmplNoIsValid", adBoolean, adParamOutput)
                .Parameters.Append .CreateParameter("@EmplName", adVarChar, adParamOutput, 20)
                .Parameters.Append .CreateParameter("@OrderNoIsValid", adBoolean, adParamOutput)
                .Parameters.Append .CreateParameter("@OrderQty", adInteger, adParamOutput)
                .Parameters.Append .CreateParameter("@SeqNumbIsValid", adBoolean, adParamOutput)
                .Parameters.Append .CreateParameter("@MachineStd", adNumeric, adParamOutput)
                .Parameters("@MachineStd").Precision = 8
                .Parameters("@MachineStd").NumericScale = 1
                .Parameters.Append .CreateParameter("@CCID", adVarChar, adParamOutput, 5)
                .Parameters.Append .CreateParameter("@WCID", adVarChar, adParamOutput, 5)
                .Parameters.Append .CreateParameter("@SetStd", adNumeric, adParamOutput)
                .Parameters("@SetStd").Precision = 6
                .Parameters("@SetStd").NumericScale = 1
                .Parameters.Append .CreateParameter("@ErrorMsg", adVarChar, adParamOutput, 150)
                .Execute
            End With
            
            With Entry
                .EmplNoIsValid = cmd.Parameters("@EmplNoIsValid").Value
                .EmplName = cmd.Parameters("@EmplName").Value
                .OrderNoIsValid = cmd.Parameters("@OrderNoIsValid").Value
                .OrderQty = cmd.Parameters("@OrderQty").Value
                .SeqNumbIsValid = cmd.Parameters("@SeqNumbIsValid").Value
                .MchStd = cmd.Parameters("@MachineStd").Value
                .CCID = cmd.Parameters("@CCID").Value
                .WCID = cmd.Parameters("@WCID").Value
                .SetStd = cmd.Parameters("@SetStd").Value
                
                If Len(.ErrorMsg) > 0 Then
                    .ErrorMsg = .ErrorMsg & " " & cmd.Parameters("@ErrorMsg").Value
                Else
                    .ErrorMsg = cmd.Parameters("@ErrorMsg").Value
                End If
            
                'Flag invalid if anything didn't match database
                If Not .EmplNoIsValid Or Not .OrderNoIsValid Or Not .SeqNumbIsValid Then
                    GoTo Invalid
                End If
                
            End With
        
        End If

    End If
    
Valid:
    ValidateEntry = True
    GoTo Cleanup
    Exit Function
    
Invalid:
    ValidateEntry = False
    GoTo Cleanup
    Exit Function
    
Cleanup:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        MsgBox "ERROR: " & Err.Number & vbCrLf & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "SOURCE: " & Err.Source
    End If
End Function

Public Function ProcessEntry(Entry As MassEntry) As Boolean
    '//Assume the worst
    ProcessEntry = False
On Error GoTo Handler
    
'//Build global records for ErrorChk function
    '~~~~~~~~~gEmployee~~~~~~~~~
    Call GetEmployee(CLng(Entry.EmplNo))
    'gEmployee
    'Reset act date for employee to entry record (GetEmployee sets the activity date from last clock-in)
    'If this is not done, ErrorChk might update the employee's current ActDate in the Attendance table
    gEmployee.Actdate = Entry.Actdate
    
    '~~~~~~~~~gJobMaster~~~~~~~~~
    Call GetJobMaster(Entry.OrderNo)
    'gJobMaster
    
    '~~~~~~~~~gJobOpers~~~~~~~~~
    Call GetJobOper(Entry.OrderNo, CInt(Entry.SeqNumb))
    'gJobOpers
    
    '~~~~~~~~~gLaborIn[]~~~~~~~~~
    With gLaborIn(1)
    
        If Left(gJobOpers.Oper, 3) = "645" Then
            .RecType = "F"
            .machHours = 0#
            .setupHours = Entry.MachHrs
        Else
            .RecType = "D"
            .machHours = Entry.MachHrs
            .setupHours = 0#
        End If
        
        .UnplnReas = Entry.UnplnReas
        .NumbMachs = 1
        
        If Len(.UnplnReas) > 0 Then
            .SeqNo = 0
        Else
            .SeqNo = Entry.SeqNumb
        End If
        
        .WCID = Entry.WCID
        .Scrap(1).Code = Entry.ScrapCd
        .Scrap(1).Qty = CLng(IIf((Entry.ScrapQty = ""), 0, Entry.ScrapQty))
        .ActQty = CLng(Entry.ActQty)
        
    End With
    
'//Run ErrorChk against single record
    Call ErrorChk(1)
    
'//Build gLabor object to import record into DC system
    With gLabor
        .EmplNo = Entry.EmplNo
        .Actdate = Entry.Actdate
        .OrderNo = Entry.OrderNo
        .SeqNumb = CInt(Entry.SeqNumb)
        .IndirCd = ""
        .WCID = Entry.WCID
        .Oper = gJobOpers.Oper
        .SetStd = gJobOpers.SetStd
        .MchStd = gJobOpers.MchStd
        .EOrSStd = gJobOpers.EOrSStd
        .ActQty = CLng(Entry.ActQty)
        .ActHrs = 0#
        .ScrapCd1 = Entry.ScrapCd
        .ScrapQty1 = IIf((Entry.ScrapQty = ""), 0, Entry.ScrapQty)
        .CCID = Entry.CCID
        .MachRec = "N"
        .PayRate = 0#
        .Earnings = 0#
        .EmplEff = 0#
        .EarnHrs = 0#
        .EarnType = ""
        .ShiftNo = gEmployee.GenShift
        .OperStat = gJobOpers.OperStatus
        .IncenCd = 0
        .RecType = gLaborIn(1).RecType
        .MMachCd = gLaborIn(1).NumbMachs
        .ActHrsUnrnd = 0#
        .MachHrs = CCur(Entry.MachHrs)
        .ErrorLst = gLaborIn(1).ErrorLst
        .UnplnReas = Entry.UnplnReas
        .AddonHrs = 0#
        .NumbLoads = 0
        .CycleTime = 0
        .batchMeth = 0
    End With
    
'//Check if the entry matches a record already in the database.
    If CheckDuplicateEntry(gLabor) = True Then
        Entry.IsDuplicate = True
    Else
        Entry.IsDuplicate = False
    End If
    
'//Send labor record to database
    If Entry.ForceProcess Or Not Entry.IsDuplicate Then
        Call Ins_Labor_Rec
        ProcessEntry = True
    End If
    
    
    GoTo Cleanup
    
Handler:
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Number & vbCrLf & "Description: " & Err.Description
    End If
    
Cleanup:
    Init_labor_rec      'gLabor
    Init_varbs          'gLaborIn
    initJobOpers        'gJobOpers
End Function

Public Sub CopyEntry(original As MassEntry, copy As MassEntry)
    With copy
        .Actdate = Format$(original.Actdate, gDateformat)
        .ActQty = original.ActQty
        .EmplName = original.EmplName
        .EmplNo = original.EmplNo
        .EmplNoIsValid = original.EmplNoIsValid
        .ErrorMsg = original.ErrorMsg
        .IsDuplicate = original.IsDuplicate
        .ForceProcess = original.ForceProcess
        .MachHrs = original.MachHrs
        .MchStd = original.MchStd
        .OrderNo = original.OrderNo
        .OrderNoIsValid = original.OrderNoIsValid
        .OrderQty = original.OrderQty
        .ScrapCd = original.ScrapCd
        .ScrapCdIsValid = original.ScrapCdIsValid
        .ScrapQty = original.ScrapQty
        .SeqNumb = original.SeqNumb
        .SeqNumbIsValid = original.SeqNumbIsValid
        .CCID = original.CCID
        .WCID = original.WCID
        .SetStd = original.SetStd
        .UnplnReas = original.UnplnReas
        .UnplnReasIsValid = original.UnplnReasIsValid
    End With
End Sub

Private Function CheckDuplicateEntry(labor As LaborType) As Boolean
On Error GoTo Handler
    
    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT COUNT(trans_date) AS [Records] " & _
            "FROM [Labor] " & _
            "WHERE empl_no=" & labor.EmplNo & _
            " AND act_date='" & labor.Actdate & _
            "' AND order_no='" & labor.OrderNo & _
            "' AND seq_numb=" & labor.SeqNumb & _
            " AND wc_id='" & labor.WCID & _
            "' AND oper_cd='" & labor.Oper & _
            "' AND act_qty=" & labor.ActQty & _
            " AND act_hrs=" & labor.ActHrs & _
            " AND scrap_qty_1=" & labor.ScrapQty1 & _
            " AND scrap_cd_1='" & labor.ScrapCd1 & _
            "' AND cc_id='" & labor.CCID & _
            "' AND mach_hrs=" & labor.MachHrs & _
            " AND unpln_reas='" & labor.UnplnReas & "'"
    
    Set rs = GetRecordsetData(sql)
    
    If rs!records > 0 Then
        CheckDuplicateEntry = True
    Else
        CheckDuplicateEntry = False
    End If
    
Cleanup:
    rs.Close
    Set rs = Nothing
    Exit Function

Handler:
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Number & vbCrLf & "Description: " & Err.Description
    End If
End Function
