Attribute VB_Name = "Inputs"
Option Explicit
Dim mintLoaded As Integer

Const vbColorWhite = &H8000005
Public Function getScrapCounts(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.getScrapCounts
'Description      : Update the OperStatus in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox
'Objects Used     : None
'Forms Loaded     : ScrapWind
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    getScrapCounts = False
    LoopCnt = iEntry    'pass record/machine as global to form
    ScrapWind.Show 1

    If gSelIndx >= 0 Then
        getScrapCounts = True
    End If
End Function

Public Function getBoxNumbers(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.getScrapCounts
'Description      : Update the OperStatus in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox
'Objects Used     : None
'Forms Loaded     : ScrapWind
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    getBoxNumbers = False
    LoopCnt = iEntry    'pass record/machine as global to form
    frmBoxEntry.Show 1

    If gSelIndx >= 0 Then
        getBoxNumbers = True
    End If
End Function


Public Function inpCloseIndirAns(iEntry As Integer) As Boolean
    Dim strEntry As String
    
    inpCloseIndirAns = False
    
    Set gEntry = New clsEntry
    gEntry.Title = LoadResString(gLanguageOffset + 149)

    gEntry.Add "N", LoadResString(gLanguageOffset + 263)
    gEntry.Add "Y", LoadResString(gLanguageOffset + 264)
    gEntry.Add "E", LoadResString(gLanguageOffset + 144)

    strEntry = basGetEntry()
    If gEntry.Escaped Then
        Set gEntry = Nothing
        gDone = True
        Exit Function
    End If
    Set gEntry = Nothing
    Select Case strEntry
    Case "N"
       inpCloseIndirAns = False
    Case "Y"
       inpCloseIndirAns = True
    Case "E"
       inpCloseIndirAns = False
       gDone = True
       Exit Function
    End Select
End Function




Public Function inpEndOfShift() As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpEndOfShift
'Description      : Update the empClkOutOk global variable.
'Parameters       : None
'Procedures Used  : basGetEntry
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strEntry As String
    
    empClkOutOK = False
    inpEndOfShift = False
    Set gEntry = New clsEntry
    gEntry.Title = "End Of Shift?"
    gEntry.Add "Y", "Yes - End of Shift"
    gEntry.Add "N", "No"
    gEntry.Add "E", LoadResString(gLanguageOffset + 144)
    strEntry = basGetEntry()
    If gEntry.Escaped Then
        Set gEntry = Nothing
        gDone = True
        Exit Function
    End If
    Set gEntry = Nothing
    Select Case strEntry
    Case "Y"
        empClkOutOK = True
    Case "N"
        empClkOutOK = False
    Case "E"
        gDone = True
        
        Exit Function
    End Select
    inpEndOfShift = True
End Function

Public Function AnotherRec() As Integer
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.AnotherRec
'Description      : Menu Entry to ask to enter another labor record.
'Parameters       : None
'Procedures Used  : basGetEntry
'Objects Used     : gEntry(clsEntry)
'Forms Loaded     : None
'Return Parameters: Boolean True if answered Yes
'Tables Updates   : None
'==============================================================================
    Dim strMore As String
    
    AnotherRec = 0
    
    Set gEntry = New clsEntry
    gEntry.Title = LoadResString(gLanguageOffset + 271)
    gEntry.Add "N", LoadResString(gLanguageOffset + 273)
    gEntry.Add "Y", LoadResString(gLanguageOffset + 272)
    gEntry.Add "E", LoadResString(gLanguageOffset + 144)
    
    strMore = basGetEntry()
    If gEntry.Escaped Then
        Set gEntry = Nothing
        AnotherRec = -1
        gDone = True
        Exit Function
    End If
    Set gEntry = Nothing
    
    Select Case strMore
    Case "Y"
        AnotherRec = 1
    Case "N"
        AnotherRec = 0
    Case "E"
        AnotherRec = -1
        gDone = True
        Exit Function
    End Select

End Function




Public Function getUnplnData(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.getUnplnData
'Description      : Enter data required for an unplanned operation:
'                   WCID, Oper, FixStd, MchStd & UnplnReas
'Parameters       : index of record/machine being entered
'Procedures Used  : initJobOpers, inpWCID, GetWorkCenter, inpOperNumb,
'                   inpFixStd, inpMchStd, inpUnplnReas
'Objects Used     : None
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    getUnplnData = False
    Call initJobOpers
    
    Screen.ActiveForm.lblJobNo.Caption = gJobMaster.OrderNo
    Do Until gDone Or getUnplnData
        'gDone = True        'Default for exit to Main Menu
        If Not inpWCID(iEntry) Then
            gDone = True
            Exit Function
        End If
        Screen.ActiveForm.lblWCID.Caption = gLaborIn(iEntry).WCID
        gLaborIn(iEntry).CCID = gWorkCenter.CCID
        If Not GetCostCenter(gWorkCenter.CCID) Then
            gCostCenter.PaidLunch = "N"
        Else
            Screen.ActiveForm.lblCCID.Caption = gWorkCenter.CCID
        End If
        
        
        If Not GetWorkCenter(gLaborIn(iEntry).WCID) Then
            gCostCenter.PaidLunch = "N"
        End If
        
        If Not inpOperNumb(iEntry) Then
            gDone = True
            Exit Function
        End If

        Screen.ActiveForm.lblOperCd.Caption = gLaborIn(iEntry).Oper
        gLaborIn(iEntry).MchStd = 0
        gLaborIn(iEntry).SetStd = 0
        
        If Not inpFixStd(iEntry) Then
            gDone = True
            Exit Function
        End If

        Screen.ActiveForm.lblFixStd.Caption = gLaborIn(iEntry).SetStd
        If gLaborIn(iEntry).SetStd = 0 Then
            If Not inpMchStd(iEntry) Then
                gDone = True
                Exit Function
            End If
        End If
        Screen.ActiveForm.lblMachStd.Caption = gLaborIn(iEntry).MchStd
        
        If Not inpUnplnReas(iEntry) Then Exit Function
        Screen.ActiveForm.lblUnplnReas.Caption = gLaborIn(iEntry).UnplnReas
        
        getUnplnData = True 'Exit loop
    Loop
End Function






Public Function inpJobAndSeq(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpJobAndSeq
'Description      : Enter the Job and Sequence Number, and any Unplanned Labor
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, GetJobMaster, ValidNumber,
'                   getUnplnData, GetJobOper, getWorkCenter
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strEntry As String
    Dim strDataId As String
    Dim strSeqNo As String
    
    inpJobAndSeq = False        'to force loop back till done
    
    Do Until inpJobAndSeq Or gDone
        Do Until inpJobAndSeq Or gDone       'Extra loop for exit do stmts
       '     ErrBox "If labor is unplanned, key in only job number and Enter"
            
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 165)
            gEntry.EntryType = 0        'Any Printable Characters
            strEntry = basGetEntry()
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            strDataId = UCase(Left(strEntry, 2))
            If strDataId = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If strDataId = "NU" Then
                ErrBox LoadResString(gLanguageOffset + 28), 2
                ErrorBeep
                Exit Do           'Loop Again
            End If
            
            If Len(strDataId) < 1 Then
                ErrBox LoadResString(gLanguageOffset + 78), 2
                ErrorBeep
                Exit Do           'Loop Again
            End If
                
            If Len(strEntry) > 11 Then
                If strDataId = "JS" Then
                    gLaborIn(iEntry).OrderNo = UCase(Mid(strEntry, 3, 7))
                    strSeqNo = Mid(strEntry, 10, 4)
                Else
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            Else
                gLaborIn(iEntry).OrderNo = UCase(Left(strEntry, 7))
                strSeqNo = Mid(strEntry, 8, 4)
            End If
            
            If Not GetJobMaster(gLaborIn(iEntry).OrderNo) Then
                If Left(gLaborIn(iEntry).OrderNo, 2) = "09" Then
                    Screen.ActiveForm.lblErrMsg.Caption = LoadResString(gLanguageOffset + 236)
                    Screen.ActiveForm.lblErrMsg.BackColor = vbColorWhite
                    strSeqNo = "000 "
                    gLaborIn(iEntry).MatlWgt = 0
                    gJobMaster.Actdate = 0
                    gJobMaster.OrderNo = gLaborIn(iEntry).OrderNo
                    gJobMaster.OrderType = "SO"
                Else
                    ErrBox LoadResString(gLanguageOffset + 61), 2
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            Else
                If gJobMaster.OrderStat = "C" Then
                       ErrBox "Order status is complete", 2
                       ErrorBeep
                       Exit Do           'Loop Again
                Else
                       gLaborIn(iEntry).MatlWgt = gJobMaster.MatlWgt
                End If
            End If
            
            If Len(strSeqNo) = 0 Then
                strSeqNo = "000 "
            Else
                If Not ValidNumber(strSeqNo) Then
                    ErrBox LoadResString(gLanguageOffset + 111), 2
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            End If
            
            gLaborIn(iEntry).SeqNo = CInt(strSeqNo)
            If gLaborIn(iEntry).SeqNo = 990 _
                    Or gLaborIn(iEntry).SeqNo = 0 Then
                If div = bristol Then        'no scrap for unplanned oper at Bristol
                    ErrBox LoadResString(gLanguageOffset + 304), 2
                    ErrorBeep
                    Exit Do           'Loop Again
                Else
                    gLaborIn(iEntry).SeqNo = 0
                    gLaborIn(iEntry).EOrS = "S"
                   Screen.ActiveForm.lblErrMsg.Caption = ""
                    If Not getUnplnData(iEntry) Then
                        Exit Do           'Loop Again
                    End If
                    inpJobAndSeq = True     'Done - Unplanned
                End If
            Else
                If Not GetJobOper(gLaborIn(iEntry).OrderNo, _
                        gLaborIn(iEntry).SeqNo) Then
                    ErrBox LoadResString(gLanguageOffset + 87), 2
                    ErrorBeep
                    Exit Do
                Else
                    If gJobOpers.WCID = "-----" Then
                        ErrBox LoadResString(gLanguageOffset + 38), 2
                        ErrorBeep
                        Exit Do
                    ElseIf Not GetWorkCenter(gJobOpers.WCID) Then
                          ErrBox LoadResString(gLanguageOffset + 103), 2
                          ErrorBeep
                          Exit Do
'                    ElseIf gJobOpers.WCID = "OSSRV" And div = bristol Then
 '                        ErrBox "Cannot enter labor or scrap with this oper.", 2
  '                       ErrorBeep
   '                      Exit Do
    '                ElseIf Left(gJobOpers.WCID, 1) = "V" And div = bristol Then
     '                    ErrBox "Cannot enter labor or scrap against this oper.", 2
      '                   ErrorBeep
       '                  Exit Do
                     End If
                    gLaborIn(iEntry).UnplnReas = "  "
                    gLaborIn(iEntry).WCID = gJobOpers.WCID
                    gLaborIn(iEntry).Oper = gJobOpers.Oper
                    gLaborIn(iEntry).EOrS = gJobOpers.EorSStd
                    gLaborIn(iEntry).SetStd = gJobOpers.SetStd
                    gLaborIn(iEntry).MchStd = gJobOpers.MchStd  'labor.run_std????
                    gLaborIn(iEntry).MchStd = gJobOpers.MchStd
                    inpJobAndSeq = True     'Done - Planned
                End If
            End If
        Loop
    Loop
    Screen.ActiveForm.lblErrMsg.Caption = ""
End Function
Public Function inpWorkOrder(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpJobAndSeq
'Description      : Enter the Job and Sequence Number, and any Unplanned Labor
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, GetJobMaster, ValidNumber,
'                   getUnplnData, GetJobOper, getWorkCenter
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strEntry As String
    Dim strDataId As String
    Dim strSeqNo As String
    
    inpWorkOrder = False        'to force loop back till done
    
    Do Until inpWorkOrder Or gDone
        Do Until inpWorkOrder Or gDone       'Extra loop for exit do stmts
       '     ErrBox "If labor is unplanned, key in only job number and Enter"
            
            Set gEntry = New clsEntry
            gEntry.Title = "Enter Work Order"
            gEntry.EntryType = 0        'Any Printable Characters
            strEntry = basGetEntry()
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            strDataId = UCase(Left(strEntry, 2))
            If strDataId = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If strDataId = "NU" Then
                ErrBox LoadResString(gLanguageOffset + 28), 2
                ErrorBeep
                Exit Do           'Loop Again
            End If
            
            If Len(strDataId) < 1 Then
                ErrBox LoadResString(gLanguageOffset + 78), 2
                ErrorBeep
                Exit Do           'Loop Again
            End If
                
            If Len(strEntry) > 11 Then
                If strDataId = "JS" Then
                    gLaborIn(iEntry).OrderNo = UCase(Mid(strEntry, 3, 7))
                    strSeqNo = "0000"
                Else
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            Else
                gLaborIn(iEntry).OrderNo = UCase(Left(strEntry, 7))
                strSeqNo = "0000"
            End If
            
            If Not GetJobMaster(gLaborIn(iEntry).OrderNo) Then
                    ErrBox LoadResString(gLanguageOffset + 61), 2
                    ErrorBeep
                    Exit Do           'Loop Again
            Else
                If gJobMaster.OrderStat = "C" Then
                       ErrBox "Order status is complete", 2
                       ErrorBeep
                       Exit Do           'Loop Again
                End If
            End If
            
            If Len(strSeqNo) = 0 Then
                strSeqNo = "0000"
            Else
                If Not ValidNumber(strSeqNo) Then
                    ErrBox LoadResString(gLanguageOffset + 111), 2
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            End If
           inpWorkOrder = True     'Done - Unplanned
        Loop
    Loop
    Screen.ActiveForm.lblErrMsg.Caption = ""
End Function



Public Function inpWcAns(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpWcAns
'Description      : Ask if Work Center is OK, else ask for new one and verify
'                   it is ok, then verify Cost Center
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, inpWCID, GetWorkCenter, GetCostCenter
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strOK As String
    
    inpWcAns = False
    
    If gLaborIn(iEntry).WCID = "----" _
            Or gLaborIn(iEntry).WCID = "-----" _
            Or div = corry Then
        If Not inpWCID(iEntry) Then Exit Function
        Screen.ActiveForm.lblWCID.Caption = gLaborIn(iEntry).WCID
    Else
        Set gEntry = New clsEntry
        gEntry.Title = LoadResString(gLanguageOffset + 166)
        If div = stroms Then
              gEntry.Add "J", LoadResString(gLanguageOffset + 265)
              gEntry.Add "N", LoadResString(gLanguageOffset + 266)
              gEntry.Add "E", LoadResString(gLanguageOffset + 144)
         ElseIf div = monterrey Then
              gEntry.Add "S", LoadResString(gLanguageOffset + 265)
              gEntry.Add "N", LoadResString(gLanguageOffset + 266)
              gEntry.Add "E", LoadResString(gLanguageOffset + 144)
         ElseIf div = remesa Then
              gEntry.Add "S", LoadResString(gLanguageOffset + 265)
              gEntry.Add "N", LoadResString(gLanguageOffset + 266)
              gEntry.Add "E", LoadResString(gLanguageOffset + 144)
         Else
              gEntry.Add "Y", LoadResString(gLanguageOffset + 265)
              gEntry.Add "N", LoadResString(gLanguageOffset + 266)
              gEntry.Add "E", LoadResString(gLanguageOffset + 144)
          End If
          
        strOK = basGetEntry()
        If gEntry.Escaped Then
            Set gEntry = Nothing
            gDone = True
            Exit Function
        End If
        Set gEntry = Nothing
        If div = stroms Then
              If strOK = "J" Then
                 strOK = "Y"
             End If
        ElseIf div = monterrey Then
              If strOK = "S" Then
                  strOK = "Y"
             End If
        ElseIf div = remesa Then
              If strOK = "S" Then
                  strOK = "Y"
             End If
        End If
        Select Case strOK
        Case "Y"
            'Done - no changes needed.
            'need_wc_id = "N" Rem this is obsolete, field is not used !!!!!
        Case "N"
            If Not inpWCID(iEntry) Then   'Allow Escape up
                gDone = True
                Exit Function
            End If
            Screen.ActiveForm.lblWCID.Caption = gLaborIn(iEntry).WCID
        Case "E"
            gDone = True
            Exit Function
        End Select
    End If
    
    If GetWorkCenter(gLaborIn(iEntry).WCID) Then
        gLaborIn(iEntry).CCID = gWorkCenter.CCID
        If Not GetCostCenter(gWorkCenter.CCID) Then
            gCostCenter.PaidLunch = "N"
        Else
            Screen.ActiveForm.lblCCID.Caption = gWorkCenter.CCID
        End If
    Else
        gWorkCenter.PaidLunch = "N"
        gCostCenter.PaidLunch = "N"
    End If
    
    inpWcAns = True
End Function

Public Function inpWCID(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpWCID
'Description      : Enter the WCID in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, GetWorkCenter
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strWCID As String
    
    inpWCID = False
    gDone = False
    
    Do Until inpWCID Or gDone
      Do Until inpWCID Or gDone
        Set gEntry = New clsEntry
        gEntry.Title = LoadResString(gLanguageOffset + 206)
        gEntry.EntryType = 0    'Any Displayable Characters
        strWCID = UCase(basGetEntry())
        If gEntry.Escaped Then
            Set gEntry = Nothing
            gDone = True
            Exit Function
        End If
        Set gEntry = Nothing
        
        If Len(strWCID) = 0 Then
            ErrBox LoadResString(gLanguageOffset + 78), 2
            ErrorBeep
            Exit Do   'go Back and Ask Again
        End If
        
        If Left(strWCID, 2) = "ZZ" Then
            gDone = True
            Exit Function
        End If
        
        If Len(strWCID) > 5 Then
            If Left(strWCID, 2) = "WC" Then
                gLaborIn(iEntry).WCID = Mid(strWCID, 3, 5)
            Else
                ErrorBeep
            End If
        Else
            gLaborIn(iEntry).WCID = Left(strWCID, 5)
        End If
        
        
        If Not GetWorkCenter(gLaborIn(iEntry).WCID) Then
            ErrBox LoadResString(gLanguageOffset + 104), 2
            ErrorBeep
        Else
            If gLaborIn(iEntry).WCID = "----" _
                    Or gLaborIn(iEntry).WCID = "-----" Then
                ErrBox LoadResString(gLanguageOffset + 104), 2
                ErrorBeep
            Else
                inpWCID = True  'Got one - Exit loop and function
            End If
        End If
      Loop
    Loop

End Function

Public Function inpAddonHours(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpAddonHours
'Description      : Enter the addonHours in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim bEntered As Boolean
    
    inpAddonHours = False
    strNumb = ""
    
    Do Until inpAddonHours Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpAddonHours Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 207)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Len(strwork) = 0 Then
                gLaborIn(iEntry).addonHours = 0
                inpAddonHours = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 1)
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strwork) > 4 Then
            ErrBox LoadResString(gLanguageOffset + 92), 2
            ErrorBeep
            bEntered = False
        Else
            If Not CheckHrs(strNumb) Then
                ErrBox LoadResString(gLanguageOffset + 111), 2
                ErrorBeep
                bEntered = False
            Else
                gLaborIn(iEntry).addonHours = CCur(strNumb)
                inpAddonHours = True
            End If
        End If
    Loop
End Function

Public Function inpIndirHrs(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpIndirHrs
'Description      : Enter the inHours in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, CheckHrs
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim bEntered As Boolean
    
    
    inpIndirHrs = False
    strNumb = ""
    gDone = False
    Do Until inpIndirHrs Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpIndirHrs Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 208)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Len(strwork) = 0 Then
                strwork = "0.0"
            End If
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 1)
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strwork) > 4 Then
             ErrBox LoadResString(gLanguageOffset + 92), 2
            ErrorBeep
            bEntered = False
        Else
            If Not CheckHrs(strNumb) Then
                ErrBox LoadResString(gLanguageOffset + 111), 2
                ErrorBeep
                bEntered = False
            Else
                gLaborIn(iEntry).inHours = CCur(strNumb)
                If gLaborIn(iEntry).inHours < 0.1 Then
                    ErrBox LoadResString(gLanguageOffset + 53), 2
                    ErrorBeep
                    bEntered = False
                    inpIndirHrs = False
                ElseIf gLaborIn(iEntry).inHours > 23.9 Then
                       'User Entered too many hours
                       ErrBox LoadResString(gLanguageOffset + 54), 2
                       ErrorBeep
                       bEntered = False
                       inpIndirHrs = False
                Else
                    inpIndirHrs = True
                End If
                
            End If
        End If
    Loop
End Function

Public Function CheckHrs(strNumb As String) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.CheckHrs
'Description      : validate an entry with hours entered
'Parameters       : string to validate
'Procedures Used  : None
'Objects Used     : None
'Forms Loaded     : None
'Return Parameters: True if valid hours string
'Tables Updates   : None
'==============================================================================
    Dim iDecs, iDecPos As Integer
    Dim I As Integer
    Dim strchar As String
      
    iDecs = 0
    iDecPos = -1
    CheckHrs = False
    
     strNumb = Trim(strNumb)
    If Len(strNumb) = 0 Then Exit Function
    
    For I = 1 To Len(strNumb)
        strchar = Mid(strNumb, I, 1)
        If strchar = "," Then
           strNumb = Mid(strNumb, 1, I - 1) & "." & Mid(strNumb, I + 1, Len(strNumb))
           Exit For
        End If
    Next I
    
    For I = 1 To Len(strNumb)
        strchar = Mid(strNumb, I, 1)
        If strchar = "." Then
            iDecPos = I
            iDecs = iDecs + 1
            If iDecs > 1 Then Exit Function
        Else
            If strchar < "0" Or strchar > "9" Then
                Exit Function
            End If
        End If
    Next I
    
    If (iDecPos < (Len(strNumb) - 1) And (iDecPos <> -1)) Then
        Exit Function
    End If
    
   CheckHrs = True

End Function

Public Function inpProdQty(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpProdQty
'Description      : Enter the ActQty in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, AllDigits
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim I As Integer
    Dim iChar As String
    Dim bEntered As Boolean
    
    inpProdQty = False
    strNumb = ""
    gDone = False
    
    Do Until inpProdQty Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpProdQty Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 175)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 7)
                If Len(strNumb) > 6 Then
                    bEntered = True
                    Exit Do
                End If
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strNumb) = 0 Then
            strNumb = "0"
        End If

        If Not AllDigits(strNumb) Then
            ErrBox LoadResString(gLanguageOffset + 111), 2
            ErrorBeep
            bEntered = False
        ElseIf Len(Trim(strNumb)) < 8 Then
            gLaborIn(iEntry).ActQty = CLng(strNumb)
            If CheckValidQty(gLaborIn(iEntry).OrderNo, gLaborIn(iEntry).SeqNo, gLaborIn(iEntry).ActQty, "E") Then
                   inpProdQty = True
            Else
                   ErrBox LoadResString(gLanguageOffset + 85), 2
                   ErrorBeep
                   inpProdQty = False
                   bEntered = False
            End If
        Else
            ErrBox LoadResString(gLanguageOffset + 50), 2
            ErrorBeep
            inpProdQty = False
            bEntered = False
       End If
    Loop
    
End Function
Public Function inpProdQtyReqd(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpProdQtyReqd
'Description      : Enter the ActQty in gLaborIn(iEntry) must be greater than zero
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, AllDigits
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim I As Integer
    Dim iChar As String
    Dim bEntered As Boolean
    
    inpProdQtyReqd = False
    strNumb = ""
    gDone = False
    
    Do Until inpProdQtyReqd Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpProdQtyReqd Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 175)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 7)
                If Len(strNumb) > 6 Then
                    bEntered = True
                    Exit Do
                End If
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strNumb) = 0 Then
            strNumb = "0"
        End If

        If Not AllDigits(strNumb) Then
            ErrBox LoadResString(gLanguageOffset + 111), 2
            ErrorBeep
            bEntered = False
        ElseIf Len(Trim(strNumb)) < 8 Then
            gLaborIn(iEntry).ActQty = CLng(strNumb)
            If gLaborIn(iEntry).ActQty > 0 Then 'quantity greater than zero is required
                   inpProdQtyReqd = True
            Else
                   ErrBox "A quantity must be entered", 2
                   ErrorBeep
                   inpProdQtyReqd = False
                   bEntered = False
            End If
        Else
            ErrBox LoadResString(gLanguageOffset + 50), 2
            ErrorBeep
            inpProdQtyReqd = False
            bEntered = False
       End If
    Loop
    
End Function


Public Function inpSendOutQty(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpProdQty
'Description      : Enter the ActQty in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, AllDigits
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim I As Integer
    Dim iChar As String
    Dim bEntered As Boolean
    
    inpSendOutQty = False
    strNumb = ""
    gDone = False
    
    Do Until inpSendOutQty Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpSendOutQty Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 288)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 7)
                If Len(strNumb) > 6 Then
                    bEntered = True
                    Exit Do
                End If
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strNumb) = 0 Then
            strNumb = "0"
        End If

        If Not AllDigits(strNumb) Then
            ErrBox LoadResString(gLanguageOffset + 111), 2
            ErrorBeep
            bEntered = False
        ElseIf Len(Trim(strNumb)) < 8 Then
            gLaborIn(iEntry).ActQty = CLng(strNumb)
            If gLaborIn(iEntry).ActQty = 0 Then
               ErrBox LoadResString(gLanguageOffset + 289), 2
               ErrorBeep
               bEntered = False
            ElseIf CheckValidQty(gLaborIn(iEntry).OrderNo, gLaborIn(iEntry).SeqNo, gLaborIn(iEntry).ActQty, "E") Then
                   inpSendOutQty = True
            Else
                   ErrBox LoadResString(gLanguageOffset + 85), 2
                   ErrorBeep
                   inpSendOutQty = False
                   bEntered = False
            End If
        Else
            ErrBox LoadResString(gLanguageOffset + 50), 2
            ErrorBeep
            inpSendOutQty = False
            bEntered = False
       End If
    Loop
    
End Function



Public Function inpAssyQty(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpProdQty
'Description      : Enter the ActQty in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, AllDigits
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim I As Integer
    Dim iChar As String
    Dim bEntered As Boolean
    
    inpAssyQty = False
    strNumb = ""
    gDone = False
    
    Do Until inpAssyQty Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpAssyQty Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 175)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 7)
                If Len(strNumb) > 6 Then
                    bEntered = True
                    Exit Do
                End If
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strNumb) = 0 Then
            ErrBox LoadResString(gLanguageOffset + 57), 2
            ErrorBeep
            bEntered = False
       ElseIf Not AllDigits(strNumb) Then
            ErrBox LoadResString(gLanguageOffset + 111), 2
            ErrorBeep
            bEntered = False
       ElseIf Len(Trim(strNumb)) < 8 Then
            gLaborIn(iEntry).ActQty = CLng(strNumb)
            inpAssyQty = True
            If gLaborIn(iEntry).ActQty <= 0 Then
                   ErrBox LoadResString(gLanguageOffset + 57), 2
                   ErrorBeep
                   inpAssyQty = False
                   bEntered = False
           End If
        ElseIf Len(Trim(strNumb)) >= 8 Then
            ErrBox LoadResString(gLanguageOffset + 50), 2
            ErrorBeep
            inpAssyQty = False
            bEntered = False
       End If
       If inpAssyQty = True And (gLaborIn(iEntry).ActQty > (gJobMaster.OrderQty - gJobOpers.ActQty)) Then
              ErrBox LoadResString(gLanguageOffset + 274), 2
             ErrorBeep
             inpAssyQty = False
             bEntered = False
     End If
     If inpAssyQty = True And _
        (gLaborIn(iEntry).ActQty < (gJobMaster.OrderQty - gJobOpers.ActQty)) Then
              ErrBox LoadResString(gLanguageOffset + 275), 2
             ErrorBeep
     End If

    
    
    Loop
    
End Function

Public Function CallCmpl(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.CallCmpl
'Description      : Update the OperStatus in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox
'Objects Used     : gEntry(clsEntry)
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strEntry As String
    
    CallCmpl = False
    
    Set gEntry = New clsEntry
    gEntry.Title = LoadResString(gLanguageOffset + 245)
    gEntry.Add "N", LoadResString(gLanguageOffset + 246)
    gEntry.Add "Y", LoadResString(gLanguageOffset + 247)
    gEntry.Add "E", LoadResString(gLanguageOffset + 144)
    strEntry = basGetEntry()
    If gEntry.Escaped Then
        Set gEntry = Nothing
        gDone = True
        Exit Function
    End If
    Set gEntry = Nothing
    Select Case strEntry
    Case "N"
        gLaborIn(iEntry).OperStatus = "P"
    Case "Y"
        gLaborIn(iEntry).OperStatus = "C"
    Case "E"
        gDone = True
        Exit Function
    End Select
    CallCmpl = True
End Function

Public Function ReEnter(I As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.ReEnter
'Description      : Menu to pick which fields on screen to change.  This code
'                   is parameter driven to only show options that are valid
'                   for a particular division and other circumstances.
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, ErrorChk, inpJobAndSeq,
'                   DispValues, inpWCID, GetCostCenter, inpProdQty,
'                   getScrapCounts, inpSetHours, inpMchHours, inpAddonHours,
'                   inpEndOfShift, CallCmpl, inpNoOfLoads, CalcBtErnHrs,
'                   CalcBatchHrs, inpCycleTime, LdBatchIncn, GetBatchIncn
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strEntrySub As String
    Dim strEntryEmpNo As String
    
    ReEnter = False
    gDone = False
    
    Do Until gDone Or ReEnter
        Set gEntry = New clsEntry
        gEntry.Title = LoadResString(gLanguageOffset + 192)
        gEntry.Add "A", LoadResString(gLanguageOffset + 193)
        gEntry.Add "E", LoadResString(gLanguageOffset + 194)
        gEntry.Add "J", LoadResString(gLanguageOffset + 195)
        gEntry.Add "W", LoadResString(gLanguageOffset + 196)
        gEntry.Add "Q", LoadResString(gLanguageOffset + 197)
        gEntry.Add "S", LoadResString(gLanguageOffset + 198)
        gEntry.Add "F", LoadResString(gLanguageOffset + 209)
        gEntry.Add "M", LoadResString(gLanguageOffset + 210)
        gEntry.Add "B", LoadResString(gLanguageOffset + 211)
        gEntry.Add "C", LoadResString(gLanguageOffset + 212)
        gEntry.Add "O", LoadResString(gLanguageOffset + 199)
        gEntry.Add "L", LoadResString(gLanguageOffset + 213)
        gEntry.Add "T", LoadResString(gLanguageOffset + 214)
        gEntry.Add "H", LoadResString(gLanguageOffset + 215)
        strEntrySub = basGetEntry()
        If gEntry.Escaped Then
            Set gEntry = Nothing
            gDone = True
            Exit Function
        End If
        Set gEntry = Nothing
        Select Case strEntrySub
        Case "A"
            Call ErrorChk(I)
            ReEnter = True
        Case "E"
            gDone = True
            Exit Function
        Case "J"
            If Not inpJobAndSeq(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Call DispValues(I)
                        End If
        Case "W"
            If Not inpWCID(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Screen.ActiveForm.lblWCID.Caption = gLaborIn(I).WCID
                gLaborIn(I).CCID = gWorkCenter.CCID
                If Not GetCostCenter(gWorkCenter.CCID) Then
                    gCostCenter.PaidLunch = "N"
                End If
                Screen.ActiveForm.lblCCID.Caption = gLaborIn(I).CCID
            End If
        Case "Q"
            If Not inpProdQty(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Screen.ActiveForm.lblProdQty.Caption = gLaborIn(I).ActQty
            End If
        Case "S"
            If Not gScrapColl Then
                ErrorBeep
            Else
                If Not getScrapCounts(I) Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                Else
                    Dim iScrap As Long
                    Dim j As Integer
                    iScrap = 0
                    For j = 1 To 8
                        iScrap = iScrap + gLaborIn(I).Scrap(j).Qty
                    Next j
                    Screen.ActiveForm.lblScrapQty.Caption = CStr(iScrap)
                End If
            End If
        Case "F"
            If gLaborIn(I).NoOfLoads > 0 _
                    Or gLaborIn(I).CycleTime > 0 Then
                ErrBox LoadResString(gLanguageOffset + 8), 2
                ErrorBeep
            Else
                If Not okToPrompt _
                        Or gLaborIn(I).RecType <> "F" Then
                    ErrorBeep
                Else
                    If Not inpSetHours(I) Then
                        gDone = False       'be sure to loop back
                        ErrorBeep
                    Else
                        Screen.ActiveForm.lblSetupHrs.Caption = gLaborIn(I).setupHours
                    End If
                End If
            End If
        Case "M"
            If gLaborIn(I).NoOfLoads > 0 _
                    Or gLaborIn(I).CycleTime > 0 Then
                ErrBox LoadResString(gLanguageOffset + 8), 2
                ErrorBeep
            Else
                If Not okToPrompt _
                        Or gLaborIn(I).RecType <> "D" Then
                    ErrorBeep
                Else
                    If Not inpMchHours(I) Then
                        gDone = False       'be sure to loop back
                        ErrorBeep
                    Else
                        Screen.ActiveForm.lblMachHrs.Caption = gLaborIn(I).machHours
                    End If
                End If
            End If
        Case "B"
            If Not inpAddonHours(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Screen.ActiveForm.lblAddonHrs.Caption = gLaborIn(I).addonHours
            End If
        Case "C"
            If gEosFlag = 0 Then
                ErrBox LoadResString(gLanguageOffset + 86), 2
                ErrorBeep
            Else
                If Not inpEndOfShift() Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                End If
            End If
        Case "O"
            If Not CallCmpl(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Screen.ActiveForm.lblOperStatus.Caption = gLaborIn(I).OperStatus
            End If
        Case "L"
            If Not okToPrompt Then
                ErrorBeep
            Else
                If Not inpNoOfLoads(I) Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                Else
                    Screen.ActiveForm.lblNumbOfLoads.Caption = gLaborIn(I).NoOfLoads
                    Call CalcBtErnHrs(I)
                    Call CalcBatchHrs(I)
                End If
            End If
        Case "T"
            If Not okToPrompt Then
                ErrorBeep
            Else
                If gLaborIn(I).NoOfLoads = 0 Then
                    ErrorBeep
                    ErrBox LoadResString(gLanguageOffset + 79), 2
                Else
                    If Not inpCycleTime(I) Then
                        gDone = False       'be sure to loop back
                        ErrorBeep
                    Else
                        Screen.ActiveForm.lblCycleTime.Caption = gLaborIn(I).CycleTime
                        Call CalcBatchHrs(I)
                    End If
                End If
            End If
        Case "H"
            If Not okToPrompt Then
                ErrorBeep
            Else
                If gLaborIn(I).NoOfLoads = 0 Then
                    ErrorBeep
                    ErrBox LoadResString(gLanguageOffset + 79), 2
                Else
                    LdBatchIncn (gLaborIn(I).WCID)
                    If Not getBatchIncn(I) Then
                        gDone = False       'be sure to loop back
                        ErrorBeep
                    Else
                        Screen.ActiveForm.lblBatchMethod.Caption = gLaborIn(I).biMethod
                        Call CalcBtErnHrs(I)
                    End If
                End If
            End If
        End Select
    Loop

End Function

Public Sub DispValues(iEntry As Integer)

    Screen.ActiveForm.lblPartNo.Caption = gJobMaster.PartNo
    Screen.ActiveForm.lblCustName.Caption = gJobMaster.CustName
    Screen.ActiveForm.lblJobNo.Caption = gJobMaster.OrderNo
    Screen.ActiveForm.lblOrdQty.Caption = gJobMaster.OrderQty
    Screen.ActiveForm.lblPartRev.Caption = gJobMaster.PartRev
    Screen.ActiveForm.lblOrdType.Caption = gJobMaster.OrderType
    If Not GetWorkCenter(gLaborIn(iEntry).WCID) Then
        ErrBox LoadResString(gLanguageOffset + 103), 2
        ErrorBeep
    Else
        gLaborIn(iEntry).CCID = gWorkCenter.CCID
    End If
    Screen.ActiveForm.lblSeqNo.Caption = gLaborIn(iEntry).SeqNo
    Screen.ActiveForm.lblWCID.Caption = gLaborIn(iEntry).WCID
    Screen.ActiveForm.lblCCID.Caption = gLaborIn(iEntry).CCID
    Screen.ActiveForm.lblOperCd.Caption = gLaborIn(iEntry).Oper
    Screen.ActiveForm.lblUnplnReas.Caption = gLaborIn(iEntry).UnplnReas
    Screen.ActiveForm.lblFixStd.Caption = gLaborIn(iEntry).SetStd
    Screen.ActiveForm.lblMachStd.Caption = gLaborIn(iEntry).MchStd
    Screen.ActiveForm.lblEorS.Caption = gLaborIn(iEntry).EOrS
End Sub

Public Sub CalcBatchHrs(iEntry As Integer)
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.CalcBatchHrs
'Description      : Update the setupHours or machHours in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : ErrBox
'Objects Used     : None
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    Dim cMachHrs As Currency
    
    Dim cCycleTime As Currency
    Dim cNoOfLoads As Currency
    cCycleTime = gLaborIn(iEntry).CycleTime
    cNoOfLoads = gLaborIn(iEntry).NoOfLoads
    cMachHrs = (cCycleTime * cNoOfLoads) / 60
    cMachHrs = Format(cMachHrs, "00.0")
    If gLaborIn(iEntry).RecType = "F" Then
        gLaborIn(iEntry).setupHours = cMachHrs
        Screen.ActiveForm.lblSetupHrs.Caption = Format(gLaborIn(iEntry).setupHours, gOneDecPos)
    Else
        gLaborIn(iEntry).machHours = cMachHrs
        Screen.ActiveForm.lblMachHrs.Caption = Format(gLaborIn(iEntry).machHours, gOneDecPos)
    End If
    If (gLaborIn(iEntry).machHours = 0 _
            And gLaborIn(iEntry).setupHours = 0) Then
        ErrBox LoadResString(gLanguageOffset + 101), 2
    End If
End Sub


Public Sub CalcBtErnHrs(iEntry As Integer)
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.CalcBtErnHrs
'Description      : Update the inHours in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : ErrBox
'Objects Used     : None
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    Dim cMachHrs As Currency
    Dim cEarnHrs As Currency
    
    If Not GetBICode(gLaborIn(iEntry).biMethod, _
            gLaborIn(iEntry).WCID) Then
        gLaborIn(iEntry).inHours = 0
    Else
        cEarnHrs = (gBatchIncen.IncenMin _
            * gLaborIn(iEntry).NoOfLoads) / 60
        gLaborIn(iEntry).inHours = CCur(Format(cEarnHrs, "00.00"))
    End If
End Sub




Public Function inpNumbMachs() As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpNumbMachs
'Description      : Enter the NumbMachs in gLaborIn(1)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strEntry As String
    inpNumbMachs = False
    
    Set gEntry = New clsEntry
    gEntry.Title = "Incentive Code :"
    gEntry.Add "0", "Single Machine - Day Rate"
    gEntry.Add "1", "Single Machine - Incentive"
    gEntry.Add "2", "Two Machines"
    gEntry.Add "3", "Three Machines"
    gEntry.Add "4", "Four Machines"
    gEntry.Add "5", "Five Machines"
    gEntry.Add "6", "Six Machines"
    gEntry.Add "7", "Seven Machines"
    gEntry.Add "8", "Eight Machines"
    gEntry.Add "9", "Nine Machines"
    'JC 01/26/99 What they are now
    gEntry.Add "N", "Ten Machines"
    gEntry.Add "E", "Eleven Machines"
    gEntry.Add "T", "Twelve Machines"
    
    'JC 01/26/99 What they were before
    'gEntry.Add "T", "Ten Machines"
    'gEntry.Add "L", "Eleven Machines"
    'gEntry.Add "W", "Twelve Machines"
    strEntry = basGetEntry()
    If gEntry.Escaped Then
        Set gEntry = Nothing
        gDone = True
        Exit Function
    End If
    Set gEntry = Nothing
    Select Case strEntry
    Case "0"
        gLaborIn(1).NumbMachs = 1
        gLaborIn(1).incenRec = 0
    Case "1" To "9"
        gLaborIn(1).NumbMachs = CInt(strEntry)
        gLaborIn(1).incenRec = CInt(strEntry)
    Case "N"
        gLaborIn(1).NumbMachs = 10
        gLaborIn(1).incenRec = 10
    Case "E"
        gLaborIn(1).NumbMachs = 11
        gLaborIn(1).incenRec = 11
    Case "T"
        gLaborIn(1).NumbMachs = 12
        gLaborIn(1).incenRec = 12
    End Select
    
    inpNumbMachs = True

End Function

Public Function inpMchHours(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpMchHours
'Description      : Enter the machHours in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, CheckHrs
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim I As Integer
    Dim iChar As String
    Dim bEntered As Boolean
    
    inpMchHours = False
    strNumb = ""
    gDone = False
    
    Do Until inpMchHours Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpMchHours Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 216)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Len(Trim$(strwork)) = 0 Then
                strwork = "0.0"
'                bEntered = True
'                Exit Do
            End If
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 1)
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strNumb) > 4 Then
            ErrBox LoadResString(gLanguageOffset + 92), 2
            ErrorBeep
            bEntered = False
        Else
            If Not CheckHrs(strNumb) Then
                ErrBox LoadResString(gLanguageOffset + 111), 2
                ErrorBeep
                bEntered = False
            Else
                gLaborIn(iEntry).machHours = CCur(strNumb)
                If gLaborIn(iEntry).machHours > 99.9 Then
                    ErrBox "Number too large", 2
                    ErrorBeep
                    bEntered = False
                ElseIf gLaborIn(iEntry).machHours = 0 And gbInpSetAndMach = False Then
                    ErrBox LoadResString(gLanguageOffset + 109), 2
                    bEntered = False
                    inpMchHours = False
                ElseIf gLaborIn(iEntry).machHours <= 99.9 And div = singapore Then
'                       (Left(gLaborIn(iEntry).Oper, 3) = "890" Or Left(gLaborIn(iEntry).Oper, 3) = "880") Then
                           inpMchHours = True
                ElseIf gLaborIn(iEntry).machHours <= 99.9 And div = pv_singapore Then
'                       (Left(gLaborIn(iEntry).Oper, 3) = "890" Or Left(gLaborIn(iEntry).Oper, 3) = "880") Then
                           inpMchHours = True
                ElseIf gLaborIn(iEntry).machHours <= 99.9 And _
                       div = bristol Then
                         inpMchHours = True
                ElseIf gLaborIn(iEntry).machHours > 23.9 And _
                       div <> bristol Then
                    'User Entered too many hours
                    ErrBox LoadResString(gLanguageOffset + 65), 2
                    ErrorBeep
                    bEntered = False
                    inpMchHours = False
                Else
                    inpMchHours = True
                End If
          End If
          
        End If
    Loop
End Function

Public Function inpSetHours(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpSetHours
'Description      : Enter the setupHours in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, CheckHrs
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim I As Integer
    Dim iChar As String
    Dim bEntered As Boolean
    
    inpSetHours = False
    strNumb = ""
    gDone = False
Do Until inpSetHours Or gDone
    Do Until inpSetHours Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpSetHours Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 217)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Len(Trim$(strwork)) = 0 Then
                strwork = "0.0"
            End If
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 1)
                If Len(strNumb) > 3 Then
                    bEntered = True
                    Exit Do
                End If
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strNumb) > 4 Then
            ErrBox LoadResString(gLanguageOffset + 92), 2
            ErrorBeep
            bEntered = False
            Exit Do
       End If
       If Not CheckHrs(strNumb) Then
            ErrBox LoadResString(gLanguageOffset + 111), 2
            ErrorBeep
            bEntered = False
            Exit Do
       Else
            gLaborIn(iEntry).setupHours = Val(strNumb)    'changed conversion from ccur to val
       End If
       If gLaborIn(iEntry).setupHours = 0 And _
             gbInpSetAndMach = False And gCompanyCd <> stroms Then
             ErrBox LoadResString(gLanguageOffset + 110), 2
             bEntered = False
             inpSetHours = False
             Exit Do
        ElseIf gLaborIn(iEntry).setupHours > 23.9 Then
              'User Entered too many hours
              ErrBox LoadResString(gLanguageOffset + 88), 2
              ErrorBeep
              bEntered = False
              inpSetHours = False
              Exit Do
        ElseIf div = stroms Then
                      If gLaborIn(iEntry).setupHours >= Elapsed_hrs(gOpenLabor.StartDateTime, gSysDateTime) Then
                          ErrBox LoadResString(gLanguageOffset + 270), 2
                          ErrorBeep
                          bEntered = False
                          inpSetHours = False
                          Exit Do
                     Else
                          inpSetHours = True
                      End If
          Else
               inpSetHours = True
          End If
    Loop
Loop
End Function
Public Function inpDDCHours(iEntry As Integer) As Boolean
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim I As Integer
    Dim iChar As String
    Dim bEntered As Boolean
    
    inpDDCHours = False
    strNumb = ""
    gDone = False
Do Until inpDDCHours Or gDone
    Do Until inpDDCHours Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpDDCHours Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = "Teclear Horas a Cargo Directo"
'            gEntry.Title = LoadResString(gLanguageOffset + 217)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Len(Trim$(strwork)) = 0 Then
                strwork = "0.0"
            End If
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 1)
                If Len(strNumb) > 3 Then
                    bEntered = True
                    Exit Do
                End If
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strNumb) > 4 Then
            ErrBox LoadResString(gLanguageOffset + 92), 2
            ErrorBeep
            bEntered = False
            Exit Do
       End If
       If Not CheckHrs(strNumb) Then
            ErrBox LoadResString(gLanguageOffset + 111), 2
            ErrorBeep
            bEntered = False
            Exit Do
       Else
            gLaborIn(iEntry).inHours = Val(strNumb)    'changed conversion from ccur to val
       End If
       If gLaborIn(iEntry).inHours = 0 And _
             gbInpSetAndMach = False And gCompanyCd <> stroms Then
             ErrBox LoadResString(gLanguageOffset + 110), 2
             bEntered = False
             inpDDCHours = False
             Exit Do
        ElseIf gLaborIn(iEntry).inHours > 23.9 Then
              'User Entered too many hours
              ErrBox LoadResString(gLanguageOffset + 88), 2
              ErrorBeep
              bEntered = False
              inpDDCHours = False
              Exit Do
          Else
               inpDDCHours = True
          End If
    Loop
Loop
End Function



Public Function inpOperNumb(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpOperNumb
'Description      : Enter the Oper in gLaborIn(iEntry) from Combo Box Lookup
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim bEntered As Boolean
    
    inpOperNumb = False
    strNumb = ""
    gDone = False
    
    Do Until inpOperNumb Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpOperNumb Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 267)
            gEntry.EntryType = 0    'Any Displayable Characters
            gEntry.LuTable = gcolOperCodes
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            'If Len(strWork) > 3 Then
                If Left(strwork, 2) = "OC" Then
                    gLaborIn(iEntry).Oper = Mid(strwork, 3, 3)
                    bEntered = True
            '    Else
            '        ErrorBeep       'Loop back again
                End If
            'Else
                gLaborIn(iEntry).Oper = strwork     'Left(strWork, 3)
                bEntered = True
            'End If
        Loop
        
        If Len(gLaborIn(iEntry).Oper) > 0 Then
            inpOperNumb = True
        Else
            bEntered = False
        End If
    Loop

End Function

Public Function inpUnplnReas(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpUnplnReas
'Description      : Enter the UnplnReas in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim bEntered As Boolean
    
    inpUnplnReas = False
    strNumb = ""
    gDone = False
    Do Until inpUnplnReas Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpUnplnReas Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 268)
            gEntry.EntryType = 0    'Any Displayable Characters
            gEntry.LuTable = gcolUnplnCodes
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            Rem removed - using combo list box for entry
            'If Len(strWork) > 2 Then
            '    ErrorBeep       'Loop back again
            'Else
                gLaborIn(iEntry).UnplnReas = Left(strwork, 2)
                bEntered = True
            'End If
        Loop
        
        If Len(gLaborIn(iEntry).UnplnReas) > 0 Then
'            If Not ValidUnplnCd(gLaborIn(iEntry).UnplnReas) Then
'                ErrBox "Unplanned reason code not valid", 2
'                ErrorBeep       'Loop back again
'            Else
                'mMaintDesc = Mid(gDivParam.ParamData, 3)
                inpUnplnReas = True
'            End If
        Else
            bEntered = False
        End If
    Loop
End Function


Public Function inpNoOfLoads(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpNoOfLoads
'Description      : Enter the NoOfLoads in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, AllDigits
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim I As Integer
    Dim iChar As String
    Dim bEntered As Boolean
    
    inpNoOfLoads = False
    strNumb = ""
    gDone = False
    
    Do Until inpNoOfLoads Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpNoOfLoads Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 218)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 3)
                If Len(strNumb) > 2 Then
                    bEntered = True
                    Exit Do
                End If
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strNumb) = 0 Then
            gLaborIn(iEntry).NoOfLoads = 0
            inpNoOfLoads = True
        Else
            If Not AllDigits(strNumb) Then
                ErrBox LoadResString(gLanguageOffset + 111), 2
                ErrorBeep
                bEntered = False
            Else
                If Len(Trim(strNumb)) < 4 Then
                    gLaborIn(iEntry).NoOfLoads = CLng(strNumb)
                    inpNoOfLoads = True
                'If gLaborIn(iEntry).NoOfLoads > 999 Then
                Else
                    ErrBox LoadResString(gLanguageOffset + 51), 2
                    ErrorBeep
                    inpNoOfLoads = False
                    bEntered = False
                End If
            End If
        End If
    Loop
End Function

Public Function inpCycleTime(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpCycleTime
'Description      : Update the CycleTime in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, AllDigits
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim I As Integer
    Dim iChar As String
    Dim bEntered As Boolean
    
    inpCycleTime = False
    strNumb = ""
    
    Do Until inpCycleTime Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpCycleTime Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 219)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 4)
                If Len(strNumb) > 3 Then
                    bEntered = True
                    Exit Do
                End If
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strNumb) = 0 Then
            gLaborIn(iEntry).CycleTime = 0
            inpCycleTime = True
        Else
            If Not AllDigits(strNumb) Then
                ErrBox LoadResString(gLanguageOffset + 111), 2
                ErrorBeep
                bEntered = False
            Else
                If Len(Trim(strNumb)) < 5 Then
                    gLaborIn(iEntry).CycleTime = CLng(strNumb)
                    inpCycleTime = True
                    If gLaborIn(iEntry).CycleTime > 2000 Then
                        ErrBox LoadResString(gLanguageOffset + 42), 2
                        ErrorBeep
                        inpCycleTime = False    'Ask again
                        bEntered = False
                    End If
                Else
                    ErrBox LoadResString(gLanguageOffset + 42), 2
                    ErrorBeep
                    inpCycleTime = False    'Ask again
                    bEntered = False
                End If
            End If
        End If
    Loop
End Function

Public Function getBatchIncn(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.getBatchIncn
'Description      : Enter the biMethod in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox
'Objects Used     : gEntry(clsEntry)
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim iBiSel As Integer
    Dim bEntered As Boolean
    Dim strwork As String
    
    Do Until getBatchIncn Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until getBatchIncn Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 220)
            gEntry.EntryType = 0    'Any Displayable Characters
            gEntry.LuTable = gcolBICodes
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            If Len(strwork) > 0 Then
                iBiSel = CInt(Left(Trim(strwork), _
                    InStr(Trim(strwork), " ") - 1))
            Else
                iBiSel = 0
            End If
            bEntered = True
        Loop
        If iBiSel <> 0 Then
            getBatchIncn = True
        Else
            ErrBox LoadResString(gLanguageOffset + 107), 2
            bEntered = False
        End If
    Loop
            
    gLaborIn(iEntry).biMethod = iBiSel

End Function


Public Function getIndirIncn(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.getIndirIncn
'Description      : Combo Box Entry of the incenRec in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox
'Objects Used     : gEntry(clsEntry), gcolIICodes(clsCodeTable)
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strIiSel As String
    Dim bEntered As Boolean
    Dim strwork As String
    Dim AnIIItem As New clsIIItem
    
    Do Until getIndirIncn Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until getIndirIncn Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 221)
            gEntry.EntryType = 0    'Any Displayable Characters
            gEntry.LuTable = gcolIICodes
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            If Len(strwork) > 0 Then
                strIiSel = Left(strwork, 30)
            Else
                strIiSel = ""
            End If
            bEntered = True
        Loop
        If strIiSel <> "" Then
            getIndirIncn = True
        Else
            ErrBox LoadResString(gLanguageOffset + 107), 2
            bEntered = False
        End If
    Loop
           
    gcolCcIICodes.getIIItem strIiSel, AnIIItem
    gLaborIn(iEntry).incenRec = AnIIItem.IncenType
    gLaborIn(iEntry).NumbMachs = AnIIItem.MachNumb
    gLaborIn(iEntry).SetStd = AnIIItem.WrkStd
    gLaborIn(iEntry).Oper = AnIIItem.Desc

    
    
    'gLaborIn(iEntry).incenRec = iBiSel

End Function


Public Function inpCCID(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpCCID
'Description      : Update the CCID in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, getCostCenter
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strCCID As String
    
    inpCCID = False
    gDone = False
    
    Do Until inpCCID Or gDone
      Do Until inpCCID Or gDone
        Set gEntry = New clsEntry
        gEntry.Title = LoadResString(gLanguageOffset + 147)
        gEntry.EntryType = 0    'Any Displayable Characters
        strCCID = UCase(basGetEntry())
        If gEntry.Escaped Then
            Set gEntry = Nothing
            gDone = True
            Exit Function
        End If
        Set gEntry = Nothing
        
        If Len(strCCID) = 0 Then
            ErrBox LoadResString(gLanguageOffset + 78), 2
            ErrorBeep
            Exit Do   'go Back and Ask Again
        End If
        
        If Left(strCCID, 2) = "ZZ" Then
            gDone = True
            Exit Function
        End If
        
        If Len(strCCID) > 5 Then
            If Left(strCCID, 2) = "CC" Then
                gLaborIn(iEntry).CCID = Mid(strCCID, 3, 5)
            Else
                ErrorBeep
            End If
        Else
            gLaborIn(iEntry).CCID = Left(strCCID, 5)
        End If
        
        
        If Not GetCostCenter(gLaborIn(iEntry).CCID) Then
            ErrBox LoadResString(gLanguageOffset + 41), 2
            ErrorBeep
        Else
            If gLaborIn(iEntry).CCID = "----" _
                    Or gLaborIn(iEntry).CCID = "-----" Then
                ErrBox LoadResString(gLanguageOffset + 41), 2
                ErrorBeep
            Else
                inpCCID = True  'Got one - Exit loop and function
            End If
        End If
      Loop
    Loop

End Function



Public Function inpEOJob(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpEOJob
'Description      : Update the OrderNo in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strEOJob As String
    
    inpEOJob = False
    gDone = False
    
    Do Until inpEOJob Or gDone
      Do Until inpEOJob Or gDone
        Set gEntry = New clsEntry
        gEntry.Title = LoadResString(gLanguageOffset + 222)
        gEntry.EntryType = 0    'Any Displayable Characters
        strEOJob = UCase(basGetEntry())
        If gEntry.Escaped Then
            Set gEntry = Nothing
            gDone = True
            Exit Function
        End If
        Set gEntry = Nothing
        
        If Len(strEOJob) = 0 Then
            ErrBox LoadResString(gLanguageOffset + 78), 2
            ErrorBeep
            Exit Do   'go Back and Ask Again must make an entry
        End If
        
        If Left(strEOJob, 2) = "ZZ" Then
            gDone = True
            Exit Function
        End If
        
        If Len(strEOJob) > 7 Then
            If Left(strEOJob, 2) = "EO" Then
                gLaborIn(iEntry).OrderNo = Mid(strEOJob, 3, 7)
            Else
                ErrBox LoadResString(gLanguageOffset + 15), 2
                ErrorBeep
            End If
        Else
            gLaborIn(iEntry).OrderNo = Left(strEOJob, 7)
            inpEOJob = True  'Got one - Exit loop and function
        End If
       Loop
     Loop
End Function

Public Function inpPMJob(iEntry As Integer) As Boolean
    Dim strPMJob As String
    Dim lPMJob As String
    Dim Start_Pos, x, Y As Integer
    inpPMJob = False
    
    gDone = False
    Do Until inpPMJob Or gDone
      Do Until inpPMJob Or gDone
    
        Set gEntry = New clsEntry
        gEntry.Title = LoadResString(gLanguageOffset + 223)
        gEntry.EntryType = 0    'Any Displayable Characters
        strPMJob = UCase(basGetEntry())
        If gEntry.Escaped Then
            Set gEntry = Nothing
            gDone = True
            Exit Function
        End If
        Set gEntry = Nothing
        
        If Len(strPMJob) = 0 Then
            ErrBox LoadResString(gLanguageOffset + 78), 2
            ErrorBeep
            Exit Do
        End If
        
        If Left(strPMJob, 2) = "ZZ" Then
            gDone = True
            Exit Function
        End If
        
        If Len(strPMJob) > 7 Then
            If Left(strPMJob, 2) = "PM" Then
                gLaborIn(iEntry).OrderNo = Mid(strPMJob, 3, 7)
            Else
                ErrBox LoadResString(gLanguageOffset + 62), 2
                ErrorBeep
            End If
        Else
            Start_Pos = Len(strPMJob)
            Select Case Start_Pos
             Case 1
               gLaborIn(iEntry).OrderNo = "000000" & strPMJob
             Case 2
               gLaborIn(iEntry).OrderNo = "00000" & strPMJob
             Case 3
                    If IsNumeric(strPMJob) Then
                           gLaborIn(iEntry).OrderNo = "DEPT" & strPMJob
                    Else
                           gLaborIn(iEntry).OrderNo = "0000" & strPMJob
                    End If
             Case 4
               gLaborIn(iEntry).OrderNo = "000" & strPMJob
             Case 5
               gLaborIn(iEntry).OrderNo = "00" & strPMJob
             Case 6
               gLaborIn(iEntry).OrderNo = "0" & strPMJob
             Case Else
               gLaborIn(iEntry).OrderNo = strPMJob
            End Select
            inpPMJob = True  'Got one - Exit loop and function
        End If
       Loop
     Loop
End Function

Public Function inpIndirCd(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpIndirCd
'Description      : Select the IndirCd in gLaborIn(iEntry) from combo box
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim bEntered As Boolean
    
    inpIndirCd = False
    strNumb = ""
    gSelDesc = ""
    gDone = False
    Do Until inpIndirCd Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpIndirCd Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 148)
            gEntry.EntryType = 0    'Any Displayable Characters
            If UCase$(Trim$(Environ("OS"))) = "WINDOWS_NT" Then
                'If Windows NT Then we want to Use the Combo Box.
                'Windows 3.1 Combo Box does not behave the same way as NT.
                gEntry.LuTable = gcolIndirCodes
            End If
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                
                Exit Function
            End If
            
            gLaborIn(iEntry).IndirCd = Left(strwork, 2)
            gSelDesc = Mid(strwork, 3, 23)
            bEntered = True
        Loop
        
        If Len(gLaborIn(iEntry).IndirCd) > 0 Then
            If Not ValidIndirCd(gLaborIn(iEntry).IndirCd) Then
                ErrBox LoadResString(gLanguageOffset + 55), 2
                ErrorBeep       'Loop back again
            Else
                gSelDesc = gcolIndirCodes.getDesc(gLaborIn(iEntry).IndirCd)
                inpIndirCd = True
            End If
        Else
            bEntered = False
        End If
    Loop
End Function

Public Function inpTlmntCd(iEntry As Integer) As Boolean
    Dim strwork As String
    Dim strNumb As String
    Dim bEntered As Boolean
    
    inpTlmntCd = False
    strNumb = ""
    gDone = False
    Do Until inpTlmntCd Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpTlmntCd Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 269)
            gEntry.EntryType = 0    'Any Displayable Characters
            gEntry.LuTable = gcolTlmntCodes
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            gLaborIn(iEntry).UnplnReas = Left(strwork, 2)
            bEntered = True
        Loop
        
        If Len(gLaborIn(iEntry).UnplnReas) > 0 Then
   '         If Not ValidTlmntCd(gLaborIn(iEntry).UnplnReas) Then
   '             ErrBox "Maint Reason code not valid", 2
   '             ErrorBeep       'Loop back again
   '         Else
                inpTlmntCd = True
   '         End If
        Else
            bEntered = False
        End If
    Loop
End Function





Public Function inpComment(iEntry As Integer) As Boolean
    Dim strComment As String
    
    inpComment = False
    
    Do Until inpComment Or gDone
      Do Until inpComment Or gDone
        Set gEntry = New clsEntry
        gEntry.Title = LoadResString(gLanguageOffset + 224)
        gEntry.EntryType = 0    'Any Displayable Characters
        strComment = UCase(basGetEntry())
        If gEntry.Escaped Then
            Set gEntry = Nothing
            gDone = True
            Exit Function
        End If
        Set gEntry = Nothing
        
        If Len(strComment) = 0 Then
            ErrBox LoadResString(gLanguageOffset + 68), 2
            ErrorBeep
            Exit Do   'go Back and Ask Again
        ElseIf Left(strComment, 2) = "ZZ" Then
            gDone = True
            Exit Function
        ElseIf Len(strComment) > 80 Then
            ErrBox LoadResString(gLanguageOffset + 91), 2
            ErrorBeep
            Exit Do   'go Back and Ask Again
        Else
           mComment = strComment
           inpComment = True
        End If
   
      Loop
    Loop
End Function



Public Function inpFixStd(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpFixStd
'Description      : Enter the SetStd in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, CheckHrs
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim I As Integer
    Dim iChar As String
    Dim bEntered As Boolean
    
    inpFixStd = False
    strNumb = ""
    gDone = False
    
    Do Until inpFixStd Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpFixStd Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 225)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 1)
                If Len(strNumb) > 4 Then
                    bEntered = True
                    Exit Do
                End If
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strNumb) > 4 Then
            ErrorBeep
            bEntered = False
        Else
            If Len(strNumb) = 0 Then
                gLaborIn(iEntry).SetStd = 0
                inpFixStd = True
            Else
                If Not CheckHrs(strNumb) Then
                   ErrBox LoadResString(gLanguageOffset + 111), 2
                   ErrorBeep
                   bEntered = False
                Else
                    gLaborIn(iEntry).SetStd = CCur(strNumb)
                    inpFixStd = True
                End If
            End If
        End If
    Loop
End Function

Public Function inpMchStd(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpMchStd
'Description      : Enter the MchStd in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, AllDigits
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strwork As String
    Dim strNumb As String
    Dim iPosit As Integer
    Dim I As Integer
    Dim iChar As String
    Dim bEntered As Boolean
    
    inpMchStd = False
    strNumb = ""
    gDone = False
    
    Do Until inpMchStd Or gDone
        bEntered = False
        'Loop through numbers entered (zz9)
        Do Until inpMchStd Or gDone Or bEntered
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 226)
            gEntry.EntryType = 0    'Any Displayable Characters
            strwork = UCase(basGetEntry())
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            If Left(strwork, 2) = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If Left(strwork, 2) = "NU" Then
                If Mid(strwork, 3, 1) = "Z" Then
                    bEntered = True
                    Exit Do
                End If
                strNumb = strNumb & Mid(strwork, 3, 6)
                If Len(strNumb) > 5 Then
                    bEntered = True
                    Exit Do
                End If
            Else
                strNumb = strwork
                bEntered = True
            End If
        Loop
        
        If Len(strNumb) = 0 Then
            gLaborIn(iEntry).MchStd = 0
            inpMchStd = True
        Else
            If Not AllDigits(strNumb) Then
                ErrBox LoadResString(gLanguageOffset + 111), 2
                ErrorBeep
                bEntered = False
            Else
                If Len(Trim(strNumb)) < 7 Then
                    gLaborIn(iEntry).MchStd = CLng(strNumb)
                    inpMchStd = True
                'If gLaborIn(iEntry).MchStd > 999999 Then
                Else
                    ErrBox LoadResString(gLanguageOffset + 52), 2
                    ErrorBeep
                    inpMchStd = False    'Ask again
                    bEntered = False
                End If
            End If
        End If
    Loop
End Function

Public Function inpJobAndSeqStds(iEntry As Integer) As Boolean
    Dim strEntry As String
    Dim strDataId As String
    Dim strSeqNo As String
    
    inpJobAndSeqStds = False        'to force loop back till done
    
    Do Until inpJobAndSeqStds Or gDone
        Do Until inpJobAndSeqStds Or gDone       'Extra loop for exit do stmts
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 165)
            gEntry.EntryType = 0        'Any Printable Characters
            strEntry = basGetEntry()
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            strDataId = UCase(Left(strEntry, 2))
            If strDataId = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If strDataId = "NU" Then
                ErrBox LoadResString(gLanguageOffset + 28), 2
                ErrorBeep
                Exit Do           'Loop Again
            End If
            
            If Len(strDataId) < 1 Then
                ErrBox LoadResString(gLanguageOffset + 78), 2
                ErrorBeep
                Exit Do           'Loop Again
            End If
                
            If Len(strEntry) > 11 Then
                If strDataId = "JS" Then
                    gLaborIn(iEntry).OrderNo = UCase(Mid(strEntry, 3, 7))
                    strSeqNo = Mid(strEntry, 10, 4)
                Else
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            Else
                gLaborIn(iEntry).OrderNo = UCase(Left(strEntry, 7))
                strSeqNo = Mid(strEntry, 8, 4)
            End If
            
            If Not GetJobMaster(gLaborIn(iEntry).OrderNo) Then
                If Left(gLaborIn(iEntry).OrderNo, 2) = "09" Then
                    Screen.ActiveForm.lblErrMsg.Caption = LoadResString(gLanguageOffset + 236)
                    Screen.ActiveForm.lblErrMsg.BackColor = vbColorWhite
                    strSeqNo = "000 "
                    gLaborIn(iEntry).MatlWgt = 0
                    gJobMaster.Actdate = 0
                    gJobMaster.OrderNo = gLaborIn(iEntry).OrderNo
                    gJobMaster.OrderType = "SO"
                Else
                    ErrBox LoadResString(gLanguageOffset + 61), 2
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            Else
                gLaborIn(iEntry).MatlWgt = gJobMaster.MatlWgt
            End If
            
            If Len(strSeqNo) = 0 Then
                ErrBox LoadResString(gLanguageOffset + 69), 2
                ErrorBeep
                Exit Do
            Else
                If Not ValidNumber(strSeqNo) Then
                    ErrBox LoadResString(gLanguageOffset + 111), 2
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            End If
            
            gLaborIn(iEntry).SeqNo = CInt(strSeqNo)
            If gLaborIn(iEntry).SeqNo = 990 _
                Or gLaborIn(iEntry).SeqNo = 0 Then
                  ErrBox LoadResString(gLanguageOffset + 97), 2
                  ErrorBeep
                  Exit Do           'Loop Again
            Else
                If Not GetJobOper(gLaborIn(iEntry).OrderNo, _
                        gLaborIn(iEntry).SeqNo) Then
                    ErrBox LoadResString(gLanguageOffset + 87), 2
                    ErrorBeep
                    Exit Do
                Else
                    If gJobOpers.WCID = "-----" Then
                        ErrBox LoadResString(gLanguageOffset + 38), 2
                        ErrorBeep
                        Exit Do
                    End If
                    If Not GetWorkCenter(gJobOpers.WCID) Then
                        ErrBox LoadResString(gLanguageOffset + 103), 2
                        ErrorBeep
                        Exit Do
                    End If
                    gLaborIn(iEntry).UnplnReas = "  "
                    gLaborIn(iEntry).WCID = gJobOpers.WCID
                    gLaborIn(iEntry).Oper = gJobOpers.Oper
                    gLaborIn(iEntry).EOrS = gJobOpers.EorSStd
                    gLaborIn(iEntry).SetStd = gJobOpers.SetStd
                    gLaborIn(iEntry).MchStd = gJobOpers.MchStd  'labor.run_std????
                    gLaborIn(iEntry).MchStd = gJobOpers.MchStd
                    inpJobAndSeqStds = True     'Done - Planned
                End If
            End If
        Loop
    Loop
    Screen.ActiveForm.lblErrMsg.Caption = ""
End Function



Public Function inpEmplNo(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Inputs.inpEmplNo
'Description      : Update the EmplNo in gLaborIn(iEntry)
'Parameters       : index of record/machine being entered
'Procedures Used  : basGetEntry, ErrBox, ErrorBeep, getEmployee,
'                   Ck_for_Attd_Rec
'Objects Used     : gEntry
'Forms Loaded     : None
'Return Parameters: True if didn't Escape
'Tables Updates   : None
'==============================================================================
    Dim strEmpl As String
    
    inpEmplNo = False
    
  Do Until inpEmplNo Or gDone
     Do Until inpEmplNo Or gDone
        Set gEntry = New clsEntry
        gEntry.Title = LoadResString(gLanguageOffset + 127)
        gEntry.EntryType = 1    'Any Displayable Characters
        strEmpl = basGetEntry()
        If gEntry.Escaped Then
            Set gEntry = Nothing
            gDone = True
            Exit Function
        End If
        Set gEntry = Nothing
        
        If Len(strEmpl) = 0 Then
            ErrBox LoadResString(gLanguageOffset + 78), 2
            ErrorBeep
            Exit Do   'go Back and Ask Again
        End If
        
        If Left(strEmpl, 2) = "ZZ" Then
            gDone = True
            Exit Function
        End If
        
        If Len(strEmpl) > 6 Then
            If Left(strEmpl, 2) = "AT" Or _
               Left(strEmpl, 2) = "LB" Then
                gLaborIn(iEntry).EmplNo = CLng(Mid(strEmpl, 3, 6))
            Else
                ErrorBeep
            End If
        Else
            gLaborIn(iEntry).EmplNo = CLng(Left(Trim(strEmpl), 6))
        End If
        
        If Not GetEmployee(gLaborIn(iEntry).EmplNo) Then
            ErrBox LoadResString(gLanguageOffset + 44), 2
            ErrorBeep
            Exit Do
        End If
        If gEmployee.AttdStatus = 2 Or _
             gEmployee.AttdStatus = 3 Or _
             gEmployee.AttdStatus = 5 Or _
             gEmployee.AttdStatus = 8 Then
                ErrBox LoadResString(gLanguageOffset + 76), 2
                ErrorBeep
                Exit Do
        End If
        If Not Ck_for_Attd_Rec(gEmployee.EmplNo, gEmployee.Actdate) Then
             ErrBox LoadResString(gLanguageOffset + 76), 2
             ErrorBeep
             Exit Do
        End If
        inpEmplNo = True  'Got one - Exit loop and function
      Loop
    Loop
End Function

