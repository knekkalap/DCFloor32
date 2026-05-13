Attribute VB_Name = "ATTENDNC"
Option Explicit
Public Function Clock_out_rnd() As Currency
  Dim endTimeSecs, ComparTime, MinStartTimeSecs, MaxStartTimeSecs As Long
  Dim shiftNum, shiftGrp As Integer
  Dim lMM, lHH, lSS As Integer
  Dim ltimeMMSS, ltimeDec As Currency
  Dim lTimeChar, tempTimeChar As String
  
  shiftNum = gEmployee.GenShift
  shiftGrp = gEmployee.shiftGrp
  
  endTimeSecs = ConvTimeToSecs(gShiftRecData(shiftNum).endTime)
  If empClkOutOK Then
     lMM = Mid(gShiftRecData(shiftNum).endTime, 3, 2)
     lHH = Mid(gShiftRecData(shiftNum).endTime, 1, 2)
     lSS = Mid(gShiftRecData(shiftNum).endTime, 5, 2)
  Else
     lMM = CInt(Mid(gSysTimeStrp, 3, 2))
     lHH = CInt(Mid(gSysTimeStrp, 1, 2))
     lSS = CInt(Mid(gSysTimeStrp, 5, 2))
     ComparTime = ConvTimeToSecs(gSysTimeStrp)
  End If
  
If HalfHourRounding Then
   ltimeMMSS = lMM + (lSS * 0.01)
    If ltimeMMSS < (0 + gCOGraceAft) Then
        ltimeDec = lHH + 0
        tempTimeChar = Format(ltimeDec, "00.00")
        lTimeChar = Left(tempTimeChar, 2) & "0000"
    ElseIf (ltimeMMSS >= (30 - gCOGraceBef) And _
        ltimeMMSS < (30 + gCOGraceAft)) Then
         ltimeDec = lHH + 0.5
         tempTimeChar = Format(ltimeDec, "00.00")
         lTimeChar = Left(tempTimeChar, 2) & "3000"
    ElseIf (ltimeMMSS >= (60 - gCOGraceBef)) Then
         ltimeDec = lHH + 1#
         tempTimeChar = Format(ltimeDec, "00.00")
         lTimeChar = Left(tempTimeChar, 2) & "0000"
    Else
         ltimeDec = RndTimeHnds(gSysTimeStrp)
         lTimeChar = gSysTimeStrp
     End If
 ElseIf QtrHourRounding Then
 'Currently only Hyson Products uses Qtr hour rounding - mwp
      ltimeMMSS = lMM + (lSS * 0.01)
      If ltimeMMSS < (0 + gCOGraceAft) Then
          ltimeDec = lHH + 0
          tempTimeChar = Format(ltimeDec, "00.00")
          lTimeChar = Left(tempTimeChar, 2) & "0000"
      ElseIf (ltimeMMSS >= (15 - gCOGraceBef) And _
          ltimeMMSS < (15 + gCOGraceAft)) Then
         ' ltimeDec = lHH + 0.25
          ltimeDec = lHH + 0.3
          tempTimeChar = Format(ltimeDec, "00.00")
          lTimeChar = Left(tempTimeChar, 2) & "1500"
      ElseIf (ltimeMMSS >= (30 - gCOGraceBef) And _
          ltimeMMSS < (30 + gCOGraceAft)) Then
          ltimeDec = lHH + 0.5
          tempTimeChar = Format(ltimeDec, "00.00")
          lTimeChar = Left(tempTimeChar, 2) & "3000"
      ElseIf (ltimeMMSS >= (45 - gCOGraceBef) And _
          ltimeMMSS < (45 + gCOGraceAft)) Then
        '  ltimeDec = lHH + 0.75
          ltimeDec = lHH + 0.8
          tempTimeChar = Format(ltimeDec, "00.00")
          lTimeChar = Left(tempTimeChar, 2) & "4500"
      ElseIf (ltimeMMSS >= (60 - gCOGraceBef)) Then
          ltimeDec = lHH + 1#
          tempTimeChar = Format(ltimeDec, "00.00")
          lTimeChar = Left(tempTimeChar, 2) & "0000"
      Else
         ltimeDec = RndTimeHnds(gSysTimeStrp)
         lTimeChar = gSysTimeStrp
     End If
 End If
     
     If ltimeDec > 24 Then
        ltimeDec = 0#
        lTimeChar = "000000"
     End If
  
  MinStartTimeSecs = endTimeSecs - (gCOGraceBef * 60)
  If MinStartTimeSecs < 0 Then
      MinStartTimeSecs = 86400 + MinStartTimeSecs
  End If
  MaxStartTimeSecs = endTimeSecs + (gCOGraceAft * 60)
  If MaxStartTimeSecs > 86400 Then
      MaxStartTimeSecs = MaxStartTimeSecs - 86400
  End If
  
  If ComparTime > MaxStartTimeSecs Then
      Call RecordError("P")
  ElseIf ComparTime < MinStartTimeSecs Then
           Call RecordError("O")
  End If
  
  Clock_out_rnd = ltimeDec
End Function




Public Function Alt_Clock_out_rnd() As Currency
  Dim endTimeSecs, ComparTime, MinStartTimeSecs, MaxStartTimeSecs As Long
  Dim shiftNum, shiftGrp As Integer
  Dim lMM, lHH, lSS As Integer
  Dim ltimeMMSS, ltimeDec As Currency
  Dim lTimeChar, tempTimeChar As String
  
  shiftNum = gEmployee.GenShift
  shiftGrp = gEmployee.shiftGrp
  
  endTimeSecs = ConvTimeToSecs(gShiftRecData(shiftNum).endTime)
  
'  ltimeMMSS = lMM + (lSS * 0.01)
'  If ltimeMMSS < (0 + gCOGraceAft) Then
'      ltimeDec = lHH + 0
'      tempTimeChar = Format(ltimeDec, "00.00")
'      lTimeChar = Left(tempTimeChar, 2) & "0000"
'  ElseIf (ltimeMMSS >= (30 - gCOGraceBef) And _
'         ltimeMMSS < (30 + gCOGraceAft)) Then
'            ltimeDec = lHH + 0.5
'            tempTimeChar = Format(ltimeDec, "00.00")
'            lTimeChar = Left(tempTimeChar, 2) & "3000"
'  ElseIf (ltimeMMSS >= (60 - gCOGraceBef)) Then
'           ltimeDec = lHH + 1#
'           tempTimeChar = Format(ltimeDec, "00.00")
'           lTimeChar = Left(tempTimeChar, 2) & "0000"
'  Else
'      ltimeDec = RndTimeHnds(gSysTimeStrp)
'      lTimeChar = gSysTimeStrp
'  End If
'  If ltimeDec > 24 Then
'     ltimeDec = 0#
'     lTimeChar = "000000"
'  End If
  
  MinStartTimeSecs = endTimeSecs - (gCOGraceBef * 60)
  If MinStartTimeSecs < 0 Then
      MinStartTimeSecs = 86400 + MinStartTimeSecs
  End If
  MaxStartTimeSecs = endTimeSecs + (gCOGraceAft * 60)
  If MaxStartTimeSecs > 86400 Then
      MaxStartTimeSecs = MaxStartTimeSecs - 86400
  End If
  
  ComparTime = ConvTimeToSecs(gSysTimeStrp)
  
  If ComparTime >= MinStartTimeSecs And ComparTime <= MaxStartTimeSecs Then
        ltimeDec = RndTimeHnds(gShiftRecData(shiftNum).endTime)
  ElseIf ComparTime > MaxStartTimeSecs Then
         ltimeDec = RndTimeHnds(gSysTimeStrp)
         Call RecordError("P")
  ElseIf ComparTime < MinStartTimeSecs Then
         ltimeDec = RndTimeHnds(gSysTimeStrp)
         Call RecordError("O")
  End If

  Alt_Clock_out_rnd = ltimeDec
End Function

Public Function Alt_Clock_in_rnd() As Currency
  Dim startTimeSecs, ComparTime, MinStartTimeSecs, MaxStartTimeSecs As Long
  Dim shiftNum, shiftGrp As Integer
  Dim lMM, lHH, lSS As Integer
  Dim ltimeMMSS, ltimeDec As Currency
  Dim lTimeChar, tempTimeChar As String
  
  shiftNum = AttdShift
  shiftGrp = gEmployee.shiftGrp
  
  'lMM = CInt(Mid(gSysTimeStrp, 3, 2))
  'lHH = CInt(Mid(gSysTimeStrp, 1, 2))
  'lSS = CInt(Mid(gSysTimeStrp, 5, 2))
 
 ' ltimeMMSS = lMM + (lSS * 0.01)
 ' If ltimeMMSS < (0 + gCIGraceAft) Then
 '     ltimeDec = lHH + 0
 '     tempTimeChar = Format(ltimeDec, "00.00")
 '     lTimeChar = Left(tempTimeChar, 2) & "0000"
 ' ElseIf (ltimeMMSS >= (30 - gCIGraceBef) And _
 '        ltimeMMSS < (30 + gCIGraceAft)) Then
 '           ltimeDec = lHH + 0.5
 '           tempTimeChar = Format(ltimeDec, "00.00")
 '           lTimeChar = Left(tempTimeChar, 2) & "3000"
 ' ElseIf (ltimeMMSS >= (60 - gCIGraceBef)) Then
 '          ltimeDec = lHH + 1#
 '          tempTimeChar = Format(ltimeDec, "00.00")
 '          lTimeChar = Left(tempTimeChar, 2) & "0000"
 ' Else
 '     ltimeDec = RndTimeHnds(gSysTimeStrp)
 '     lTimeChar = gSysTimeStrp
 ' End If
 ' If ltimeDec > 24 Then
 '    ltimeDec = 0#
 '    lTimeChar = "000000"
 ' End If
  
  startTimeSecs = ConvTimeToSecs(gShiftRecData(shiftNum).startTime)
  MinStartTimeSecs = startTimeSecs - (gCIGraceBef * 60)
  If MinStartTimeSecs < 0 Then
      MinStartTimeSecs = 86400 + MinStartTimeSecs
  End If
  MaxStartTimeSecs = startTimeSecs + (gCIGraceAft * 60)
  If MaxStartTimeSecs > 86400 Then
      MaxStartTimeSecs = MaxStartTimeSecs - 86400
  End If
  
  ComparTime = ConvTimeToSecs(gSysTimeStrp)
  
'  uEmployee.LaborStartDate = CDate(gSysDateStr & " " & _
'                                   Left(lTimeChar, 2) & ":" & _
'                                   Mid(lTimeChar, 3, 2) & ":" & _
'                                   Mid(lTimeChar, 5, 2))
  If ComparTime >= MinStartTimeSecs And ComparTime <= MaxStartTimeSecs Then
       ltimeDec = RndTimeHnds(gShiftRecData(shiftNum).startTime)
  ElseIf ComparTime > MaxStartTimeSecs Then
         Call RecordError("I")
         ltimeDec = RndTimeHnds(gSysTimeStrp)
  ElseIf ComparTime < MinStartTimeSecs Then
           Call RecordError("J")
           ltimeDec = RndTimeHnds(gSysTimeStrp)
  End If
  
  Alt_Clock_in_rnd = ltimeDec
End Function

Public Function Burl_Clock_in_rnd() As Currency
  Dim startTimeSecs, ComparTime, MinStartTimeSecs, MaxStartTimeSecs As Long
  Dim shiftNum, shiftGrp As Integer
  Dim lMM, lHH, lSS As Integer
  Dim ltimeMMSS, ltimeDec As Currency
  Dim lTimeChar, tempTimeChar As String
  Dim booDateIncr As Boolean
  
  shiftNum = AttdShift
  shiftGrp = gEmployee.shiftGrp
  
  lMM = CInt(Mid(gSysTimeStrp, 3, 2))
  lHH = CInt(Mid(gSysTimeStrp, 1, 2))
  lSS = CInt(Mid(gSysTimeStrp, 5, 2))
If HalfHourRounding Then
   ltimeMMSS = lMM + (lSS * 0.01)
   If ltimeMMSS < (0 + gCIGraceAft) Then
      ltimeDec = lHH + 0
      tempTimeChar = Format(ltimeDec, "00.00")
      lTimeChar = Left(tempTimeChar, 2) & "0000"
   ElseIf (ltimeMMSS >= (30 - gCIGraceBef) And _
         ltimeMMSS < (30 + gCIGraceAft)) Then
            ltimeDec = lHH + 0.5
            tempTimeChar = Format(ltimeDec, "00.00")
            lTimeChar = Left(tempTimeChar, 2) & "3000"
   ElseIf (ltimeMMSS >= (60 - gCIGraceBef)) Then
           ltimeDec = lHH + 1#
           tempTimeChar = Format(ltimeDec, "00.00")
           lTimeChar = Left(tempTimeChar, 2) & "0000"
   Else
       ltimeDec = RndTimeHnds(gSysTimeStrp)
       lTimeChar = gSysTimeStrp
   End If
 ElseIf QtrHourRounding Then
 'Currently only Hyson Products uses Qtr hour rounding - mwp
   ltimeMMSS = lMM + (lSS * 0.01)
   If ltimeMMSS < (0 + gCIGraceAft) Then
      ltimeDec = lHH + 0
      tempTimeChar = Format(ltimeDec, "00.00")
      lTimeChar = Left(tempTimeChar, 2) & "0000"
   ElseIf (ltimeMMSS >= (15 - gCIGraceBef) And _
         ltimeMMSS < (15 + gCIGraceAft)) Then
           ' ltimeDec = lHH + 0.25
            ltimeDec = lHH + 0.3
            tempTimeChar = Format(ltimeDec, "00.00")
            lTimeChar = Left(tempTimeChar, 2) & "1500"
   ElseIf (ltimeMMSS >= (30 - gCIGraceBef) And _
         ltimeMMSS < (30 + gCIGraceAft)) Then
            ltimeDec = lHH + 0.5
            tempTimeChar = Format(ltimeDec, "00.00")
            lTimeChar = Left(tempTimeChar, 2) & "3000"
   ElseIf (ltimeMMSS >= (45 - gCIGraceBef) And _
         ltimeMMSS < (45 + gCIGraceAft)) Then
          '  ltimeDec = lHH + 0.75
            ltimeDec = lHH + 0.8
            tempTimeChar = Format(ltimeDec, "00.00")
            lTimeChar = Left(tempTimeChar, 2) & "4500"
   ElseIf (ltimeMMSS >= (60 - gCIGraceBef)) Then
           ltimeDec = lHH + 1#
           tempTimeChar = Format(ltimeDec, "00.00")
           lTimeChar = Left(tempTimeChar, 2) & "0000"
   Else
       ltimeDec = RndTimeHnds(gSysTimeStrp)
       lTimeChar = gSysTimeStrp
   End If
End If
 
  booDateIncr = False
  If ltimeDec >= 24 Then
    booDateIncr = True
     ltimeDec = 0#
     lTimeChar = "000000"
  End If
  
  startTimeSecs = ConvTimeToSecs(gShiftRecData(shiftNum).startTime)
  MinStartTimeSecs = startTimeSecs - (gCIGraceBef * 60)
  If MinStartTimeSecs < 0 Then
      MinStartTimeSecs = 86400 + MinStartTimeSecs
  End If
  MaxStartTimeSecs = startTimeSecs + (gCIGraceAft * 60)
  If MaxStartTimeSecs > 86400 Then
      MaxStartTimeSecs = MaxStartTimeSecs - 86400
  End If
  
  ComparTime = ConvTimeToSecs(gSysTimeStrp)
  
  uEmployee.LaborStartDate = CDate(gSysDateStr & " " & _
                                   Left(lTimeChar, 2) & ":" & _
                                   Mid(lTimeChar, 3, 2) & ":" & _
                                   Mid(lTimeChar, 5, 2))

'in case the rounded time above is set to midnight, the date must be incremented
' by one
 
 If booDateIncr Then
     uEmployee.LaborStartDate = DateAdd("d", 1, uEmployee.LaborStartDate)
End If
'If clock in time earlier than designated shift start time, use half-hour rounding
'rules, otherwise, re-round to tenth of an hour using special rules - mwp 10.8.03
  If ComparTime > MaxStartTimeSecs Then
       ltimeDec = BurlRndTimeHnds(gSysTimeStrp)
      Call RecordError("I")
  ElseIf ComparTime < MinStartTimeSecs Then
           Call RecordError("J")
  End If
  
  Burl_Clock_in_rnd = ltimeDec
End Function




Public Function Clock_in_rnd() As Currency
  Dim startTimeSecs, ComparTime, MinStartTimeSecs, MaxStartTimeSecs As Long
  Dim shiftNum, shiftGrp As Integer
  Dim lMM, lHH, lSS As Integer
  Dim ltimeMMSS, ltimeDec As Currency
  Dim lTimeChar, tempTimeChar As String
  Dim booDateIncr As Boolean
  
  shiftNum = AttdShift
  shiftGrp = gEmployee.shiftGrp
  
  lMM = CInt(Mid(gSysTimeStrp, 3, 2))
  lHH = CInt(Mid(gSysTimeStrp, 1, 2))
  lSS = CInt(Mid(gSysTimeStrp, 5, 2))
If HalfHourRounding Then
   ltimeMMSS = lMM + (lSS * 0.01)
   If ltimeMMSS < (0 + gCIGraceAft) Then
      ltimeDec = lHH + 0
      tempTimeChar = Format(ltimeDec, "00.00")
      lTimeChar = Left(tempTimeChar, 2) & "0000"
   ElseIf (ltimeMMSS >= (30 - gCIGraceBef) And _
         ltimeMMSS < (30 + gCIGraceAft)) Then
            ltimeDec = lHH + 0.5
            tempTimeChar = Format(ltimeDec, "00.00")
            lTimeChar = Left(tempTimeChar, 2) & "3000"
   ElseIf (ltimeMMSS >= (60 - gCIGraceBef)) Then
           ltimeDec = lHH + 1#
           tempTimeChar = Format(ltimeDec, "00.00")
           lTimeChar = Left(tempTimeChar, 2) & "0000"
   Else
       ltimeDec = RndTimeHnds(gSysTimeStrp)
       lTimeChar = gSysTimeStrp
   End If
 ElseIf QtrHourRounding Then
 'Currently only Hyson Products uses Qtr hour rounding - mwp
   ltimeMMSS = lMM + (lSS * 0.01)
   If ltimeMMSS < (0 + gCIGraceAft) Then
      ltimeDec = lHH + 0
      tempTimeChar = Format(ltimeDec, "00.00")
      lTimeChar = Left(tempTimeChar, 2) & "0000"
   ElseIf (ltimeMMSS >= (15 - gCIGraceBef) And _
         ltimeMMSS < (15 + gCIGraceAft)) Then
           ' ltimeDec = lHH + 0.25
            ltimeDec = lHH + 0.3
            tempTimeChar = Format(ltimeDec, "00.00")
            lTimeChar = Left(tempTimeChar, 2) & "1500"
   ElseIf (ltimeMMSS >= (30 - gCIGraceBef) And _
         ltimeMMSS < (30 + gCIGraceAft)) Then
            ltimeDec = lHH + 0.5
            tempTimeChar = Format(ltimeDec, "00.00")
            lTimeChar = Left(tempTimeChar, 2) & "3000"
   ElseIf (ltimeMMSS >= (45 - gCIGraceBef) And _
         ltimeMMSS < (45 + gCIGraceAft)) Then
          '  ltimeDec = lHH + 0.75
            ltimeDec = lHH + 0.8
            tempTimeChar = Format(ltimeDec, "00.00")
            lTimeChar = Left(tempTimeChar, 2) & "4500"
   ElseIf (ltimeMMSS >= (60 - gCIGraceBef)) Then
           ltimeDec = lHH + 1#
           tempTimeChar = Format(ltimeDec, "00.00")
           lTimeChar = Left(tempTimeChar, 2) & "0000"
   Else
       ltimeDec = RndTimeHnds(gSysTimeStrp)
       lTimeChar = gSysTimeStrp
   End If
End If
 
  booDateIncr = False
  If ltimeDec >= 24 Then
    booDateIncr = True
     ltimeDec = 0#
     lTimeChar = "000000"
  End If
  
  startTimeSecs = ConvTimeToSecs(gShiftRecData(shiftNum).startTime)
  MinStartTimeSecs = startTimeSecs - (gCIGraceBef * 60)
  If MinStartTimeSecs < 0 Then
      MinStartTimeSecs = 86400 + MinStartTimeSecs
  End If
  MaxStartTimeSecs = startTimeSecs + (gCIGraceAft * 60)
  If MaxStartTimeSecs > 86400 Then
      MaxStartTimeSecs = MaxStartTimeSecs - 86400
  End If
  
  ComparTime = ConvTimeToSecs(gSysTimeStrp)
  
  uEmployee.LaborStartDate = CDate(gSysDateStr & " " & _
                                   Left(lTimeChar, 2) & ":" & _
                                   Mid(lTimeChar, 3, 2) & ":" & _
                                   Mid(lTimeChar, 5, 2))

'in case the rounded time above is set to midnight, the date must be incremented
' by one
 
 If booDateIncr Then
     uEmployee.LaborStartDate = DateAdd("d", 1, uEmployee.LaborStartDate)
End If
'
  If ComparTime > MaxStartTimeSecs Then
      Call RecordError("I")
  ElseIf ComparTime < MinStartTimeSecs Then
           Call RecordError("J")
  End If
  
  Clock_in_rnd = ltimeDec
End Function

Public Sub Attd_Badge()
 Call Init_varbs
 Call GetSysDateTime
 gSysTimeSecs = ConvTimeToSecs(gSysTimeStrp)

 If Clock_Err() = True Then
     ErrBox LoadResString(gLanguageOffset + 93), 4
     Exit Sub
 End If
 
 If gEmployee.ClassWrkr = 4 And div = stroms Then
     ErrBox LoadResString(gLanguageOffset + 292), 4
    Exit Sub
End If
     
 Select Case gEmployee.AttdStatus
    Case 0
            Call ClockIn
    Case 1
  ' the next condition will disable toggling lunch ringings at Hyson products
            If Check_For_Lunch() = True And div <> hyson_products Then
               Call Clk_out_to_lunch
            Else
              Call ClockOut
            End If
    Case 2
            Call ClockIn
            If Ck_for_Attd_Rec(gEmployee.EmplNo, AttdActDate) Then
                If gEmployee.AttdStatus <> 8 Then
                   ErrBox LoadResString(gLanguageOffset + 9), 4
                   Exit Sub
                End If
            End If
    Case 3
            Call Clk_Back_From_Lunch
    Case 4
            Call ClockOut
    Case 5
            Call ClockOut
    Case 6
            If Check_For_Lunch() = True And div <> hyson_products Then
               Call Clk_out_to_lunch
            Else
              Call ClockOut
            End If
 End Select

 If gClockOutFlag = True Then
      frmInpPrompt.Show 1
      If inpClockOut = False Then
         Exit Sub
      End If
 End If
 Call Upd_Empl_Attd
 
End Sub

Public Sub Lunch_out_proc()
   If gEmployee.AttdStatus = 2 Or gEmployee.AttdStatus = 8 Then
        ErrBox LoadResString(gLanguageOffset + 34), 4
        Exit Sub
   End If
   If gEmployee.AttdStatus = 3 Then
        ErrBox LoadResString(gLanguageOffset + 33), 4
        Exit Sub
   End If
   If gEmployee.AttdStatus = 4 Then
        ErrBox LoadResString(gLanguageOffset + 29), 4
        Exit Sub
   End If
   If gEmployee.AttdStatus = 7 Then
        ErrBox "Cannot Use During Call-Back", 4
        Exit Sub
   End If
   If gEmployee.AttdStatus = 5 Then
        ErrBox LoadResString(gLanguageOffset + 32), 4
        Exit Sub
   End If
   
   If div <> syracuse Then
   
       If Not Check_For_Lunch() Then
            ErrBox LoadResString(gLanguageOffset + 77), 4
            Exit Sub
       End If
       
   End If
   
  Call Clk_out_to_lunch
  Call Upd_Empl_Attd
End Sub

Public Sub Lunch_In_Proc()
   If gEmployee.AttdStatus = 2 Or gEmployee.AttdStatus = 8 Then
        ErrBox LoadResString(gLanguageOffset + 34), 4
        Exit Sub
   End If
   If gEmployee.AttdStatus <> 3 Then
        ErrBox LoadResString(gLanguageOffset + 35), 4
        Exit Sub
   End If
   If Ck_for_Auto_Clock() = True Then
       Call ClockIn
       Call RecordError("A")
   Else
       Call Clk_Back_From_Lunch
   End If
   Call Upd_Empl_Attd
End Sub

Public Sub Susp_Process()
   If gEmployee.AttdStatus = 2 Or gEmployee.AttdStatus = 8 Then
        ErrBox LoadResString(gLanguageOffset + 34), 4
        Exit Sub
   End If
   If gEmployee.AttdStatus = 3 Then
        ErrBox LoadResString(gLanguageOffset + 33), 4
        Exit Sub
   End If
   If gEmployee.AttdStatus = 7 Then
        ErrBox "Cannot Use During Call-Back", 4
        Exit Sub
   End If
   If gEmployee.AttdStatus = 5 Then
        ErrBox LoadResString(gLanguageOffset + 32), 4
        Exit Sub
   End If
   
  Call Susp_Attd
  Call Upd_Empl_Attd
End Sub

Public Sub Resume_Process()
   If gEmployee.AttdStatus = 2 Or gEmployee.AttdStatus = 8 Then
        ErrBox LoadResString(gLanguageOffset + 34), 4
        Exit Sub
   End If
   If gEmployee.AttdStatus <> 5 Then
        ErrBox LoadResString(gLanguageOffset + 31), 4
        Exit Sub
   End If
   If Ck_for_Auto_Clock() = True Then
       Call ClockIn
       Call RecordError("A")
   Else
       Call Resm_Attd
   End If
   Call Upd_Empl_Attd
End Sub

Public Sub Att_Clock_in()
  Call ClockIn
  If Ck_for_Attd_Rec(gEmployee.EmplNo, AttdActDate) Then
     If gEmployee.AttdStatus <> 8 Then
        ErrBox LoadResString(gLanguageOffset + 9), 4
        Exit Sub
     End If
  End If
  AttdDayCd = DatePart("w", gSysDate)
'  If gCompanyCd = 17 then
'    If AttdDayCd = 1 Or AttdDayCd = 7 Then
'      gEmployee.PdLunch = "Y"
'    Else
'      If AttdShift = 2 Or AttdShift = 3 Then
'          gEmployee.PdLunch = "Y"
'      End If
'    End If
'  End If
  Call Upd_Empl_Attd
End Sub



Public Sub Att_Clock_out()
    If gEmployee.AttdStatus = 7 Then
        ErrBox LoadResString(gLanguageOffset + 98), 4
        Exit Sub
    End If
    Call ClockOut
    If Already_clk_out() Then
       ErrBox LoadResString(gLanguageOffset + 30), 4
       Exit Sub
    End If
    If gEmployee.AttdStatus = 5 Then
        Call RecordError("S")
    End If
    Call Upd_Empl_Attd
End Sub

Public Sub CI_start()
  If Ck_for_Attd_Rec(gEmployee.EmplNo, gSysDate) Then
     AttdActDate = gEmployee.Actdate
     AttdShift = gEmployee.GenShift
  Else
     Call ClockIn
  End If
  AttdStatus = 7
  AttdDayCd = DatePart("w", gSysDate)
  Call Upd_Empl_Attd
End Sub

Public Sub CI_End()
  If gEmployee.AttdStatus <> 7 Then
     ErrBox LoadResString(gLanguageOffset + 36), 4
     Exit Sub
  End If
  AttdActDate = gEmployee.Actdate
  AttdShift = gEmployee.GenShift
  AttdStatus = 8
  Call Upd_Empl_Attd
End Sub

Public Function Check_For_Lunch() As Boolean
Dim lunch_in As Long
Dim lunch_out As Long
Dim Lunch_grace As Integer
Dim shift_num As Integer
Dim Shift_Grp As Integer

Check_For_Lunch = False

shift_num = gEmployee.GenShift
Shift_Grp = gEmployee.shiftGrp

lunch_out = ConvTimeToSecs(gShiftRecData(shift_num).LunchOut)
lunch_in = ConvTimeToSecs(gShiftRecData(shift_num).LunchIn)

Lunch_grace = gShiftRecData(shift_num).lunchGrace * 60
lunch_out = lunch_out - Lunch_grace

If lunch_out < 0 Then
   lunch_out = lunch_out + 86400
End If

If lunch_out > lunch_in Then
   If ((gSysTimeSecs >= lunch_out And gSysTimeSecs <= 86400) Or _
       (gSysTimeSecs <= lunch_in And gSysTimeSecs >= 0)) Then
         Check_For_Lunch = True
    End If
Else
    If (gSysTimeSecs >= lunch_out And gSysTimeSecs <= lunch_in) Then
          Check_For_Lunch = True
    End If
End If
End Function

Public Sub ClockIn()
Dim ClockInDate As Variant
Dim shiftGrp As Integer

ClockInDate = gSysDate
shiftGrp = gEmployee.shiftGrp
'gSysTimeStrp = 160500
If gEmployee.AsgnShift <> "0" Then
    AttdShift = gEmployee.AsgnShift
Else
    If gSysTimeStrp >= gShiftRecData(1).StParam And _
       gSysTimeStrp <= gShiftRecData(1).EndParam Then
        AttdShift = 1
   ElseIf gShiftRecData(2).EndParam < gShiftRecData(2).StParam Then
       If gSysTimeStrp >= gShiftRecData(2).StParam Or _
          gSysTimeStrp <= gShiftRecData(2).EndParam Then
             AttdShift = 2
       End If
   ElseIf gSysTimeStrp >= gShiftRecData(2).StParam And _
       gSysTimeStrp <= gShiftRecData(2).EndParam Then
        AttdShift = 2
   ElseIf gShiftRecData(3).EndParam < gShiftRecData(3).StParam Then
       If gSysTimeStrp >= gShiftRecData(3).StParam Or _
          gSysTimeStrp <= gShiftRecData(3).EndParam Then
             AttdShift = 3
       End If
   ElseIf gSysTimeStrp >= gShiftRecData(3).StParam And _
        gSysTimeStrp <= gShiftRecData(3).EndParam Then
           AttdShift = 3
   End If
End If


'If AttdShift = 3 And gEmployee.ClassWrkr <> 4 Then
'   If gEmployee.AsgnShift <> 0 Then
'      ClockInDate = DateAdd("d", 1, gSysDate)
'   ElseIf gShiftData(shiftGrp, 3).StParam > gShiftData(shiftGrp, 3).EndParam And _
'            gSysTimeStrp >= gShiftData(shiftGrp, 3).StParam Then
'            ClockInDate = DateAdd("d", 1, gSysDate)
'   End If
'End If

If AttdShift = 3 And gEmployee.ClassWrkr <> 4 Then
     If gEmployee.AsgnShift <> 0 Then
        ' Unconditionally set activity date to next day when employee has assigned shift = 3
         ClockInDate = DateAdd("d", 1, gSysDate)
    ElseIf gShiftRecData(3).StParam > gShiftRecData(3).EndParam And _
                    (gSysTimeStrp <= "235959" And gSysTimeStrp >= gShiftRecData(3).StParam) Then
       ' When employee has no assigned shift - floating -  only adjust activity date if current time is
       ' between shift start param and midnight; after midnight, date is not adjusted.
                      ClockInDate = DateAdd("d", 1, gSysDate)
    End If
End If

AttdActDate = CDate(ClockInDate)
AttdStatus = 1
If AttdShift = 0 Then
    AttdShift = 1
End If

End Sub

Public Sub Ck_End_Of_shift()
  Dim shiftEnd As String
  Dim Eos1Time As Long
  Dim Eos2Time As Long
  Dim ClockOutTime As Long
  Dim ClockOutEos1 As Long
  Dim clockOutEos2 As Long
  Dim shiftNum As Integer
  Dim shiftGrp As Integer
  
  shiftGrp = gEmployee.shiftGrp
  shiftNum = gEmployee.GenShift
  
  shiftEnd = gShiftRecData(shiftNum).endTime
  Eos1Time = CLng((gShiftRecData(shiftNum).EOSStart) * 60)
  Eos2Time = CLng((gShiftRecData(shiftNum).EOSEnd) * 60)
  If Eos1Time = 0 And Eos2Time = 0 Then
    gClockOutFlag = False
  Else
     ClockOutTime = ConvTimeToSecs(shiftEnd)
     clockOutEos2 = ClockOutTime + Eos2Time
     If clockOutEos2 > 86399 Then
        clockOutEos2 = clockOutEos2 - 86400
     End If
     ClockOutEos1 = ClockOutTime - Eos1Time
     If ClockOutEos1 < 0 Then
         ClockOutEos1 = 86400 + ClockOutEos1
     End If
     If clockOutEos2 < ClockOutEos1 Then
        If gSysTimeSecs >= ClockOutEos1 Or gSysTimeSecs <= clockOutEos2 Then
            gClockOutFlag = False
        Else
            gClockOutFlag = True
        End If
     Else
        If gSysTimeSecs >= ClockOutEos1 And gSysTimeSecs <= clockOutEos2 Then
             gClockOutFlag = False
        Else
              gClockOutFlag = True
        End If
     End If
   End If
End Sub

Public Sub ClockOut()
   If Ck_for_Auto_Clock() = True Then
       Call ClockIn
       Call RecordError("A")
   Else
       Call Ck_End_Of_shift
       AttdActDate = gEmployee.Actdate
       AttdShift = gEmployee.GenShift
       AttdStatus = 2
    End If
End Sub

Public Function Ck_for_Auto_Clock() As Boolean
  Dim ShiftBegin, shiftEnd, ACT, cClocktimeSecs As Long
  Dim shiftNum, shiftGrp As Integer
  Dim cClockOutDate As Date
  
  Ck_for_Auto_Clock = False
  If gEmployee.AttdStatus = 0 Then
     Ck_for_Auto_Clock = True
     Exit Function
  End If
  
  shiftNum = gEmployee.GenShift
  shiftGrp = gEmployee.shiftGrp
  
  ShiftBegin = CLng(gShiftRecData(shiftNum).startTime)
  shiftEnd = CLng(gShiftRecData(shiftNum).endTime)
  ACT = CLng(gShiftRecData(shiftNum).ACT)
  If ACT = 0 Then
     Ck_for_Auto_Clock = False
     Exit Function
  End If
  ACT = ACT * 60
  cClocktimeSecs = ConvTimeToSecs(gShiftRecData(shiftNum).endTime)
  If ShiftBegin > shiftEnd Then
      cClockOutDate = DateAdd("d", 1, gEmployee.ShiftStartDate)
  Else
      cClockOutDate = gEmployee.ShiftStartDate
  End If
  'next stmt strips time from shift start date
  cClockOutDate = CDate(Format(cClockOutDate, "MM/DD/YYYY"))
  cClocktimeSecs = cClocktimeSecs + ACT
  If cClocktimeSecs > 86399 Then
      cClockOutDate = DateAdd("d", 1, cClockOutDate)
      cClocktimeSecs = cClocktimeSecs - 86400
  End If
  
  If gSysDate > cClockOutDate Then
    Ck_for_Auto_Clock = True
  ElseIf gSysDate = cClockOutDate And _
          gSysTimeSecs > cClocktimeSecs Then
            Ck_for_Auto_Clock = True
  End If
End Function

Public Sub Clk_out_to_lunch()
   AttdActDate = gEmployee.Actdate
   AttdShift = gEmployee.GenShift
   AttdStatus = 3
End Sub

Public Sub Clk_Back_From_Lunch()
  If Ck_for_Auto_Clock() = True Then
      Call ClockIn
      Call RecordError("A")
  Else
      AttdActDate = gEmployee.Actdate
      AttdShift = gEmployee.GenShift
      AttdStatus = 4
  End If
End Sub

Public Sub Susp_Attd()
   AttdActDate = gEmployee.Actdate
   AttdShift = gEmployee.GenShift
   AttdStatus = 5
End Sub

Public Sub Resm_Attd()
   AttdActDate = gEmployee.Actdate
   AttdShift = gEmployee.GenShift
   AttdStatus = 6
End Sub

Public Sub Upd_Empl_Attd()
   'uEmployee and gEmployee both defined as udt employeeType
   uEmployee = gEmployee
   Call GetSysDateTime
   Select Case AttdStatus
      Case 1
          Attd_Disp = LoadResString(gLanguageOffset + 39) & gSysTime & "--"
          uEmployee.Actdate = AttdActDate
          uEmployee.GenShift = AttdShift
          uEmployee.AttdStartDate = gSysDateTime
          uEmployee.AttdEndDate = dBlankDate
          uEmployee.AttdStatus = AttdStatus
          uEmployee.LaborStartDate = uEmployee.AttdStartDate
          uEmployee.ShiftStartDate = uEmployee.AttdStartDate
      Case 2
          Attd_Disp = LoadResString(gLanguageOffset + 40) & gSysTime & "--"
          uEmployee.AttdEndDate = gSysDateTime
          uEmployee.AttdStatus = AttdStatus
      Case 3
          Attd_Disp = LoadResString(gLanguageOffset + 82) & gSysTime & "--"
          uEmployee.AttdEndDate = gSysDateTime
          uEmployee.AttdStatus = AttdStatus
      Case 4
          Attd_Disp = LoadResString(gLanguageOffset + 14) & gSysTime & "--"
          uEmployee.AttdStartDate = gSysDateTime
          uEmployee.AttdStatus = AttdStatus
      Case 5
          Attd_Disp = LoadResString(gLanguageOffset + 13) & gSysTime & "--"
          uEmployee.AttdEndDate = gSysDateTime
          uEmployee.AttdStatus = AttdStatus
      Case 6
          Attd_Disp = LoadResString(gLanguageOffset + 12) & gSysTime & "--"
          uEmployee.AttdStartDate = gSysDateTime
          uEmployee.AttdStatus = AttdStatus
      Case 7
          Attd_Disp = LoadResString(gLanguageOffset + 17) & gSysTime & "--"
          uEmployee.Actdate = AttdActDate
          uEmployee.GenShift = AttdShift
          uEmployee.AttdStartDate = gSysDateTime
          uEmployee.AttdEndDate = dBlankDate
          uEmployee.AttdStatus = AttdStatus
          uEmployee.LaborStartDate = uEmployee.AttdStartDate
      Case 8
          Attd_Disp = LoadResString(gLanguageOffset + 16) & gSysTime & "--"
          uEmployee.AttdEndDate = gSysDateTime
          uEmployee.AttdStatus = AttdStatus
    End Select
    Call Attd_upd
End Sub

Public Function Already_clk_out() As Boolean
   If (gEmployee.AttdStatus = 2 Or gEmployee.AttdStatus = 8) And _
      gEmployee.Actdate = AttdActDate Then
        Already_clk_out = True
   Else
        Already_clk_out = False
   End If
End Function

Public Sub Check_Lunch_rings()
 Dim actLunchEnd, plnLunchEnd As Long
 Dim shiftNum, shiftGrp As Integer
 Dim TimeDiff As Currency
 Dim timeStr As String
 
 'for some reason, this function is performed only if the employee-paid-lunch flag is Y
 
 shiftNum = gEmployee.GenShift
 shiftGrp = gEmployee.shiftGrp
 If gAttendance.LunchStart = dBlankDate And _
    gAttendance.LunchEnd = dBlankDate Then
      Exit Sub
 End If
 
 TimeDiff = Calc_Elapsed_time(gAttendance.LunchOutRnd, gAttendance.LunchInRnd)
 If TimeDiff = -1# Then
      Call RecordError("L")
      Exit Sub
 End If
 If IsDate(gAttendance.LunchEnd) And gAttendance.LunchEnd <> dBlankDate Then
      timeStr = ConvTimeToString(gAttendance.LunchEnd)
      actLunchEnd = ConvTimeToSecs(timeStr)
 End If
 plnLunchEnd = ConvTimeToSecs(gShiftRecData(shiftNum).LunchIn)
 If actLunchEnd > plnLunchEnd Then
    Call RecordError("K")
 End If
End Sub

Public Function Lunch_Out_Rnd()
  Dim lunchOutStart, LunchOutEnd As Long
  Dim shiftNum, shiftGrp As Integer
  Dim lunchGrace As Integer
  Dim lMM, lHH, lSS As Integer
  Dim ltimeMMSS, ltimeDec As Currency
  Dim lTimeChar, tempTimeChar As String
    
  shiftNum = AttdShift
  shiftGrp = gEmployee.shiftGrp
  
  LunchOutEnd = ConvTimeToSecs(gShiftRecData(shiftNum).LunchOut)
  lunchGrace = gShiftRecData(shiftNum).lunchGrace * 60
  lunchOutStart = LunchOutEnd - lunchGrace
  If lunchOutStart < 0 Then
      lunchOutStart = lunchOutStart + 86400
  End If
  LunchOutEnd = ConvTimeToSecs(gShiftRecData(shiftNum).LunchIn)
  If lunchOutStart > LunchOutEnd Then
     If ((gSysTimeSecs >= lunchOutStart And gSysTimeSecs <= 86400) Or _
         (gSysTimeSecs <= LunchOutEnd And gSysTimeSecs >= 0)) Then
           ltimeDec = RndTimeHnds(gShiftRecData(shiftNum).LunchOut)
           uEmployee.AttdEndDate = CDate(ConvToDateTime(gSysDateStr, gShiftRecData(shiftNum).LunchOut))
     End If
  ElseIf (gSysTimeSecs >= lunchOutStart And _
          gSysTimeSecs <= LunchOutEnd) Then
           ltimeDec = RndTimeHnds(gShiftRecData(shiftNum).LunchOut)
           uEmployee.AttdEndDate = CDate(ConvToDateTime(gSysDateStr, gShiftRecData(shiftNum).LunchOut))
  Else
      ltimeDec = RndTimeHnds(gSysTimeStrp)
      uEmployee.AttdEndDate = gSysDateTime
  End If
  
Lunch_Out_Rnd = ltimeDec

End Function

Public Function Lunch_in_Rnd()
  
  Dim lunchInStart, LunchInEnd As Long
  Dim shiftNum, shiftGrp As Integer
  Dim lunchGrace As Integer
  Dim ltimeDec As Currency
  Dim lTimeChar, tempTimeChar As String
    
  shiftNum = AttdShift
  shiftGrp = gEmployee.shiftGrp
  
  lunchInStart = ConvTimeToSecs(gShiftRecData(shiftNum).LunchOut)
  LunchInEnd = ConvTimeToSecs(gShiftRecData(shiftNum).LunchIn)
  
  If lunchInStart > LunchInEnd Then
     If ((gSysTimeSecs >= lunchInStart And gSysTimeSecs <= 86400) Or _
         (gSysTimeSecs <= LunchInEnd And gSysTimeSecs >= 0)) Then
           ltimeDec = RndTimeHnds(gShiftRecData(shiftNum).LunchIn)
           uEmployee.AttdStartDate = CDate(ConvToDateTime(gSysDateStr, gShiftRecData(shiftNum).LunchIn))
     End If
  ElseIf (gSysTimeSecs >= lunchInStart And _
          gSysTimeSecs <= LunchInEnd) Then
           ltimeDec = RndTimeHnds(gShiftRecData(shiftNum).LunchIn)
           
           uEmployee.AttdStartDate = CDate(ConvToDateTime(gSysDateStr, gShiftRecData(shiftNum).LunchIn))
  Else
      ltimeDec = RndTimeHnds(gSysTimeStrp)
      uEmployee.AttdStartDate = gSysDateTime
  End If
  
'Below statement added 4/3/12 as a cure-all for too much time being deducted for lunch in Syracuse - MWP
If gCompanyCd = 3 Then
    Lunch_in_Rnd = gAttendance.LunchOutRnd + 0.5
    If Lunch_in_Rnd > 23.9 Then
         Lunch_in_Rnd = Lunch_in_Rnd - 24#
    End If
Else
   Lunch_in_Rnd = ltimeDec
End If

End Function

Public Function Elapsed_hrs(ByVal fromDate As Date, ByVal ToDate As Date) As Currency
'If the fromDate is less then the ToDate the function
'returns a minus 1.

Dim DiffSecs As Long

If fromDate > ToDate Then
    Elapsed_hrs = -1
    Exit Function
End If
DiffSecs = DateDiff("s", fromDate, ToDate)
Elapsed_hrs = CCur(Format(DiffSecs / 3600, "000000.0"))

End Function

Public Sub Adj_Labor_Start()
'This procedure adjusts the labor start date/time, to
'accommodate lunch time and suspend times

Dim lbr_start As Date
Dim NumbOfSecs As Long

NumbOfSecs = DateDiff("s", uEmployee.AttdEndDate, uEmployee.AttdStartDate)
uEmployee.LaborStartDate = DateAdd("s", NumbOfSecs, uEmployee.LaborStartDate)

End Sub

Public Sub Build_Auto_Indir()
' For class 3 employees that are full-time indirect workers,
' (janitors, shipping/receiving, etc), automatically build a
' indirect record when the employee clocks out for the day. If
' the employee has previously rung labor for the same date/shift,
' do not build the record.
  
  If gEmployee.ClassWrkr = 3 And (Trim(gEmployee.IndirCd) <> "") Then
       If First_Labor_Rec(gEmployee.EmplNo, gSysDate) = True Then
           AttdIndirLbr = True
           If gCompanyCd = bristol Then
'              Call Labor_Earnings
              uEmployee.LaborStartDate = gSysDateTime
           Else
              gLabor.Earnings = 0
              gLabor.PayRate = 0
           End If
           gAttendance.CCID = gEmployee.IndirCCID
       End If
  End If
End Sub

Public Sub Attd_upd()
Let gAttendance = nAttendance
If Ck_for_Attd_Rec(gEmployee.EmplNo, AttdActDate) Then
    Call Build_Attd_Rec("U")
    Call Upd_Attd_rec
Else
' need to initialize these fields to some large negative
' in place of null; init to zero may represent midnight

   Let gAttendance.ClockInRnd = INIT_RND_TM
   Let gAttendance.ClockOutRnd = INIT_RND_TM
   Let gAttendance.LunchInRnd = INIT_RND_TM
   Let gAttendance.LunchOutRnd = INIT_RND_TM
   Let gAttendance.SuspLbr1Rnd = INIT_RND_TM
   Let gAttendance.ResmLbr1Rnd = INIT_RND_TM
   Let gAttendance.SuspLbr2Rnd = INIT_RND_TM
   Let gAttendance.ResmLbr2Rnd = INIT_RND_TM
   Let gAttendance.CIStartRnd = INIT_RND_TM
   Let gAttendance.CIEndRnd = INIT_RND_TM
   Call Build_Attd_Rec("I")
   Call Ins_Attd_rec
End If
End Sub

Public Function Build_Attd_Rec(ByVal buildType As String)
Dim hhmmRnd As Currency
Dim shiftNum, shiftGrp As Integer

If buildType = "I" Then
     If AttdActDate = dBlankDate Then
         AttdActDate = Date
     End If
     gAttendance.companyCd = gCompanyCd
     gAttendance.EmplNo = gEmployee.EmplNo
     gAttendance.ShiftCd = AttdShift
     gAttendance.CCID = gEmployee.HomeCCID
     gAttendance.Actdate = AttdActDate
     gAttendance.HrsUnrnd = 0
     gAttendance.ActHrs = 0
     gAttendance.CallInHrs = 0
     gAttendance.ErrorLst = "          "
     gAttendance.PdLunch = "N"
     gAttendance.ErrorFlag = " "
     If AttdStatus <> 1 Then
         Call RecordError("F")
     End If
  End If
  AttdIndirLbr = False
  shiftNum = AttdShift
  shiftGrp = gEmployee.shiftGrp
  Select Case AttdStatus
    Case 1
       If gCompanyCd = 12 Or gCompanyCd = 76 Then
           gAttendance.ClockInRnd = Alt_Clock_in_rnd()
       Else
           gAttendance.ClockInRnd = Clock_in_rnd()
       End If
       gAttendance.ClockIn = gSysDateTime
       If gEmployee.PdLunch = "Y" Then
            gAttendance.PdLunch = gEmployee.PdLunch
       Else
            gAttendance.PdLunch = "N"
            If gAutoLunch Then
                SetLunchTimes
            End If
       End If
    Case 2
       If gCompanyCd = 12 Or gCompanyCd = 76 Then
           gAttendance.ClockOutRnd = Alt_Clock_out_rnd()
       Else
           gAttendance.ClockOutRnd = Clock_out_rnd()
       End If
       gAttendance.ClockOut = gSysDateTime
       Calc_Attd_Hrs
       If gAttdAdj(shiftGrp) > 0# Then
          gAttendance.ActHrs = gAttendance.ActHrs - gAttdAdj(shiftGrp)
       End If
       Build_Auto_Indir
    Case 3
       If gCompanyCd = 3 And shiftGrp = 3 Then
            gAttendance.LunchOutRnd = RndTimeHnds(gSysTimeStrp)
       Else
             gAttendance.LunchOutRnd = Lunch_Out_Rnd()
       End If
       gAttendance.LunchStart = gSysDateTime
       gAttendance.LunchInRnd = INIT_RND_TM
       gAttendance.LunchEnd = dBlankDate
    Case 4
       If gCompanyCd = 3 And shiftGrp = 3 Then
            gAttendance.LunchInRnd = gAttendance.LunchOutRnd + 0.5
            If gAttendance.LunchInRnd > 23.9 Then
                   gAttendance.LunchInRnd = gAttendance.LunchInRnd - 24#
            End If
'            RndTimeHnds (gSysTimeStrp)
       Else
             gAttendance.LunchInRnd = Lunch_in_Rnd()
       End If
       gAttendance.LunchEnd = gSysDateTime
       Adj_Labor_Start
    Case 5
       If gAttendance.SuspLbr1 = dBlankDate And gAttendance.ResmLbr1 = dBlankDate Then
           gAttendance.SuspLbr1Rnd = RndTimeHnds(gSysTimeStrp)
           gAttendance.SuspLbr1 = gSysDateTime
       Else
          gAttendance.SuspLbr2Rnd = RndTimeHnds(gSysTimeStrp)
          gAttendance.SuspLbr2 = gSysDateTime
       End If
    Case 6
       Adj_Labor_Start
       If gAttendance.ResmLbr1 = dBlankDate Then
           gAttendance.ResmLbr1Rnd = RndTimeHnds(gSysTimeStrp)
           gAttendance.ResmLbr1 = gSysDateTime
       Else
          gAttendance.ResmLbr2Rnd = RndTimeHnds(gSysTimeStrp)
          gAttendance.ResmLbr2 = gSysDateTime
       End If
    Case 7
       gAttendance.CIStartRnd = RndTimeHnds(gSysTimeStrp)
       gAttendance.CIStart = gSysDateTime
    Case 8
       gAttendance.CIEndRnd = RndTimeHnds(gSysTimeStrp)
       gAttendance.CIEnd = gSysDateTime
       Calc_Attd_Hrs
       If gAttdAdj(shiftGrp) > 0# Then
          gAttendance.ActHrs = gAttendance.ActHrs - gAttdAdj(shiftGrp)
       End If
     End Select
    gAttendance.ErrorLst = RTrim(gAttendance.ErrorLst)
    If (Len(Trim(gAttendance.ErrorLst)) + Len(Trim(gLaborIn(LoopCnt).ErrorLst))) <= 10 Then
        If Len(Trim(gLaborIn(LoopCnt).ErrorLst)) > 0 Then
            gAttendance.ErrorLst = gAttendance.ErrorLst & gLaborIn(LoopCnt).ErrorLst
'Corry has elected not to flag attendance records for review if there is an error
'This rule change was established as of 2/24/98 - MWP. Stroms added on 11/7/00, as well.

            If (gCompanyCd <> 7 And gCompanyCd <> 77) Then
                gAttendance.ErrorFlag = "Y"
            End If
        End If
    End If
End Function




Public Sub SetLunchTimes()
Dim shiftNum, shiftGrp As Integer
Dim sDateStr As String

shiftNum = AttdShift
shiftGrp = gEmployee.shiftGrp

sDateStr = AttdActDate
gAttendance.LunchStart = CDate(ConvToDateTime(sDateStr, gShiftRecData(shiftNum).LunchOut))
gAttendance.LunchOutRnd = RndTimeHnds(gShiftRecData(shiftNum).LunchOut)
gAttendance.LunchEnd = CDate(ConvToDateTime(sDateStr, gShiftRecData(shiftNum).LunchIn))
gAttendance.LunchInRnd = RndTimeHnds(gShiftRecData(shiftNum).LunchIn)

End Sub

