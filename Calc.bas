Attribute VB_Name = "Calc"
Option Explicit

Public Function RndTimeHnds(ByVal inTimeStr As String) As Currency
  Dim lHH, lMM, lSS, lCompTime As Integer
  Dim lOutTime As Currency
  
  If Len(inTimeStr) < 6 Then
      RndTimeHnds = 0#
      Exit Function
   End If
   
  lHH = CInt(Left(inTimeStr, 2))
  lMM = CInt(Mid(inTimeStr, 3, 2))
  RndTimeHnds = 0#
  Select Case lMM
     Case Is < 3
          lOutTime = lHH + 0#
     Case Is < 9
          lOutTime = lHH + 0.1
     Case Is < 15
          lOutTime = lHH + 0.2
     Case Is < 21
          lOutTime = lHH + 0.3
     Case Is < 27
          lOutTime = lHH + 0.4
     Case Is < 33
          lOutTime = lHH + 0.5
     Case Is < 39
          lOutTime = lHH + 0.6
     Case Is < 45
          lOutTime = lHH + 0.7
     Case Is < 51
          lOutTime = lHH + 0.8
     Case Is < 57
          lOutTime = lHH + 0.9
     Case Else
          lOutTime = lHH + 1#
  End Select
  If lOutTime > 24# Then
      lOutTime = 0
  End If
  RndTimeHnds = lOutTime
End Function

Public Function BurlRndTimeHnds(ByVal inTimeStr As String) As Currency
  Dim lHH, lMM, lSS, lCompTime As Integer
  Dim lOutTime As Currency
  
  If Len(inTimeStr) < 6 Then
      BurlRndTimeHnds = 0#
      Exit Function
   End If
   
  lHH = CInt(Left(inTimeStr, 2))
  lMM = CInt(Mid(inTimeStr, 3, 2))
  BurlRndTimeHnds = 0#
  Select Case lMM
'     Case Is < 3
 '         lOutTime = lHH + 0#
     Case Is < 6
          lOutTime = lHH + 0.1
     Case Is < 12
          lOutTime = lHH + 0.2
     Case Is < 18
          lOutTime = lHH + 0.3
     Case Is < 24
          lOutTime = lHH + 0.4
     Case Is < 30
          lOutTime = lHH + 0.5
     Case Is < 36
          lOutTime = lHH + 0.6
     Case Is < 42
          lOutTime = lHH + 0.7
     Case Is < 48
          lOutTime = lHH + 0.8
     Case Is < 54
          lOutTime = lHH + 0.9
     Case Else
          lOutTime = lHH + 1#
  End Select
  If lOutTime > 24# Then
      lOutTime = 0
  End If
  BurlRndTimeHnds = lOutTime
End Function


Public Function SpectrumRndTimeHnds(ByVal inTimeStr As String) As Currency
  Dim lHH, lMM, lSS, lCompTime As Integer
  Dim lOutTime As Currency
  Dim lBaseValue As Currency
  Dim lIncrValue As Currency
  If Len(inTimeStr) < 6 Then
      SpectrumRndTimeHnds = 0#
      Exit Function
   End If
   
  lHH = CInt(Left(inTimeStr, 2))
  lMM = CInt(Mid(inTimeStr, 3, 2))
  If lMM = 0 Then
    SpectrumRndTimeHnds = lHH + 0#
    Exit Function
 End If
  'Could derive hundredths of an hour by dividing minutes by 60,
  'putting result into a field with two decimal pos. The method below is consistent.
  
  SpectrumRndTimeHnds = 0#
  Select Case lMM
      Case 1, 16, 31, 46
         lBaseValue = 0.02
     Case 2, 17, 32, 47
         lBaseValue = 0.03
     Case 3, 18, 33, 48
         lBaseValue = 0.05
     Case 4, 19, 34, 49
         lBaseValue = 0.07
     Case 5, 20, 35, 50
         lBaseValue = 0.08
     Case 6, 21, 36, 51
         lBaseValue = 0.1
     Case 7, 22, 37, 52
         lBaseValue = 0.12
     Case 8, 23, 38, 53
         lBaseValue = 0.13
     Case 9, 24, 39, 54
         lBaseValue = 0.15
     Case 10, 25, 40, 55
         lBaseValue = 0.17
     Case 11, 26, 41, 56
         lBaseValue = 0.18
     Case 12, 27, 42, 57
         lBaseValue = 0.2
     Case 13, 28, 43, 58
         lBaseValue = 0.22
     Case 14, 29, 44, 59
         lBaseValue = 0.23
     Case 15, 30, 45
         lBaseValue = 0.25
     Case Else
         lBaseValue = 0#
  End Select
  
  Select Case lMM
      Case Is < 16
         lIncrValue = 0#
      Case Is < 31
         lIncrValue = 0.25
      Case Is < 46
         lIncrValue = 0.5
      Case Is < 60
         lIncrValue = 0.75
      Case Else
         lIncrValue = 0#
  End Select
  SpectrumRndTimeHnds = lHH + lBaseValue + lIncrValue
 
End Function


Public Function AllDigits(strNumb As String) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Calc.AllDigits
'Description      : Function to test if a string has all digits 0-9,
'                   and nothing else.
'Parameters       : Input string to check.
'Procedures Used  : None
'Forms Loaded     : None
'Return Parameters: True if all digits, otherwise False
'Tables Updates   : None
'==============================================================================
    Dim I As Integer
    Dim iChar As String
    'Assume valid
    AllDigits = True
    If Len(strNumb) = 0 Then
        AllDigits = False
        Exit Function
    End If
    For I = 1 To Len(strNumb)
        iChar = Mid(strNumb, I, 1)
        If iChar < "0" Or iChar > "9" Then
            AllDigits = False
            Exit For
        End If
    Next I
End Function
Public Function AllDigitSpaces(strNumb As String) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Calc.AllDigitSpacess
'Description      : Function to test if a string has all digits 0-9,
'                   and spaces, and nothing else.
'Parameters       : Input string to check.
'Procedures Used  : None
'Forms Loaded     : None
'Return Parameters: True if all digits or spaces, otherwise False
'Tables Updates   : None
'==============================================================================
    'Function to test if a string has all digits 0-9, and spaces
    '   and nothing else.
    Dim I As Integer
    Dim iChar As String
    'Assume valid
    AllDigitSpaces = True
    If Len(strNumb) = 0 Then
        AllDigitSpaces = False
        Exit Function
    End If
    For I = 1 To Len(strNumb)
        iChar = Mid(strNumb, I, 1)
        If iChar <> " " And (iChar < "0" Or iChar > "9") Then
            AllDigitSpaces = False
            Exit For
        End If
    Next I
End Function



Public Function Calc_Attd_Hrs() As Currency
Dim lAttdHrs, lCallInHrs As Currency
Dim totTime, diffTm As Currency
Dim timeStr As String
Dim shiftNum, shiftGrp As Integer
 Dim actLunchEnd, plnLunchEnd As Long

lAttdHrs = 0
totTime = 0
lCallInHrs = 0

 shiftNum = gEmployee.GenShift
 shiftGrp = gEmployee.shiftGrp

'first calculate the difference between clock in and clock out
diffTm = Calc_Elapsed_time(gAttendance.ClockInRnd, gAttendance.ClockOutRnd)
If diffTm <> -1 Then
   totTime = totTime + diffTm
Else
   gAttendance.HrsUnrnd = totTime
   gAttendance.ActHrs = lAttdHrs
   gAttendance.CallInHrs = lCallInHrs
   Exit Function
End If
If gAttendance.PdLunch = "N" Then
' first test to see if employee clocked out before start of lunch. This should happen only
' when system is set up for AUTOLUNCH - the lunch times are automatically set when
' the employee clocks in.
  If gAttendance.ClockOutRnd > gAttendance.ClockInRnd Then
     If (gAttendance.ClockOutRnd < gAttendance.LunchOutRnd Or _
         gAttendance.ClockInRnd > gAttendance.LunchOutRnd) Then
           gAttendance.LunchOutRnd = INIT_RND_TM
           gAttendance.LunchInRnd = INIT_RND_TM
           gAttendance.LunchStart = dBlankDate
           gAttendance.LunchEnd = dBlankDate
     End If
  End If
  If (gAttendance.ClockInRnd > gAttendance.LunchOutRnd And _
     gAttendance.ClockOutRnd < gAttendance.LunchOutRnd) Then
           gAttendance.LunchOutRnd = INIT_RND_TM
           gAttendance.LunchInRnd = INIT_RND_TM
           gAttendance.LunchStart = dBlankDate
           gAttendance.LunchEnd = dBlankDate
  End If
    
    diffTm = Calc_Elapsed_time(gAttendance.LunchOutRnd, gAttendance.LunchInRnd)
    If diffTm <> -1 Then
       totTime = totTime - diffTm
    Else
       Call RecordError("L")
       gAttendance.HrsUnrnd = totTime
       gAttendance.ActHrs = lAttdHrs
       gAttendance.CallInHrs = lCallInHrs
       Exit Function
    End If
    If IsDate(gAttendance.LunchEnd) And gAttendance.LunchEnd <> dBlankDate Then
         timeStr = ConvTimeToString(gAttendance.LunchEnd)
        actLunchEnd = ConvTimeToSecs(timeStr)
        plnLunchEnd = ConvTimeToSecs(gShiftRecData(shiftNum).LunchIn)
        If actLunchEnd > plnLunchEnd Then
 ' attendance hours calc continues despite error  code K since this only indicates a late
 ' back from lunch.
            Call RecordError("K")
        End If
    End If

Else
    #If Not gbJobPrint Then
        Call Check_Lunch_rings
    #End If
End If

diffTm = Calc_Elapsed_time(gAttendance.SuspLbr1Rnd, gAttendance.ResmLbr1Rnd)
If diffTm <> -1 Then
    totTime = totTime - diffTm
Else
    Call RecordError("X")
    gAttendance.HrsUnrnd = totTime
    gAttendance.ActHrs = lAttdHrs
    gAttendance.CallInHrs = lCallInHrs
    Exit Function
End If

diffTm = Calc_Elapsed_time(gAttendance.SuspLbr2Rnd, gAttendance.ResmLbr2Rnd)
If diffTm <> -1 Then
    totTime = totTime - diffTm
Else
    Call RecordError("X")
    gAttendance.HrsUnrnd = totTime
    gAttendance.ActHrs = lAttdHrs
    gAttendance.CallInHrs = lCallInHrs
    Exit Function
End If

diffTm = Calc_Elapsed_time(gAttendance.CIStartRnd, gAttendance.CIEndRnd)
If diffTm <> -1 Then
    totTime = totTime + diffTm
    lCallInHrs = diffTm
Else
    Call RecordError("Y")
    gAttendance.HrsUnrnd = totTime
    gAttendance.ActHrs = lAttdHrs
    gAttendance.CallInHrs = lCallInHrs
    Exit Function
End If

If totTime < 0 Then
    totTime = 0
End If
'next statement should round totTime to one decimal,
'and assign the result to lAttdHrs.
lAttdHrs = CCur(Format(totTime, "00.0"))

gAttendance.HrsUnrnd = totTime
gAttendance.ActHrs = lAttdHrs
gAttendance.CallInHrs = lCallInHrs

End Function

Public Function Calc_Elapsed_time(ByVal startTime As Currency, ByVal endTime As Currency) As Currency
  Calc_Elapsed_time = 0
  If startTime = INIT_RND_TM Then
     If endTime = INIT_RND_TM Then
         Calc_Elapsed_time = 0
     Else
         Calc_Elapsed_time = -1#
      End If
  Else
    If endTime = INIT_RND_TM Then
          Calc_Elapsed_time = -1#
    Else
       If startTime > endTime And _
        ((gEmployee.GenShift = 1 Or gEmployee.GenShift = 2 Or gEmployee.GenShift = 3) Or AttdStatus = 8) Then
           endTime = endTime + 24#
       End If
       Calc_Elapsed_time = endTime - startTime
    End If
  End If
End Function



Public Function CalcActHrs() As Currency
  Dim shiftNum, shiftGrp As Integer
  Dim shiftEnd As String
  Dim endTimeTths, startTimeTths As Currency
  Dim sLaborStart As String
  
  CalcActHrs = 0#
  
  Call GetSysDateTime
  If empClkOutOK Then
     Let shiftNum = gEmployee.GenShift
     Let shiftGrp = gEmployee.shiftGrp
     Let shiftEnd = gShiftRecData(shiftNum).endTime
     endTimeTths = ConvTimeTths(shiftEnd)
  Else
     endTimeTths = ConvTimeTths(gSysTimeStrp)
  End If
  sLaborStart = Format(gEmployee.LaborStartDate, "HHMMSS")
  startTimeTths = ConvTimeTths(sLaborStart)
  If startTimeTths > endTimeTths Then
      endTimeTths = endTimeTths + 24#
  End If
  CalcActHrs = endTimeTths - startTimeTths
End Function



Public Function ConvTimeTths(ByVal sTimeStr As String) As Currency

Dim iMins As Integer
Dim cCurrFld As Currency
Dim sOutMins As String * 4
Dim sDecMins As String * 2
Dim lOutTime As String
Dim iOutTime As Long
ConvTimeTths = 0#

If Len(sTimeStr) < 6 Then
    Exit Function
End If


iMins = CInt(Mid(sTimeStr, 3, 2))
cCurrFld = iMins / 60
sOutMins = Format(cCurrFld, "0.00")
sDecMins = Mid(sOutMins, 3, 2)
lOutTime = Left(sTimeStr, 2) & Left(sDecMins, 2) & Mid(sTimeStr, 5, 2)
iOutTime = CLng(lOutTime)
cCurrFld = iOutTime / 10000
sOutMins = Format(cCurrFld, "00.0")
ConvTimeTths = CCur(sOutMins)

End Function



Public Function ConvTimeToSecs(ByVal inStrTime As String) As Long
Dim iHrs As Long
Dim iMins As Long
Dim iSecs As Long

If Len(inStrTime) < 6 Then
     ConvTimeToSecs = 0
     Exit Function
End If

iHrs = CLng(Mid(inStrTime, 1, 2))
iMins = CLng(Mid(inStrTime, 3, 2))
iSecs = CLng(Mid(inStrTime, 5, 2))

ConvTimeToSecs = (iHrs * 3600) + (iMins * 60) + iSecs

End Function


Public Function Clock_Err() As Boolean

' This function determines if a number of seconds have elapsed since
' the last attendance ringing by the same employee. This is done by
' computing the difference (in seconds) between the current time and
' the time of the last attendance ringing. The number of
' seconds is defined by a parameter ,  gGracePer, that is kept with
' the clock in and clock out grace times in the div param record
' GRACEPER"

Dim DiffSecs As Long
If gEmployee.AttdStatus = 0 Then
  Clock_Err = False
  Exit Function
End If
If (gEmployee.AttdStatus = 1 Or gEmployee.AttdStatus = 4 Or _
    gEmployee.AttdStatus = 6 Or gEmployee.AttdStatus = 7) Then
     DiffSecs = DateDiff("s", gEmployee.AttdStartDate, gSysDateTime)
Else
     DiffSecs = DateDiff("s", gEmployee.AttdEndDate, gSysDateTime)
End If
If DiffSecs > (gGracePer * 60) Then
     Clock_Err = False
Else
     Clock_Err = True
End If
   
End Function
Public Function ConvSqlDate(ByVal sqlDateStr As Variant) As String
Dim convStr As Variant
Dim DateLen As Long
Dim SQLYear As Integer
Dim SQLMonth As Integer
Dim SQLDay As Integer
Dim SQLDate As String
Dim SQLDate2 As Date


If IsNull(sqlDateStr) Then
    ConvSqlDate = dBlankDate
    Exit Function
End If

sqlDateStr = RTrim(sqlDateStr)
DateLen = Len(sqlDateStr)
If DateLen > 19 Then
     'date converted using SQL style 109 MMM dd yyyy hh:mm:ss:tttXX
      convStr = Trim(Left(sqlDateStr, 20)) & " " & Right(sqlDateStr, 2)
ElseIf DateLen = 19 Then
     'date converted using SQL style 120 yyyy-mm-dd hh:mm:ss
      convStr = Trim(Left(sqlDateStr, 19))
End If
If IsDate(convStr) Then
    ConvSqlDate = convStr
Else
    ConvSqlDate = dBlankDate
End If

End Function
Public Function ConvSqlVals(ByVal sqlValStr As Variant) As Currency
Dim convStr As Variant
If IsNull(sqlValStr) Then
    ConvSqlVals = INIT_RND_TM
    Exit Function
End If
sqlValStr = Val(sqlValStr)
ConvSqlVals = CCur(sqlValStr)

End Function


Public Function ConvTimeToString(ByVal inpDate As Date) As String
  ' returns a string with the time formatted as HHMMSS
  Dim timeStr As String
  ConvTimeToString = " "
  timeStr = Format(inpDate, "HH:MM:SS")
  ConvTimeToString = Left(timeStr, 2) & Mid(timeStr, 4, 2) & Mid(timeStr, 7, 2)

End Function




Public Function ConvToDateTime(ByVal inDate As String, ByVal inTime As String) As String
'This function accepts two strings - one is a date string
'that is formatted mm/dd/yyyy and the other is a time string
'that is formatted HHMMSS
'MsgBox (inDate & " " & inTime)

If Len(inTime) < 6 Then
     inTime = "000000"
End If
ConvToDateTime = ""
ConvToDateTime = (inDate & " " & Left(inTime, 2) & ":" & Mid(inTime, 3, 2) & ":" & Mid(inTime, 5, 2))

End Function


Public Function ValidNumber(strIn As String) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Calc.ValidNumber
'Description      : Function to test if a string has all digits 0-9,
'                   or spaces, and nothing else, that is a valid number.
'Parameters       : Input string to check.
'Called From      : frmDirLabor.inpLabJobSeq, Inputs.inpJobAndSeq
'Procedures Used  : AllDigitSpaces
'Forms Loaded     : None
'Return Parameters: True if valid number of digits and spaces, otherwise False
'Tables Updates   : None
'==============================================================================
    'Make sure is a valid number using IsNumeric (allows signs . e trailing and
    '   and leading spaces etc), then make sure number only has digits and
    '   leading and/or trailing spaces
    ValidNumber = IsNumeric(strIn) And AllDigitSpaces(strIn)
        
End Function




