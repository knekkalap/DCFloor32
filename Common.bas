Attribute VB_Name = "Common"
Option Explicit

Public gEntry As clsEntry


Public Function isLetter(ByVal strchar As String) As Boolean
strchar = Trim(UCase(strchar))

If strchar = "A" Or strchar = "B" Or strchar = "C" Or strchar = "D" Or _
   strchar = "E" Or strchar = "F" Or strchar = "G" Or strchar = "H" Or _
   strchar = "I" Or strchar = "J" Or strchar = "K" Or strchar = "L" Or _
   strchar = "M" Or strchar = "N" Or strchar = "O" Or strchar = "P" Or _
   strchar = "Q" Or strchar = "R" Or strchar = "S" Or strchar = "T" Or _
   strchar = "U" Or strchar = "V" Or strchar = "W" Or strchar = "X" Or strchar = "Y" Or strchar = "Z" Then
        isLetter = True
Else
        isLetter = False
End If

End Function

Public Function QuoteSet(ByVal strQuote) As String

'Parses given string for single quotes. Transact sqlStr will fail when there
'is a single quote in the middle of a string such as an apostrophe.
'As a result, another single quote must be added before the existing
'single quote.  2 single quotes together will be interpreted
'as 1 single quote by TSQL.

    Dim iLen As Integer
    Dim iPos As Integer
        
If strQuote = "" Or IsNull(strQuote) Then
    QuoteSet = " "
    Exit Function
End If

QuoteSet = strQuote
iLen = Len(QuoteSet)
iPos = 1
Do
    iPos = InStr(iPos, QuoteSet, "'")
    If iPos > 0 Then
        QuoteSet = Left(QuoteSet, iPos) & "'" & Right(QuoteSet, iLen - iPos)
        iLen = iLen + 1
        iPos = iPos + 2
     End If
Loop Until iPos = 0 Or iPos > iLen

End Function

Public Function basGetEntry() As String
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Common.basGetEntry
'Description      : Function to initiate the data entry process using the
'                   gEntry object and the frmDCEntry Form.  In the menu
'                   process, returns the code. In the entry process, returns
'                   the entered text.  In the Code lookup process, returns
'                   the code selected from the collection supplied, or a null
'                   string if the collection is empty.
'Parameters       : None
'Called From      : frmInpEmplNo.basEntry, frmDirLabor.inpLabJobSeq,
'                   frmDirLabor.DirLaborReEnter, frmNotepad.NotepadProcess,
'                   frmLaborChk.LaborChkProcess, frmJobPrint.InpJobNum,
'                   Inputs.inpEndOfShift, Inputs.AnotherRec,
'                   Inputs.inpJobAndSeq, Inputs.inpWcAns, Inputs.inpWCID,
'                   Inputs.inpAddonHours, Inputs.inpIndirHrs,
'                   Inputs.inpProdQty, Inputs.CallCmpl, Inputs.ReEnter,
'                   Inputs.inpNumbMachs, Inputs.inpMchHours,
'                   Inputs.inpSetHours, Inputs.inpOperNumb,
'                   Inputs.inpUnplnReas, Inputs.inpNoOfLoads,
'                   Inputs.inpCycleTime, Inputs.getBatchIncn,
'                   Inputs.getIndirIncn, Inputs.inpCCID, Inputs.inpEOJob,
'                   Inputs.inpPMJob, Inputs.inpIndirCd, Inputs.inpTlmntCd,
'                   Inputs.inpComment, Inputs.inpFixStd, Inputs.inpMchStd,
'                   Inputs.inpJobAndSeqStds, Inputs.inpEmplNo,
'Objects Used     : gEntry
'Forms Loaded     : frmDCEntry
'Return Parameters: string from entry system.
'Tables Updates   : None
'==============================================================================
'    Menu Example
'    Set gEntry = New clsEntry      'put a new clsEntry object into gEntry
'    gEntry.Title = "What is it?"   'Prompt for the menu
'    gEntry.Add "A", "An Option"    'First menu option
'    gEntry.Add "B", "But Another"  'Second menu option
'    strEntry = basGetEntry()       'prompt the user for the menu selection
'    If gEntry.Escaped Then         'detect if esc key was hit
'        'Some code to handle escape
'        Set gEntry = Nothing       'always set use count to zero on gEntry
'        Exit Function
'    End If
'    Set gEntry = Nothing           'always set use count to zero on gEntry
'    'do something with strEntry (the code A or B from Menu)

'   Other options
' -     If Menu has a line like:
'           gEntry.AddScanItem "Y"
'       Then when the letter Y is the (first) letter entered, then
'           the user will be allowed to enter more letters until return.
'           This will allow scanning of a string beginning with Y.
'
' -     If no .Add functions are used, then there will be no menu.
'
' -     If gEntry has a line like:
'           gEntry.LuTable = colCodes
'       Where colCodes is an object of type clsCodeTable that has had
'           Codes and Descriptions added to it with the .Add method.
'       Then the Entry system will bring up a combo box populated with
'           the codes from the colCodes Collection Class.  Only values
'           in colCodes will be allowed to be entered, except when esc
'           has been hit (see .Escaped), or when there are no entries
'           in the collection, where it will return with a null string ""
'
' -     If Neither  the .Add or .LuTable methods have been used, then
'           the user will be given a freeform text box to enter text.
'
' -     If gEntry has a line like:
'           gEntry.EntryType = 1
'       Then Only numeric digits will be allowed, and all others will be
'           ignored.
'
' -     If gEntry has a line like:
'           gEntry.MainEsc = True
'       Then the application will be shut down when the user enters
'           the Ctrl-Alt-X key combination.
    
    Rem Code to prevent empty lookup collection to start free form entry
    Rem     Will return a null string if no possible choices
If gEntry.IsLookup And gEntry.CountLookupItems = 0 Then
        basGetEntry = ""
        Exit Function
    End If
    frmDCEntry.Show 1
    basGetEntry = gEntry.Response
    Unload frmDCEntry
End Function

Public Function basGetEntryValue() As Variant
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Common.basGetEntryValue
'Description      : Function to initiate the data entry process using the
'                   gEntry object and the frmDCEntry Form.  In the menu
'                   process, returns the value associated with a menu entry
'                   (specified in the optional argument in the .Add function),
'                   unless there is no value associated with it, in which case
'                   the code will be returned as a string as in basGetEntry().
'                   In the free form entry process, returns the entered text.
'                   In the Code lookup process, returns the code selected from
'                   the collection supplied, or a null string if the collection
'                   is empty.
'                   This would have been part of the Activex Document if VB5.
'Parameters       : None
'Called From      : frmLabor.basEntry, frmDCMain.basEntry
'Objects Used     : gEntry
'Forms Loaded     : frmDCEntry
'Return Parameters: string from entry system.
'Tables Updates   : None
'==============================================================================
'
'See basGetEntryValue for examples of using the entry system
'

    Rem Code to prevent empty combo box / free form entry
    Rem     Will return a null string if no possible choices

    If gEntry.IsLookup And gEntry.CountLookupItems = 0 Then
        basGetEntryValue = ""
        Exit Function
    End If
    frmDCEntry.Show 1
    basGetEntryValue = gEntry.ResponseValue
'    GetServerTime
    Unload frmDCEntry
End Function


Public Sub ErrorBeep()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Common.ErrorBeep
'Description      : Function to send a sound to the terminal.
'                   Currently sends one beep.
'Parameters       : None
'Procedures Used  : None
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
' Function to send a sound to the terminal.  Currently sends one beep
    Beep
End Sub

Public Function checkClockOut()
    Dim lBeginTime As Long, lEndTime As Long
    Dim lClockOutMins As Long
    
    gEosFlag = False
    lEndTime = ConvTimeToSecs(gShiftRecData(gEmployee.GenShift).endTime)
    gSysTimeSecs = ConvTimeToSecs(gSysTimeStrp)
    
    lClockOutMins = gClockOutMins * 60
    lBeginTime = lEndTime - lClockOutMins
    
    Rem begin time holds the earliest time
    Rem end time holds the end of shift time
    
    If lBeginTime < 0 Then
        lBeginTime = lBeginTime + 86400
    End If
    
    If lBeginTime > lEndTime Then
        If ((gSysTimeSecs >= lBeginTime And gSysTimeSecs <= 86400) _
          Or (gSysTimeSecs <= lEndTime And gSysTimeSecs >= 0)) Then
            gEosFlag = True
        End If
    Else
        If gSysTimeSecs >= lBeginTime _
          And gSysTimeSecs <= lEndTime Then
            gEosFlag = True
        End If
    End If
    
End Function
