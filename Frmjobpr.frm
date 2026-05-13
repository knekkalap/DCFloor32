VERSION 5.00
Begin VB.Form frmJobPrint 
   Caption         =   "Data Collection"
   ClientHeight    =   6555
   ClientLeft      =   5625
   ClientTop       =   930
   ClientWidth     =   8985
   ControlBox      =   0   'False
   Icon            =   "Frmjobpr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6555
   ScaleWidth      =   8985
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   8460
      Top             =   2925
   End
   Begin VB.ListBox lstJobSeq 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   540
      TabIndex        =   0
      Top             =   3780
      Width           =   7815
   End
   Begin VB.PictureBox MSCommPrinter 
      Height          =   480
      Left            =   7830
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Print Job Documents"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   2340
      TabIndex        =   4
      Top             =   1800
      Width           =   4830
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   1485
      Picture         =   "Frmjobpr.frx":0442
      Stretch         =   -1  'True
      Top             =   1665
      Width           =   780
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   135
      TabIndex        =   3
      Top             =   6075
      Width           =   90
   End
   Begin VB.Label Label2 
      Caption         =   $"Frmjobpr.frx":44D8
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1845
      TabIndex        =   2
      Top             =   2520
      Width           =   5055
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Sequence       WrkCtr       Operation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   540
      TabIndex        =   1
      Top             =   3465
      Width           =   7800
   End
End
Attribute VB_Name = "frmJobPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mintLoaded As Integer

'Private MyJobOpers()    As JobOpersType
Private colJobOpers      As New Collection
'Private myPrinter       As Object
Private myPrinter       As New CPrinter 'Instance of the Printer Object



Private Sub Disp_JobOpers(colJobOpers As Collection)
'==============================================================================
'Name : John Canny
'Date : 06/12/98
'Comments : Adds the records from the JobOpers Struct to the listbox
'Parameters : Public JobOpers(iNumRecs) As JobOpersType

'Returns : VOID
'==============================================================================
Dim sText           As String       'Text to put in List box
Dim sTmp            As String       'Temp Storage for Strings
Dim MyJobDoc        As Object


    frmJobPrint!lstJobSeq.Clear
    For Each MyJobDoc In colJobOpers
        sTmp = Str$(MyJobDoc.intSeq_Num)
        sText = sTmp & Space(7 - Len(sTmp)) & Space(2)
        sText = sText & MyJobDoc.strWrkCtr & Space(5 - Len(MyJobDoc.strWrkCtr)) & Space(2)
        sText = sText & MyJobDoc.strOperation & Space(30 - Len(MyJobDoc.strOperation))
        frmJobPrint!lstJobSeq.AddItem sText
        Set MyJobDoc = Nothing
    Next
    
End Sub


Public Function inpJobNum() As Boolean
'==============================================================================
'Name : John Canny
'Date : 06/15/98
'Comments : Prompts the User for a Job Number, Will display the results if valid
'Parameters : None
'Returns : False for Error, True for good to go
'==============================================================================
Dim strEntry        As String
Dim strDataId       As String
'Dim strSeqNo        As String
Dim sJobNo          As String       'Job Number / Work Order Number
Dim MyJobDoc        As Object       'Instance of the JobDoc info to be displayed _
                                    or Printed. Used to add two special Job Seq to Collection
    
    inpJobNum = False        'to force loop back till done
    
    Do Until inpJobNum Or gDone
        Do Until inpJobNum Or gDone       'Extra loop for exit do stmts
            
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 262)
            gEntry.EntryType = 0        'Any Printable Characters
            strEntry = basGetEntry()
            If gEntry.Escaped Then
                If myPrinter.StillPrinting = False Then
                    Set gEntry = Nothing
                    gDone = True
                    Unload Me
                    Exit Function
                End If
            End If
            Set gEntry = Nothing
           
           ' almost any key stroke will disable return timer at this point
            
            strDataId = UCase(Left(strEntry, 2))
            
            If Len(strDataId) < 1 Then
                Call ErrBox(LoadResString(gLanguageOffset + 78), 2)
                ErrorBeep
                Exit Do           'Loop Again
            End If
                
            If Len(strEntry) > 11 Then
'                If strDataId = "JS" Then
'                    gLaborIn(iEntry).OrderNo = UCase(Mid(strEntry, 3, 7))
'                    strSeqNo = Mid(strEntry, 10, 4)
'                Else
'                    ErrorBeep
'                    Exit Do           'Loop Again
'                End If
            Else
                sJobNo = UCase(Left(strEntry, 7))
            End If
            
            If GetJobMaster(sJobNo) Then
                'OK Lets get the Job_Opers Info
                Set colJobOpers = Nothing
                Set MyJobDoc = New CJObDoc
                
                'JC 09/25/98 Since there is a problem printing _
                to the Old Intermec printers, and they do not seem _
                to send a XOFF, probably due to the emulation mode _
                they are in, we can not send too much data to them _
                at one time. So for now commenting this next section _
                out
                
                'These are special Hard Coded Workeorders
                With MyJobDoc
                    'Allows the user to print all of them
                    .strPartNum = gJobMaster.PartNo
                    .strOrderNum = sJobNo
                    '.intSeq_Num =
                    .strWrkCtr = "ALL"
                    .strOperation = "ALL"
                End With
                colJobOpers.Add MyJobDoc
                Set MyJobDoc = Nothing
                
                Set MyJobDoc = New CJObDoc
                With MyJobDoc
                    'Allows the user to UnPlannedn Tag
                    .strPartNum = gJobMaster.PartNo
                    .strOrderNum = sJobNo
                    .intSeq_Num = 0
                    .strWrkCtr = Space(5)
                    .strOperation = "UnPlanned"
                End With
                colJobOpers.Add MyJobDoc
                Set MyJobDoc = Nothing
                
                If GetJobOpers(sJobNo, colJobOpers) Then
                    Call Disp_JobOpers(colJobOpers)
                    inpJobNum = True
                Else
                    Call ErrBox(LoadResString(gLanguageOffset + 63), 2)
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            Else
                Call ErrBox(LoadResString(gLanguageOffset + 61), 2)
                ErrorBeep
                Exit Do           'Loop Again
            End If
            
        Loop
    Loop
lblErrMsg.Caption = ""
End Function


Sub JobDoc_Print()
'==============================================================================
'Name : John Canny
'Date : 06/15/98
'Comments : Will print the selected job Sequence or all if the first one is selected
'Parameters : None
'Returns : Void
'==============================================================================
Dim intIndex    As Integer      'Index for the record selected in the List Box
Dim lngRc       As Long         'Var for checking Return Parms
Dim blnAllFound As Boolean      'Flag to indicate we have skipped / passed _
                                the first item in the collection which is the All Record
Dim MyJobDoc    As Object       'Instance of the Job Doc to print
'Dim myPrinter   As New CPrinter 'Instance of the Printer Object

'Set myPrinter = New CPrinter
    
    blnAllFound = False         'Assume we have not found the first All Item
    Me.MousePointer = vbHourglass
    lngRc = myPrinter.Printer_Find  'Check and Make Sure We have a valid printer
    If lngRc = vbError Then
        'Problem with the printer setup
        GoTo ExitPoint
    End If
    Me.MousePointer = vbDefault
    
    intIndex = lstJobSeq.ListIndex
    If intIndex < 0 Then
        'No Items Selected
        GoTo ExitPoint
    End If
    
    
    'JC 09/25/98 Since there is a problem printing _
    to the Old Intermec printers, and they do not seem _
    to send a XOFF, probably due to the emulation mode _
    they are in, we can not send too much data to them _
    at one time. So for now commenting this next section _
    out
    
    If intIndex = 0 Then
        'We have to Print All the Job Docs
        For Each MyJobDoc In colJobOpers
            If blnAllFound = True Then
                Me.MousePointer = vbHourglass
                lngRc = myPrinter.Print_Doc(MyJobDoc)   'Print the JobDoc
                Me.MousePointer = vbDefault
            Else
                blnAllFound = True
            End If
            Set MyJobDoc = Nothing
        Next
    Else
        'We have to Print only one.
        Set MyJobDoc = New CJObDoc
        intIndex = intIndex + 1
        MyJobDoc.strPartNum = gJobMaster.PartNo
        MyJobDoc.strOrderNum = colJobOpers(intIndex).strOrderNum
        MyJobDoc.intSeq_Num = colJobOpers(intIndex).intSeq_Num
        MyJobDoc.strWrkCtr = colJobOpers(intIndex).strWrkCtr
        MyJobDoc.strOperation = colJobOpers(intIndex).strOperation
        MyJobDoc.curSetStd = colJobOpers(intIndex).curSetStd
        MyJobDoc.lngRunStd = colJobOpers(intIndex).lngRunStd
        MyJobDoc.strEorSstd = colJobOpers(intIndex).strEorSstd
        
        Me.MousePointer = vbHourglass
        lngRc = myPrinter.Print_Doc(MyJobDoc)   'Print the JobDoc
        Me.MousePointer = vbDefault
        Set MyJobDoc = Nothing
    End If

ExitPoint:
Me.MousePointer = vbDefault
Set MyJobDoc = Nothing

End Sub


Private Sub Form_Activate()

Dim brc     As Boolean

    'frmJobPrint.lstJobSeq.ListIndex = 0
    gSelIndx = 0
    
    If gDone Then
        Unload Me
    End If
    
    If mintLoaded <> 1 Then
        mintLoaded = 1
        brc = inpJobNum()
        'mintLoaded = 0
        'Unload Me
    End If
    

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Const iPage = 9     'The numnber of items to skip through a list box _
                    When the user presses the PageUp or PageDown
Dim iAt As Integer

    gSelIndx = frmJobPrint.lstJobSeq.ListIndex
    
    If myPrinter.StillPrinting = False Then
        'If We are done Printing You may go
        
        Select Case KeyCode
        
            Case vbKeyEscape
                Unload Me
                Exit Sub
                
            Case vbKeyReturn, vbKeyF1
                'Do Nothing
                
            Case vbKeyF6
                'Print the Selected item
                If lstJobSeq.ListIndex >= 0 Then
                    JobDoc_Print
                Else
                    Call ErrBox(LoadResString(gLanguageOffset + 70), 2)
                    ErrorBeep
                End If
                
            Case vbKeyDown
                'If gSelIndx < frmJobPrint.lstJobSeq.ListCount - 1 Then
                '    frmJobPrint.lstJobSeq.ListIndex = gSelIndx + 1
                '    gSelIndx = gSelIndx + 1
                'Else
                '    frmJobPrint.lstJobSeq.ListIndex = frmJobPrint.lstJobSeq.ListCount - 2
                '    gSelIndx = frmJobPrint.lstJobSeq.ListCount - 2
                'End If
                'KeyCode = 0
                
            Case vbKeyUp
            '    If gSelIndx > 0 Then
            '        frmJobPrint.lstJobSeq.ListIndex = gSelIndx - 1
            '        gSelIndx = gSelIndx - 1
            '    Else
            '        frmJobPrint.lstJobSeq.ListIndex = 0
            '        gSelIndx = 0
            '    End If
            '    KeyCode = 0
                
            Case vbKeyPageDown
            '    If gSelIndx < frmJobPrint.lstJobSeq.ListCount - iPage - 1 Then
            '        frmJobPrint.lstJobSeq.ListIndex = gSelIndx + iPage
            '        gSelIndx = gSelIndx + iPage
            '    Else
            '        frmJobPrint.lstJobSeq.ListIndex = frmJobPrint.lstJobSeq.ListCount - 2
            '        gSelIndx = frmJobPrint.lstJobSeq.ListCount - 2
            '    End If
            '    KeyCode = 0
                
            Case vbKeyPageUp
            '    If gSelIndx > iPage Then
            '        frmJobPrint.lstJobSeq.ListIndex = gSelIndx - iPage
            '       gSelIndx = gSelIndx - iPage
            '    Else
            '        frmJobPrint.lstJobSeq.ListIndex = 0
            '        gSelIndx = 0
            '    End If
            '    KeyCode = 0
                
            'Case Else
            '    'Don't let anything else happen to field
            '    KeyCode = 0
        End Select
    Else
        KeyCode = 0
    End If


End Sub

Private Sub Form_Load()

mintLoaded = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set colJobOpers = Nothing
    Set myPrinter = Nothing
End Sub


Private Sub MSCommPrinter_OnComm()
    DoEvents
End Sub



Private Sub Timer1_Timer()

    Timer1.Enabled = False
    DoEvents
    myPrinter.StillPrinting = False
    Beep
    
End Sub


