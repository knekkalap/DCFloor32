VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLaborChk 
   Caption         =   "Data Collection Employee Daily Summary"
   ClientHeight    =   8220
   ClientLeft      =   2340
   ClientTop       =   570
   ClientWidth     =   11130
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8220
   ScaleWidth      =   11130
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   2415
      Left            =   0
      TabIndex        =   21
      Top             =   2640
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   65535
   End
   Begin VB.Label lblESCCaption 
      AutoSize        =   -1  'True
      Caption         =   "Press ESC key to Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   20
      Top             =   6030
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   135
      Picture         =   "Laborchk.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   870
   End
   Begin VB.Label lblAttdStat 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   3015
      TabIndex        =   19
      Top             =   2250
      Width           =   4170
   End
   Begin VB.Label lblLunchEnd 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   6750
      TabIndex        =   18
      Top             =   1755
      Width           =   60
   End
   Begin VB.Label lblClockOut 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   6750
      TabIndex        =   17
      Top             =   1350
      Width           =   75
   End
   Begin VB.Label lblLunchStart 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   2400
      TabIndex        =   16
      Top             =   1755
      Width           =   60
   End
   Begin VB.Label lblClockIn 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   2400
      TabIndex        =   15
      Top             =   1350
      Width           =   60
   End
   Begin VB.Label lblLECaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Lunch End:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   5280
      TabIndex        =   14
      Top             =   1755
      Width           =   1350
   End
   Begin VB.Label lblLSCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Lunch Start:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   870
      TabIndex        =   13
      Top             =   1755
      Width           =   1425
   End
   Begin VB.Label lblCOCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Clock Out:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   5430
      TabIndex        =   12
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label lblCICaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Clock In:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1275
      TabIndex        =   11
      Top             =   1350
      Width           =   1020
   End
   Begin VB.Label lblInst3Caption 
      AutoSize        =   -1  'True
      Caption         =   "Use Up- and Down-arrow keys to review all bookings "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   5715
      Width           =   6090
   End
   Begin VB.Label lblInst2Caption 
      AutoSize        =   -1  'True
      Caption         =   "Use Left- and Right-arrow keys to scroll across entire row"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   6540
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   8
      Top             =   6600
      Width           =   45
   End
   Begin VB.Label lblInst1Caption 
      AutoSize        =   -1  'True
      Caption         =   "While Clocked In, Att Hrs equals current time minus Clock-In time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   5085
      Width           =   7365
   End
   Begin VB.Label lblShift 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   9210
      TabIndex        =   6
      Top             =   810
      Width           =   60
   End
   Begin VB.Label lblShiftCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Shift:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   8505
      TabIndex        =   5
      Top             =   810
      Width           =   615
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   7095
      TabIndex        =   4
      Top             =   810
      Width           =   60
   End
   Begin VB.Label LblDateCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   6360
      TabIndex        =   3
      Top             =   810
      Width           =   615
   End
   Begin VB.Label lblEmployee 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   1800
      TabIndex        =   2
      Top             =   810
      Width           =   60
   End
   Begin VB.Label lblEmplCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Employee:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   465
      TabIndex        =   1
      Top             =   810
      Width           =   1230
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Data Collection Employee Activity Review"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   480
      Left            =   1200
      TabIndex        =   0
      Top             =   135
      Width           =   8520
   End
End
Attribute VB_Name = "frmLaborChk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mintLoaded As Integer

Private Sub Form_Activate()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmLaborChk.Form_Activate
'Description      : When form is activated for the first time after the
'                   form is loaded, the initialization process is called.
'Parameters       : None
'Called From      : Activate Event for form
'Procedures Used  : LaborChkProcess
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    'If gDone set, then unload forms all the way back to frmDCMain
    If gDone Then
        Unload Me
    End If
    
    'Only do this once, the first time after the form is loaded
    If mintLoaded <> 1 Then
        'Prevent this from happening again, until form is unloaded
        mintLoaded = 1
        Call LaborChkProcess
        'Don't unload now, we will wait for keystroke events to decide what
        '   to do.  This is an event driven form, see events for the
        '   frmLaborChk.Form object(Keydown).
        'mintLoaded = 0
        'Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmLaborChk.Form_KeyDown
'Description      : Process the Esc keystroke event.  Scrolling events will be
'                   handled by the grid control automatically.  The escape key
'                   will be set to unload form.  KeyPreview property is set
'                   on the form, so that the form captures keystrokes first.
'Parameters       : KeyCode and Shift indicator
'Called From      : KeyDown Event
'Procedures Used  : None
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    If KeyCode = vbKeyEscape Then
        Unload Me
        mintLoaded = 0
        Exit Sub
    End If
End Sub


Public Sub LaborChkProcess()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmLaborChk.LaborChkProcess
'Description      : Form to display Labor information for the requested
'                   employee number, for the last day of activity in the
'                   database for that employee.
'                   Initialize the Labor Check form.  The Grid will have it's
'                   column headers set and column widths set.  The Employee
'                   number will be requested, then the employee information
'                   will be displayed.  The Labor information will be filled
'                   into the grid control.  The form will be left in event
'                   driven mode, to allow viewing and scrolling, until done.
'                   All  actions will be done by keystroke events after this.
'                   KeyPreview property is set on the form, so that the form
'                   captures keystrokes first.  Keystroke events are controlled
'                   in the frmLaborChk.Form_Keydown Event handling function.
'Parameters       : None
'Called From      : frmLaborChk.Form_Activate (only first time after load)
'Procedures Used  : gEntry Object, basGetEntry, Valid_Empl_badge,
'                   Ck_for_Attd_Rec, RndTimeHnds, Calc_Elapsed_time,
'                   Calc_Attd_Hrs, InitPaEmplyData, GetPaEmplyData
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    Dim datestr As String
    Dim strRow As String
    Dim strEntryEmpNo As String
    Dim bGotOne As Boolean
    Dim iWid As Integer
    Dim cCurTime As Currency
    Dim iLaborRecs As Integer
    Dim I As Integer
    Dim lTotIndHrs As Currency
    Dim lTotSetHrs As Currency
    Dim lTotMachHrs As Currency
    Dim lTotDirHrs As Currency
    
    
    Rem Column Titles and Column Width setup
    iWid = Grid1.Width / 60    'twips per char as % of width
                                ' less apx scroll bar width
    Grid1.Cols = 15
    Grid1.Font = "Arial Bold"
    Grid1.Row = 0
    
    Grid1.Col = 0
    Grid1.ColWidth(0) = 350
    Grid1.Text = LoadResString(gLanguageOffset + 248)
    Grid1.ColAlignment(0) = 4
    
    Grid1.Col = 1
    Grid1.ColWidth(1) = 850
    Grid1.Text = LoadResString(gLanguageOffset + 249)
    Grid1.ColAlignment(1) = 1
    
    Grid1.Col = 2
    Grid1.ColWidth(2) = 650
    Grid1.Text = LoadResString(gLanguageOffset + 250)
    Grid1.ColAlignment(2) = 1
    
    Grid1.Col = 3
    Grid1.ColWidth(3) = 450
    Grid1.Text = LoadResString(gLanguageOffset + 251)
    Grid1.ColAlignment(3) = 4
    
    Grid1.Col = 4
    Grid1.ColWidth(4) = 650
    Grid1.Text = LoadResString(gLanguageOffset + 252)
    Grid1.ColAlignment(4) = 1
    
    Grid1.Col = 5
    Grid1.ColWidth(5) = 650
    Grid1.Text = LoadResString(gLanguageOffset + 253)
    Grid1.ColAlignment(5) = 1
    
    Grid1.Col = 6
    Grid1.ColWidth(6) = 850
    Grid1.Text = LoadResString(gLanguageOffset + 254)
    Grid1.ColAlignment(6) = 1
    
    Grid1.Col = 7
    Grid1.ColWidth(7) = 850
    Grid1.Text = LoadResString(gLanguageOffset + 255)
    Grid1.ColAlignment(7) = 1
    
    Grid1.Col = 8
    Grid1.ColWidth(8) = 700
    Grid1.ColAlignment(8) = 1
    If div = singapore Or div = pv_singapore Then
        Grid1.Text = "Direct"
    Else
        Grid1.Text = LoadResString(gLanguageOffset + 256)
    End If
    
    Grid1.Col = 9
    Grid1.ColWidth(9) = 700
    Grid1.Text = LoadResString(gLanguageOffset + 257)
    Grid1.ColAlignment(9) = 1
       
    Grid1.Col = 10
    Grid1.ColWidth(10) = 700
    Grid1.Text = LoadResString(gLanguageOffset + 258)
    Grid1.ColAlignment(10) = 1
   
    Grid1.Col = 11
    Grid1.ColWidth(11) = 750
    Grid1.ColAlignment(11) = 1
    If div = singapore Or div = pv_singapore Then
        Grid1.Text = "M/C"
   Else
        Grid1.Text = LoadResString(gLanguageOffset + 259)
   End If
  
  If div = corry Then
    Grid1.Col = 12
    Grid1.ColWidth(12) = 500
    Grid1.Text = "Inc Cd"
    Grid1.ColAlignment(12) = 4
  Else
    Grid1.Col = 12
    Grid1.ColWidth(12) = 600
    Grid1.Text = LoadResString(gLanguageOffset + 260)
    Grid1.ColAlignment(12) = 1
  End If
   
  If div <> stroms Then
    Grid1.Col = 13
    Grid1.ColWidth(13) = 1
    Grid1.Col = 14
    Grid1.ColWidth(14) = 1
  Else
    Grid1.Col = 13
    Grid1.ColWidth(13) = 1200
    Grid1.Text = "Started"
    Grid1.ColAlignment(13) = 1
    Grid1.Col = 14
    Grid1.ColWidth(14) = 1200
    Grid1.Text = "Cmpl"
    Grid1.ColAlignment(14) = 1
  End If
  
  Grid1.Width = Grid1.ColWidth(0) + Grid1.ColWidth(1) + Grid1.ColWidth(2) + Grid1.ColWidth(3) + _
                Grid1.ColWidth(4) + Grid1.ColWidth(5) + Grid1.ColWidth(6) + Grid1.ColWidth(7) + _
                Grid1.ColWidth(8) + Grid1.ColWidth(9) + Grid1.ColWidth(10) + Grid1.ColWidth(11) + _
                Grid1.ColWidth(12) + Grid1.ColWidth(13) + Grid1.ColWidth(14)
        
    bGotOne = False
    lblEmployee.Caption = gEmployee.EmplNo & " " _
        & gEmployee.EmplName
    
    If Ck_for_Attd_Rec(gEmployee.EmplNo, _
            gEmployee.Actdate) = True Then
        If gAttendance.ClockIn <> dBlankDate Then
            lblClockIn.Caption = Format(gAttendance.ClockIn, gDateformat) & "   " & _
                                          Format(gAttendance.ClockIn, gTimeFormat)
       Else
            lblClockIn.Caption = ""
       End If
       If gAttendance.ClockOut <> dBlankDate Then
            lblClockOut.Caption = Format(gAttendance.ClockOut, gDateformat) & "   " & _
                                            Format(gAttendance.ClockOut, gTimeFormat)
       Else
            lblClockOut.Caption = ""
       End If
       If gAttendance.LunchStart <> dBlankDate Then
            lblLunchStart.Caption = Format(gAttendance.LunchStart, gDateformat) & "   " & _
                                                 Format(gAttendance.LunchStart, gTimeFormat)
       Else
            lblLunchStart.Caption = ""
       End If
       If gAttendance.LunchEnd <> dBlankDate Then
            lblLunchEnd.Caption = Format(gAttendance.LunchEnd, gDateformat) & "   " & _
                                               Format(gAttendance.LunchEnd, gTimeFormat)
       Else
            lblLunchEnd.Caption = ""
       End If
       
       Select Case gEmployee.AttdStatus
        Case 1: lblAttdStat.Caption = LoadResString(gLanguageOffset + 227)
        Case 2: lblAttdStat.Caption = LoadResString(gLanguageOffset + 228)
        Case 3: lblAttdStat.Caption = LoadResString(gLanguageOffset + 229)
        Case 4: lblAttdStat.Caption = LoadResString(gLanguageOffset + 230)
        Case 5: lblAttdStat.Caption = LoadResString(gLanguageOffset + 231)
        Case 6: lblAttdStat.Caption = LoadResString(gLanguageOffset + 232)
        Case 7: lblAttdStat.Caption = LoadResString(gLanguageOffset + 233)
        Case 8: lblAttdStat.Caption = LoadResString(gLanguageOffset + 228)
        Case Else
                    lblAttdStat.Caption = ""
        End Select
        
        lblDate.Caption = Format(gAttendance.Actdate, gDateformat)
        lblShift.Caption = CStr(gAttendance.ShiftCd)
        
        Grid1.Row = 1
        Grid1.Col = 4
        Grid1.Text = gAttendance.CCID
        Grid1.Col = 8       'Att Hrs
        Grid1.Text = Format(0, gOneDecPos)
        If div <> singapore And div <> pv_singapore Then
            If (gAttendance.ActHrs = 0 Or _
                    gAttendance.ActHrs = Null) Then
                gAttendance.ClockOutRnd = RndTimeHnds(gSysTimeStrp)
                If (gAttendance.ClockOutRnd <= gAttendance.ClockInRnd) Then
                      Grid1.Text = Format(0, "##.0")
                ElseIf (gAttendance.ClockOutRnd <= gAttendance.LunchOutRnd) And _
                           gAttendance.ClockOutRnd <> INIT_RND_TM Then
                                   Grid1.Text _
                                    = Format(Calc_Elapsed_time(gAttendance.ClockInRnd, _
                                      gAttendance.ClockOutRnd), "#0.0")
                Else
                    Calc_Attd_Hrs
                    Grid1.Text = Format(gAttendance.ActHrs, "#0.0")
                End If
            Else
                Grid1.Text = Format(gAttendance.ActHrs, "#0.0")
            End If
       End If 'div <> singapore
    Else
        lblDate.Caption = Format(gEmployee.Actdate, gDateformat)
        lblShift.Caption = CStr(gEmployee.AsgnShift)
        lblClockIn.Caption = ""
        lblClockOut.Caption = ""
        lblLunchStart.Caption = ""
        lblLunchEnd.Caption = ""
        Grid1.Row = 1
        Grid1.Col = 4
        Grid1.Text = gEmployee.HomeCCID
        Grid1.Col = 8
        Grid1.Text = Format(0, gOneDecPos)
        
        gEmployee.Actdate = gSysDate
    End If
    
    Call InitPaEmplyData
        
    lTotIndHrs = 0
    lTotSetHrs = 0
    lTotMachHrs = 0
    lTotDirHrs = 0
    I = 0
    iLaborRecs = GetPaEmplyData()
    While I <= iLaborRecs - 1 And I < 250
        strRow = gPaEmplyData(I).RecType
        
        
        strRow = strRow & Chr(9) & gPaEmplyData(I).OrderNo
        If (gPaEmplyData(I).RecType = "I" Or gPaEmplyData(I).RecType = "C" Or _
            (gPaEmplyData(I).RecType = "O" And gPaEmplyData(I).IndirCd <> "  ")) Then
            strRow = strRow & Chr(9) & ""
        Else
            strRow = strRow & Chr(9) & gPaEmplyData(I).SeqNo
        End If
        strRow = strRow & Chr(9) & gPaEmplyData(I).IndirCd
        strRow = strRow & Chr(9) & gPaEmplyData(I).CCID
        strRow = strRow & Chr(9) & gPaEmplyData(I).WCID
        
        If (gPaEmplyData(I).RecType = "I" Or gPaEmplyData(I).RecType = "C" Or _
            gPaEmplyData(I).RecType = "S") Then
            strRow = strRow & Chr(9) & ""
        ElseIf gPaEmplyData(I).RecType = "O" Then
            strRow = strRow & Chr(9) & LoadResString(gLanguageOffset + 261)
        Else
            strRow = strRow & Chr(9) & gPaEmplyData(I).ActQty
        End If
        
        If (gPaEmplyData(I).RecType = "I" Or gPaEmplyData(I).RecType = "C" Or _
             gPaEmplyData(I).RecType = "O") Then
            strRow = strRow & Chr(9) & ""
        Else
            strRow = strRow & Chr(9) & gPaEmplyData(I).scrapQty
        End If
        If gPaEmplyData(I).RecType = "S" Then
           GoTo NextI
        End If
        
        If div = singapore Then
            If Left(gPaEmplyData(I).Oper_Cd, 3) <> "120" And _
               Left(gPaEmplyData(I).Oper_Cd, 3) <> "110" And _
               Left(gPaEmplyData(I).Oper_Cd, 3) <> "100" Then
               If gPaEmplyData(I).RecType = "F" Then
                   strRow = strRow & Chr(9) & Format(gPaEmplyData(I).ActHrs, "##0.0")
                   lTotDirHrs = lTotDirHrs + gPaEmplyData(I).ActHrs
               Else
                   strRow = strRow & Chr(9) & Format(gPaEmplyData(I).MachHrs, "##0.0")
                   lTotDirHrs = lTotDirHrs + gPaEmplyData(I).MachHrs
               End If
            Else
               strRow = strRow & Chr(9) & ""
            End If
        ElseIf div = pv_singapore Then
            If Left(gPaEmplyData(I).Oper_Cd, 3) <> "110" And _
             Left(gPaEmplyData(I).Oper_Cd, 3) <> "880" And _
             Left(gPaEmplyData(I).Oper_Cd, 3) <> "890" Then
               If gPaEmplyData(I).RecType = "F" Then
                   strRow = strRow & Chr(9) & Format(gPaEmplyData(I).ActHrs, "##0.0")
                   lTotDirHrs = lTotDirHrs + gPaEmplyData(I).ActHrs
               Else
                   strRow = strRow & Chr(9) & Format(gPaEmplyData(I).MachHrs, "##0.0")
                   lTotDirHrs = lTotDirHrs + gPaEmplyData(I).MachHrs
               End If
            Else
               strRow = strRow & Chr(9) & ""
            End If
        Else 'all other divisions
            strRow = strRow & Chr(9) & ""   'att hrs column
        End If
        
        If gPaEmplyData(I).RecType = "I" Then
           strRow = strRow & Chr(9) & Format(gPaEmplyData(I).ActHrs, "##0.0") & Chr(9) & "" _
                     & Chr(9) & ""
            lTotIndHrs = lTotIndHrs + gPaEmplyData(I).ActHrs
        ElseIf (gPaEmplyData(I).RecType = "F" Or gPaEmplyData(I).RecType = "C") Then
           strRow = strRow & Chr(9) & "" & Chr(9) & Format(gPaEmplyData(I).ActHrs, "##0.0") _
                     & Chr(9) & ""
           lTotSetHrs = lTotSetHrs + gPaEmplyData(I).ActHrs
        ElseIf gPaEmplyData(I).RecType = "D" Then
            If div = singapore Then
                If Left(gPaEmplyData(I).Oper_Cd, 3) = "120" Or _
                   Left(gPaEmplyData(I).Oper_Cd, 3) = "110" Or _
                   Left(gPaEmplyData(I).Oper_Cd, 3) = "100" Then
                   strRow = strRow & Chr(9) & "" & Chr(9) & "" & Chr(9) _
                             & Format(gPaEmplyData(I).MachHrs, "##0.0")
                   lTotMachHrs = lTotMachHrs + gPaEmplyData(I).MachHrs
                Else
                   strRow = strRow & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                End If
            ElseIf div = pv_singapore Then
                If Left(gPaEmplyData(I).Oper_Cd, 3) = "110" Or _
                   Left(gPaEmplyData(I).Oper_Cd, 3) = "880" Or _
                   Left(gPaEmplyData(I).Oper_Cd, 3) = "890" Then
                   strRow = strRow & Chr(9) & "" & Chr(9) & "" & Chr(9) _
                             & Format(gPaEmplyData(I).MachHrs, "##0.0")
                   lTotMachHrs = lTotMachHrs + gPaEmplyData(I).MachHrs
                Else
                   strRow = strRow & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                End If
            Else
                strRow = strRow & Chr(9) & "" & Chr(9) & "" & Chr(9) _
                          & Format(gPaEmplyData(I).MachHrs, "##0.0")
                 lTotMachHrs = lTotMachHrs + gPaEmplyData(I).MachHrs
            End If
        Else
           strRow = strRow & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
        End If
        If div <> corry Then
           If (gPaEmplyData(I).RecType = "I" Or gPaEmplyData(I).RecType = "O" Or _
              gPaEmplyData(I).RecType = "S") Then
               strRow = strRow & Chr(9) & ""
           Else
               strRow = strRow & Chr(9) & gPaEmplyData(I).EmplEff
           End If
        Else
        If (gPaEmplyData(I).RecType = "I" Or gPaEmplyData(I).RecType = "C" Or _
             gPaEmplyData(I).RecType = "O") Then
               strRow = strRow & Chr(9) & ""
           Else
               strRow = strRow & Chr(9) & gPaEmplyData(I).IncenCd
           End If
        End If
        If (gPaEmplyData(I).RecType = "O") Then
              strRow = strRow & Chr(9) & gPaEmplyData(I).transdate
        Else
              strRow = strRow & Chr(9) & ""
        End If
        If (gPaEmplyData(I).RecType <> "O") Then
              strRow = strRow & Chr(9) & gPaEmplyData(I).transdate
'        Else
 '             strRow = strRow & Chr(9) & ""
        End If
NextI:
        I = I + 1
        Grid1.AddItem strRow, I + 1
    Wend

'write out totals to appropriate columns
    Grid1.Row = 1
    Grid1.Col = 8
    If lTotDirHrs > 0 And (div = singapore Or div = pv_singapore) Then
        Grid1.Text = Format(lTotDirHrs, "##0.0")
    End If
    Grid1.Row = 1
    Grid1.Col = 9
    If lTotIndHrs > 0 Then
        Grid1.Text = Format(lTotIndHrs, "##0.0")
   End If
   Grid1.Col = 10
   If lTotSetHrs > 0 Then
        Grid1.Text = Format(lTotSetHrs, "##0.0")
   End If
   Grid1.Col = 11
   If lTotMachHrs > 0 Then
       Grid1.Text = Format(lTotMachHrs, "##0.0")
   End If
    
End Sub


Private Sub Form_Load()
'Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
'MyForm.Height = Me.Height ' Remember the current size
'MyForm.Width = Me.Width


lblTitle.Caption = LoadResString(gLanguageOffset + 279)
lblEmplCaption.Caption = LoadResString(gLanguageOffset + 278)
lblDateCaption.Caption = LoadResString(gLanguageOffset + 114)
lblShiftCaption.Caption = LoadResString(gLanguageOffset + 277)
lblCICaption.Caption = LoadResString(gLanguageOffset + 280)
lblCOCaption.Caption = LoadResString(gLanguageOffset + 281)
lblLSCaption.Caption = LoadResString(gLanguageOffset + 283)
lblLECaption.Caption = LoadResString(gLanguageOffset + 282)
lblInst1Caption.Caption = LoadResString(gLanguageOffset + 284)
lblInst2Caption.Caption = LoadResString(gLanguageOffset + 285)
lblInst3Caption.Caption = LoadResString(gLanguageOffset + 286)
lblESCCaption.Caption = LoadResString(gLanguageOffset + 144)

End Sub

