VERSION 5.00
Begin VB.Form frmDCMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DC SHOP FLOOR"
   ClientHeight    =   8640
   ClientLeft      =   3060
   ClientTop       =   1935
   ClientWidth     =   8970
   ControlBox      =   0   'False
   Icon            =   "Dcmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8640
   ScaleWidth      =   8970
   WindowState     =   2  'Maximized
   Begin VB.Timer ReturnTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7830
      Top             =   6210
   End
   Begin VB.Timer tmrDCMain 
      Interval        =   1000
      Left            =   8460
      Top             =   6210
   End
   Begin VB.Label lblSysError 
      Alignment       =   2  'Center
      Caption         =   "Notify MIS of Error"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Label lblDCMenuSel 
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
      Height          =   300
      Index           =   11
      Left            =   3000
      TabIndex        =   18
      Top             =   6200
      Width           =   1005
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   960
      Left            =   990
      Picture         =   "Dcmain.frx":0442
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   960
   End
   Begin VB.Label lblDCMenuSel 
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
      Height          =   300
      Index           =   10
      Left            =   3000
      TabIndex        =   17
      Top             =   5880
      Width           =   1005
   End
   Begin VB.Label lblSysErrorDet 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4395
      TabIndex        =   16
      Top             =   270
      Width           =   60
   End
   Begin VB.Label lblDCMenuSel 
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
      Height          =   300
      Index           =   9
      Left            =   3000
      TabIndex        =   15
      Top             =   5600
      Width           =   1005
   End
   Begin VB.Label lblDCMenuSel 
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
      Height          =   300
      Index           =   8
      Left            =   3000
      TabIndex        =   14
      Top             =   5300
      Width           =   1005
   End
   Begin VB.Label lblDCMenuSel 
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
      Height          =   300
      Index           =   7
      Left            =   3000
      TabIndex        =   13
      Top             =   5000
      Width           =   1005
   End
   Begin VB.Label lblDCMenuSel 
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
      Height          =   300
      Index           =   6
      Left            =   3000
      TabIndex        =   12
      Top             =   4700
      Width           =   1005
   End
   Begin VB.Label lblDCMenuSel 
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
      Height          =   300
      Index           =   5
      Left            =   3000
      TabIndex        =   11
      Top             =   4400
      Width           =   1005
   End
   Begin VB.Label lblDCMenuSel 
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
      Height          =   300
      Index           =   4
      Left            =   3000
      TabIndex        =   10
      Top             =   4100
      Width           =   1005
   End
   Begin VB.Label lblDCMenuSel 
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
      Height          =   300
      Index           =   3
      Left            =   3000
      TabIndex        =   9
      Top             =   3800
      Width           =   1005
   End
   Begin VB.Label lblDCMenuSel 
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
      Height          =   300
      Index           =   2
      Left            =   3000
      TabIndex        =   8
      Top             =   3500
      Width           =   1005
   End
   Begin VB.Label lblDCMenuSel 
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
      Height          =   300
      Index           =   1
      Left            =   3000
      TabIndex        =   7
      Top             =   3200
      Width           =   60
   End
   Begin VB.Label lblTimeDisp 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   660
      Left            =   5640
      TabIndex        =   6
      Top             =   2295
      Width           =   150
   End
   Begin VB.Label lblDateDisp 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   660
      Left            =   1200
      TabIndex        =   5
      Top             =   2295
      Width           =   150
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   330
   End
   Begin VB.Label lblTimeCaption 
      Caption         =   "TIME:"
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
      Left            =   4455
      TabIndex        =   3
      Top             =   2475
      Width           =   960
   End
   Begin VB.Label lblDateCaption 
      Caption         =   "DATE:"
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
      Left            =   270
      TabIndex        =   2
      Top             =   2475
      Width           =   1020
   End
   Begin VB.Label lblDCMenuSel 
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
      Height          =   300
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   2900
      Width           =   60
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Data Collection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   840
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   5265
   End
End
Attribute VB_Name = "frmDCMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' NOTE ABOUT THIS FORM:
'CURRENTLY,  THE ANSONIA DIVISION (SPECTRUM) REQUIRES THAT THE
'TASK BAR BE AVAILABLE WHILE THIS PROGRAM IS RUNNING, SO THE
'SHOP EMPLOYEES CAN UTILIZE THE MATTEC SOFTWARE
' THIS IS ACCOMPLISHED BY SETTING TWO PROPERTIES ON THIS FORM -
'CHANGE THE MinButton and MaxButton TO TRUE WHEN COMPILING THIS
'PROGRAM FOR SPECTRUM, AND SET BOTH PROPERTIES TO FALSE FOR
'ALL OTHER DIVISIONS. THESE PROPERTIES CANNOT BE SET AT RUN TIME,
'ONLY DESIGN TIME.


Rem Module Specific constants
Const mmNoSelection = 0
Const mmAttCi = 1
Const mmAttCo = 2
Const mmCoLun = 3
Const mmCbLun = 4
Const mmSusAtt = 5
Const mmResAtt = 6
Const mmCallStart = 7
Const mmCallEnd = 8
Const mmLab = 9
Const mmNote = 10
Const mmLabChk = 11
Const mmStand = 12
Const mmPrint = 13
Const mmMoveTicket = 14
Const mmAssyCmpl = 15
Const mmScrap = 16
Const mmOTAcc = 17
Const mmUpldLbr = 18
Const mmExit = 999



Const DcKey1 = &H48
Const DcKey2 = &H49
Const DCKey3 = &H50
Dim inp_str As String
Dim BadgeInput As Boolean
Dim StrPos As Integer
Dim AttActive, LaborActive As Boolean

Dim mintLoaded As Integer
Dim mintTimerCount  As Integer
Private Sub Form_Activate()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDCMain.Form_Activate
'Description      : Form Activation of Main processing for Application.
'                   Calls frmDCMain.basEntry if it isn't currently loaded.
'Parameters       : None
'Called From      : Event activating this Main Menu Form.
'Procedures Used  : frmDCMain.basEntry
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    LaborActive = False
    AttActive = False
    BadgeInput = False
    Attd_Disp = ""
    With Me
        .Caption = UCase(RTrim(gCompanyName)) & " on " & UCase(RTrim(gSQLServer)) & "/" _
                          & UCase(RTrim(gSQLDatabase)) & " " & App.FileDescription & "  -  " & "32bit Version " & App.Major & "." & App.Minor & "." & App.Revision
    End With

    If mintLoaded <> 1 Then ' if I am not loaded, load me up
        mintLoaded = 1  ' indicate that I am now loaded
        gDone = False   ' preset to not abort entry
        Call basEntry   ' show me the menu
        mintLoaded = 0  ' allow the menu to come up again
        'Dont unload me here, I stay up running
    End If
End Sub


Private Sub Form_Load()
'   frmMain.Hide
   BadgeInput = False
   KeyPreview = True
   Attd_Disp = ""
'   Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
'   MyForm.Height = Me.Height ' Remember the current size
'   MyForm.Width = Me.Width
    
 '    Me.Top = (Screen.Height - Me.Height) / 2
 '    Me.Left = (Screen.Width - Me.Width) / 2

    LblDateCaption.Caption = LoadResString(gLanguageOffset + 114)
    lblTimeCaption.Caption = LoadResString(gLanguageOffset + 115)
    lblDateDisp = Format(Date, gDateformat)
    lblTimeDisp = Format(Time, gTimeFormat)

    
'   With Me
'       .Caption = RtrimgCompanyName & " " & App.FileDescription & "  -  " & "Version " & App.Major & "." & App.Minor & "." & App.Revision
'   End With

End Sub

Private Sub ReturnTimer_Timer()
    If mintReturnCount = 0 Then
'       ErrBox "Returning to main menu.....", 2
'       SendKeys "{ESC}", True
        DoEvents
        mintReturnCount = G_RETURNCOUNT    '3 Minutes   = 60 Secs  * 5Min * 2PolInterval
        ' leave set to zero so it will cycle as many times as needed to
        ' get back to the main menu.
    Else
        mintReturnCount = mintReturnCount - 1
    End If

    
End Sub
Private Sub tmrDCMain_Timer()
Dim datestr         As String
Dim strServerTime   As String
    
    
'    Debug.Print "Seconds Left : " & mintTimerCount
    If mintTimerCount = 0 Then
        'MsgBox "Getting time from server now...", vbInformation
'        strServerTime = GetServerTime()
'        If strServerTime <> "" Then
'            Time = Format$(strServerTime, "HH:MM:SS AMPM")
'            Date = Format$(strServerTime, "MM/DD/YYYY")
            GetServerTime
            mintTimerCount = G_SETDATETIME    '5 Minutes   = 60 Secs  * 5Min * 2PolInterval
'        End If
    Else
        mintTimerCount = mintTimerCount - 1
   End If
   
   If Right(gTimeFormat, 1) = "t" Then
          gTimeFormat = "hh:mm:ss AMPM"
   End If
   
    datestr = Format(Date, gDateformat)
    frmDCMain.lblDateDisp = datestr
    frmDCMain.lblTimeDisp = Format(Time, gTimeFormat)    'dtmNow
End Sub



Public Sub basEntry()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDCMain.basEntry
'Description      : Main processing for Application.  Creates a Menu based
'                   upon the division.  Allows for scanning card for labor
'                   or attendance.
'Parameters       : None
'Called From      : frmDCMain.Form_Activate
'Procedures Used  : gEntry.Title, gEntry.Add, gEntry.AddScanItem,
'                   basgetEntryValue, DBOpen, MainError, AllDigits,
'                   Init_varbs, InitPaEmplyData, UnloadApp, Valid_Empl_badge,
'                   Attd_Badge, ErrBox
'Forms Loaded     : frmAttendance, frmInpEmplNo, frmNotepad, frmLaborChk,
'                   frmDCStdsChgs, frmJobPrint, frmLabor
'Return Parameters: None
'Tables Updates   : None
'Dependencies     : gEntry global variable (object of type clsEntry)
'==============================================================================
    Dim vEntry As Variant
    Dim iSelect As Integer
    
    ' Start Main Menu Process
    Set gEntry = New clsEntry
   If Not tmrDCMain.Enabled Then
       tmrDCMain.Enabled = True
       'MsgBox "enabling time reset timer...", vbInformation
  End If
    gEntry.Title = LoadResString(gLanguageOffset + 113)
    
    gEntry.MainEsc = True      ' Allow Application to Exit
    Select Case div
        Case bristol
            '===============================================
            '   Bristol
            '===============================================
            gEntry.Add "1", LoadResString(gLanguageOffset + 128), mmAttCi
            lblDCMenuSel(0) = LoadResString(gLanguageOffset + 116)
            gEntry.Add "2", LoadResString(gLanguageOffset + 129), mmAttCo
            lblDCMenuSel(1) = LoadResString(gLanguageOffset + 117)
            gEntry.Add "3", LoadResString(gLanguageOffset + 130), mmCoLun
            lblDCMenuSel(2) = LoadResString(gLanguageOffset + 118)
            gEntry.Add "4", LoadResString(gLanguageOffset + 131), mmCbLun
            lblDCMenuSel(3) = LoadResString(gLanguageOffset + 119)
            gEntry.Add "5", LoadResString(gLanguageOffset + 132), mmSusAtt
            lblDCMenuSel(4) = LoadResString(gLanguageOffset + 120)
            gEntry.Add "6", LoadResString(gLanguageOffset + 133), mmResAtt
            lblDCMenuSel(5) = LoadResString(gLanguageOffset + 121)
            gEntry.Add "7", LoadResString(gLanguageOffset + 134), mmCallStart
            lblDCMenuSel(6) = LoadResString(gLanguageOffset + 122)
            gEntry.Add "8", LoadResString(gLanguageOffset + 135), mmCallEnd
            lblDCMenuSel(7) = LoadResString(gLanguageOffset + 123)
            gEntry.Add "9", Right(LoadResString(gLanguageOffset + 124), Len(LoadResString(gLanguageOffset + 124)) - 3), mmLab
            lblDCMenuSel(8) = LoadResString(gLanguageOffset + 124)
'            gEntry.Add "N", "Shop Floor Notepad", mmNote
'            lblDCMenuSel(9) = "N. Shop Floor Notepad"
            gEntry.Add "T", "Print Move Tickets", mmMoveTicket
            lblDCMenuSel(9) = "T. Print Move Tags"
            
            gEntry.AddScanItem "A"
            gEntry.AddScanItem "L"
            gEntry.AddScanItem "E"
            QtrHourRounding = False
            HalfHourRounding = True
            gErrBoxDuration = 1
        Case syracuse
            '===============================================
            '   Syracuse
            '===============================================
            gEntry.Add "", "", mmNoSelection
            gEntry.Add "1", LoadResString(gLanguageOffset + 128), mmAttCi
            lblDCMenuSel(0) = LoadResString(gLanguageOffset + 116)
            gEntry.Add "2", LoadResString(gLanguageOffset + 129), mmAttCo
            lblDCMenuSel(1) = LoadResString(gLanguageOffset + 117)
            gEntry.Add "3", LoadResString(gLanguageOffset + 130), mmCoLun
            lblDCMenuSel(2) = LoadResString(gLanguageOffset + 118)
            gEntry.Add "4", LoadResString(gLanguageOffset + 131), mmCbLun
            lblDCMenuSel(3) = LoadResString(gLanguageOffset + 119)
            gEntry.Add "5", LoadResString(gLanguageOffset + 132), mmSusAtt
            lblDCMenuSel(4) = LoadResString(gLanguageOffset + 120)
            gEntry.Add "6", LoadResString(gLanguageOffset + 133), mmResAtt
            lblDCMenuSel(5) = LoadResString(gLanguageOffset + 121)
            gEntry.Add "7", LoadResString(gLanguageOffset + 134), mmCallStart
            lblDCMenuSel(6) = LoadResString(gLanguageOffset + 122)
            gEntry.Add "8", LoadResString(gLanguageOffset + 135), mmCallEnd
            lblDCMenuSel(7) = LoadResString(gLanguageOffset + 123)
            gEntry.Add "9", Right(LoadResString(gLanguageOffset + 124), Len(LoadResString(gLanguageOffset + 124)) - 3), mmLab
            lblDCMenuSel(8) = LoadResString(gLanguageOffset + 124)
            gEntry.Add "R", "Review Labor / Attendance", mmLabChk
            lblDCMenuSel(9) = "R. Review Labor / Attendance"
            gEntry.Add "M", "Enter a Message", mmNote
            lblDCMenuSel(10) = "M. Enter a Message"
            
            gEntry.AddScanItem "A"
            QtrHourRounding = False
            HalfHourRounding = True
            gErrBoxDuration = 1
        Case remesa
            '===============================================
            '    Mexico City
            '===============================================
            gEntry.Add "1", LoadResString(gLanguageOffset + 128), mmAttCi
            lblDCMenuSel(0) = LoadResString(gLanguageOffset + 116)
            gEntry.Add "2", LoadResString(gLanguageOffset + 129), mmAttCo
            lblDCMenuSel(1) = LoadResString(gLanguageOffset + 117)
            gEntry.Add "3", LoadResString(gLanguageOffset + 130), mmCoLun
            lblDCMenuSel(2) = LoadResString(gLanguageOffset + 118)
            gEntry.Add "4", LoadResString(gLanguageOffset + 131), mmCbLun
            lblDCMenuSel(3) = LoadResString(gLanguageOffset + 119)
            gEntry.Add "5", LoadResString(gLanguageOffset + 132), mmSusAtt
            lblDCMenuSel(4) = LoadResString(gLanguageOffset + 120)
            gEntry.Add "6", LoadResString(gLanguageOffset + 133), mmResAtt
            lblDCMenuSel(5) = LoadResString(gLanguageOffset + 121)
            gEntry.Add "7", LoadResString(gLanguageOffset + 134), mmCallStart
            lblDCMenuSel(6) = LoadResString(gLanguageOffset + 122)
            gEntry.Add "8", LoadResString(gLanguageOffset + 135), mmCallEnd
            lblDCMenuSel(7) = LoadResString(gLanguageOffset + 123)
            gEntry.Add "9", Right(LoadResString(gLanguageOffset + 124), Len(LoadResString(gLanguageOffset + 124)) - 3), mmLab
            lblDCMenuSel(8) = LoadResString(gLanguageOffset + 124)
            gEntry.AddScanItem "A"
            gEntry.AddScanItem "L"
'            gEntry.AddScanItem "E"
            QtrHourRounding = False
            HalfHourRounding = True
            gErrBoxDuration = 1
        Case corry
            '===============================================
            '   Corry
            '===============================================
            
            gEntry.Add "", "", mmNoSelection
            'Corry can not press shift to exit screen saver sooo
            'The user pressed enter with out making menu selection
                        
            gEntry.Add "1", LoadResString(gLanguageOffset + 128), mmAttCi
            lblDCMenuSel(0) = LoadResString(gLanguageOffset + 116)
            gEntry.Add "2", LoadResString(gLanguageOffset + 129), mmAttCo
            lblDCMenuSel(1) = LoadResString(gLanguageOffset + 117)
            gEntry.Add "3", LoadResString(gLanguageOffset + 130), mmCoLun
            lblDCMenuSel(2) = LoadResString(gLanguageOffset + 118)
            gEntry.Add "4", LoadResString(gLanguageOffset + 131), mmCbLun
            lblDCMenuSel(3) = LoadResString(gLanguageOffset + 119)
            gEntry.Add "5", LoadResString(gLanguageOffset + 132), mmSusAtt
            lblDCMenuSel(4) = LoadResString(gLanguageOffset + 120)
            gEntry.Add "6", LoadResString(gLanguageOffset + 133), mmResAtt
            lblDCMenuSel(5) = LoadResString(gLanguageOffset + 121)
            gEntry.Add "7", LoadResString(gLanguageOffset + 134), mmCallStart
            lblDCMenuSel(6) = LoadResString(gLanguageOffset + 122)
            gEntry.Add "8", LoadResString(gLanguageOffset + 135), mmCallEnd
            lblDCMenuSel(7) = LoadResString(gLanguageOffset + 123)
            gEntry.Add "9", Right(LoadResString(gLanguageOffset + 124), Len(LoadResString(gLanguageOffset + 124)) - 3), mmLab
            lblDCMenuSel(8) = LoadResString(gLanguageOffset + 124)
            gEntry.Add "S", "Standards", mmStand
            lblDCMenuSel(9) = "S. Standards"
            gEntry.Add "R", "Review Labor / Attendance", mmLabChk
            lblDCMenuSel(10) = "R. Review Labor / Attendance"
            gEntry.Add "O", "Overtime Acceptance", mmOTAcc
            lblDCMenuSel(11) = "O. Overtime Acceptance"
            gEntry.AddScanItem "A"
            gEntry.AddScanItem "L"
            QtrHourRounding = False
            HalfHourRounding = True
            gErrBoxDuration = 1
        Case saline
            '===============================================
            '   Saline
            '===============================================
            gEntry.Add "", "", mmNoSelection
            gEntry.Add "1", LoadResString(gLanguageOffset + 128), mmAttCi
            lblDCMenuSel(0) = LoadResString(gLanguageOffset + 116)
            gEntry.Add "2", LoadResString(gLanguageOffset + 129), mmAttCo
            lblDCMenuSel(1) = LoadResString(gLanguageOffset + 117)
            gEntry.Add "3", LoadResString(gLanguageOffset + 130), mmCoLun
            lblDCMenuSel(2) = LoadResString(gLanguageOffset + 118)
            gEntry.Add "4", LoadResString(gLanguageOffset + 131), mmCbLun
            lblDCMenuSel(3) = LoadResString(gLanguageOffset + 119)
            gEntry.Add "5", LoadResString(gLanguageOffset + 132), mmSusAtt
            lblDCMenuSel(4) = LoadResString(gLanguageOffset + 120)
            gEntry.Add "6", LoadResString(gLanguageOffset + 133), mmResAtt
            lblDCMenuSel(5) = LoadResString(gLanguageOffset + 121)
            gEntry.Add "7", LoadResString(gLanguageOffset + 134), mmCallStart
            lblDCMenuSel(6) = LoadResString(gLanguageOffset + 122)
            gEntry.Add "8", LoadResString(gLanguageOffset + 135), mmCallEnd
            lblDCMenuSel(7) = LoadResString(gLanguageOffset + 123)
            gEntry.Add "9", Right(LoadResString(gLanguageOffset + 124), Len(LoadResString(gLanguageOffset + 124)) - 3), mmLab
            lblDCMenuSel(8) = LoadResString(gLanguageOffset + 124)
            gEntry.Add "R", "Review Labor / Attendance", mmLabChk
            lblDCMenuSel(9) = "R. Review Labor / Attendance"
            gEntry.AddScanItem "A"
            gEntry.AddScanItem "L"
            gEntry.AddScanItem "E"
            QtrHourRounding = False
            HalfHourRounding = True
            gErrBoxDuration = 1
        Case singapore
            '===============================================
            '   AS Singapore
            '===============================================
            gEntry.Add "1", LoadResString(gLanguageOffset + 128), mmAttCi
            lblDCMenuSel(0) = LoadResString(gLanguageOffset + 116)
            gEntry.Add "2", LoadResString(gLanguageOffset + 129), mmAttCo
            lblDCMenuSel(1) = LoadResString(gLanguageOffset + 117)
            gEntry.Add "3", LoadResString(gLanguageOffset + 130), mmCoLun
            lblDCMenuSel(2) = LoadResString(gLanguageOffset + 118)
            gEntry.Add "4", LoadResString(gLanguageOffset + 131), mmCbLun
            lblDCMenuSel(3) = LoadResString(gLanguageOffset + 119)
            gEntry.Add "5", LoadResString(gLanguageOffset + 132), mmSusAtt
            lblDCMenuSel(4) = LoadResString(gLanguageOffset + 120)
            gEntry.Add "6", LoadResString(gLanguageOffset + 133), mmResAtt
            lblDCMenuSel(5) = LoadResString(gLanguageOffset + 121)
            gEntry.Add "7", LoadResString(gLanguageOffset + 134), mmCallStart
            lblDCMenuSel(6) = LoadResString(gLanguageOffset + 122)
            gEntry.Add "8", LoadResString(gLanguageOffset + 135), mmCallEnd
            lblDCMenuSel(7) = LoadResString(gLanguageOffset + 123)
            gEntry.Add "9", Right(LoadResString(gLanguageOffset + 124), Len(LoadResString(gLanguageOffset + 124)) - 3), mmLab
            lblDCMenuSel(8) = LoadResString(gLanguageOffset + 124)
            gEntry.Add "R", Right(LoadResString(gLanguageOffset + 125), Len(LoadResString(gLanguageOffset + 125)) - 3), mmLabChk
            lblDCMenuSel(9) = LoadResString(gLanguageOffset + 125)
            gEntry.Add "U", Right(LoadResString(gLanguageOffset + 310), Len(LoadResString(gLanguageOffset + 310)) - 3), mmUpldLbr
            lblDCMenuSel(10) = LoadResString(gLanguageOffset + 310)

            gEntry.AddScanItem "A"
            gEntry.AddScanItem "L"
'            gEntry.AddScanItem "E"
            QtrHourRounding = False
            HalfHourRounding = True
            gErrBoxDuration = 1
        Case pv_singapore 'div 60
            '===============================================
            '   PV Singapore
            '===============================================
            gEntry.Add "1", LoadResString(gLanguageOffset + 128), mmAttCi
            lblDCMenuSel(0) = LoadResString(gLanguageOffset + 116)
            gEntry.Add "2", LoadResString(gLanguageOffset + 129), mmAttCo
            lblDCMenuSel(1) = LoadResString(gLanguageOffset + 117)
            gEntry.Add "3", LoadResString(gLanguageOffset + 130), mmCoLun
            lblDCMenuSel(2) = LoadResString(gLanguageOffset + 118)
            gEntry.Add "4", LoadResString(gLanguageOffset + 131), mmCbLun
            lblDCMenuSel(3) = LoadResString(gLanguageOffset + 119)
            gEntry.Add "5", LoadResString(gLanguageOffset + 132), mmSusAtt
            lblDCMenuSel(4) = LoadResString(gLanguageOffset + 120)
            gEntry.Add "6", LoadResString(gLanguageOffset + 133), mmResAtt
            lblDCMenuSel(5) = LoadResString(gLanguageOffset + 121)
            gEntry.Add "7", LoadResString(gLanguageOffset + 134), mmCallStart
            lblDCMenuSel(6) = LoadResString(gLanguageOffset + 122)
            gEntry.Add "8", LoadResString(gLanguageOffset + 135), mmCallEnd
            lblDCMenuSel(7) = LoadResString(gLanguageOffset + 123)
            gEntry.Add "9", Right(LoadResString(gLanguageOffset + 124), Len(LoadResString(gLanguageOffset + 124)) - 3), mmLab
            lblDCMenuSel(8) = LoadResString(gLanguageOffset + 124)
            gEntry.Add "R", Right(LoadResString(gLanguageOffset + 125), Len(LoadResString(gLanguageOffset + 125)) - 3), mmLabChk
            lblDCMenuSel(9) = LoadResString(gLanguageOffset + 125)

            gEntry.AddScanItem "A"
            gEntry.AddScanItem "L"
'            gEntry.AddScanItem "E"
            QtrHourRounding = False
            HalfHourRounding = True
            gErrBoxDuration = 1
            
        Case milw
            '===============================================
            '   Milwaukee
            '===============================================
            gEntry.Add "", "", mmNoSelection
            gEntry.Add "1", LoadResString(gLanguageOffset + 128), mmAttCi
            lblDCMenuSel(0) = LoadResString(gLanguageOffset + 116)
            gEntry.Add "2", LoadResString(gLanguageOffset + 129), mmAttCo
            lblDCMenuSel(1) = LoadResString(gLanguageOffset + 117)
            gEntry.Add "3", LoadResString(gLanguageOffset + 132), mmSusAtt
            lblDCMenuSel(2) = "3. Suspend Attendance"
            gEntry.Add "4", LoadResString(gLanguageOffset + 133), mmResAtt
            lblDCMenuSel(3) = "4. Resume Attendance"
            gEntry.Add "5", "Check Labor Entries", mmLabChk
            lblDCMenuSel(4) = "5. Check Labor Entries"
            gEntry.Add "6", "Shop Floor Notepad", mmNote
            lblDCMenuSel(5) = "6. Shop Floor Notepad"
            gEntry.Add "7", "Labor Entry", mmLab
            lblDCMenuSel(6) = "7. Labor Entry"
            
'            gEntry.Add "C", "Assembly Complete", mmAssyCmpl
'            lblDCMenuSel(7) = "C. Assembly Complete"

            lblDCMenuSel(7) = ""
            lblDCMenuSel(8) = ""
            lblDCMenuSel(9) = ""
            lblDCMenuSel(10) = ""
            gEntry.AddScanItem "A"
            gEntry.AddScanItem "L"
            QtrHourRounding = False
            HalfHourRounding = True
            gErrBoxDuration = 2
        
        Case hyson_products
            '===============================================
            '   Hyson products - Brecksville
            '===============================================
            gEntry.Add "", "", mmNoSelection
            gEntry.Add "1", LoadResString(gLanguageOffset + 128), mmAttCi
            lblDCMenuSel(0) = LoadResString(gLanguageOffset + 116)
            gEntry.Add "2", LoadResString(gLanguageOffset + 129), mmAttCo
            lblDCMenuSel(1) = LoadResString(gLanguageOffset + 117)
            gEntry.Add "3", LoadResString(gLanguageOffset + 130), mmCoLun
            lblDCMenuSel(2) = LoadResString(gLanguageOffset + 118)
            gEntry.Add "4", LoadResString(gLanguageOffset + 131), mmCbLun
            lblDCMenuSel(3) = LoadResString(gLanguageOffset + 119)
            gEntry.Add "5", LoadResString(gLanguageOffset + 132), mmSusAtt
            lblDCMenuSel(4) = LoadResString(gLanguageOffset + 120)
            gEntry.Add "6", LoadResString(gLanguageOffset + 133), mmResAtt
            lblDCMenuSel(5) = LoadResString(gLanguageOffset + 121)
            gEntry.Add "7", LoadResString(gLanguageOffset + 134), mmCallStart
            lblDCMenuSel(6) = LoadResString(gLanguageOffset + 122)
            gEntry.Add "8", LoadResString(gLanguageOffset + 135), mmCallEnd
            lblDCMenuSel(7) = LoadResString(gLanguageOffset + 123)
            gEntry.Add "9", Right(LoadResString(gLanguageOffset + 124), Len(LoadResString(gLanguageOffset + 124)) - 3), mmLab
            lblDCMenuSel(8) = LoadResString(gLanguageOffset + 124)
            gEntry.Add "R", "Check Labor / Attendance Entries", mmLabChk
            lblDCMenuSel(9) = "R. Check Labor / Attendance Entries"
            gEntry.Add "C", "Assembly Complete", mmAssyCmpl
            lblDCMenuSel(10) = "C. Assembly Complete"
           gEntry.AddScanItem "A"
           gEntry.AddScanItem "L"
           gEntry.AddScanItem "E"
            QtrHourRounding = True
            HalfHourRounding = False
            gErrBoxDuration = 2
        Case Synventive
            '===============================================
            '   Hyson products - Brecksville
            '===============================================
            gEntry.Add "", "", mmNoSelection
            gEntry.Add "1", LoadResString(gLanguageOffset + 128), mmAttCi
            lblDCMenuSel(0) = LoadResString(gLanguageOffset + 116)
            gEntry.Add "2", LoadResString(gLanguageOffset + 129), mmAttCo
            lblDCMenuSel(1) = LoadResString(gLanguageOffset + 117)
            gEntry.Add "3", LoadResString(gLanguageOffset + 130), mmCoLun
            lblDCMenuSel(2) = LoadResString(gLanguageOffset + 118)
            gEntry.Add "4", LoadResString(gLanguageOffset + 131), mmCbLun
            lblDCMenuSel(3) = LoadResString(gLanguageOffset + 119)
            gEntry.Add "5", LoadResString(gLanguageOffset + 132), mmSusAtt
            lblDCMenuSel(4) = LoadResString(gLanguageOffset + 120)
            gEntry.Add "6", LoadResString(gLanguageOffset + 133), mmResAtt
            lblDCMenuSel(5) = LoadResString(gLanguageOffset + 121)
           gEntry.AddScanItem "A"
           gEntry.AddScanItem "L"
           gEntry.AddScanItem "E"
            QtrHourRounding = False
            HalfHourRounding = True
            gErrBoxDuration = 2
    Case stroms
            '===============================================
            '   Stroms
            '===============================================
            gEntry.Add "", "", mmNoSelection
            gEntry.Add "1", LoadResString(gLanguageOffset + 128), mmAttCi
            lblDCMenuSel(0) = LoadResString(gLanguageOffset + 116)
            gEntry.Add "2", LoadResString(gLanguageOffset + 129), mmAttCo
            lblDCMenuSel(1) = LoadResString(gLanguageOffset + 117)
            gEntry.Add "3", LoadResString(gLanguageOffset + 130), mmCoLun
            lblDCMenuSel(2) = LoadResString(gLanguageOffset + 118)
            gEntry.Add "4", LoadResString(gLanguageOffset + 131), mmCbLun
            lblDCMenuSel(3) = LoadResString(gLanguageOffset + 119)
            gEntry.Add "5", LoadResString(gLanguageOffset + 132), mmSusAtt
            lblDCMenuSel(4) = LoadResString(gLanguageOffset + 120)
            gEntry.Add "6", LoadResString(gLanguageOffset + 133), mmResAtt
            lblDCMenuSel(5) = LoadResString(gLanguageOffset + 121)
            gEntry.Add "7", LoadResString(gLanguageOffset + 134), mmCallStart
            lblDCMenuSel(6) = LoadResString(gLanguageOffset + 122)
            gEntry.Add "8", LoadResString(gLanguageOffset + 135), mmCallEnd
            lblDCMenuSel(7) = LoadResString(gLanguageOffset + 123)
            gEntry.Add "9", Right(LoadResString(gLanguageOffset + 124), Len(LoadResString(gLanguageOffset + 124)) - 3), mmLab
            lblDCMenuSel(8) = LoadResString(gLanguageOffset + 124)
            gEntry.Add "R", Right(LoadResString(gLanguageOffset + 125), Len(LoadResString(gLanguageOffset + 125)) - 3), mmLabChk
            lblDCMenuSel(9) = LoadResString(gLanguageOffset + 125)
            gEntry.Add "C", Right(LoadResString(gLanguageOffset + 126), Len(LoadResString(gLanguageOffset + 126)) - 3), mmAssyCmpl
            lblDCMenuSel(10) = LoadResString(gLanguageOffset + 126)
            gEntry.AddScanItem "A"
            gEntry.AddScanItem "L"
            gEntry.AddScanItem "E"
            QtrHourRounding = True
            HalfHourRounding = False
            gErrBoxDuration = 3
    End Select
    
    vEntry = basGetEntryValue()
    
    Set gEntry = Nothing
    
    If ConnectToDB() = vbError Then
        gbConnected = False
        Call MainError("Cannot connect to DB")
        Exit Sub
    End If
    
    gbConnected = True
    
    Rem Only act on selection if got Connection !!!
    If gbConnected Then
     
     ' shut off timer until transaction is complete...
     tmrDCMain.Enabled = False
      
        If AllDigits(CStr(vEntry)) Then
            Init_varbs
           Select Case vEntry
                'Show for DC_Floor App
                Case mmNoSelection
                    'We have just come off the screen saver and the user pressed enter
                
                Case mmAttCi
                    gstrLastOper = "Attendance Clock In"
                    Attd_sel = vEntry
                    frmAttendance.Show 1
                Case mmAttCo
                    gstrLastOper = "Attendance Clock Out"
                    Attd_sel = vEntry
                    frmAttendance.Show 1
                Case mmCoLun
                    gstrLastOper = "Clock Out for Lunch"
                    Attd_sel = vEntry
                    frmAttendance.Show 1
                Case mmCbLun
                     gstrLastOper = "Clock back for Lunch"
                    Attd_sel = vEntry
                    frmAttendance.Show 1
                Case mmSusAtt
                    gstrLastOper = "Suspend Attendance"
                    Attd_sel = vEntry
                   frmAttendance.Show 1
                Case mmResAtt
                    gstrLastOper = "Resume Attendance"
                    Attd_sel = vEntry
                    frmAttendance.Show 1
                Case mmCallStart
                    gstrLastOper = "Start Callback"
                    Attd_sel = vEntry
                    frmAttendance.Show 1
                Case mmCallEnd
                    gstrLastOper = "End Callback"
                    Attd_sel = vEntry
                    frmAttendance.Show 1
                Case mmLab
                    gstrLastOper = "Labor"
                    'Start Labor Menu Process
                    frmInpEmplNo.Tag = 1
                    frmLabor.Tag = 0
                    frmInpEmplNo.Show 1
                Case mmNote
                    gstrLastOper = "Notepad"
                    Call InitPaEmplyData
                    frmInpEmplNo.Tag = 3
                    frmInpEmplNo.Show 1
                Case mmLabChk
                    gstrLastOper = "Check Labor Entries"
                    Call InitPaEmplyData
                    frmInpEmplNo.Tag = 2
                    frmInpEmplNo.Show 1
                Case mmStand
                    gstrLastOper = "Standards"
                    frmInpEmplNo.Tag = 4
                    frmInpEmplNo.Show 1
                Case mmAssyCmpl
                    gstrLastOper = "Assembly Complete"
                    frmInpEmplNo.Tag = 5
                    frmInpEmplNo.Show 1
                Case mmMoveTicket
                    gstrLastOper = "Print Move Tickets"
                    frmMoveTag.Show 1
                Case mmScrap
                    gstrLastOper = "Scrap Reporting"
                    frmInpEmplNo.Tag = 6
                    frmInpEmplNo.Show 1
                Case mmOTAcc
                    gstrLastOper = "Overtime Acceptance"
                    frmInpEmplNo.Tag = 7
                    frmInpEmplNo.Show 1
                Case mmUpldLbr
                    gstrLastOper = "Upload Labor Spreadsheet"
                    frmInpEmplNo.Tag = 8
                    frmInpEmplNo.Show 1
            Case Else
                ErrBox LoadResString(gLanguageOffset + 58) & vEntry, 4
                
            End Select
        Else
                If Left(vEntry, 2) = "LB" Then
                    Call GetSysDateTime     'save current time
                    If Valid_Empl_badge(Mid(vEntry, 3, 6)) Then
                        Select Case gEmployee.AttdStatus
                            Case 2, 5, 8
                                ErrBox LoadResString(gLanguageOffset + 76), 4
                            Case 3
                                ErrBox LoadResString(gLanguageOffset + 2), 4
                            Case Else
                                Call GetSysDateTime     'save current time
                                If Elapsed_hrs(gEmployee.AttdStartDate, gSysDateTime) >= 24# Then
                                    ErrBox LoadResString(gLanguageOffset + 5), 4
                                Else
                                    gstrLastOper = "Scanned Labor"
                                    AttActive = False
                                    LaborActive = True
                                    Init_varbs
                                    frmLabor.Tag = 0
                                    frmLabor.Show 1
                                End If
                        End Select
                    End If
                ElseIf Left(vEntry, 2) = "EG" Then
                    Call GetSysDateTime     'save current time
                    If Valid_Empl_badge(Mid(vEntry, 3, 6)) Then
                           Call GetSysDateTime     'save current time
                           gEmployee.Actdate = gSysDate 'Needed since employee does not clock in or out
                           gEmployee.GenShift = 1    'Also needed since no clock in/out
                            gstrLastOper = "Scanned Labor"
                            AttActive = False
                            LaborActive = True
                            Init_varbs
                            Select Case div
                                Case stroms
                                    frmLabor.Tag = 0
                                Case Else
                                    frmLabor.Tag = 1
                            End Select
                            frmLabor.Show 1
                    End If
                ElseIf Left(vEntry, 2) = "LO" Then
                    Call GetSysDateTime     'save current time
'                    If Valid_Empl_badge(Mid(vEntry, 3, 6)) Then
                           Call GetSysDateTime     'save current time
                           gEmployee.Actdate = gSysDate 'Needed since employee does not clock in or out
                           gEmployee.GenShift = 1    'Also needed since no clock in/out
                            gstrLastOper = "Scanned Labor"
                            AttActive = False
                            LaborActive = True
                            Init_varbs
 '                           Select Case div
 '                               Case stroms
 '                                   frmLabor.Tag = 0
 '                               Case Else
 '                                   frmLabor.Tag = 1
 '                           End Select
                            frmLabor.Show 1
  '                  End If
                ElseIf Left(vEntry, 2) = "AT" Then
                    If Valid_Empl_badge(Mid(vEntry, 3, 6)) Then
                        gstrLastOper = "Scanned Attendance"
                        AttActive = True
                        LaborActive = False
                        Call Attd_Badge
                        If Len(Attd_Disp) > 1 Then
                           Call ErrBox(RTrim(gEmployee.EmplName) & " " & Attd_Disp, 3)
                            Attd_Disp = ""
                        End If
                        Init_varbs
                    End If
                End If
        End If

    DisconnectDB
'     If gSQLConnection.State = adStateOpen Then
'       gSQLConnection.Close
'     End If
    
    End If
End Sub
