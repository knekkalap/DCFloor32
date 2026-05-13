VERSION 5.00
Begin VB.Form frmNotepad 
   Caption         =   "Data Collection Messages"
   ClientHeight    =   6300
   ClientLeft      =   1050
   ClientTop       =   240
   ClientWidth     =   9900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6300
   ScaleWidth      =   9900
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtNote 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   450
      MaxLength       =   2048
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1800
      Width           =   8775
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   3510
      Picture         =   "Frmnotep.frx":0000
      Stretch         =   -1  'True
      Top             =   135
      Width           =   870
   End
   Begin VB.Label Label4 
      Caption         =   "Press F1 to save message; ESC to quit without saving"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   450
      TabIndex        =   10
      Top             =   5175
      Width           =   8775
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
      Left            =   225
      TabIndex        =   9
      Top             =   5760
      Width           =   45
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2025
      TabIndex        =   8
      Top             =   1035
      Width           =   60
   End
   Begin VB.Label Label3 
      Caption         =   "Message - (Enter Up to Two Thousand Characters)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   495
      TabIndex        =   7
      Top             =   1440
      Width           =   5205
   End
   Begin VB.Label Label2 
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
      Height          =   330
      Left            =   525
      TabIndex        =   6
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Data Collection Message"
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
      Height          =   495
      Left            =   4410
      TabIndex        =   5
      Top             =   240
      Width           =   5070
   End
   Begin VB.Label Label11 
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
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   225
      Width           =   870
   End
   Begin VB.Label Label12 
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   2
      Top             =   615
      Width           =   825
   End
   Begin VB.Label lblDateDisp 
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1425
      TabIndex        =   1
      Top             =   255
      Width           =   60
   End
   Begin VB.Label lblTimeDisp 
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   615
      Width           =   60
   End
End
Attribute VB_Name = "frmNotepad"
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
'Procedure Name   : frmNotepad.Form_Activate
'Description      : When form is activated for the first time after the
'                   form is loaded, the initialization process is called.
'Parameters       : None
'Called From      : Activate Event for form
'Procedures Used  : NotepadProcess
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
        Call NotepadProcess
        'Don't unload now, we will wait for keystroke events to decide what
        '   to do.  This is an event driven form, see events for the
        '   frmNotepad.Form object(Keydown).
        'mintLoaded = 0
        'Unload Me
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmNotepad.Form_KeyDown
'Description      : Process the Esc and F1 keystroke events.  When F1 key is
'                   hit, the text will be fixed for single quotes, then
'                   inserted into the database. KeyPreview property is set
'                   on the form, so that the form captures keystrokes first.
'Parameters       : KeyCode and Shift indicator
'Called From      : KeyDown Event
'Procedures Used  : FixNoteQuotes, InsertNote
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : InsertNote updates the Notepad Table
'==============================================================================
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
        Exit Sub
    Case vbKeyF1
        gNotepad.MsgText = FixNoteQuotes(txtNote.Text)
        Call InsertNote
        Unload Me
        Exit Sub
    End Select
End Sub

Public Function FixNoteQuotes(ByVal strQuote) As String
'==============================================================================
'Author           : Chris Connors, Associated Spring / David Taylor, PRT Group
'Date             : June, 1998
'Procedure Name   : frmNotepad.FixNoteQuotes(String)
'Description      : Fix the text string before update to database.  All single
'                   quotes are doubled.  All double quotes are left alone.
'Parameters       : String to be Checked / Fixed
'Called From      : frmNotepad.Form_KeyDown (Just before update)
'Procedures Used  : None
'Forms Loaded     : None
'Return Parameters: Fixed String
'Tables Updates   : None
'==============================================================================

'Parses given string for single quotes. Transact SQL will fail when there
'is a single quote in the middle of a string such as an apostrophe.
'As a result, another single quote must be added before the existing
'single quote.  2 single quotes together will be interpreted
'as 1 single quote by TSQL.

Dim iLen As Integer
Dim iPos As Integer

FixNoteQuotes = strQuote

Rem Fix single quotes
iLen = Len(FixNoteQuotes)
iPos = 1
Do
    iPos = InStr(iPos, FixNoteQuotes, "'")
    If iPos > 0 Then
        FixNoteQuotes = Left(FixNoteQuotes, iPos) & "'" _
            & Right(FixNoteQuotes, iLen - iPos)
        iLen = iLen + 1
        iPos = iPos + 2
    End If
Loop Until iPos = 0 Or iPos > iLen

End Function




Public Sub NotepadProcess()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmNotepad.NotepadProcess
'Description      : Initialize the Notepad form with Employee Info and date
'                   and time.  All other actions are done by keystroke events.
'                   KeyPreview property is set on the form, so that the form
'                   captures keystrokes first.  Update or escape is controlled
    '                   in the frmNotepad.Form_Keydown Event handling function.
'Parameters       : None
'Called From      : frmNotepad.Form_Activate (only first time after load)
'Procedures Used  : gEntry.Title, gEntry.EntryType, gEntry.Escaped,
'                   basGetEntry, Valid_Empl_badge
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    Dim datestr As String
    Dim strEntryEmpNo As String
    Dim bGotOne As Boolean
    
    gNotepad.companyCd = gCompanyCd
    gNotepad.PrintFlag = " "
    gNotepad.Actdate = Now
    datestr = Format(gNotepad.Actdate, gDateformat)
    lblDateDisp = Format(gNotepad.Actdate, gDateformat)
    lblTimeDisp = Format(gNotepad.Actdate, gTimeFormat)
    
    lblEmployee.Caption = gEmployee.EmplNo & " " & gEmployee.EmplName
    gNotepad.EmplNo = gEmployee.EmplNo
    

End Sub

Private Sub Form_Load()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmNotepad.Form_Load
'Description      : Indicate that the Initialization process needs to be done
'                   on the form, once the form is activated.
'Parameters       : None
'Called From      : Form Load Event
'Procedures Used  : None
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    mintLoaded = 0
End Sub


