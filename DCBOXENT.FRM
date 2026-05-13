VERSION 5.00
Begin VB.Form frmBoxEntry 
   Caption         =   "Data Collection Box Number Entry"
   ClientHeight    =   8460
   ClientLeft      =   3405
   ClientTop       =   2760
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8460
   ScaleWidth      =   8220
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtBoxNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3165
      MaxLength       =   10
      TabIndex        =   0
      Top             =   3180
      Width           =   1215
   End
   Begin VB.TextBox txtBoxNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3165
      MaxLength       =   10
      TabIndex        =   1
      Top             =   3660
      Width           =   1215
   End
   Begin VB.TextBox txtBoxNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3165
      MaxLength       =   10
      TabIndex        =   2
      Top             =   4150
      Width           =   1215
   End
   Begin VB.TextBox txtBoxNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3165
      MaxLength       =   10
      TabIndex        =   3
      Top             =   4620
      Width           =   1215
   End
   Begin VB.TextBox txtBoxNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3165
      MaxLength       =   10
      TabIndex        =   4
      Top             =   5100
      Width           =   1215
   End
   Begin VB.TextBox txtBoxNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3165
      MaxLength       =   10
      TabIndex        =   5
      Top             =   5580
      Width           =   1215
   End
   Begin VB.TextBox txtBoxNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3165
      MaxLength       =   10
      TabIndex        =   6
      Top             =   6060
      Width           =   1215
   End
   Begin VB.TextBox txtBoxNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   3165
      MaxLength       =   10
      TabIndex        =   8
      Top             =   6540
      Width           =   1215
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   9
      Top             =   7320
      Width           =   75
   End
   Begin VB.Label lblFormTitle 
      Caption         =   "Box Number Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   645
      Left            =   1560
      TabIndex        =   7
      Top             =   450
      Width           =   5565
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Box Number 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   15
      Top             =   3165
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Box Number 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   3660
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Box Number 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   17
      Top             =   4155
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Box Number 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   18
      Top             =   4650
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Box Number 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   19
      Top             =   5145
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Box Number 6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   20
      Top             =   5595
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Box Number 7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   21
      Top             =   6045
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Box Number 8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   22
      Top             =   6540
      Width           =   2460
   End
   Begin VB.Label lblQuantityCaption 
      Alignment       =   2  'Center
      Caption         =   "Box "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3165
      TabIndex        =   14
      Top             =   2805
      Width           =   1290
   End
   Begin VB.Label lblInstrLine1Caption 
      Caption         =   "Use Arrows or Enter to select  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   8115
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   10
      Top             =   6945
      Width           =   75
   End
   Begin VB.Label lblInstrLine2Caption 
      Caption         =   "Press F1 when done   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   6885
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   360
      Picture         =   "DCBOXENT.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblInstrLine3Caption 
      Caption         =   "Press Esc to exit without saving anything"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   8535
   End
End
Attribute VB_Name = "frmBoxEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub CheckKeys(index As Integer, KeyCode As Integer, Shift As Integer)
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : ScrapWind.CheckKeys
'Description      : Handle Keydown Events from Text Box control arrays.  If
'                   navigation keys, make sure description is up to date
'                   when positioning.  Esc key to exit with no update, and
'                   F1 key to exit with update.  F6 key to lookup a scrap code
'                   using the frmLookup form.
'                   Uses module variable miColFocus for positioning logic.
'                   In order to correctly determine which row we are on when
'                   the F6 code lookup action is taken, we must capture the
'                   keystroke from the control arrays, to get which control
'                   (or row) we are on.  See ScrapWind.txtScrapCd_Keydown
'                   and ScrapWind.txtScrapQty_Keydown Events.
'Parameters       : Parameters from Control Array Keydown Event.
'Called From      : ScrapWind.txtScrapCd_Keydown, ScrapWind.txtScrapQty_Keydown
'Procedures Used  : AllDigits, ErrorBeep, ErrBox, ValidScrapCd
'Forms Loaded     : frmLookup
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    Dim vCode As Variant
    Dim I As Integer
    Dim iDownUp As Integer
    Dim strwork As String
    
    Select Case KeyCode
    Case vbKeyEscape
        gSelIndx = -1
        Unload Me
        Exit Sub
    
    Case vbKeyDown, vbKeyReturn, vbKeyTab
        'Make sure we stay in control array bounds
        If index < 7 Then
                txtBoxNumber(index + 1).SetFocus
        End If
        KeyCode = 0     'Prevent Value change in combo box
    
    Case vbKeyUp
        'Make sure we stay in control array bounds
        If index > 0 Then
                txtBoxNumber(index - 1).SetFocus
        End If
        KeyCode = 0     'Prevent Value change in combo box
    
    Case vbKeyRight
        txtBoxNumber(index).SetFocus
        KeyCode = 0     'Prevent Value change in combo box
    
    Case vbKeyLeft
        txtBoxNumber(index).SetFocus
        KeyCode = 0     'Prevent Value change in combo box
    
    Case vbKeyF1
        'Loop through Control array to make sure at least one Box Number is entered
        gSelIndx = 0
        For I = 0 To 7
            If txtBoxNumber(I).Text = " " Then txtBoxNumber(I).Text = ""
            If txtBoxNumber(I).Text = "" Then
                'Skip checks on this row - not used
            ElseIf Not AllDigits(txtBoxNumber(I).Text) _
                    Or Len(Trim(txtBoxNumber(I).Text)) > 2 Then
                  ErrorBeep
                  Call ErrBox("Box Number is invalid", 2)
                  txtBoxNumber(I).SetFocus
                  KeyCode = 0
                  Exit Sub
            ElseIf CLng(txtBoxNumber(I).Text) < 1 Or CLng(txtBoxNumber(I).Text) > 99 Then
                  ErrorBeep
                  Call ErrBox("Box Number is invalid", 2)
                  txtBoxNumber(I).SetFocus
                  KeyCode = 0
                  Exit Sub
            Else
                gSelIndx = gSelIndx + 1   'increment count of entries
            End If
        Next I
        
        If gSelIndx = 0 Then
                  ErrorBeep
                  Call ErrBox("Must enter at least one box number", 2)
                  txtBoxNumber(0).SetFocus
                  Exit Sub
        End If
        
        
        Rem Update the fields before exiting
        For I = 1 To 8
            If txtBoxNumber(I - 1).Text <> "" Then
                gLaborIn(LoopCnt).BoxNumber(I) = CLng(txtBoxNumber(I - 1).Text)
            End If
        Next I
       
        Unload Me
        Exit Sub
        
    Case vbKeyF6
               'do nothing    key not valid on this input screen
    End Select

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Code to have enter key tab to next field (and not to beep)
    If KeyCode = vbKeyReturn Then KeyCode = vbKeyTab
End Sub
Private Sub txtBoxNumber_GotFocus(index As Integer)
        
        lblInfo = "Enter a box number"
End Sub
Private Sub txtBoxNumber_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Call CheckKeys(index, KeyCode, Shift)
End Sub
Private Sub txtBoxNumber_LostFocus(index As Integer)
    lblInfo = ""
     If txtBoxNumber(index).Text = " " Then txtBoxNumber(index).Text = ""
     If txtBoxNumber(index).Text = "" Then
     ElseIf Not AllDigits(txtBoxNumber(index).Text) Then
                  Call ErrBox(" Box number must be all numbers", 2)
                  txtBoxNumber(index).SetFocus
     ElseIf Len(Trim(txtBoxNumber(index).Text)) > 2 Then
                  Call ErrBox("Box number too long", 2)
                  txtBoxNumber(index).SetFocus
     ElseIf DuplicateBoxNumb(CLng(txtBoxNumber(index).Text), index) Then
                  Call ErrBox("Duplicate box number entered", 2)
                  txtBoxNumber(index).SetFocus
    End If
End Sub

Public Function DuplicateBoxNumb(BoxNbr As Long, index As Integer) As Boolean
Dim x As Integer
Dim Y As Integer
  DuplicateBoxNumb = False
For x = 0 To 7
  If x <> index Then
     If IsNumeric(txtBoxNumber(x)) Then
        If CLng(txtBoxNumber(x)) = BoxNbr Then
              DuplicateBoxNumb = True
             Exit Function
        End If
  End If
  End If
  Next x

End Function

Private Sub Form_Load()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : ScrapWind.Form_Load
'Description      : When form is Loaded, the text boxes are initialized with
'                   any data existing in the current machine.  Form_Activate
'                   event is not needed, because initialization in not
'                   interactive here, and the form then goes immediately into
'                   event driven mode, where actions are done by keystroke
'                   events. KeyPreview property is not set on this form.
'                   In order to correctly determine which row we are on when
'                   the F6 code lookup action is taken, we must capture the
'                   keystroke from the control arrays, to get which control
'                   (or row) we are on.  See ScrapWind.txtScrapCd_Keydown
'                   and ScrapWind.txtScrapQty_Keydown Events (which call
'                   ScrapWind.CheckKeys).
'Parameters       : None
'Called From      : Load Event for form
'Procedures Used  : NotepadProcess
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    Dim I As Integer
    
    Rem Update the fields before exiting
        For I = 0 To 7
            If gLaborIn(LoopCnt).BoxNumber(I + 1) <> 0 Then
                txtBoxNumber(I).Text = Format(gLaborIn(LoopCnt).BoxNumber(I + 1), "00")
            End If
        Next I
     '   txtBoxNumber(0).SetFocus
     '   SendKeys "%{DOWN}"

End Sub


