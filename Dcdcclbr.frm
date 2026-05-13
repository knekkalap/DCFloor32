VERSION 5.00
Begin VB.Form frmDCCLabor 
   Caption         =   "Data Collection - Direct Costcenter Charge "
   ClientHeight    =   5850
   ClientLeft      =   600
   ClientTop       =   330
   ClientWidth     =   8805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5850
   ScaleWidth      =   8805
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   735
      Left            =   765
      Picture         =   "Dcdcclbr.frx":0000
      Stretch         =   -1  'True
      Top             =   1305
      Width           =   870
   End
   Begin VB.Label lblccidDesc 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   240
      Left            =   3195
      TabIndex        =   10
      Top             =   2700
      Width           =   60
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      Caption         =   " "
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
      Left            =   450
      TabIndex        =   9
      Top             =   5130
      Width           =   60
   End
   Begin VB.Label lblIndHrs 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   240
      Left            =   2475
      TabIndex        =   8
      Top             =   3150
      Width           =   60
   End
   Begin VB.Label lblChargeHoursCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Charge Hours"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1080
      TabIndex        =   7
      Top             =   3150
      Width           =   1290
   End
   Begin VB.Label lblCCID 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   240
      Left            =   2475
      TabIndex        =   6
      Top             =   2700
      Width           =   60
   End
   Begin VB.Label lblCCIDCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cost Center"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1260
      TabIndex        =   5
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label lblEmplName 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   240
      Left            =   4680
      TabIndex        =   4
      Top             =   2205
      Width           =   60
   End
   Begin VB.Label lblNameCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   4005
      TabIndex        =   3
      Top             =   2205
      Width           =   555
   End
   Begin VB.Label lblEmplNo 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   240
      Left            =   2475
      TabIndex        =   2
      Top             =   2205
      Width           =   60
   End
   Begin VB.Label lblEmplNoCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Empl No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1530
      TabIndex        =   1
      Top             =   2205
      Width           =   855
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      Caption         =   "Direct Costcenter Charge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   510
      Left            =   1710
      TabIndex        =   0
      Top             =   1440
      Width           =   6210
   End
End
Attribute VB_Name = "frmDCCLabor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintLoaded As Integer

Const vbColorWhite = &H8000005



Public Sub DCProcess()
    
    Dim strEntry As String
    
    Dim I As Integer
        
    Dim CalcHours As Currency
    Dim strRecType As String
    Dim j As Integer
    
    Call Init_varbs     'Initialize gLaborIn Array
    Call checkClockOut
    I = 1
'gEOSFlag is set to true by function CheckClockOut
'when the current time is within a window before
'and after the shift Clock-Out time (from the shift
'parameter record). The upper and lower bounds of
'the window are defined by the values in EOSStart
'and EOSEnd fields, also in the shift param rec.
'If EOSStart and EOSEnd are zero, gEOSFlag is set
'to false.

'InpEndOfShift is set to True when the user answers
'the prompt "Clocking Out?". Additionally, the
'boolean variable empClkOutOK is set to true if
'the user answers Yes to the prompt, or false if
'the user answers No to the prompt.

        
    gLaborIn(I).MachRec = "N"
    gLaborIn(I).RecType = "C"
    gLaborIn(I).OperStatus = " "
          
   If Not inpCCID(I) Then Exit Sub
   Screen.ActiveForm.lblCCID = gLaborIn(I).CCID
   Screen.ActiveForm.lblccidDesc = gCostCenter.Descr
  
   If div = corry Then
     If gEosFlag Then
        If Not inpEndOfShift() Then
            gDone = True
            Unload Me
            Exit Sub
        End If
    End If
 End If
   
  If div = corry Then
     gLaborIn(I).inHours = CalcActHrs()
 Else
   If Not inpDDCHours(I) Then Exit Sub
 End If
    
   Call Build_Labor_Recs
    
   If div = corry Then
   If empClkOutOK Then
      gLaborIn(I).ErrorLst = ""
      Call ClockOut
      Call Upd_Empl_Attd
      gDone = True
   End If
End If
  Unload Me
    
End Sub

Private Sub Form_Activate()
   lblEmplNo = gEmployee.EmplNo
   lblEmplName = gEmployee.EmplName
    
    If gDone Then
        Unload Me
    End If
    
    If mintLoaded <> 1 Then
        mintLoaded = 1
        Call DCProcess
        mintLoaded = 0
        Unload Me
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then
       'frmDirLabor.Hide
       'frmDCMain.Show
       Unload Me
   End If
End Sub

Private Sub List1_Click()

End Sub

Private Sub Form_Load()
lblFormTitle.Caption = LoadResString(gLanguageOffset + 234)
lblEmplNoCaption.Caption = LoadResString(gLanguageOffset + 136)
lblNameCaption.Caption = LoadResString(gLanguageOffset + 137)
lblChargeHoursCaption.Caption = LoadResString(gLanguageOffset + 235)
lblCCIDCaption.Caption = LoadResString(gLanguageOffset + 146)

End Sub


