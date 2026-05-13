VERSION 5.00
Begin VB.Form frmEndIndirect 
   Caption         =   "Data Collection - End Indirect Labor"
   ClientHeight    =   8685
   ClientLeft      =   2310
   ClientTop       =   2265
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8685
   ScaleWidth      =   9615
   WindowState     =   2  'Maximized
   Begin VB.Line Line2 
      X1              =   315
      X2              =   9360
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Line Line1 
      X1              =   315
      X2              =   9360
      Y1              =   3330
      Y2              =   3330
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1845
      Picture         =   "Frmendin.frx":0000
      Stretch         =   -1  'True
      Top             =   2430
      Width           =   870
   End
   Begin VB.Label lblccidDesc 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   3375
      TabIndex        =   17
      Top             =   4500
      Width           =   60
   End
   Begin VB.Label lblIndirDesc 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   3375
      TabIndex        =   16
      Top             =   4035
      Width           =   60
   End
   Begin VB.Label lblCCID 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   2520
      TabIndex        =   15
      Top             =   4500
      Width           =   60
   End
   Begin VB.Label lblccidCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cost Center"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   270
      TabIndex        =   14
      Top             =   4500
      Width           =   2190
   End
   Begin VB.Label lblIndirCd 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   2520
      TabIndex        =   13
      Top             =   4035
      Width           =   60
   End
   Begin VB.Label lblIndirectCodeCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Indirect Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   225
      TabIndex        =   12
      Top             =   4035
      Width           =   2250
   End
   Begin VB.Label lblEmplName 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   5220
      TabIndex        =   11
      Top             =   3465
      Width           =   60
   End
   Begin VB.Label lblNameCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3780
      TabIndex        =   10
      Top             =   3465
      Width           =   1320
   End
   Begin VB.Label lblEmplNo 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   2250
      TabIndex        =   9
      Top             =   3495
      Width           =   60
   End
   Begin VB.Label lblEmplNoCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Empl No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   270
      TabIndex        =   8
      Top             =   3510
      Width           =   1890
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      Caption         =   "End Indirect Labor"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   2520
      Width           =   5610
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
      Left            =   495
      TabIndex        =   6
      Top             =   5715
      Width           =   60
   End
   Begin VB.Label lblVarbData3 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   2475
      TabIndex        =   5
      Top             =   5355
      Width           =   60
   End
   Begin VB.Label lblVarbFld3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   2295
      TabIndex        =   4
      Top             =   5355
      Width           =   60
   End
   Begin VB.Label lblVarbData1 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   2475
      TabIndex        =   3
      Top             =   4905
      Width           =   60
   End
   Begin VB.Label lblVarbFld1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   2295
      TabIndex        =   2
      Top             =   4905
      Width           =   60
   End
   Begin VB.Label lblVarbData2 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   6075
      TabIndex        =   1
      Top             =   4905
      Width           =   60
   End
   Begin VB.Label lblVarbFld2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   5895
      TabIndex        =   0
      Top             =   4905
      Width           =   60
   End
End
Attribute VB_Name = "frmEndIndirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mintLoaded As Integer

Const vbColorWhite = &H8000005

Public Sub EndIndProcess()
    
    Dim strEntry As String
    
    Dim I As Integer
        
    Dim lScrapTot As Long
    Dim CalcHours As Currency
    Dim strRecType As String
    Dim j As Integer
    
    Call Init_varbs     'Initialize gLaborIn Array
   Call checkClockOut
    I = 1
    
    gLaborIn(I).MachRec = "N"
    gLaborIn(I).RecType = "I"
    gLaborIn(I).OperStatus = " "
        
   ListOK = True
   If Not inpCloseIndirAns(I) Then Exit Sub
       
   Call CloseIndirect(gEmployee.EmplNo, gEmployee.Actdate)
    
   Unload Me
    
End Sub

Private Sub Form_Activate()
   lblEmplNo = gEmployee.EmplNo
   lblEmplName = gEmployee.EmplName
    lblIndirCd = gOpenLabor.IndirCd
    lblCCID = gOpenLabor.CCID
    lblIndirDesc = gcolIndirCodes.getDesc(gOpenLabor.IndirCd)
    GetCostCenter (gOpenLabor.CCID)
    lblccidDesc = gCostCenter.Descr
    
    If gDone Then
'       gDone = False
       Unload Me
    End If
    
    If mintLoaded <> 1 Then
        mintLoaded = 1
        Call EndIndProcess
        mintLoaded = 0
        Unload Me
    End If


End Sub

Private Sub Form_Load()
lblFormTitle.Caption = LoadResString(gLanguageOffset + 204)
lblEmplNoCaption.Caption = LoadResString(gLanguageOffset + 136)
lblNameCaption.Caption = LoadResString(gLanguageOffset + 137)
lblIndirectCodeCaption.Caption = LoadResString(gLanguageOffset + 145)
lblCCIDCaption.Caption = LoadResString(gLanguageOffset + 146)

End Sub


