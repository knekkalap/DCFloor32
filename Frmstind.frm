VERSION 5.00
Begin VB.Form frmStartIndir 
   Caption         =   "Data Collection - Start Indirect Labor"
   ClientHeight    =   6660
   ClientLeft      =   2460
   ClientTop       =   1395
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6660
   ScaleWidth      =   10245
   WindowState     =   2  'Maximized
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
      Left            =   6885
      TabIndex        =   17
      Top             =   4230
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
      Left            =   7065
      TabIndex        =   16
      Top             =   4230
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
      Left            =   3285
      TabIndex        =   15
      Top             =   4230
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
      Left            =   3465
      TabIndex        =   14
      Top             =   4230
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
      Left            =   3285
      TabIndex        =   13
      Top             =   4680
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
      Left            =   3465
      TabIndex        =   12
      Top             =   4680
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
      Left            =   1485
      TabIndex        =   11
      Top             =   5355
      Width           =   60
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      Caption         =   "Start Indirect Labor"
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
      Left            =   3780
      TabIndex        =   10
      Top             =   2070
      Width           =   4020
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
      Height          =   375
      Left            =   2205
      TabIndex        =   9
      Top             =   2800
      Width           =   945
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
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2800
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
      Height          =   375
      Left            =   4950
      TabIndex        =   7
      Top             =   2800
      Width           =   600
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
      Height          =   375
      Left            =   5670
      TabIndex        =   6
      Top             =   2800
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
      Height          =   375
      Left            =   1260
      TabIndex        =   5
      Top             =   3360
      Width           =   2160
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
      Height          =   375
      Left            =   3510
      TabIndex        =   4
      Top             =   3360
      Width           =   60
   End
   Begin VB.Label lblCCIDCaption 
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
      Height          =   375
      Left            =   1125
      TabIndex        =   3
      Top             =   3825
      Width           =   2280
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
      Height          =   375
      Left            =   3510
      TabIndex        =   2
      Top             =   3825
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
      Height          =   375
      Left            =   4365
      TabIndex        =   1
      Top             =   3360
      Width           =   60
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
      Height          =   375
      Left            =   4365
      TabIndex        =   0
      Top             =   3825
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   2835
      Picture         =   "Frmstind.frx":0000
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   870
   End
End
Attribute VB_Name = "frmStartIndir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintLoaded As Integer

Const vbColorWhite = &H8000005

Public Sub StartIndProcess()
    
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
   If Not inpIndirCd(I) Then Exit Sub
   lblIndirCd = gLaborIn(I).IndirCd
   lblIndirDesc = gSelDesc
   ListOK = False
   
   If div = milw Then
     If gLaborIn(I).IndirCd = "12" Then
       If Not inpWCID(I) Then Exit Sub
       lblVarbData1 = gLaborIn(I).WCID
       If Not inpTlmntCd(I) Then Exit Sub
       lblVarbData2 = gLaborIn(I).UnplnReas
       gLaborIn(I).CCID = gWorkCenter.CCID
       lblCCID = gLaborIn(I).CCID
       If GetCostCenter(gWorkCenter.CCID) Then
            lblccidDesc = gCostCenter.Descr
       End If
     Else
         If Not inpCCID(I) Then Exit Sub
         lblCCID = gLaborIn(I).CCID
         lblccidDesc = gCostCenter.Descr
     End If
   Else
      If Not inpCCID(I) Then Exit Sub
      lblCCID = gLaborIn(I).CCID
      lblccidDesc = gCostCenter.Descr
   End If
   
'   If Not inpCCID(i) Then Exit Sub
'   lblCCID = gLaborIn(i).CCID
'   lblccidDesc = gCostCenter.Descr
   
   If div = corry Then
     If gLaborIn(I).IndirCd = "50" Then
       If Not inpEOJob(I) Then Exit Sub
       lblVarbData3 = gLaborIn(I).OrderNo
     ElseIf gLaborIn(I).IndirCd = "49" Then
          If LdIndirIncn(gLaborIn(I).CCID) <> 0 Then
             If Not getIndirIncn(I) Then Exit Sub
             lblVarbData1 = gLaborIn(I).incenRec
             If Not CallCmpl(I) Then Exit Sub
             lblVarbData2 = gLaborIn(I).OperStatus
          Else
             gLaborIn(I).incenRec = 0
             lblVarbData1 = gLaborIn(I).incenRec
             gLaborIn(I).NumbMachs = 0
             gLaborIn(I).SetStd = 0#
             gLaborIn(I).OperStatus = " "
          End If
     End If
   End If
   If div = bristol Then
      If Left(gLaborIn(I).IndirCd, 1) = "A" Or _
         Left(gLaborIn(I).IndirCd, 1) = "B" Then
           If Not inpPMJob(I) Then Exit Sub
           lblVarbData1 = gLaborIn(I).OrderNo
      End If
   End If
 
 If div = corry Then
    If gEosFlag Then
        If Not inpEndOfShift() Then
            gDone = True
            Unload Me
            Exit Sub
        End If
    End If
End If
  
     
 Call Build_OpenLabor
    
'   If div = corry Then
'      If empClkOutOK Then
'         gLaborIn(i).ErrorLst = ""
'         Call ClockOut
'         Call Upd_Empl_Attd
'         gDone = True
'      End If
'   End If
    
  Unload Me
    
End Sub


Private Sub Form_Activate()
   lblEmplNo = gEmployee.EmplNo
   lblEmplName = gEmployee.EmplName
    
    If gDone Then
'       gDone = False
       Unload Me
    End If
    
    If mintLoaded <> 1 Then
        mintLoaded = 1
        Call StartIndProcess
        mintLoaded = 0
        Unload Me
    End If

End Sub

Private Sub Form_Load()
lblFormTitle.Caption = LoadResString(gLanguageOffset + 202)
lblEmplNoCaption.Caption = LoadResString(gLanguageOffset + 136)
lblNameCaption.Caption = LoadResString(gLanguageOffset + 137)
lblIndirectCodeCaption.Caption = LoadResString(gLanguageOffset + 145)
lblCCIDCaption.Caption = LoadResString(gLanguageOffset + 146)
End Sub


