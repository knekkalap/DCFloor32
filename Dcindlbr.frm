VERSION 5.00
Begin VB.Form frmIndLabor 
   Caption         =   "Data Collection - Labor Entry"
   ClientHeight    =   7965
   ClientLeft      =   3465
   ClientTop       =   1530
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7965
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1935
      Picture         =   "Dcindlbr.frx":0000
      Stretch         =   -1  'True
      Top             =   1305
      Width           =   870
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
      Left            =   2610
      TabIndex        =   19
      Top             =   4050
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
      Left            =   2430
      TabIndex        =   18
      Top             =   4050
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
      Left            =   6345
      TabIndex        =   17
      Top             =   3600
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
      Left            =   6165
      TabIndex        =   16
      Top             =   3600
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
      Left            =   2610
      TabIndex        =   15
      Top             =   3600
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
      Left            =   2430
      TabIndex        =   14
      Top             =   3600
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
      Height          =   270
      Left            =   3465
      TabIndex        =   13
      Top             =   3150
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
      Left            =   3465
      TabIndex        =   12
      Top             =   2685
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
      Left            =   540
      TabIndex        =   11
      Top             =   4950
      Width           =   60
   End
   Begin VB.Label lblIndHrs 
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
      Left            =   2610
      TabIndex        =   10
      Top             =   4500
      Width           =   60
   End
   Begin VB.Label lblIndirectHoursCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Indirect Hours"
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
      Left            =   990
      TabIndex        =   9
      Top             =   4500
      Width           =   1515
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
      Left            =   2610
      TabIndex        =   8
      Top             =   3150
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
      Height          =   270
      Left            =   1215
      TabIndex        =   7
      Top             =   3150
      Width           =   1290
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
      Left            =   2610
      TabIndex        =   6
      Top             =   2685
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
      Left            =   1080
      TabIndex        =   5
      Top             =   2685
      Width           =   1440
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
      Left            =   4770
      TabIndex        =   4
      Top             =   2145
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
      Left            =   4005
      TabIndex        =   3
      Top             =   2145
      Width           =   600
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
      Left            =   2340
      TabIndex        =   2
      Top             =   2145
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
      Left            =   1305
      TabIndex        =   1
      Top             =   2160
      Width           =   945
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      Caption         =   "Indirect Labor"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   1395
      Width           =   3735
   End
End
Attribute VB_Name = "frmIndLabor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintLoaded As Integer

Const vbColorWhite = &H8000005


Public Sub INProcess()
    
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
    If div = hyson_products Then
     If gLaborIn(I).IndirCd = "01" Or _
        gLaborIn(I).IndirCd = "19" Or _
        gLaborIn(I).IndirCd = "20" Then
        frmIndLabor.lblVarbFld1.Caption = "JBA Work Order"
        If Not inpWorkOrder(I) Then Exit Sub
        lblVarbData1 = gLaborIn(I).OrderNo
     End If
    End If
   
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
  
  If div = corry Then
       gLaborIn(I).inHours = CalcActHrs()
  ElseIf div = hyson_products And Not gbInpIndirHrs Then
       gLaborIn(I).inHours = CalcActHrs()
  ElseIf div = stroms Then
       gLaborIn(I).inHours = CalcActHrs()
  ElseIf Not inpIndirHrs(I) Then
         Exit Sub
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
'       gDone = False
       Unload Me
    End If
    
    If mintLoaded <> 1 Then
        mintLoaded = 1
        Call INProcess
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

Private Sub Form_Load()
lblFormTitle.Caption = LoadResString(gLanguageOffset + 238)
lblEmplNoCaption.Caption = LoadResString(gLanguageOffset + 136)
lblNameCaption.Caption = LoadResString(gLanguageOffset + 137)
lblIndirectCodeCaption.Caption = LoadResString(gLanguageOffset + 145)
lblCCIDCaption.Caption = LoadResString(gLanguageOffset + 146)
lblIndirectHoursCaption = LoadResString(gLanguageOffset + 239)
End Sub


