VERSION 5.00
Begin VB.Form frmDCStdsChgs 
   Caption         =   "Data Collection - Standards Changes"
   ClientHeight    =   6270
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   10665
   WindowState     =   2  'Maximized
   Begin VB.Label lblMchStdType 
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   8550
      TabIndex        =   34
      Top             =   4995
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1620
      Picture         =   "Frmdcstd.frx":0000
      Stretch         =   -1  'True
      Top             =   1395
      Width           =   870
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   270
      TabIndex        =   33
      Top             =   5490
      Width           =   60
   End
   Begin VB.Label lblEorS 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   4860
      TabIndex        =   32
      Top             =   4995
      Width           =   60
   End
   Begin VB.Label lblMachStd 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   8370
      TabIndex        =   31
      Top             =   4995
      Width           =   60
   End
   Begin VB.Label lblFixStd 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   1980
      TabIndex        =   30
      Top             =   4995
      Width           =   60
   End
   Begin VB.Label lblMachineStdCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Mach. Std."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   6300
      TabIndex        =   29
      Top             =   4995
      Width           =   1110
   End
   Begin VB.Label lblESCAption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "E/S"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4365
      TabIndex        =   28
      Top             =   4995
      Width           =   360
   End
   Begin VB.Label lblFixedStdCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fixed Std."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   225
      TabIndex        =   27
      Top             =   4995
      Width           =   1650
   End
   Begin VB.Label lblCCID 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   4860
      TabIndex        =   26
      Top             =   4410
      Width           =   60
   End
   Begin VB.Label lblOperStatus 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   7695
      TabIndex        =   25
      Top             =   4410
      Width           =   60
   End
   Begin VB.Label lblCCIDCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cost Center"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3555
      TabIndex        =   24
      Top             =   4410
      Width           =   1290
   End
   Begin VB.Label lblStatusCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Status"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   7020
      TabIndex        =   23
      Top             =   4410
      Width           =   675
   End
   Begin VB.Label lblOperCd 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   4860
      TabIndex        =   22
      Top             =   4050
      Width           =   60
   End
   Begin VB.Label lblOperCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Oper."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4185
      TabIndex        =   21
      Top             =   4050
      Width           =   600
   End
   Begin VB.Label lblWCID 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   1980
      TabIndex        =   20
      Top             =   4410
      Width           =   60
   End
   Begin VB.Label lblSeqNo 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   1980
      TabIndex        =   19
      Top             =   4050
      Width           =   60
   End
   Begin VB.Label lblWCIDCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Workcenter "
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   270
      TabIndex        =   18
      Top             =   4410
      Width           =   1605
   End
   Begin VB.Label lblSeqNoCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Seq No."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   180
      TabIndex        =   17
      Top             =   4050
      Width           =   1695
   End
   Begin VB.Label lblOrdQty 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   7965
      TabIndex        =   16
      Top             =   2745
      Width           =   60
   End
   Begin VB.Label lblOrderQtyCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Order Qty"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   6345
      TabIndex        =   15
      Top             =   2790
      Width           =   1545
   End
   Begin VB.Label lblCustName 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   3240
      TabIndex        =   14
      Top             =   3150
      Width           =   60
   End
   Begin VB.Label lblCustomerCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   2115
      TabIndex        =   13
      Top             =   3150
      Width           =   1035
   End
   Begin VB.Label lblPartRev 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   8010
      TabIndex        =   12
      Top             =   3105
      Width           =   60
   End
   Begin VB.Label lblRevCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Rev."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   7425
      TabIndex        =   11
      Top             =   3150
      Width           =   495
   End
   Begin VB.Label lblOrdType 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   1260
      TabIndex        =   10
      Top             =   3165
      Width           =   60
   End
   Begin VB.Label lblTypeCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Type"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   585
      TabIndex        =   9
      Top             =   3165
      Width           =   525
   End
   Begin VB.Label lblPartNo 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   3240
      TabIndex        =   8
      Top             =   2805
      Width           =   60
   End
   Begin VB.Label lblPartCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Part"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   2520
      TabIndex        =   7
      Top             =   2790
      Width           =   615
   End
   Begin VB.Label lblJobNo 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   1260
      TabIndex        =   6
      Top             =   2805
      Width           =   60
   End
   Begin VB.Label lblJobNoCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Job No."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   315
      TabIndex        =   5
      Top             =   2790
      Width           =   810
   End
   Begin VB.Label lblEmplName 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   4815
      TabIndex        =   4
      Top             =   2265
      Width           =   60
   End
   Begin VB.Label lblNameCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4095
      TabIndex        =   3
      Top             =   2265
      Width           =   600
   End
   Begin VB.Label lblEmplNo 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   2655
      TabIndex        =   2
      Top             =   2265
      Width           =   60
   End
   Begin VB.Label lblEmplNoCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Empl No."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1665
      TabIndex        =   1
      Top             =   2295
      Width           =   945
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      Caption         =   "Standards Changes"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1530
      Width           =   5175
   End
End
Attribute VB_Name = "frmDCStdsChgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintLoaded As Integer

Const vbColorWhite = &H8000005



Public Sub StdProcess()
    Dim strEntry As String
    
    Dim I As Integer
        
    Dim lScrapTot As Long
    Dim strRecType As String
    Dim j As Integer
   
        
    gLaborIn(1).MachRec = "N"
    I = 1
        
    'If Not inpEmplNo(i) Then Exit Sub
    lblEmplNo.Caption = gEmployee.EmplNo
    lblEmplName.Caption = gEmployee.EmplName
    DoEvents
        
    If Not inpJobAndSeqStds(I) Then Exit Sub
    lblPartNo.Caption = gJobMaster.PartNo
    lblCustName.Caption = gJobMaster.CustName
    lblJobNo.Caption = gJobMaster.OrderNo
    lblOrdQty.Caption = gJobMaster.OrderQty
    lblPartRev.Caption = gJobMaster.PartRev
    lblOrdType.Caption = gJobMaster.OrderType
        
    lblSeqNo.Caption = gLaborIn(I).SeqNo
    lblWCID.Caption = gLaborIn(I).WCID
    lblCCID.Caption = gLaborIn(I).CCID
    lblOperCd.Caption = gLaborIn(I).Oper
  '  lblFixStd.Caption = gLaborIn(i).SetStd
   ' lblMachStd.Caption = gLaborIn(i).MchStd
    lblEorS.Caption = gLaborIn(I).EOrS
    If gLaborIn(I).SetStd > 0 Then
         lblFixStd.Caption = Format(gLaborIn(I).SetStd, gOneDecPos)
    Else
        lblFixStd.Caption = ""
    End If
    If gLaborIn(I).MchStd > 0 Then
        If gJobOpers.TimeBasisCd = "1" Then
           lblMachStd.Caption = Format(gLaborIn(I).MchStd, gOneDecPos)
           lblMchStdType.Caption = "hrs/pc"
       ElseIf gJobOpers.TimeBasisCd = "4" Then
           lblMachStd.Caption = Format(gLaborIn(I).MchStd, "#######.0")
           lblMchStdType.Caption = "pcs/hr"
       ElseIf gJobOpers.TimeBasisCd = "5" Then
          lblMachStd.Caption = Format(gLaborIn(I).MchStd, gOneDecPos)
          lblMchStdType.Caption = "hrs"
       Else
          lblMachStd.Caption = Format(gLaborIn(I).MchStd, "#######.0")
          lblMchStdType.Caption = ""
       End If
    Else
       lblMachStd.Caption = ""
       lblMchStdType.Caption = ""
    End If
  
    
    
    If gJobOpers.TimeBasisCd = "1" And gLaborIn(I).MchStd > 0 Then
          lblMachStd.Caption = Format(gLaborIn(I).MchStd, gOneDecPos)
          lblMchStdType.Caption = "hrs/pc"
    ElseIf gJobOpers.TimeBasisCd = "4" And gLaborIn(I).MchStd > 0 Then
         lblMachStd.Caption = Format(gLaborIn(I).MchStd, "#######")
         lblMchStdType.Caption = "pcs/hr"
    ElseIf gJobOpers.TimeBasisCd = "5" And gLaborIn(I).MchStd > 0 Then
         lblMachStd.Caption = Format(gLaborIn(I).MchStd, gOneDecPos)
         lblMchStdType.Caption = "hrs"
    Else
         lblMachStd.Caption = Format(gLaborIn(I).MchStd, "#######")
         lblMchStdType.Caption = ""
    End If
 
    lblEorS.Caption = gLaborIn(I).EOrS
        
    If gLaborIn(I).SetStd > 0 Then
       If Not inpFixStd(I) Then Exit Sub
    Else
       If Not inpMchStd(I) Then Exit Sub
    End If
    
    If Not inpComment(I) Then Exit Sub
    
    gLaborIn(I).OperStatus = "P"
    
    LoopCnt = 1
    Call GetSysDateTime
    gStdChg.companyCd = gCompanyCd
    gStdChg.EmplNo = gEmployee.EmplNo
    gStdChg.OrderNo = gJobMaster.OrderNo
    gStdChg.SeqNumb = gJobOpers.OperSeqNo
    gStdChg.OrigFixStd = gJobOpers.SetStd
    gStdChg.RevFixedStd = gLaborIn(1).SetStd
    gStdChg.origMchStd = gJobOpers.MchStd
    gStdChg.RevMchStd = gLaborIn(1).MchStd
    gStdChg.ReasonText = mComment
    gStdChg.ShiftCd = gEmployee.GenShift
    gStdChg.Actdate = gEmployee.Actdate
    gStdChg.transdate = gSysDateTime
    gStdChg.ProcessFlg = " "
    gStdChg.ProcessDate = dBlankDate
       
    Call Ins_Std_chg
    
    Unload Me
End Sub

Private Sub Form_Activate()
    If gDone Then
        Unload Me
    End If
    
    If mintLoaded <> 1 Then
        mintLoaded = 1
        Call StdProcess
        mintLoaded = 0
        Unload Me
    End If

End Sub

Private Sub Form_Load()
'lblFormTitle.Caption = LoadResString(gLanguageOffset + 205)
lblEmplNoCaption.Caption = LoadResString(gLanguageOffset + 136)
lblNameCaption.Caption = LoadResString(gLanguageOffset + 137)
lblJobNoCaption.Caption = LoadResString(gLanguageOffset + 150)
lblPartCaption.Caption = LoadResString(gLanguageOffset + 151)
lblOrderQtyCaption.Caption = LoadResString(gLanguageOffset + 152)
lblTypeCaption.Caption = LoadResString(gLanguageOffset + 153)
lblCustomerCaption.Caption = LoadResString(gLanguageOffset + 154)
lblRevCaption.Caption = LoadResString(gLanguageOffset + 155)
lblSeqNoCaption.Caption = LoadResString(gLanguageOffset + 156)
lblOperCaption.Caption = LoadResString(gLanguageOffset + 157)
lblWCIDCaption.Caption = LoadResString(gLanguageOffset + 158)
lblCCIDCaption.Caption = LoadResString(gLanguageOffset + 159)
lblStatusCaption.Caption = LoadResString(gLanguageOffset + 160)
lblFixedStdCaption.Caption = LoadResString(gLanguageOffset + 162)
lblESCaption.Caption = LoadResString(gLanguageOffset + 163)
lblMachineStdCaption.Caption = LoadResString(gLanguageOffset + 164)

End Sub


