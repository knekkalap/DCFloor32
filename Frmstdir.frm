VERSION 5.00
Begin VB.Form frmStartDirect 
   Caption         =   "Data Collection - Start Direct Labor"
   ClientHeight    =   8340
   ClientLeft      =   2370
   ClientTop       =   1485
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8340
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin VB.Line Line2 
      X1              =   90
      X2              =   11295
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Label lblErrMsg 
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
      Height          =   270
      Left            =   270
      TabIndex        =   38
      Top             =   6165
      Width           =   60
   End
   Begin VB.Label lblInfo 
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
      Height          =   375
      Left            =   270
      TabIndex        =   37
      Top             =   5805
      Width           =   60
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      Caption         =   "Start Direct Labor"
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
      Left            =   4590
      TabIndex        =   36
      Top             =   2430
      Width           =   3675
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
      Left            =   450
      TabIndex        =   35
      Top             =   3100
      Width           =   1800
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
      Left            =   2430
      TabIndex        =   34
      Top             =   3100
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
      Left            =   4230
      TabIndex        =   33
      Top             =   3100
      Width           =   1230
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
      Left            =   5580
      TabIndex        =   32
      Top             =   3100
      Width           =   60
   End
   Begin VB.Label lblJobNoCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Job No."
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
      Left            =   225
      TabIndex        =   31
      Top             =   3600
      Width           =   1665
   End
   Begin VB.Label lblJobNo 
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
      Left            =   2025
      TabIndex        =   30
      Top             =   3600
      Width           =   60
   End
   Begin VB.Label lblPartCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Part"
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
      Left            =   3375
      TabIndex        =   29
      Top             =   3600
      Width           =   1560
   End
   Begin VB.Label lblPartNo 
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
      Left            =   5085
      TabIndex        =   28
      Top             =   3600
      Width           =   60
   End
   Begin VB.Label lblTypeCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Type"
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
      Left            =   225
      TabIndex        =   27
      Top             =   4000
      Width           =   1650
   End
   Begin VB.Label lblOrdType 
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
      Left            =   2025
      TabIndex        =   26
      Top             =   4000
      Width           =   60
   End
   Begin VB.Label lblRevCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Rev."
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
      Left            =   9225
      TabIndex        =   25
      Top             =   4000
      Width           =   1125
   End
   Begin VB.Label lblPartRev 
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
      Left            =   10440
      TabIndex        =   24
      Top             =   4000
      Width           =   60
   End
   Begin VB.Label lblCustomerCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Customer"
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
      Left            =   3330
      TabIndex        =   23
      Top             =   4000
      Width           =   1620
   End
   Begin VB.Label lblCustName 
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
      Left            =   5085
      TabIndex        =   22
      Top             =   4000
      Width           =   60
   End
   Begin VB.Label lblOrderQtyCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Order Qty"
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
      Left            =   8685
      TabIndex        =   21
      Top             =   3600
      Width           =   1635
   End
   Begin VB.Label lblOrdQty 
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
      Left            =   10440
      TabIndex        =   20
      Top             =   3600
      Width           =   60
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11280
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblSeqNoCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Seq No."
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
      Left            =   720
      TabIndex        =   19
      Top             =   4600
      Width           =   1575
   End
   Begin VB.Label lblWCIDCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Workcenter ID"
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
      Left            =   120
      TabIndex        =   18
      Top             =   5000
      Width           =   2175
   End
   Begin VB.Label lblSeqNo 
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
      Left            =   2520
      TabIndex        =   17
      Top             =   4600
      Width           =   135
   End
   Begin VB.Label lblWCID 
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
      Left            =   2520
      TabIndex        =   16
      Top             =   5000
      Width           =   135
   End
   Begin VB.Label lblOperCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Oper."
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
      Left            =   5520
      TabIndex        =   15
      Top             =   4600
      Width           =   1455
   End
   Begin VB.Label lblUnplanReasCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Unplan Reas."
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
      Left            =   7440
      TabIndex        =   14
      Top             =   5000
      Width           =   2175
   End
   Begin VB.Label lblOperCd 
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
      Left            =   7080
      TabIndex        =   13
      Top             =   4600
      Width           =   135
   End
   Begin VB.Label lblUnplnReas 
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
      Left            =   9720
      TabIndex        =   12
      Top             =   5000
      Width           =   135
   End
   Begin VB.Label lblStatusCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Status"
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
      Left            =   5760
      TabIndex        =   11
      Top             =   5000
      Width           =   1215
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
      Left            =   3240
      TabIndex        =   10
      Top             =   5000
      Width           =   1695
   End
   Begin VB.Label lblOperStatus 
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
      Left            =   7080
      TabIndex        =   9
      Top             =   5000
      Width           =   135
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
      Left            =   5040
      TabIndex        =   8
      Top             =   5000
      Width           =   135
   End
   Begin VB.Label lblFixedStdCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fixed Std."
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
      Left            =   720
      TabIndex        =   7
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblESCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "E/S"
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
      Left            =   5880
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblMachineStdCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Machine Std."
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
      Left            =   7440
      TabIndex        =   5
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label lblFixStd 
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
      Left            =   2520
      TabIndex        =   4
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label lblMachStd 
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
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   10680
      TabIndex        =   3
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label lblEorS 
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
      Left            =   7080
      TabIndex        =   2
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label lblUnplanned 
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
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   4600
      Width           =   15
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   3645
      Picture         =   "Frmstdir.frx":0000
      Stretch         =   -1  'True
      Top             =   2295
      Width           =   870
   End
   Begin VB.Label lblMchStdType 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   0
      Top             =   5400
      Width           =   135
   End
End
Attribute VB_Name = "frmStartDirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mintLoaded As Integer

Const vbColorWhite = &H8000005


Public Function getStartUnplnData(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDirLabor.getLabUnplnData(i)
'Description      : Routine for the entry of Unplanned Data, for a single
'                   machine entry for direct labor for all divisions. Boolean
'                   values are globally set by division and type of
'                   Labor entry, to determine which options are used for the
'                   labor process.
'Parameters       : None
'Called From      : frmDirLabor.DirLaborProcess
'Procedures Used  : ErrorChk, inpJobAndSeq, dispValues, inpWCID, GetCostCenter,
'                   inpProdQty, getScrapCounts, inpSetHrs, inpMchHours, inpAddonHours,
'                   inpEndOfShift, CallCmpl, inpNoOfLoads, CalcBtErnHrs, CalcBatchHrs,
'                   inpCycleTime, CalcBatchHrs, getBatchIncn, ErrorBeep, ErrBox
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    getStartUnplnData = False
    Call initJobOpers
    
    Screen.ActiveForm.lblJobNo.Caption = gJobMaster.OrderNo
    Do Until gDone Or getStartUnplnData
        'gDone = True        'Default for exit to Main Menu
        If Not inpWCID(iEntry) Then
            gDone = True
            Exit Function
        End If
        Screen.ActiveForm.lblWCID.Caption = gLaborIn(iEntry).WCID
        'JC New Here
        gLaborIn(iEntry).CCID = gWorkCenter.CCID
        If Not GetCostCenter(gWorkCenter.CCID) Then
            gCostCenter.PaidLunch = "N"
        Else
            Screen.ActiveForm.lblCCID.Caption = gWorkCenter.CCID
        End If
        
        
        If Not GetWorkCenter(gLaborIn(iEntry).WCID) Then
            gCostCenter.PaidLunch = "N"
        End If
        
        If Not inpOperNumb(iEntry) Then
            gDone = True
            Exit Function
        End If

        Screen.ActiveForm.lblOperCd.Caption = gLaborIn(iEntry).Oper
        gLaborIn(iEntry).MchStd = 0
        gLaborIn(iEntry).SetStd = 0
        
        If gbInpStds Then
          If Not inpFixStd(iEntry) Then
              gDone = True
              Exit Function
          End If
          Screen.ActiveForm.lblFixStd.Caption = gLaborIn(iEntry).SetStd
          If gLaborIn(iEntry).SetStd = 0 Then
              If Not inpMchStd(iEntry) Then
                  gDone = True
                  Exit Function
              End If
          End If
          Screen.ActiveForm.lblMachStd.Caption = gLaborIn(iEntry).MchStd
        End If
        If Not inpUnplnReas(iEntry) Then Exit Function
        Screen.ActiveForm.lblUnplnReas.Caption = gLaborIn(iEntry).UnplnReas
        
        getStartUnplnData = True 'Exit loop
    Loop
End Function







Public Sub ClearForm()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDirLabor.ClearForm
'Description      : Clears the label fields that display data
'Parameters       : None
'Called From      : DirLaborProcess (after an entry is done)
'Procedures Used  : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
        lblCCID.Caption = ""
        lblPartNo.Caption = ""
        lblCustName.Caption = ""
        lblJobNo.Caption = ""
        lblOrdQty.Caption = ""
        lblPartRev.Caption = ""
        lblOrdType.Caption = ""
        
        lblSeqNo.Caption = ""
        lblWCID.Caption = ""
        lblOperCd.Caption = ""
        lblUnplnReas.Caption = ""
        lblFixStd.Caption = ""
        lblMachStd.Caption = ""
        lblMchStdType.Caption = ""
        lblEorS.Caption = ""

End Sub

Public Function inpStartJobSeq(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDirLabor.inpLabJobSeq(i)
'Description      : Routine for the entry of Job and Seq No, for a single
'                   machine entry for direct labor for all divisions. Boolean
'                   values are globally set by division and type of
'                   Labor entry, to determine which options are used for the
'                   labor process.
'Parameters       : None
'Called From      : frmDirLabor.DirLaborProcess
'Procedures Used  : GetJobMaster, getLabUnplnData, GetJobOper, GetWorkCenter,
'                   ValidNumber, basGetEntry, ErrorBeep, ErrBox
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    Dim strEntry As String
    Dim strDataId As String
    Dim strSeqNo As String
    
    inpStartJobSeq = False        'to force loop back till done
    
    Do Until inpStartJobSeq Or gDone
        Do Until inpStartJobSeq Or gDone       'Extra loop for exit do stmts
         '   ErrBox "If labor is unplanned, key in only job number and Enter"
            
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 165)
            gEntry.EntryType = 0        'Any Printable Characters
            strEntry = basGetEntry()
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            strDataId = UCase(Left(strEntry, 2))
            If strDataId = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If strDataId = "NU" Then
                ErrBox LoadResString(gLanguageOffset + 28), 2
                ErrorBeep
                Exit Do           'Loop Again
            End If
            
            If Len(strDataId) < 1 Then
                ErrBox LoadResString(gLanguageOffset + 78), 2
                ErrorBeep
                Exit Do           'Loop Again
            End If
                
            If Len(strEntry) > 11 Then
                If strDataId = "JS" Then
                    gLaborIn(iEntry).OrderNo = UCase(Mid(strEntry, 3, 7))
                    strSeqNo = Mid(strEntry, 10, 4)
                Else
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            Else
                gLaborIn(iEntry).OrderNo = UCase(Left(strEntry, 7))
                strSeqNo = Mid(strEntry, 8, 4)
            End If
            
         If Not GetJobMaster(gLaborIn(iEntry).OrderNo) Then
                If Left(gLaborIn(iEntry).OrderNo, 2) = "09" Then
                    Screen.ActiveForm.lblErrMsg.Caption = LoadResString(gLanguageOffset + 236)
                    Screen.ActiveForm.lblErrMsg.BackColor = vbColorWhite
                    strSeqNo = "000 "
                    gLaborIn(iEntry).MatlWgt = 0
                    gJobMaster.Actdate = 0
                    gJobMaster.OrderNo = gLaborIn(iEntry).OrderNo
                    gJobMaster.OrderType = "SO"
                Else
                    ErrBox LoadResString(gLanguageOffset + 61), 2
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
         ElseIf gJobMaster.OrderStat = "C" Then
                ErrBox LoadResString(gLanguageOffset + 293), 2
                ErrorBeep
                Exit Do           'Loop Again
         Else
                gLaborIn(iEntry).MatlWgt = gJobMaster.MatlWgt
         End If
            
            If Len(strSeqNo) = 0 Then
                strSeqNo = "000 "
            Else
                If Not ValidNumber(strSeqNo) Then
                    ErrBox LoadResString(gLanguageOffset + 111), 2
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            End If
            
            gLaborIn(iEntry).SeqNo = CInt(strSeqNo)
            If gLaborIn(iEntry).SeqNo = 990 _
               Or gLaborIn(iEntry).SeqNo = 0 Then
                 gLaborIn(iEntry).SeqNo = 0
                 gLaborIn(iEntry).EOrS = "S"
                 lblErrMsg.Caption = ""
                 lblUnplanned = LoadResString(gLanguageOffset + 240)
                 If Not getStartUnplnData(iEntry) Then
                    Exit Do           'Loop Again
                 End If
                 inpStartJobSeq = True     'Done - Unplanned
            Else
                lblUnplanned = ""
                If Not GetJobOper(gLaborIn(iEntry).OrderNo, _
                        gLaborIn(iEntry).SeqNo) Then
                    ErrBox LoadResString(gLanguageOffset + 87), 2
                    ErrorBeep
                    Exit Do
                ElseIf gJobOpers.WCID = "-----" Then
                        ErrBox LoadResString(gLanguageOffset + 38), 2
                        ErrorBeep
                        Exit Do
                ElseIf Not GetWorkCenter(gJobOpers.WCID) Then
                        ErrBox LoadResString(gLanguageOffset + 103), 2
                        ErrorBeep
                        Exit Do
                ElseIf gJobOpers.OperStatus = "C" Then
                        ErrBox LoadResString(gLanguageOffset + 80), 2
                        ErrorBeep
                        Exit Do
                ElseIf OpenDirectLabor(gEmployee.EmplNo, gLaborIn(iEntry).OrderNo, gLaborIn(iEntry).SeqNo, gEmployee.Actdate) Then
                        ErrBox LoadResString(gLanguageOffset + 108), 2
                        ErrorBeep
                        Exit Do
                ElseIf OkayToStartLabor(gEmployee.EmplNo, gLaborIn(iEntry).OrderNo, gLaborIn(iEntry).SeqNo, gEmployee.Actdate) = False Then
                        ErrBox LoadResString(gLanguageOffset + 295), 2
                        ErrorBeep
                        Exit Do
                ElseIf Not CheckValidQty(gLaborIn(iEntry).OrderNo, gLaborIn(iEntry).SeqNo, 0, "S") Then
                       ErrBox LoadResString(gLanguageOffset + 291), 2
                       ErrorBeep
                       Exit Do
                End If
                gLaborIn(iEntry).UnplnReas = "  "
                gLaborIn(iEntry).WCID = gJobOpers.WCID
                gLaborIn(iEntry).Oper = gJobOpers.Oper
                gLaborIn(iEntry).EOrS = gJobOpers.EorSStd
                gLaborIn(iEntry).SetStd = gJobOpers.SetStd
                gLaborIn(iEntry).MchStd = gJobOpers.MchStd
                inpStartJobSeq = True     'Done - Planned
            End If
        Loop
    Loop
 lblErrMsg.Caption = ""
End Function


Public Sub StartDirLaborProcess()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDirLabor.DirLaborProcess
'Description      : Routine that processes direct labor for all divisions.
'                   Boolean values are globally set by division and type of
'                   Labor entry, to determine which options are used for the
'                   labor process.
'Parameters       : None
'Called From      : frmDirLabor.Form_Activate
'Procedures Used  : inpEndOfShift, inpNumbMachs, inpLabJobSeq, inpSetHours,
'                   inpMchHours, inpAddonHours, inpProdQty, getScrapCounts,
'                   CallCmpl, inpNoOfLoads, inpCycleTime, CalcBatchHrs,
'                   getBatchIncn, DirLaborReEnter, CalcActHrs, AnotherRec,
'                   Build_Labor_Recs, ClockOut, Upd_Empl_Attd, ErrorBeep,
'                   ErrBox
'Return Parameters: None
'Tables Updates   : Updates done in Build_Labor_Recs, Upd_Empl_Attd
'==============================================================================
    
    Dim strEntry As String
    
    Dim I As Integer
        
    Dim lScrapTot As Long
    Dim CalcHours As Currency
    Dim strRecType As String
    Dim j As Integer
    Dim bEnter As Boolean
    Dim bZeroIt As Boolean
    
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
    
    Call Init_varbs     'Initialize gLaborIn Array
    Call checkClockOut
   
    If gbEOShift And gEosFlag Then
        If Not inpEndOfShift() Then
            gDone = True
            Unload Me
            Exit Sub
        End If
    End If
       
    If gbAnotherRec Then
        gLaborIn(1).NumbMachs = 15
    ElseIf gbNumbMach Then
        If Not inpNumbMachs() Then
            gDone = True
            Unload Me
            Exit Sub
        End If
    Else
        gLaborIn(1).NumbMachs = 1
    End If
    
    If gbMachRecY Then
        gLaborIn(1).MachRec = "Y"
    Else
        gLaborIn(1).MachRec = "N"
    End If
    
    bEnter = True
    I = 1
    LoopCnt = I     'Keep global variable in sync.
    While (I <= gLaborIn(1).NumbMachs) And bEnter
        
        If gbNumbMach Then
            lblInfo.Caption = LoadResString(gLanguageOffset + 237) & CStr(I) _
                & " of " & CStr(gLaborIn(1).NumbMachs)
        End If
        
        lblEmplNo.Caption = gEmployee.EmplNo
        lblEmplName.Caption = gEmployee.EmplName
        
        If Not inpStartJobSeq(I) Then Exit Sub
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
        lblUnplnReas.Caption = gLaborIn(I).UnplnReas
        If div <> burl Then
            lblEorS.Caption = gLaborIn(I).EOrS
            If gLaborIn(I).SetStd > 0 Then
                lblFixStd.Caption = Format(gLaborIn(I).SetStd, "#####.0")
           Else
               lblFixStd.Caption = ""
           End If
           If gLaborIn(I).MchStd > 0 Then
              If gJobOpers.TimeBasisCd = "1" Then
                lblMachStd.Caption = Format(gLaborIn(I).MchStd, "#####.0")
                lblMchStdType.Caption = "hrs/pc"
              ElseIf gJobOpers.TimeBasisCd = "4" Then
                lblMachStd.Caption = Format(gLaborIn(I).MchStd, "#######.0")
                lblMchStdType.Caption = "pcs/hr"
              ElseIf gJobOpers.TimeBasisCd = "5" Then
                lblMachStd.Caption = Format(gLaborIn(I).MchStd, "#####.0")
                lblMchStdType.Caption = "hrs"
             Else
                lblMachStd.Caption = Format(gLaborIn(I).MchStd, "#######.0")
                lblMchStdType.Caption = ""
             End If
           Else
              lblMachStd.Caption = ""
              lblMchStdType.Caption = ""
           End If
       End If
        
        If gLaborIn(I).SeqNo <> 0 _
                And gLaborIn(I).SeqNo <> 990 Then
            
            If Not inpWcAns(I) Then Exit Sub
        End If
        lblCCID.Caption = gLaborIn(I).CCID
        
        
        If I <> 1 Then
            If gbNumbMach Then
                gLaborIn(I).MachRec = "Y"
                gLaborIn(I).NumbMachs = gLaborIn(1).NumbMachs
                gLaborIn(I).incenRec = gLaborIn(1).incenRec
            ElseIf gbAnotherRec Then
                gLaborIn(I).MachRec = "Y"
                gLaborIn(I).NumbMachs = 1
                gLaborIn(I).incenRec = 0
            End If
        End If
        
        'If rework, then clobber the RecType with R before update
        If gbRecTypeR Then
            gLaborIn(I).RecType = "R"
            Call RecordError("R")
        End If
        
        I = I + 1
        LoopCnt = I     'keep global variable in sync
        
        Rem if gbInpSetAndMach - make sure only one entry
        Rem (Just in Case)
        If gbInpSetAndMach Then
            gbAnotherRec = False
            gbNumbMach = False
        End If
        
        If gbAnotherRec Then
           Select Case AnotherRec()
              Case 0
                      'selected No when asked 'enter another record?'
                       bEnter = False
             Case 1
                      'selected Yes when asked 'enter another record?'
                      bEnter = True
             Case Else
                     ' escaped from anotherRec
                     Exit Sub
          End Select
        ElseIf Not gbNumbMach Then
            Rem Only enter one machine at a time
            bEnter = False
        End If
    
    ClearForm
    Wend
    
    'LoopCnt = gLaborIn(1).NumbMachs
    LoopCnt = LoopCnt - 1
    
    Rem Reset NumbMachs to 1 once looping is done
    If gbAnotherRec Then
        gLaborIn(1).NumbMachs = 1
    End If
    
    Call Build_OpenLabor
  
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
        Call StartDirLaborProcess
        mintLoaded = 0
        Unload Me
    End If
End Sub


Private Sub Form_Load()
lblFormTitle.Caption = LoadResString(gLanguageOffset + 203)
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
lblUnplanReasCaption.Caption = LoadResString(gLanguageOffset + 161)
lblFixedStdCaption.Caption = LoadResString(gLanguageOffset + 162)
lblESCaption.Caption = LoadResString(gLanguageOffset + 163)
lblMachineStdCaption.Caption = LoadResString(gLanguageOffset + 164)
End Sub


