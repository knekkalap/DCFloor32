VERSION 5.00
Begin VB.Form frmAssyCmpl 
   Caption         =   "Data Collection - Assembly Complete"
   ClientHeight    =   8460
   ClientLeft      =   4335
   ClientTop       =   1875
   ClientWidth     =   10575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8460
   ScaleWidth      =   10575
   WindowState     =   2  'Maximized
   Begin VB.Label lblActQty 
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
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   8280
      TabIndex        =   42
      Top             =   6030
      Width           =   60
   End
   Begin VB.Label lblActQtyCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Qty Already Reported"
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
      Height          =   255
      Left            =   4800
      TabIndex        =   41
      Top             =   6030
      Width           =   3375
   End
   Begin VB.Line Line3 
      X1              =   180
      X2              =   10080
      Y1              =   3735
      Y2              =   3735
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      Caption         =   "Assembly Complete"
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
      TabIndex        =   40
      Top             =   2565
      Width           =   6450
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
      Left            =   855
      TabIndex        =   39
      Top             =   3375
      Width           =   1665
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
      Left            =   2655
      TabIndex        =   38
      Top             =   3375
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
      Left            =   4095
      TabIndex        =   37
      Top             =   3375
      Width           =   1365
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
      TabIndex        =   36
      Top             =   3375
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
      Height          =   270
      Left            =   135
      TabIndex        =   35
      Top             =   3960
      Width           =   1395
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
      Height          =   270
      Left            =   1665
      TabIndex        =   34
      Top             =   3960
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
      Height          =   270
      Left            =   3375
      TabIndex        =   33
      Top             =   3960
      Width           =   1065
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
      Height          =   270
      Left            =   4545
      TabIndex        =   32
      Top             =   3960
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
      Height          =   270
      Left            =   180
      TabIndex        =   31
      Top             =   3960
      Width           =   1335
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
      Height          =   270
      Left            =   1665
      TabIndex        =   30
      Top             =   4335
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
      Height          =   270
      Left            =   8235
      TabIndex        =   29
      Top             =   4365
      Width           =   945
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
      Height          =   270
      Left            =   9270
      TabIndex        =   28
      Top             =   4365
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
      Height          =   270
      Left            =   3420
      TabIndex        =   27
      Top             =   4365
      Width           =   1035
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
      Height          =   270
      Left            =   4545
      TabIndex        =   26
      Top             =   4365
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
      Height          =   270
      Left            =   7425
      TabIndex        =   25
      Top             =   3960
      Width           =   1680
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
      Height          =   270
      Left            =   9270
      TabIndex        =   24
      Top             =   3960
      Width           =   60
   End
   Begin VB.Line Line1 
      X1              =   180
      X2              =   10035
      Y1              =   4770
      Y2              =   4770
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
      Height          =   270
      Left            =   495
      TabIndex        =   23
      Top             =   4950
      Width           =   1875
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
      Height          =   270
      Left            =   495
      TabIndex        =   22
      Top             =   5310
      Width           =   1875
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
      Height          =   270
      Left            =   2520
      TabIndex        =   21
      Top             =   4950
      Width           =   60
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
      Height          =   270
      Left            =   2520
      TabIndex        =   20
      Top             =   5310
      Width           =   60
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
      Height          =   270
      Left            =   5055
      TabIndex        =   19
      Top             =   4950
      Width           =   600
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
      Height          =   270
      Left            =   5760
      TabIndex        =   18
      Top             =   4950
      Width           =   60
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
      Height          =   270
      Left            =   7245
      TabIndex        =   17
      Top             =   5310
      Width           =   900
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
      Left            =   3735
      TabIndex        =   16
      Top             =   5310
      Width           =   1920
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
      Height          =   270
      Left            =   8235
      TabIndex        =   15
      Top             =   5310
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
      Left            =   5760
      TabIndex        =   14
      Top             =   5310
      Width           =   60
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
      Height          =   270
      Left            =   495
      TabIndex        =   13
      Top             =   5670
      Width           =   1860
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
      Height          =   270
      Left            =   4455
      TabIndex        =   12
      Top             =   5670
      Width           =   1170
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
      Height          =   270
      Left            =   6795
      TabIndex        =   11
      Top             =   5670
      Width           =   1380
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
      Height          =   270
      Left            =   2520
      TabIndex        =   10
      Top             =   5670
      Width           =   60
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
      Height          =   270
      Left            =   9270
      TabIndex        =   9
      Top             =   5670
      Width           =   60
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
      Height          =   270
      Left            =   5760
      TabIndex        =   8
      Top             =   5670
      Width           =   60
   End
   Begin VB.Line Line2 
      X1              =   180
      X2              =   9960
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label lblQuantityCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Quantity"
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
      Left            =   600
      TabIndex        =   7
      Top             =   6540
      Width           =   1830
   End
   Begin VB.Label lblProdQty 
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
      TabIndex        =   6
      Top             =   6540
      Width           =   60
   End
   Begin VB.Label lblScrapQtyCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Scrap Quantity"
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
      Left            =   5580
      TabIndex        =   5
      Top             =   6540
      Width           =   2595
   End
   Begin VB.Label lblScrapQty 
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
      Left            =   8280
      TabIndex        =   4
      Top             =   6540
      Width           =   60
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
      Left            =   495
      TabIndex        =   3
      Top             =   7920
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
      Height          =   270
      Left            =   495
      TabIndex        =   2
      Top             =   7590
      Width           =   60
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
      Height          =   270
      Left            =   2700
      TabIndex        =   1
      Top             =   4950
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1800
      Picture         =   "Frmassyc.frx":0000
      Stretch         =   -1  'True
      Top             =   2475
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
      Height          =   270
      Left            =   9450
      TabIndex        =   0
      Top             =   5670
      Width           =   60
   End
End
Attribute VB_Name = "frmAssyCmpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintLoaded As Integer

Const vbColorWhite = &H8000005

Public Function inpAssyJobSeq(iEntry As Integer) As Boolean
'==============================================================================
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
    Dim asmbcount As Integer
    
    inpAssyJobSeq = False        'to force loop back till done
    
    Do Until inpAssyJobSeq Or gDone
        Do Until inpAssyJobSeq Or gDone       'Extra loop for exit do stmts
            
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 201)
            gEntry.EntryType = 0        'Any Printable Characters
            strEntry = basGetEntry()
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            strDataId = UCase(Left(strEntry, 2))
            
            If strDataId = "NU" Then
                Call ErrBox(LoadResString(gLanguageOffset + 28), 2)
                ErrorBeep
                Exit Do           'Loop Again
            End If
            
            If Len(strDataId) < 1 Then
                Call ErrBox(LoadResString(gLanguageOffset + 78), 2)
                ErrorBeep
                Exit Do           'Loop Again
            End If
                
            If Len(strEntry) > 11 Then
                If strDataId = "JS" Then   'data identifier for job and sequence from JBA
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
                    Call ErrBox(LoadResString(gLanguageOffset + 61), 2)
                    ErrorBeep
                    Exit Do           'Loop Again
            End If
'            If gJobMaster.OrderStat <> " " Then
 '                   Call ErrBox("Order Status is complete", 2)
  '                  ErrorBeep
   '                 Exit Do
    '        End If
            
            If Len(strSeqNo) = 0 Then
                    Call ErrBox(LoadResString(gLanguageOffset + 6), 2)
                    ErrorBeep
                    Exit Do
            ElseIf Not ValidNumber(strSeqNo) Then
                    Call ErrBox(LoadResString(gLanguageOffset + 111), 2)
                    ErrorBeep
                    Exit Do           'Loop Again
            End If
            
            gLaborIn(iEntry).SeqNo = CInt(strSeqNo)
            If gLaborIn(iEntry).SeqNo > 900 Or gLaborIn(iEntry).SeqNo = 0 Then
                    Call ErrBox(LoadResString(gLanguageOffset + 72), 2)
                    ErrorBeep
                    Exit Do
            End If
            asmbcount = 1
           If Not GetJobOper(gLaborIn(iEntry).OrderNo, gLaborIn(iEntry).SeqNo) Then
                    Call ErrBox(LoadResString(gLanguageOffset + 87), 2)
                    ErrorBeep
                    Exit Do
           ElseIf gJobOpers.WCID = "-----" Then
                    Call ErrBox(LoadResString(gLanguageOffset + 38), 2)
                    ErrorBeep
                    Exit Do
           ElseIf Not GetWorkCenter(gJobOpers.WCID) Then
                    Call ErrBox(LoadResString(gLanguageOffset + 81), 2)
                    ErrorBeep
                    Exit Do
           ElseIf gJobOpers.OperStatus = "C" Then
                     Call ErrBox(LoadResString(gLanguageOffset + 80), 2)
                     ErrorBeep
                     Exit Do
           End If
                
           asmbcount = 1
           Do While asmbcount <= 7
               If gAssyOperCode(asmbcount) = "999" Then
                     ErrBox LoadResString(gLanguageOffset + 74), 2
                    asmbcount = 999
                    ErrorBeep
                    Exit Do
              End If
              If Left(gJobOpers.Oper, 3) = gAssyOperCode(asmbcount) Then
                    Exit Do
              Else
                    asmbcount = asmbcount + 1
              End If
           Loop
           If asmbcount = 999 Then
                Exit Do
           End If
           gLaborIn(iEntry).WCID = gJobOpers.WCID
           gLaborIn(iEntry).Oper = gJobOpers.Oper
           gLaborIn(iEntry).EOrS = gJobOpers.EorSStd
           gLaborIn(iEntry).SetStd = gJobOpers.SetStd
           gLaborIn(iEntry).MchStd = gJobOpers.MchStd
           inpAssyJobSeq = True     'Done - Assembly order
        Loop
    Loop
    lblErrMsg.Caption = ""
End Function


Public Sub ClearForm()
'==============================================================================
'Procedure Name   : frmAssyCmpl.ClearForm
'Description      : Clears the label fields that display data
'Parameters       : None
'Called From      : AssyCmplProcess (after an entry is done)
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
        lblFixStd.Caption = ""
        lblMachStd.Caption = ""
        lblMchStdType.Caption = ""
        lblEorS.Caption = ""
        lblProdQty = ""
        lblScrapQty.Caption = ""
 
End Sub

Public Sub AssyCmplProcess()
'==============================================================================
'Procedure Name   : frmAssyCmpl.AssyCmplProcess
'Description      : Routine that processes processes assembly labor for Hyson
'                            products. This transaction generates a single direct labor
'                            record for each sequence in an assembly order. Hours are
'                            at standard, all operations are flagged as complete, and the
'                            order master is also called complete.
'                   Boolean values are globally set by division and type of
'                   Labor entry, to determine which options are used for the
'                   labor process.
'Parameters       : None
'Called From      : frmAssycmpl.Form_Activate
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
        
        If Not inpAssyJobSeq(I) Then Exit Sub
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
        lblActQty.Caption = gJobOpers.ActQty
        
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
       
       lblCCID.Caption = gLaborIn(I).CCID
       
       gLaborIn(I).RecType = "D"       'Mach hrs rec type?
       
'      If div = stroms Then
'           If gLaborIn(i).ActQty <> gJobMaster.OrderQty Then
 '                 If Not inpQtyAns() Then Exit Sub
  '          End If
   '   Else
           If Not inpAssyQty(I) Then Exit Sub
           lblProdQty.Caption = gLaborIn(I).ActQty
    '  End If

        
        lScrapTot = 0
        If gScrapColl Then
           If Not getScrapCounts(I) Then Exit Sub
                For j = 1 To 8
                    lScrapTot = lScrapTot _
                        + gLaborIn(I).Scrap(j).Qty
                Next
        End If
        lblScrapQty.Caption = CStr(lScrapTot)
        
'        If Not DirLaborReEnter(i) Then Exit Sub
        
        If I <> 1 Then
            If gbNumbMach Then
                gLaborIn(I).MachRec = "Y"
                gLaborIn(I).NumbMachs = gLaborIn(1).NumbMachs
                gLaborIn(I).incenRec = gLaborIn(1).incenRec
            Else
'            If gbAnotherRec Then
                gLaborIn(I).MachRec = "Y"
                gLaborIn(I).NumbMachs = 1
                gLaborIn(I).incenRec = 0
            End If
        End If
        
        I = I + 1
        LoopCnt = I     'keep global variable in sync
        
      '  If gbAnotherRec Then
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
      ' End If
    
    Me.ClearForm
    Wend
    
    'LoopCnt = gLaborIn(1).NumbMachs
    LoopCnt = LoopCnt - 1
    
    Rem Reset NumbMachs to 1 once looping is done
'    If gbAnotherRec Then
        
    gLaborIn(1).NumbMachs = 1

'    End If
  If Not gDone Then
    Call Build_Assembly_Recs
    
    If gbEOShift Then
        If empClkOutOK Then
            gLaborIn(1).ErrorLst = ""
            Call ClockOut
            Call Upd_Empl_Attd
            gDone = True
        End If
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
        Call AssyCmplProcess
        mintLoaded = 0
        Unload Me
    End If

End Sub



Private Sub Form_Load()
lblFormTitle.Caption = LoadResString(gLanguageOffset + 200)
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
lblQuantityCaption.Caption = LoadResString(gLanguageOffset + 170)
lblScrapQtyCaption.Caption = LoadResString(gLanguageOffset + 171)
lblActQtyCaption.Caption = LoadResString(gLanguageOffset + 276)

End Sub




