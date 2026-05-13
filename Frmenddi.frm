VERSION 5.00
Begin VB.Form frmEndDirect 
   Caption         =   "Data Collection - End Direct Labor"
   ClientHeight    =   8460
   ClientLeft      =   1350
   ClientTop       =   4260
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8460
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Line Line3 
      X1              =   90
      X2              =   11295
      Y1              =   3780
      Y2              =   3780
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
      Height          =   300
      Left            =   10890
      TabIndex        =   54
      Top             =   5700
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   3330
      Picture         =   "Frmenddi.frx":0000
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   870
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
      Height          =   300
      Left            =   2340
      TabIndex        =   53
      Top             =   4900
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
      Height          =   300
      Left            =   6975
      TabIndex        =   52
      Top             =   5700
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
      Height          =   300
      Left            =   10665
      TabIndex        =   51
      Top             =   5700
      Width           =   60
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
      Height          =   300
      Left            =   2340
      TabIndex        =   50
      Top             =   5700
      Width           =   60
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
      Height          =   300
      Left            =   7470
      TabIndex        =   49
      Top             =   5700
      Width           =   2100
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
      Height          =   300
      Left            =   5805
      TabIndex        =   48
      Top             =   5700
      Width           =   1080
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
      Height          =   300
      Left            =   720
      TabIndex        =   47
      Top             =   5700
      Width           =   1545
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
      Height          =   300
      Left            =   4995
      TabIndex        =   46
      Top             =   5300
      Width           =   60
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
      Height          =   300
      Left            =   6975
      TabIndex        =   45
      Top             =   5300
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
      Height          =   300
      Left            =   3240
      TabIndex        =   44
      Top             =   5300
      Width           =   1605
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
      Height          =   300
      Left            =   5760
      TabIndex        =   43
      Top             =   5300
      Width           =   1110
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
      Height          =   300
      Left            =   9675
      TabIndex        =   42
      Top             =   5300
      Width           =   60
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
      Height          =   300
      Left            =   6975
      TabIndex        =   41
      Top             =   4900
      Width           =   60
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
      Height          =   300
      Left            =   7425
      TabIndex        =   40
      Top             =   5300
      Width           =   2085
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
      Height          =   300
      Left            =   5490
      TabIndex        =   39
      Top             =   4900
      Width           =   1365
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
      Height          =   300
      Left            =   2340
      TabIndex        =   38
      Top             =   5300
      Width           =   60
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
      Height          =   300
      Left            =   2475
      TabIndex        =   37
      Top             =   4900
      Width           =   60
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
      Height          =   300
      Left            =   135
      TabIndex        =   36
      Top             =   5300
      Width           =   2100
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
      Height          =   300
      Left            =   675
      TabIndex        =   35
      Top             =   4900
      Width           =   1560
   End
   Begin VB.Line Line1 
      X1              =   135
      X2              =   11295
      Y1              =   4725
      Y2              =   4725
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
      Height          =   300
      Left            =   10440
      TabIndex        =   34
      Top             =   3900
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
      Height          =   300
      Left            =   8685
      TabIndex        =   33
      Top             =   3900
      Width           =   1635
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
      Height          =   300
      Left            =   5085
      TabIndex        =   32
      Top             =   4300
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
      Height          =   300
      Left            =   3330
      TabIndex        =   31
      Top             =   4300
      Width           =   1620
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
      Height          =   300
      Left            =   10440
      TabIndex        =   30
      Top             =   4300
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
      Height          =   300
      Left            =   9225
      TabIndex        =   29
      Top             =   4300
      Width           =   1125
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
      Height          =   300
      Left            =   2025
      TabIndex        =   28
      Top             =   4300
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
      Height          =   300
      Left            =   225
      TabIndex        =   27
      Top             =   4300
      Width           =   1650
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
      Height          =   300
      Left            =   5085
      TabIndex        =   26
      Top             =   3900
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
      Height          =   300
      Left            =   3375
      TabIndex        =   25
      Top             =   3900
      Width           =   1560
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
      Height          =   300
      Left            =   2025
      TabIndex        =   24
      Top             =   3900
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
      Height          =   300
      Left            =   240
      TabIndex        =   23
      Top             =   3900
      Width           =   1665
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
      Left            =   5580
      TabIndex        =   22
      Top             =   3420
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
      Left            =   4230
      TabIndex        =   21
      Top             =   3420
      Width           =   1230
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
      Left            =   2430
      TabIndex        =   20
      Top             =   3420
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
      Height          =   375
      Left            =   450
      TabIndex        =   19
      Top             =   3420
      Width           =   1800
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      Caption         =   "End Direct Labor"
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
      Left            =   4320
      TabIndex        =   18
      Top             =   2760
      Width           =   5505
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
      Left            =   585
      TabIndex        =   17
      Top             =   7740
      Width           =   45
   End
   Begin VB.Label lblSetupHrsCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Setup Hrs."
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
      Left            =   360
      TabIndex        =   16
      Top             =   6300
      Width           =   2250
   End
   Begin VB.Label lblAddonHrsCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Addon Hrs."
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
      Left            =   3915
      TabIndex        =   15
      Top             =   6300
      Width           =   1575
   End
   Begin VB.Label lblMachineHrsCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Machine Hrs."
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
      Left            =   6660
      TabIndex        =   14
      Top             =   6300
      Width           =   2475
   End
   Begin VB.Label lblSetupHrs 
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
      Left            =   2700
      TabIndex        =   13
      Top             =   6300
      Width           =   60
   End
   Begin VB.Label lblAddonHrs 
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
      TabIndex        =   12
      Top             =   6300
      Width           =   60
   End
   Begin VB.Label lblMachHrs 
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
      Left            =   9225
      TabIndex        =   11
      Top             =   6300
      Width           =   60
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
      Left            =   360
      TabIndex        =   10
      Top             =   6660
      Width           =   2235
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
      Left            =   2700
      TabIndex        =   9
      Top             =   6660
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
      Left            =   6660
      TabIndex        =   8
      Top             =   6660
      Width           =   2460
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
      Left            =   9225
      TabIndex        =   7
      Top             =   6660
      Width           =   60
   End
   Begin VB.Label lblNoLoadsCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Numb of Loads"
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
      Left            =   315
      TabIndex        =   6
      Top             =   7020
      Width           =   2265
   End
   Begin VB.Label lblCycleTimeCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cycle Time"
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
      Left            =   3915
      TabIndex        =   5
      Top             =   7020
      Width           =   1575
   End
   Begin VB.Label lblMethodCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Method"
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
      Left            =   6660
      TabIndex        =   4
      Top             =   7020
      Width           =   2475
   End
   Begin VB.Label lblNumbOfLoads 
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
      Left            =   2700
      TabIndex        =   3
      Top             =   7020
      Width           =   60
   End
   Begin VB.Label lblCycleTime 
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
      TabIndex        =   2
      Top             =   7020
      Width           =   60
   End
   Begin VB.Label lblBatchMethod 
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
      Left            =   9225
      TabIndex        =   1
      Top             =   7020
      Width           =   60
   End
   Begin VB.Line Line2 
      X1              =   180
      X2              =   11340
      Y1              =   6120
      Y2              =   6120
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
      Left            =   585
      TabIndex        =   0
      Top             =   7365
      Width           =   60
   End
End
Attribute VB_Name = "frmEndDirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mintLoaded As Integer

Const vbColorWhite = &H8000005


Public Function EndDirReEnter(I As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDirLabor.DirLaborReEnter(i)
'Description      : Routine that handles the re-enter an item menu for a single
'                   machine entry for direct labor for all divisions. Boolean
'                   values are globally set by division and type of
'                   Labor entry, to determine which options are used for the
'                   labor process.
'Parameters       : None
'Called From      : frmDirLabor.DirLaborProcess
'Procedures Used  : ErrorChk, inpJobAndSeq, DispValues, inpWCID, GetCostCenter,
'                   inpProdQty, getScrapCounts, inpSetHrs, inpMchHours, inpAddonHours,
'                   inpEndOfShift, CallCmpl, inpNoOfLoads, CalcBtErnHrs, CalcBatchHrs,
'                   inpCycleTime, getBatchIncn, ErrorBeep, ErrBox
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    Dim strEntrySub As String
    Dim strEntryEmpNo As String
    
    EndDirReEnter = False
    gDone = False
    
    Do Until gDone Or EndDirReEnter
        Set gEntry = New clsEntry
        gEntry.Title = LoadResString(gLanguageOffset + 192)
        gEntry.Add "A", LoadResString(gLanguageOffset + 193)
        gEntry.Add "E", LoadResString(gLanguageOffset + 194)
        gEntry.Add "J", LoadResString(gLanguageOffset + 195)
        gEntry.Add "W", LoadResString(gLanguageOffset + 196)
        If Not gbZeroQty Then
            gEntry.Add "Q", LoadResString(gLanguageOffset + 197)
        End If
        If gScrapColl And gJobMaster.OrderType <> "RS" Then
            gEntry.Add "S", LoadResString(gLanguageOffset + 198)
        End If
        If gbInpSetAndMach Or (gbInpSetOrMach _
                And gLaborIn(I).RecType = "F") _
                Or (gLaborIn(I).RecType = "F" And div = bristol) Then
            gEntry.Add "F", "change Fixed (setup) hours"
        End If
        If gbInpSetAndMach Or (gbInpSetOrMach _
                And gLaborIn(I).RecType = "D") _
                Or (gLaborIn(I).RecType = "D" And div = bristol) Then
            gEntry.Add "M", "change Machine hours"
        End If
        If gbAddonHrs Then
            gEntry.Add "B", "change addon hours"
        End If
        If gbEOShift Then
            gEntry.Add "C", "change Clock-out status"
        End If
        If gbCallCompl Then
            gEntry.Add "O", LoadResString(gLanguageOffset + 199)
        End If
        If gbNoLoads Then
            gEntry.Add "L", "change number of Loads"
        End If
        If gbNoLoads And gLaborIn(I).NoOfLoads > 0 Then
            gEntry.Add "T", "change cycle Time"
            gEntry.Add "H", "change batcH method"
        End If
        strEntrySub = basGetEntry()
        If gEntry.Escaped Then
            Set gEntry = Nothing
            gDone = True
            Exit Function
        End If
        Set gEntry = Nothing
        Select Case strEntrySub
        Case "A"
            Call ErrorChk(I)
            EndDirReEnter = True
        Case "E"
            gDone = True
            Exit Function
        Case "J"
            If Not inpEndJobSeq(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Call DispValues(I)
                        End If
        Case "W"
            If Not inpWCID(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Screen.ActiveForm.lblWCID.Caption = gLaborIn(I).WCID
                gLaborIn(I).CCID = gWorkCenter.CCID
                If Not GetCostCenter(gWorkCenter.CCID) Then
                    gCostCenter.PaidLunch = "N"
                End If
                Screen.ActiveForm.lblCCID.Caption = gLaborIn(I).CCID
            End If
        Case "Q"
            If Not inpProdQty(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Screen.ActiveForm.lblProdQty.Caption = gLaborIn(I).ActQty
            End If
        Case "S"
            If Not gScrapColl Then
                ErrorBeep
            Else
                If Not getScrapCounts(I) Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                Else
                    Dim iScrap As Long
                    Dim j As Integer
                    iScrap = 0
                    For j = 1 To 8
                        iScrap = iScrap + gLaborIn(I).Scrap(j).Qty
                    Next j
                    Screen.ActiveForm.lblScrapQty.Caption = CStr(iScrap)
                End If
            End If
        Case "F"
            If gLaborIn(I).NoOfLoads > 0 _
                    Or gLaborIn(I).CycleTime > 0 Then
                ErrBox LoadResString(gLanguageOffset + 8), 2
                ErrorBeep
            Else
                If gLaborIn(I).RecType <> "F" And _
                   Not gbInpSetAndMach Then
                    ErrorBeep
                Else
                    If Not inpSetHours(I) Then
                        gDone = False       'be sure to loop back
                        ErrorBeep
                    Else
                        Screen.ActiveForm.lblSetupHrs.Caption = gLaborIn(I).setupHours
                    End If
                End If
            End If
        Case "M"
            If gLaborIn(I).NoOfLoads > 0 _
                    Or gLaborIn(I).CycleTime > 0 Then
                ErrBox LoadResString(gLanguageOffset + 8), 2
                ErrorBeep
            Else
                If gLaborIn(I).RecType <> "D" And _
                   Not gbInpSetAndMach Then
                    ErrorBeep
                Else
                    If Not inpMchHours(I) Then
                        gDone = False       'be sure to loop back
                        ErrorBeep
                    Else
                        Screen.ActiveForm.lblMachHrs.Caption = gLaborIn(I).machHours
                    End If
                End If
            End If
        Case "B"
            If Not inpAddonHours(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Screen.ActiveForm.lblAddonHrs.Caption = gLaborIn(I).addonHours
            End If
        Case "C"
            If gEosFlag = 0 Then
                ErrBox LoadResString(gLanguageOffset + 86), 2
                ErrorBeep
            Else
                If Not inpEndOfShift() Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                End If
            End If
        Case "O"
            If Not CallCmpl(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Screen.ActiveForm.lblOperStatus.Caption = gLaborIn(I).OperStatus
            End If
        Case "L"
            If Not okToPrompt Then
                ErrorBeep
            Else
                If Not inpNoOfLoads(I) Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                Else
                    Screen.ActiveForm.lblNumbOfLoads.Caption = gLaborIn(I).NoOfLoads
                    Call CalcBtErnHrs(I)
                    Call CalcBatchHrs(I)
                End If
            End If
        Case "T"
            If Not okToPrompt Then
                ErrorBeep
            Else
                If gLaborIn(I).NoOfLoads = 0 Then
                    ErrorBeep
                    ErrBox LoadResString(gLanguageOffset + 79), 2
                Else
                    If Not inpCycleTime(I) Then
                        gDone = False       'be sure to loop back
                        ErrorBeep
                    Else
                        Screen.ActiveForm.lblCycleTime.Caption = gLaborIn(I).CycleTime
                        Call CalcBatchHrs(I)
                    End If
                End If
            End If
        Case "H"
            If Not okToPrompt Then
                ErrorBeep
            Else
                If gLaborIn(I).NoOfLoads = 0 Then
                    ErrorBeep
                    ErrBox LoadResString(gLanguageOffset + 79), 2
                Else
                    LdBatchIncn (gLaborIn(I).WCID)
                    If Not getBatchIncn(I) Then
                        gDone = False       'be sure to loop back
                        ErrorBeep
                    Else
                        Screen.ActiveForm.lblBatchMethod.Caption = gLaborIn(I).biMethod
                        Call CalcBtErnHrs(I)
                    End If
                End If
            End If
        End Select
    Loop

End Function

Public Sub EndDirLaborProcess()
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
        
        If Not inpEndJobSeq(I) Then Exit Sub
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
        lblMachHrs.Caption = Format(gLaborIn(I).machHours, gOneDecPos)
        If div <> burl Then
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
       End If
        
        If gLaborIn(I).SeqNo <> 0 _
                And gLaborIn(I).SeqNo <> 990 Then
            
            If Not inpWcAns(I) Then Exit Sub
        End If
        lblCCID.Caption = gLaborIn(I).CCID
        
        
        ' note  - this is forStromsholmen only
        If div = stroms Then
            If gLaborIn(I).SetStd > 0 And gLaborIn(I).MchStd > 0 Then
                If Not inpSetHours(I) Then Exit Sub
                lblSetupHrs.Caption = Format(gLaborIn(I).setupHours, gOneDecPos)
                gLaborIn(I).machHours = 0
           End If
       End If
        
        
        If Left(gLaborIn(I).Oper, 3) = "645" Then
            gLaborIn(I).RecType = "F"       'Setup hours rec type
            If gbInpSetOrMach Then
                If Not inpSetHours(I) Then Exit Sub
                lblSetupHrs.Caption = Format(gLaborIn(I).setupHours, gOneDecPos)
                gLaborIn(I).machHours = 0
            End If
        Else
            gLaborIn(I).RecType = "D"       'Mach hrs rec type?
            If gbInpSetOrMach Then
                If Not inpMchHours(I) Then Exit Sub
                lblMachHrs.Caption = Format(gLaborIn(I).machHours, gOneDecPos)
                gLaborIn(I).setupHours = 0
            End If
        End If
        
        If gbInpSetAndMach Then
          If Not gbRecTypeR Then
            If Not inpSetHours(I) Then Exit Sub
                lblSetupHrs.Caption = Format(gLaborIn(I).setupHours, gOneDecPos)
          Else:
            gLaborIn(I).setupHours = 0
          End If
            If Not inpMchHours(I) Then Exit Sub
            lblMachHrs.Caption = Format(gLaborIn(I).machHours, gOneDecPos)
            If gLaborIn(I).setupHours = 0 And gLaborIn(I).machHours = 0 Then
                ErrBox LoadResString(gLanguageOffset + 102), 2
            End If
        End If
       
        If gbAddonHrs Then
            If Not inpAddonHours(I) Then Exit Sub
            lblAddonHrs = Format(gLaborIn(I).addonHours, gOneDecPos)
        End If
        
        bZeroIt = False
        If gbZeroQty Then
            Select Case Left(gLaborIn(I).Oper, 3)
            Case 635, 645, 660, 665, 670, 770, 900, 920, 960
                bZeroIt = True
            End Select
        End If
        If bZeroIt Then
            gLaborIn(I).ActQty = 0
        Else
            If Not inpProdQty(I) Then Exit Sub
            lblProdQty = gLaborIn(I).ActQty
        End If
        
        lScrapTot = 0
        If gScrapColl Then
            If gJobMaster.OrderType <> "RS" Then
                If Not getScrapCounts(I) Then Exit Sub
                For j = 1 To 8
                    lScrapTot = lScrapTot _
                        + gLaborIn(I).Scrap(j).Qty
                Next
            End If
        End If
        lblScrapQty.Caption = CStr(lScrapTot)
        
        If gOKToComplete Then
            gLaborIn(I).OperStatus = "C"
'            If Not CallCmpl(i) Then Exit Sub
            lblOperStatus = gLaborIn(I).OperStatus
        Else
            Rem Don't ask Operation Complete Question
            gLaborIn(I).OperStatus = "P"
            lblOperStatus = gLaborIn(I).OperStatus
        End If
        
        If gbNoLoads And LdBatchIncn(gLaborIn(LoopCnt).WCID) <> 0 Then
            If Not inpNoOfLoads(LoopCnt) Then Exit Sub
            lblNumbOfLoads.Caption = gLaborIn(LoopCnt).NoOfLoads
            If gLaborIn(LoopCnt).NoOfLoads > 0 Then
                If Not inpCycleTime(LoopCnt) Then Exit Sub
                lblCycleTime.Caption = gLaborIn(LoopCnt).CycleTime
                Call CalcBatchHrs(LoopCnt)
                'MsgBox "Is Ld_batch_incn the same as LoadBatIncnAry ?"
                'Call LdBatchIncn(gLaborIn(LoopCnt).WCID)
                If Not getBatchIncn(LoopCnt) Then Exit Sub
                lblBatchMethod.Caption = gLaborIn(LoopCnt).biMethod
                Call CalcBtErnHrs(LoopCnt)
            End If
        Else
Rem Coded to not ask info if no Batch Incentive Codes for WCID
            gLaborIn(LoopCnt).NoOfLoads = 0
            lblNumbOfLoads.Caption = ""
            gLaborIn(LoopCnt).CycleTime = 0
            lblCycleTime.Caption = ""
            gLaborIn(LoopCnt).biMethod = 0
            lblBatchMethod.Caption = ""
            Call CalcBtErnHrs(LoopCnt)
        End If
        
'        If (Not gbInpSetOrMach) And (Not gbInpSetAndMach) Then
            Rem Calculate Hours
'            If i = 1 Then
'                CalcHours = CalcActHrs()
'            End If
'            If gLaborIn(i).RecType = "F" Then
'                gLaborIn(i).setupHours = CalcHours
'                gLaborIn(i).machHours = 0
'            Else
'                gLaborIn(i).machHours = CalcHours
'                gLaborIn(i).setupHours = 0
'            End If
'            gLaborIn(i).inHours = 0
'            lblSetupHrs.Caption = gLaborIn(i).setupHours
'            lblMachHrs.Caption = gLaborIn(i).machHours
'        End If
        
        
        If Not EndDirReEnter(I) Then Exit Sub
        
     ' currently, we will check elapsed time of the open labor; if it exceeds 23.9
    '  flag as an error, to prevent it's automatic upload...
   '   this will keep excessive hours away from JBA until someone has reviewed
   '   the records in question
   
'       If Elapsed_hrs(gOpenLabor.StartDateTime, gSysDateTime) > 23.9 Then
'          RecordError ("T")
'      End If
        
        
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
    
'    If gbInpSetAndMach Then
'        Rem First write F rec with only setup hours
'        gLaborIn(i).RecType = "F"
'        CalcHours = gLaborIn(i).machHours   'save it
'        gLaborIn(i).machHours = 0
'        Call Build_Labor_Recs
'        Rem Next - Prepare to write D rec with mach hours
'        gLaborIn(i).RecType = "D"
'        gLaborIn(i).machHours = CalcHours   'restore it
'        gLaborIn(i).setupHours = 0
'    End If
    Call Close_Direct_Labor
    
    If gbEOShift Then
        If empClkOutOK Then
            gLaborIn(1).ErrorLst = ""
            Call ClockOut
            Call Upd_Empl_Attd
            gDone = True
        End If
    End If
   
    Unload Me
    
End Sub

Public Function inpEndJobSeq(iEntry As Integer) As Boolean
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
    
    inpEndJobSeq = False        'to force loop back till done
    
    Do Until inpEndJobSeq Or gDone
        Do Until inpEndJobSeq Or gDone       'Extra loop for exit do stmts
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
                 lblUnplanned.Caption = LoadResString(gLanguageOffset + 240)
                 If Not getLabUnplnData(iEntry) Then
                    Exit Do           'Loop Again
                 End If
                 inpEndJobSeq = True     'Done - Unplanned
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
                ElseIf Not OpenDirectLabor(gEmployee.EmplNo, gLaborIn(iEntry).OrderNo, gLaborIn(iEntry).SeqNo, gEmployee.Actdate) Then
                        ErrBox LoadResString(gLanguageOffset + 106), 2
                        ErrorBeep
                        Exit Do
                ElseIf gOpenLabor.EmplNo <> gEmployee.EmplNo Then  'added this 10/16/2001 to prevent other employees from completing someone
                        ErrBox LoadResString(gLanguageOffset + 294) & CStr(gOpenLabor.EmplNo), 2    'else's work.
                        ErrorBeep
                        Exit Do
                End If
                gLaborIn(iEntry).UnplnReas = gOpenLabor.UnplnReas
                gLaborIn(iEntry).WCID = gOpenLabor.WCID
                gLaborIn(iEntry).Oper = gOpenLabor.OperCd
                gLaborIn(iEntry).EOrS = gJobOpers.EorSStd
                gLaborIn(iEntry).SetStd = gJobOpers.SetStd
                gLaborIn(iEntry).MchStd = gJobOpers.MchStd
                gLaborIn(iEntry).machHours = gOpenLabor.ElapsedHrs
                inpEndJobSeq = True     'Done - Planned
            End If
        Loop
    Loop
 lblErrMsg.Caption = ""
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
        lblSetupHrs.Caption = ""
        lblMachHrs.Caption = ""
        lblAddonHrs = ""
        lblProdQty = ""
        lblScrapQty.Caption = ""
        lblOperStatus = ""
        lblNumbOfLoads.Caption = ""
        lblCycleTime.Caption = ""
        lblBatchMethod.Caption = ""

End Sub

Public Function getLabUnplnData(iEntry As Integer) As Boolean
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
    getLabUnplnData = False
    Call initJobOpers
    
    Screen.ActiveForm.lblJobNo.Caption = gJobMaster.OrderNo
    Do Until gDone Or getLabUnplnData
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
        
        getLabUnplnData = True 'Exit loop
    Loop
End Function

Private Sub Form_Activate()
    lblEmplNo = gEmployee.EmplNo
    lblEmplName = gEmployee.EmplName
    
    If gDone Then
        Unload Me
    End If
    
    If mintLoaded <> 1 Then
        mintLoaded = 1
        Call EndDirLaborProcess
        mintLoaded = 0
        Unload Me
    End If

End Sub

Private Sub Form_Load()
lblFormTitle.Caption = LoadResString(gLanguageOffset + 205)
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
lblSetupHrsCaption.Caption = LoadResString(gLanguageOffset + 167)
lblMachineHrsCaption.Caption = LoadResString(gLanguageOffset + 168)
lblQuantityCaption.Caption = LoadResString(gLanguageOffset + 170)
lblScrapQtyCaption.Caption = LoadResString(gLanguageOffset + 171)
lblNoLoadsCaption.Caption = LoadResString(gLanguageOffset + 172)
lblCycleTimeCaption.Caption = LoadResString(gLanguageOffset + 173)
lblMethodCaption.Caption = LoadResString(gLanguageOffset + 174)
If div <> corry Then
   lblAddonHrsCaption.Visible = False
   lblNoLoadsCaption.Visible = False
   lblCycleTimeCaption.Visible = False
   lblMethodCaption.Visible = False
Else
   lblNoLoadsCaption.Caption = LoadResString(gLanguageOffset + 172)
   lblCycleTimeCaption.Caption = LoadResString(gLanguageOffset + 173)
   lblMethodCaption.Caption = LoadResString(gLanguageOffset + 174)
   lblAddonHrsCaption.Caption = LoadResString(gLanguageOffset + 169)
End If



End Sub


