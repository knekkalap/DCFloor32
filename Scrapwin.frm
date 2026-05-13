VERSION 5.00
Begin VB.Form ScrapWind 
   Caption         =   "Data Collection Scrap Code Data Entry"
   ClientHeight    =   8460
   ClientLeft      =   1890
   ClientTop       =   1665
   ClientWidth     =   11205
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8460
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbPounds 
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
      Index           =   7
      ItemData        =   "Scrapwin.frx":0000
      Left            =   2655
      List            =   "Scrapwin.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   6060
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbPounds 
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
      Index           =   6
      ItemData        =   "Scrapwin.frx":0014
      Left            =   2655
      List            =   "Scrapwin.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   5580
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbPounds 
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
      Index           =   5
      ItemData        =   "Scrapwin.frx":0028
      Left            =   2655
      List            =   "Scrapwin.frx":0032
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   5100
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbPounds 
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
      Index           =   4
      ItemData        =   "Scrapwin.frx":003C
      Left            =   2655
      List            =   "Scrapwin.frx":0046
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4620
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbPounds 
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
      Index           =   3
      ItemData        =   "Scrapwin.frx":0050
      Left            =   2655
      List            =   "Scrapwin.frx":005A
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4140
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbPounds 
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
      Index           =   2
      ItemData        =   "Scrapwin.frx":0064
      Left            =   2655
      List            =   "Scrapwin.frx":006E
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3660
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbPounds 
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
      Index           =   1
      ItemData        =   "Scrapwin.frx":0078
      Left            =   2655
      List            =   "Scrapwin.frx":0082
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3180
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbPounds 
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
      ItemData        =   "Scrapwin.frx":008C
      Left            =   2655
      List            =   "Scrapwin.frx":0096
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2685
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtScrapQty 
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
      Left            =   3285
      TabIndex        =   22
      Top             =   6060
      Width           =   1215
   End
   Begin VB.TextBox txtScrapQty 
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
      Left            =   3285
      TabIndex        =   19
      Top             =   5580
      Width           =   1215
   End
   Begin VB.TextBox txtScrapQty 
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
      Left            =   3285
      TabIndex        =   16
      Top             =   5100
      Width           =   1215
   End
   Begin VB.TextBox txtScrapQty 
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
      Left            =   3285
      TabIndex        =   13
      Top             =   4620
      Width           =   1215
   End
   Begin VB.TextBox txtScrapQty 
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
      Left            =   3285
      TabIndex        =   10
      Top             =   4140
      Width           =   1215
   End
   Begin VB.TextBox txtScrapQty 
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
      Left            =   3285
      TabIndex        =   7
      Top             =   3675
      Width           =   1215
   End
   Begin VB.TextBox txtScrapQty 
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
      Left            =   3285
      TabIndex        =   4
      Top             =   3180
      Width           =   1215
   End
   Begin VB.TextBox txtScrapQty 
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
      Left            =   3285
      TabIndex        =   1
      Top             =   2700
      Width           =   1215
   End
   Begin VB.TextBox txtScrapCd 
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
      Left            =   4575
      TabIndex        =   23
      Top             =   6060
      Width           =   735
   End
   Begin VB.TextBox txtScrapCd 
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
      Left            =   4575
      TabIndex        =   20
      Top             =   5580
      Width           =   735
   End
   Begin VB.TextBox txtScrapCd 
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
      Left            =   4575
      TabIndex        =   17
      Top             =   5100
      Width           =   735
   End
   Begin VB.TextBox txtScrapCd 
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
      Left            =   4575
      TabIndex        =   14
      Top             =   4620
      Width           =   735
   End
   Begin VB.TextBox txtScrapCd 
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
      Left            =   4575
      TabIndex        =   11
      Top             =   4140
      Width           =   735
   End
   Begin VB.TextBox txtScrapCd 
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
      Left            =   4575
      TabIndex        =   8
      Top             =   3660
      Width           =   735
   End
   Begin VB.TextBox txtScrapCd 
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
      Left            =   4575
      TabIndex        =   5
      Top             =   3180
      Width           =   735
   End
   Begin VB.TextBox txtScrapCd 
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
      Left            =   4560
      TabIndex        =   2
      Top             =   2685
      Width           =   735
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
      Height          =   375
      Left            =   120
      TabIndex        =   49
      Top             =   1800
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   1800
      Picture         =   "Scrapwin.frx":00A0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
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
      Left            =   120
      TabIndex        =   48
      Top             =   1440
      Width           =   6885
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
      Left            =   720
      TabIndex        =   47
      Top             =   6465
      Width           =   75
   End
   Begin VB.Label lblLP 
      Caption         =   "L/P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   2700
      TabIndex        =   46
      Top             =   2325
      Visible         =   0   'False
      Width           =   495
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
      Left            =   240
      TabIndex        =   45
      Top             =   6960
      Width           =   75
   End
   Begin VB.Label lblDescrCaption 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   330
      Left            =   6120
      TabIndex        =   44
      Top             =   2325
      Width           =   3255
   End
   Begin VB.Label lblScrapDesc 
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
      Left            =   5505
      TabIndex        =   43
      Top             =   6105
      Width           =   3405
   End
   Begin VB.Label lblScrapDesc 
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
      Left            =   5505
      TabIndex        =   42
      Top             =   5610
      Width           =   3405
   End
   Begin VB.Label lblScrapDesc 
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
      Left            =   5505
      TabIndex        =   41
      Top             =   5115
      Width           =   3405
   End
   Begin VB.Label lblScrapDesc 
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
      Left            =   5505
      TabIndex        =   40
      Top             =   4620
      Width           =   3405
   End
   Begin VB.Label lblScrapDesc 
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
      Left            =   5505
      TabIndex        =   39
      Top             =   4125
      Width           =   3405
   End
   Begin VB.Label lblScrapDesc 
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
      Left            =   5505
      TabIndex        =   38
      Top             =   3675
      Width           =   3405
   End
   Begin VB.Label lblScrapDesc 
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
      Left            =   5505
      TabIndex        =   37
      Top             =   3180
      Width           =   3405
   End
   Begin VB.Label lblScrapDesc 
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
      Left            =   5505
      TabIndex        =   36
      Top             =   2685
      Width           =   3405
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
      Left            =   120
      TabIndex        =   35
      Top             =   1080
      Width           =   8115
   End
   Begin VB.Label lblQuantityCaption 
      Alignment       =   2  'Center
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   3285
      TabIndex        =   34
      Top             =   2325
      Width           =   1290
   End
   Begin VB.Label lblCodeCaption 
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   4560
      TabIndex        =   33
      Top             =   2325
      Width           =   1455
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Scrap Entry 8"
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
      Left            =   120
      TabIndex        =   32
      Top             =   6060
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Scrap Entry 7"
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
      Left            =   120
      TabIndex        =   31
      Top             =   5565
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Scrap Entry 6"
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
      Left            =   120
      TabIndex        =   30
      Top             =   5115
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Scrap Entry 5"
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
      Left            =   120
      TabIndex        =   29
      Top             =   4665
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Scrap Entry 4"
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
      Left            =   120
      TabIndex        =   28
      Top             =   4170
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Scrap Entry 3"
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
      Left            =   120
      TabIndex        =   27
      Top             =   3675
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Scrap Entry 2"
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
      Left            =   120
      TabIndex        =   26
      Top             =   3180
      Width           =   2460
   End
   Begin VB.Label lblScrapEntryCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Scrap Entry 1"
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
      Left            =   120
      TabIndex        =   25
      Top             =   2685
      Width           =   2460
   End
   Begin VB.Label lblFormTitle 
      Caption         =   "Scrap  Entry"
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
      Left            =   3000
      TabIndex        =   24
      Top             =   210
      Width           =   6885
   End
End
Attribute VB_Name = "ScrapWind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim micolfocus As Integer

Private Sub cmbPounds_GotFocus(index As Integer)
    'Set indicator that we are on L/P field.
    '   This is used for the down and up navigations.
    micolfocus = 0  'on L/P column
    If cmbPounds(index).Text = "" And div = saline Then
        cmbPounds(index).Text = "L"
    ElseIf cmbPounds(index).Text = "" And div = milw Then
        cmbPounds(index).Text = "P"
    End If
    
    lblInfo.Caption = LoadResString(gLanguageOffset + 241)
End Sub

Private Sub cmbPounds_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Call CheckKeys(index, KeyCode, Shift)
End Sub


Private Sub cmbPounds_LostFocus(index As Integer)
    lblInfo = ""
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Code to have enter key tab to next field (and not to beep)
    If KeyCode = vbKeyReturn Then KeyCode = vbKeyTab
End Sub

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
    
    lblFormTitle.Caption = LoadResString(gLanguageOffset + 176)
    lblInstrLine1Caption.Caption = LoadResString(gLanguageOffset + 177)
    lblInstrLine2Caption.Caption = LoadResString(gLanguageOffset + 178)
    lblInstrLine3Caption.Caption = LoadResString(gLanguageOffset + 179)
    lblDescrCaption.Caption = LoadResString(gLanguageOffset + 181)
    lblCodeCaption.Caption = LoadResString(gLanguageOffset + 180)
    lblQuantityCaption.Caption = LoadResString(gLanguageOffset + 170)
    lblScrapEntryCaption(0).Caption = LoadResString(gLanguageOffset + 182)
    lblScrapEntryCaption(1).Caption = LoadResString(gLanguageOffset + 183)
    lblScrapEntryCaption(2).Caption = LoadResString(gLanguageOffset + 184)
    lblScrapEntryCaption(3).Caption = LoadResString(gLanguageOffset + 185)
    lblScrapEntryCaption(4).Caption = LoadResString(gLanguageOffset + 186)
    lblScrapEntryCaption(5).Caption = LoadResString(gLanguageOffset + 187)
    lblScrapEntryCaption(6).Caption = LoadResString(gLanguageOffset + 188)
    lblScrapEntryCaption(7).Caption = LoadResString(gLanguageOffset + 189)
      
         If gbScrapPounds Then
            lblLP.Visible = True
            For I = 0 To 7
                cmbPounds(I).Visible = True
            Next I
        End If
    
    Rem Update the fields before exiting
        For I = 0 To 7
            If gLaborIn(LoopCnt).Scrap(I + 1).Qty <> 0 Then
                txtScrapQty(I).Text = CStr(gLaborIn(LoopCnt).Scrap(I + 1).Qty)
                txtScrapCd(I).Text = gLaborIn(LoopCnt).Scrap(I + 1).Code
                If gbScrapPounds Then
                    cmbPounds(I).Text = gLaborIn(LoopCnt).Scrap(I + 1).ScrapType
                End If
            End If
        Next I
        
End Sub



Private Sub txtScrapCd_GotFocus(index As Integer)
   Dim strwork As String
    'Set indicator that we are on code field.
    '   This is used for the down and up navigations.
    micolfocus = 2  'on the Code column
    lblInfo.Caption = LoadResString(gLanguageOffset + 242)
    
'        Set gEntry = New clsEntry
'        gEntry.Title = "Key in or cursor to Scrap Code "
'        gEntry.EntryType = 0    'Any Displayable Characters
'        gEntry.LuTable = gcolScrapCodes
'        strwork = UCase(basGetEntry())
'        If gEntry.Escaped Then
'            Set gEntry = Nothing
'            'gDone = True       'dont abort everything when exiting code lookup
'            Exit Sub
'        End If
'        Set gEntry = Nothing
'        If Len(strwork) > 0 Then
'            gSelCode = Left(Trim(strwork), _
'                InStr(LTrim(strwork), " ") - 1)
'            gSelDesc = Mid(Trim(strwork), InStr(Trim(strwork), ""))
'            gSelIndx = 0        'Does this need to match entry in Collection?
'            txtScrapCd(Index).Text = gSelCode
'            lblScrapDesc(Index).Caption = gSelDesc
'            miColFocus = 1
'            SendKeys "{TAB}"
'        Else
'            gSelIndx = -1
'        End If

End Sub
'
Private Sub txtScrapCd_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Call CheckKeys(index, KeyCode, Shift)
End Sub


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
    
    Case vbKeyDown
        'Make sure we stay in control array bounds
        If index < 7 Then
            'If currently on Qty field, go down to Qty field below
            '   (else go down to Cd field below)
            Select Case micolfocus
            Case 0
                cmbPounds(index + 1).SetFocus
            Case 1
                txtScrapQty(index + 1).SetFocus
            Case 2
                txtScrapCd(index + 1).SetFocus
            End Select
        End If
        KeyCode = 0     'Prevent Value change in combo box
    
    Case vbKeyUp
        'Make sure we stay in control array bounds
        If index > 0 Then
            'If currently on Qty field, go up to Qty field above
            '   (same for L/P or Code fields)
            Select Case micolfocus
            Case 0
                cmbPounds(index - 1).SetFocus
            Case 1
                txtScrapQty(index - 1).SetFocus
            Case 2
                txtScrapCd(index - 1).SetFocus
            End Select
        End If
        KeyCode = 0     'Prevent Value change in combo box
    
    Case vbKeyRight
        'Only valid right action is from Qty to Cd field
        Select Case micolfocus
        Case 0
            txtScrapQty(index).SetFocus
        Case 1
            txtScrapCd(index).SetFocus
        End Select
        KeyCode = 0     'Prevent Value change in combo box
    
    Case vbKeyLeft
        'Only valid left action is from Cd field to Qty Field
        Select Case micolfocus
        Case 1
            If gbScrapPounds Then cmbPounds(index).SetFocus
        Case 2
            txtScrapQty(index).SetFocus
        End Select
        KeyCode = 0     'Prevent Value change in combo box
    
    Case vbKeyReturn
        SendKeys "{TAB}"    'Tab to next field
        KeyCode = 0
    
    Case vbKeyF1
        'Loop through Control array to make sure all codes and values
        '   are good here
        gSelIndx = 0
        For I = 0 To 7
            If txtScrapCd(I).Text = "  " Then txtScrapCd(I).Text = ""
            If txtScrapCd(I).Text = "" And txtScrapQty(I).Text = "" Then
                'Skip checks on this row - not used
            ElseIf Not AllDigits(txtScrapQty(I).Text) _
                    Or Len(Trim(txtScrapQty(I).Text)) > 7 Then
                ErrorBeep
                 Call ErrBox(LoadResString(gLanguageOffset + 60), 2)
                txtScrapQty(I).SetFocus
                KeyCode = 0
                Exit Sub
            ElseIf CLng(txtScrapQty(I).Text) <= 0 _
                And Not txtScrapCd(I).Text <> "" Then
                    ErrorBeep
                    Call ErrBox(LoadResString(gLanguageOffset + 67), 2)
                    txtScrapQty(I).SetFocus
                    KeyCode = 0
                    Exit Sub
            ElseIf CLng(txtScrapQty(I).Text) > 0 _
                And txtScrapCd(I).Text = "" Then
                    ErrorBeep
                    Call ErrBox(LoadResString(gLanguageOffset + 66), 2)
                    txtScrapCd(I).SetFocus
                    KeyCode = 0
                    Exit Sub
            ElseIf txtScrapQty(I).Text > 0 _
                And Not ValidScrapCd(txtScrapCd(I).Text) Then
                    ErrorBeep
                    Call ErrBox(LoadResString(gLanguageOffset + 59), 2)
                    txtScrapCd(I).SetFocus
                    KeyCode = 0
                    Exit Sub
            Else
                gSelIndx = gSelIndx + 1   'increment count of entries
            End If
        Next I
        
        Rem Update the fields before exiting
        For I = 1 To 8
            If txtScrapCd(I - 1).Text <> "" And txtScrapQty(I - 1).Text <> "" Then
                gLaborIn(LoopCnt).Scrap(I).Qty = CLng(txtScrapQty(I - 1).Text)
                gLaborIn(LoopCnt).Scrap(I).Code = txtScrapCd(I - 1).Text
                If gbScrapPounds Then
                    gLaborIn(LoopCnt).Scrap(I).ScrapType = cmbPounds(I - 1).Text
                End If
            End If
        Next I
        Unload Me
        Exit Sub
        
    Case vbKeyF6
        ' Removed frmLookup, and used frmDCEntry for Code Lookup Combo Box
        ''Lookup Scrap codes using frmLookup
        ''Clear the listbox and set it's caption
        'frmLookup.lstCodes.Clear
        'frmLookup.lblHeader.Caption = "Code Description"
        ''Load list box up with the code and description in gColScrapCodes
        'For i = 1 To gcolScrapCodes.Count
        '    frmLookup.lstCodes.AddItem gcolScrapCodes.getCodeDesc(i)
        'Next i
        ''Bring up lookup form Modally
        'frmLookup.Show 1
        
        
        Set gEntry = New clsEntry
        gEntry.Title = "Key in or cursor to Scrap Code "
        gEntry.EntryType = 0    'Any Displayable Characters
        gEntry.LuTable = gcolScrapCodes
        strwork = UCase(basGetEntry())
        If gEntry.Escaped Then
            Set gEntry = Nothing
            'gDone = True       'dont abort everything when exiting code lookup
            Exit Sub
        End If
        Set gEntry = Nothing
        If Len(strwork) > 0 Then
            gSelCode = Left(Trim(strwork), InStr(LTrim(strwork), " ") - 1)
            gSelDesc = Mid(Trim(strwork), InStr(Trim(strwork), ""))
            gSelIndx = 0        'Does this need to match entry in Collection?
        Else
            gSelIndx = -1
        End If
        
        
        'Out of modal form, the lookup has set gSelCode, gSelDesc.
        If gSelIndx <> -1 Then
            'If not escaped out of form, update the row on the form
            txtScrapCd(index).Text = gSelCode
            lblScrapDesc(index).Caption = gSelDesc
        '   Removed positioning code - leave at current field
        '    'Make sure that there is a qty, else set the focus on qty
        '    If txtScrapQty(Index) = "" Or txtScrapQty(Index) = "0" Then
        '        txtScrapQty(Index).SetFocus     'Enter current Qty
        '    'Else go on to the next row's qty field
        '    'ElseIf Index < 7 Then
        '    '    txtScrapQty(Index + 1).SetFocus 'Enter next Qty
        '    Else
        '        SendKeys "{TAB}"    'Tab to next field
        '        KeyCode = 0
        '    End If
        End If
    End Select

End Sub

Private Sub txtScrapCd_LostFocus(index As Integer)
    lblInfo = ""
    txtScrapCd(index).Text = UCase(txtScrapCd(index).Text)
    If txtScrapCd(index).Text = "" Then
    ElseIf Not ValidScrapCd(txtScrapCd(index).Text) Then
        ErrorBeep
        Call ErrBox(LoadResString(gLanguageOffset + 59), 2)
        ScrapWind.txtScrapCd(index).SetFocus
    Else
        lblScrapDesc(index).Caption = _
            gcolScrapCodes.getDesc(txtScrapCd(index).Text)
    End If
End Sub

Private Sub txtScrapQty_GotFocus(index As Integer)
    'Set indicator that we are on qty field.
    '   This is used for the down and up navigations.
    micolfocus = 1   'on Qty column
    If cmbPounds(index).Text = "L" Then
        lblInfo = LoadResString(gLanguageOffset + 243)
    Else
        lblInfo = LoadResString(gLanguageOffset + 244)
    End If
End Sub

Private Sub txtScrapQty_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Call CheckKeys(index, KeyCode, Shift)
End Sub

Private Sub txtScrapQty_LostFocus(index As Integer)
    lblInfo = ""
    If Not AllDigits(txtScrapQty(index).Text) _
            And txtScrapQty(index).Text <> "" Then
        ErrorBeep
         Call ErrBox(LoadResString(gLanguageOffset + 57), 2)
        txtScrapQty(index).SetFocus
    End If
End Sub


