VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmShopMessage 
   Caption         =   "Message for shop floor"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox txtShopFloorMsg 
      Height          =   2655
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4683
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmShopMessage.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Data Collection Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1440
      Width           =   5070
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   360
      Picture         =   "frmShopMessage.frx":0084
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   870
   End
   Begin VB.Label lblEmplName 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   75
   End
   Begin VB.Label Label2 
      Caption         =   "Employee:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "frmShopMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintLoaded As Integer
Public Sub ShopFloorMsgAckProcess()
Dim strAnswer As String

txtShopFloorMsg.Text = gMessage.message
    
lblEmplName = gEmployee.EmplName
gDone = False

Set gEntry = New clsEntry
'disable the ESC key so the user has to respond with one of the menu options
gEntry.ESCDisabled = True 'the only place in the app. where this is other than false
   
Select Case gMessage.msg_type
  Case 1 ' to an individual employee
      gEntry.Title = "Message Response:"
      gEntry.Add "A", "Accept"
      gEntry.Add "D", "Do not accept"
  Case Else ' type 2 and 3 for home dept / all employees respectively
      gEntry.Title = "Message Response:"
      gEntry.Add "A", "Acknowledge"
End Select
strAnswer = UCase(basGetEntry())
Set gEntry = Nothing
gMessage.msg_answer = strAnswer
Call InsEmplCommAction
    
End Sub

Private Sub Form_Activate()
    If gDone Then
        Unload Me
    End If
   
    If mintLoaded <> 1 Then
        mintLoaded = 1
        Call ShopFloorMsgAckProcess
        mintLoaded = 0
        Unload Me
    End If

End Sub

