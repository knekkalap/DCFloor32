VERSION 5.00
Begin VB.Form frmSplashLoading 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1200
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10005
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   FillColor       =   &H80000004&
   FontTransparent =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   10005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label LoadingMessage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading... Please Wait."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      TabIndex        =   0
      Top             =   10
      Width           =   10000
   End
   Begin VB.Label ProgressBar 
      BackColor       =   &H8000000D&
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   590
      Width           =   10000
   End
End
Attribute VB_Name = "frmSplashLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ProgressBar.Width = 0
End Sub
