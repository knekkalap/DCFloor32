VERSION 5.00
Begin VB.Form frmErrMsg 
   Caption         =   "Error"
   ClientHeight    =   3825
   ClientLeft      =   2025
   ClientTop       =   4680
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   10110
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Press Enter to continue......"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   3735
   End
   Begin VB.Label lblSysError 
      Alignment       =   2  'Center
      Caption         =   "A database error has occurred"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1530
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   8625
   End
End
Attribute VB_Name = "frmErrMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdContinue_Click()
 Unload Me
 End Sub
