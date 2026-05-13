VERSION 5.00
Begin VB.Form frmInpPrompt 
   Caption         =   "Data Collection"
   ClientHeight    =   2625
   ClientLeft      =   870
   ClientTop       =   1545
   ClientWidth     =   8715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   8715
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtInpPrompt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7110
      TabIndex        =   1
      Text            =   " "
      Top             =   810
      Width           =   510
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   540
      TabIndex        =   2
      Top             =   2070
      Width           =   60
   End
   Begin VB.Label lblClockOut 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Are You Clocking Out? (Press Y or N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   450
      TabIndex        =   0
      Top             =   810
      Width           =   6555
   End
End
Attribute VB_Name = "frmInpPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If KeyAscii = vbKeyY Then
    inpClockOut = True
    'frmInpPrompt.Hide
    Unload Me
    Exit Sub
  End If
  If KeyAscii = vbKeyN Then
    inpClockOut = False
    'frmInpPrompt.Hide
    Unload Me
    Exit Sub
   End If
   ErrBox LoadResString(gLanguageOffset + 45), 2
   frmInpPrompt.txtInpPrompt.Text = ""
   frmInpPrompt.txtInpPrompt.SetFocus
   
End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub txtInpPrompt_GotFocus()
  txtInpPrompt.Text = ""
End Sub

