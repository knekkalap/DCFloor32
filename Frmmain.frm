VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Data Collection"
   ClientHeight    =   4140
   ClientLeft      =   5130
   ClientTop       =   435
   ClientWidth     =   7680
   Icon            =   "Frmmain.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   7680
   WindowState     =   1  'Minimized
   Begin VB.PictureBox frmStart 
      Height          =   1005
      Left            =   135
      ScaleHeight     =   945
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   135
      Width           =   7215
      Begin VB.Frame frameStart 
         Caption         =   " "
         Height          =   690
         Left            =   45
         TabIndex        =   1
         Top             =   180
         Width           =   6765
         Begin VB.PictureBox VBSQL1 
            Height          =   285
            Left            =   315
            ScaleHeight     =   225
            ScaleWidth      =   990
            TabIndex        =   3
            Top             =   315
            Width           =   1050
         End
         Begin VB.Label Label1 
            Caption         =   "Loading Application...PLEASE WAIT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1575
            TabIndex        =   2
            Top             =   180
            Width           =   5055
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Dim strVer As String
Screen.MousePointer = MOUSE_HOUR_GLASS
Show
DoEvents

'Initialization is done in Startup/DBOpen
'This assumes one database open in this app
'strVer = SqlInit()
'If strVer = "" Then
'   MsgBox ("Unable to initialize SQLSRVR control")
'End If
Call LoadApp
'frameStart.Visible = False

Screen.MousePointer = MOUSE_NORMAL
frmDCMain.Show 1

End Sub
Private Sub Form_Unload(Cancel As Integer)
Call UnloadApp
End Sub


