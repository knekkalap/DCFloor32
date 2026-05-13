VERSION 5.00
Begin VB.Form frmWorkOrderPrompt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Move Tag Print"
   ClientHeight    =   2730
   ClientLeft      =   5805
   ClientTop       =   2535
   ClientWidth     =   6060
   Icon            =   "Frmworko.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2730
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtHeatNo 
      Height          =   375
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "O&k"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtQuantity 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtWorkOrder 
      Height          =   375
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "LRMC / Heat #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantity "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Work Order # "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmWorkOrderPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strWorkOrder As String       'The WorK Order Number for the Move Tickets
Public intQuantity  As Integer      'The Number of Move Tickets to Print
Public strLRMC As String
Public strHeatNo As String



Private Sub cmdCancel_Click()
    Me.txtWorkOrder.Text = ""
    Me.txtQuantity.Text = ""
    Me.txtHeatNo.Text = ""
    strWorkOrder = ""
    strHeatNo = ""
    Me.Visible = False
End Sub


Private Sub cmdOK_Click()
   Dim textlen As Long
Dim x As Long
Dim strchar As String
    If Len(Trim(txtWorkOrder.Text)) < 1 Then
         MsgBox "Work order invalid", vbExclamation
         txtWorkOrder.SetFocus
         Exit Sub
    ElseIf Not GetJobMaster(txtWorkOrder.Text) Then
        MsgBox "Work order not found", vbExclamation
       txtWorkOrder.SetFocus
       Exit Sub
    ElseIf gJobMaster.OrderStat = "C" Then
        MsgBox "Order status is complete", vbExclamation
        txtWorkOrder.SetFocus
        Exit Sub
   Else
         strWorkOrder = txtWorkOrder.Text
    End If
    
'     If Len(Trim(txtLRMC.Text)) > 6 Then
'          MsgBox "LRMC Invalid", vbExclamation
 '         txtLRMC.SetFocus
  '        Exit Sub
   '  Else
    '     strLRMC = txtLRMC.Text
    ' End If
    
    If IsNumeric(Left$(txtQuantity, 2)) = True Then
        intQuantity = CInt(Left$(txtQuantity, 2))
        intQuantity = (intQuantity \ 1) * 1 'The quantity can be any number 1-50 new movetag format takes up full 8.5" by 11"
'        intQuantity = (intQuantity \ 2) * 2 'The Quantity has to be a multiple of two
        If intQuantity > 50 Then
            intQuantity = 50
        End If
   Else
        MsgBox LoadResString(gLanguageOffset + 111), vbExclamation
        txtQuantity.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtHeatNo.Text)) > 15 Then
          MsgBox "Heat number too long", vbExclamation
          txtHeatNo.SetFocus
          Exit Sub
     Else
         strHeatNo = txtHeatNo.Text
    End If
     If Len(Trim(txtHeatNo.Text)) = 0 Then
            MsgBox "Must enter heat number", vbExclamation
             txtHeatNo.SetFocus
             Exit Sub
    End If
    
    textlen = Len(Trim(txtHeatNo.Text))
    For x = 1 To textlen
         strchar = Mid(txtHeatNo.Text, x, 1)
         If IsNumeric(strchar) Or isLetter(strchar) Then
        Else
            MsgBox "Invalid characters in heat number", vbExclamation
            txtHeatNo.SetFocus
            Exit Sub
       End If
   Next x
    
    Me.Visible = False

End Sub


Private Sub Form_Activate()

    txtWorkOrder = strWorkOrder
    txtQuantity = intQuantity
'    txtLRMC = strLRMC
    txtHeatNo = strHeatNo
    'frmWorkOrderPrompt.Left = Int((Screen.Width - frmWorkOrderPrompt.Width) / 2)
    'frmWorkOrderPrompt.Top = Int((Screen.Height - frmWorkOrderPrompt.Height) / 2)
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


Select Case KeyCode
    Case vbKeyEscape
        Me.txtHeatNo.Text = ""
        Me.txtWorkOrder.Text = ""
        Me.txtQuantity.Text = ""
        strWorkOrder = ""
        strHeatNo = ""
        Me.Visible = False
        'Exit Sub
    Case Else
End Select

End Sub


Private Sub txtHeatNo_GotFocus()
    txtHeatNo.SelStart = 0
    txtHeatNo.SelLength = Len(txtHeatNo)

End Sub




Private Sub txtQuantity_GotFocus()

    txtQuantity.SelStart = 0
    txtQuantity.SelLength = Len(txtQuantity)

End Sub


Private Sub txtWorkOrder_GotFocus()

    txtWorkOrder.SelStart = 0
    txtWorkOrder.SelLength = Len(txtWorkOrder)
    
End Sub


