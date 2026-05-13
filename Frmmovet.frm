VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMoveTag 
   Caption         =   "Data Collection"
   ClientHeight    =   8370
   ClientLeft      =   3795
   ClientTop       =   1455
   ClientWidth     =   8175
   ControlBox      =   0   'False
   Icon            =   "Frmmovet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8370
   ScaleWidth      =   8175
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid grdMovetags 
      Height          =   5175
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   9128
      _Version        =   393216
      Rows            =   21
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "  Order No.      Copies     LRMC/Heat No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   90
   End
   Begin VB.Label Label1 
      Caption         =   "Print Move Tags"
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
      Height          =   555
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   495
      Picture         =   "Frmmovet.frx":0442
      Stretch         =   -1  'True
      Top             =   135
      Width           =   780
   End
   Begin VB.Label Label2 
      Caption         =   $"Frmmovet.frx":11C31
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   5415
   End
End
Attribute VB_Name = "frmMoveTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub EditMoveTag()
'==========================================================
' Name : John Canny
' Date : 09/14/98
' Comment : Will prompt the user with  a dialog box to enter _
    The WorkOrder # and Qauantity of move tags to print. Will _
    also Chack the DB to make sure it is a valid WorkOrder #
'==========================================================

Dim myQuantity  As New frmWorkOrderPrompt
Dim blnRc       As Boolean
    
    grdMovetags.Col = 0
    myQuantity.strWorkOrder = grdMovetags.Text
    grdMovetags.Col = 1
    If IsNumeric(grdMovetags.Text) Then
        myQuantity.intQuantity = CInt(grdMovetags.Text)
    Else
        myQuantity.intQuantity = 0
    End If
    grdMovetags.Col = 2
'    myQuantity.strLRMC = grdMoveTags.Text
 '   grdMoveTags.Col = 3
    myQuantity.strHeatNo = grdMovetags.Text
    myQuantity.Left = Int((Screen.Width - myQuantity.Width) / 2)
    myQuantity.Top = Int((Screen.Height - myQuantity.Height) / 2)
    
    myQuantity.Show 1
    
    While myQuantity.Visible = True
        DoEvents    'Wait for the User to enter the Info
    Wend
    
    'Check if Valid WorkOrder #
    If Trim$(myQuantity.strWorkOrder) <> "" Then
        blnRc = CheckWorkOrder(myQuantity.strWorkOrder)
        If blnRc = False Then
            Call ErrBox(LoadResString(gLanguageOffset + 105) & " - " & myQuantity.strWorkOrder, 2)
            myQuantity.strWorkOrder = ""
        Else
            
            With grdMovetags
                If .Row < 0 Then
                .Row = 0
                End If
                '.Row = 1
                .Col = 0
                .Text = UCase(myQuantity.strWorkOrder)
                .Col = 1
                .Text = Str$(myQuantity.intQuantity)
                .Col = 2
  '              .Text = UCase(myQuantity.strLRMC)
   '             .Col = 3
                .Text = UCase(myQuantity.strHeatNo)
            End With
        End If
    End If
    Unload myQuantity
    Set myQuantity = Nothing

End Sub

Private Sub NewReport()

Dim nCount As Integer          'Loop Counter
Dim x As Integer
Dim I As Long
Dim j As Long
Dim copies As Integer
Dim MyApp As New CRAXDRT.Application
Dim MyRpt As New CRAXDRT.Report
Dim rptname As String

rptname = App.Path & "\movetag.rpt"
Set MyRpt = MyApp.OpenReport(rptname, 1)
'Screen.MousePointer = vbHourglass
MyRpt.DiscardSavedData

With grdMovetags
        nCount = 0
        While nCount < .Rows
            'Move through all the Rows of the Grid
            .Row = nCount
            .Col = 0
            If Trim$(.Text) <> "" Then
                'We have a WorkOrder
'                Call myWorkOrder.GetWorkOrder(.Text)
             .Col = 1
'             copies = .Text / 2  'Divide by two, because there is two to a page
              copies = .Text / 1  'Divide by two, because there is two to a page
              MyRpt.Database.Tables(1).SetLogOnInfo "BRISSQL2", "Datacoll", SQL_LOGIN_ID, SQL_LOGIN_PSWD
              MyRpt.ParameterFields(1).ClearCurrentValueAndRange
              .Col = 0
              MyRpt.ParameterFields(1).AddCurrentValue .Text
              MyRpt.ParameterFields(2).ClearCurrentValueAndRange
              MyRpt.ParameterFields(2).AddCurrentValue "   "
              MyRpt.ParameterFields(3).ClearCurrentValueAndRange
              .Col = 2
              MyRpt.ParameterFields(3).AddCurrentValue .Text
              MyRpt.PrintOut False, copies
              DoEvents
            End If
            nCount = nCount + 1
        Wend
        
        .Row = 0
        
        .Col = 0
    End With
    
Set MyRpt = Nothing
Set MyApp = Nothing
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Const iPage = 9     'The numnber of items to skip through a list box _
                    When the user presses the PageUp or PageDown
Dim iAt As Integer
    'gSelIndx = frmJobPrint.lstJobSeq.ListIndex
    
Select Case KeyCode
    Case vbKeyEscape
        Unload Me
        Exit Sub
        
    Case vbKeyReturn, vbKeyF1
        'Prompt the User to Enter a WorkOrder # and Amount
    Case vbKeyF6
        'Print
         Call NewReport
         'Unload Me
        'Exit Sub
        
    Case vbKeyDown
        'KeyCode = 0
        
    Case vbKeyUp
        'KeyCode = 0
        
    Case vbKeyPageDown
        'KeyCode = 0
        
    Case vbKeyPageUp
        'KeyCode = 0
        
    'Case Else
    '    'Don't let anything else happen to field
    '    KeyCode = 0
End Select




End Sub


Private Sub Form_Load()

    With grdMovetags
        .Row = 0
        .Col = 0
        .ColAlignment(.Col) = 1
        .ColWidth(.Col) = 1200               '(grdMoveTags.Width / 4) - 40
        .Col = 1
        .ColAlignment(.Col) = 1
        .ColWidth(.Col) = 800                 ' (grdMoveTags.Width / 4) - 40
        .Col = 2
        .ColAlignment(.Col) = 1
        .ColWidth(.Col) = 1700                '(grdMoveTags.Width / 4) - 40
        grdMovetags.Width = 3800
        Label3.Width = 3800
        
        'Put us in a good starting Position
        .Row = 0
        .Col = 0
        grdMoveTags_SelChange

    End With
    
End Sub


Private Sub grdMoveTags_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyEscape
        Unload Me
        Exit Sub
        
    Case vbKeyReturn, vbKeyF1
        'Do Nothing
        Call EditMoveTag
    Case vbKeyF6
        'Select the current item
        
    Case vbKeyDown
        'KeyCode = 0
        
    Case vbKeyUp
        'KeyCode = 0
        
    Case vbKeyPageDown
        'KeyCode = 0
        
    Case vbKeyPageUp
        'KeyCode = 0
        
    'Case Else
    '    'Don't let anything else happen to field
    '    KeyCode = 0
End Select

End Sub

Private Sub grdMoveTags_SelChange()
    'grdMovetags.SelStartCol = 0
   ' grdMovetags.SelEndCol = 2

End Sub


