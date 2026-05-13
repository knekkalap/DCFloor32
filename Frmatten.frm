VERSION 5.00
Begin VB.Form frmAttendance 
   Caption         =   "Data Collection Attendance"
   ClientHeight    =   7665
   ClientLeft      =   2235
   ClientTop       =   2565
   ClientWidth     =   11610
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7665
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrAttendance 
      Left            =   270
      Top             =   1800
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   960
      Picture         =   "Frmatten.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   180
      TabIndex        =   5
      Top             =   3120
      Width           =   75
   End
   Begin VB.Label lblAttdType 
      AutoSize        =   -1  'True
      Caption         =   " CLOCK IN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   840
      Left            =   2160
      TabIndex        =   4
      Top             =   1800
      Width           =   3675
   End
   Begin VB.Label lblEmplName 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   2760
      Width           =   4680
   End
   Begin VB.Label lblNameCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   2760
      Width           =   990
   End
   Begin VB.Label lblEmplNo 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblEmplNoCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Empl No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1605
   End
End
Attribute VB_Name = "frmAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintLoaded As Integer
Public Sub basEntry()
    Dim strEntrySub As String
    Dim strEntryEmpNo As String
    
    Set gEntry = New clsEntry
    gEntry.Title = LoadResString(gLanguageOffset + 127)
    gEntry.EntryType = 1    'Numeric only
    strEntryEmpNo = basGetEntry()
    If gEntry.Escaped Then
        gDone = True
        Set gEntry = Nothing
        Unload Me
        Exit Sub
    End If
    Set gEntry = Nothing

    strEntryEmpNo = Left$(strEntryEmpNo, 6) 'Length of Employee Number
    
    If Valid_Empl_badge(strEntryEmpNo) Then
       lblEmplNo = gEmployee.EmplNo
       lblEmplName = gEmployee.EmplName
       Call GetSysDateTime
       gSysTimeSecs = ConvTimeToSecs(gSysTimeStrp)
        
        ' write attendance transaction to error log
       If div = milw Then
              Call WriteAttdLog(gUserInfo.UserName, strEntryEmpNo, Attd_sel) '
       End If
       
'       If div = milw Then
'        If GetMsgsForShopFloor() = 1 Then
'             frmShopMessage.Show 1
'             Unload frmShopMessage
'        End If
'       End If
       
       
       If Clock_Err() = True Then
                ErrBox LoadResString(gLanguageOffset + 93), 4
      ElseIf gEmployee.ClassWrkr = 4 And div = stroms Then
                ErrBox LoadResString(gLanguageOffset + 292), 4
       Else
            Select Case Attd_sel
             Case 1
               Call Att_Clock_in
             Case 2
               Call Att_Clock_out
             Case 3
               Call Lunch_out_proc
             Case 4
               Call Lunch_In_Proc
             Case 5
               Call Susp_Process
             Case 6
               Call Resume_Process
             Case 7
               Call CI_start
             Case 8
               Call CI_End
            End Select
  
            If Len(Attd_Disp) > 1 Then
               ErrBox Attd_Disp, 3
               Attd_Disp = ""
            End If
            
            If div = milw Then
'             If Attd_sel = 1 Then
'                  gEmployee.Actdate = AttdActDate
'             End If
             If GetMsgsForShopFloor("A") = 1 Then
                  frmShopMessage.Show 1
                  Unload frmShopMessage
             End If
            End If
       
       End If
   Else
       Call GetSysDateTime
        ' write error attendance transaction to error log
       If div = milw Then
              Call WriteAttdLog(gUserInfo.UserName, "ERROR - " & strEntryEmpNo, Attd_sel)
       End If
   End If
End Sub

Private Sub Form_Activate()
    Dim strEntrySub As String
    Dim strEntryEmpNo As String
   
   Select Case Attd_sel
   Case 1
     frmAttendance.lblAttdType.Caption = LoadResString(gLanguageOffset + 128)
   Case 2
     frmAttendance.lblAttdType.Caption = LoadResString(gLanguageOffset + 129)
   Case 3
     frmAttendance.lblAttdType.Caption = LoadResString(gLanguageOffset + 130)
   Case 4
     frmAttendance.lblAttdType.Caption = LoadResString(gLanguageOffset + 131)
   Case 5
     frmAttendance.lblAttdType.Caption = LoadResString(gLanguageOffset + 132)
   Case 6
     frmAttendance.lblAttdType.Caption = LoadResString(gLanguageOffset + 133)
   Case 7
     frmAttendance.lblAttdType.Caption = LoadResString(gLanguageOffset + 134)
   Case 8
     frmAttendance.lblAttdType.Caption = LoadResString(gLanguageOffset + 135)
  End Select
'frmAttendance.lblPrompt.Caption = "Scan or Key Employee No."
'frmAttendance.txtPrompt.SetFocus

 If gDone Then
     Unload Me
 End If
    
    If mintLoaded <> 1 Then ' if I am not loaded, load me up
        mintLoaded = 1  ' indicate that I am now loaded
        gDone = False   ' preset to not abort entry
        Call basEntry   ' show me the menu
        mintLoaded = 0  ' allow the menu to come up again
        Unload Me       ' unload me to refresh the form next time
    End If

End Sub

Private Sub Form_Load()
   'Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
   'MyForm.Height = Me.Height ' Remember the current size
   'MyForm.Width = Me.Width
    
    lblEmplNoCaption.Caption = LoadResString(gLanguageOffset + 136)
    lblNameCaption.Caption = LoadResString(gLanguageOffset + 137)

End Sub


