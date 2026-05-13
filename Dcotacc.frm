VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DCOvertime 
   Caption         =   "DC Corry Overtime Acceptance"
   ClientHeight    =   7590
   ClientLeft      =   3960
   ClientTop       =   2595
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7590
   ScaleWidth      =   10425
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1455
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2566
      _Version        =   393216
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   8
      Top             =   5760
      Width           =   45
   End
   Begin VB.Label Label19 
      Caption         =   "Enter 0 to decline, the Offered hours to fully accept, or a number between 0 and the Offered hours to partially accept"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   4920
      Width           =   6135
   End
   Begin VB.Label lblCurrentDate 
      AutoSize        =   -1  'True
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   3000
      Width           =   60
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblEmplName 
      AutoSize        =   -1  'True
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   4920
      TabIndex        =   4
      Top             =   2640
      Width           =   60
   End
   Begin VB.Label lblEmplNo 
      Caption         =   "."
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
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   3825
      TabIndex        =   2
      Top             =   2640
      Width           =   750
   End
   Begin VB.Label lblEmplCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Employee Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   2205
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Overtime Acceptance "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   480
      Left            =   1680
      TabIndex        =   0
      Top             =   1695
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   600
      Picture         =   "Dcotacc.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   870
   End
End
Attribute VB_Name = "DCOvertime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintLoaded As Integer
Public Sub OTAcceptanceProcess()
    
    Dim datestr As String
    Dim strRow As String
    Dim strOK As String
    Dim strEntryEmpNo As String
    Dim bGotOne As Boolean
    Dim iWid As Integer
    Dim cCurTime As Currency
    Dim iLaborRecs As Integer
    Dim I As Integer
    Dim RecCount As Integer
    Dim strHours As String
    Dim OriginalHrs As Currency
    Dim NewHours As Currency
    Dim OTOffereddate As Date
    Dim EntryOK As Boolean
    EntryOK = False
    grid1.Visible = False
    lblEmplNo.Caption = gEmployee.EmplNo
    lblEmplName.Caption = gEmployee.EmplName
    lblCurrentDate.Caption = Format(gEmployee.Actdate, "MM/DD/YY")
    
    RecCount = GetOTOfferingData()
    If RecCount = 0 Then
         ErrBox "No Overtime offered for the next 7 days", 4, 2
         Exit Sub
    End If
     
    Rem Column Titles and Column Width setup
    
    iWid = grid1.Width / 60    'twips per char as % of width
                                ' less apx scroll bar width
    grid1.Width = 1300
    grid1.Height = 1300
    grid1.ColWidth(0) = 1000
    grid1.Cols = RecCount + 1
    grid1.Rows = 4
    grid1.Col = 0
    grid1.Font = "Arial"
    grid1.FontSize = 10
    grid1.FontBold = True
    grid1.Row = 1
    grid1.Text = "Offered"
    grid1.Row = 2
    grid1.Text = "Accepted"
    grid1.RowHeight(0) = 350
    grid1.RowHeight(1) = 350
    grid1.RowHeight(2) = 350
    grid1.RowHeight(3) = 1
    
    For I = 1 To RecCount
      grid1.ColWidth(I) = 900
      grid1.Width = grid1.Width + grid1.ColWidth(I)
      grid1.ColAlignment(I) = 2 'CENTERED
      grid1.Col = I
      grid1.Row = 0
      grid1.Text = Format(gOTOffering(I).Actdate, "MM/DD/YY")
      grid1.Row = 1
       grid1.Text = Format(gOTOffering(I).OfferedHrs, gOneDecPos)
      grid1.Row = 2
      If gOTOffering(I).AcceptedHrs = -1# Then
            grid1.Text = " "
      Else
            grid1.Text = Format(gOTOffering(I).AcceptedHrs, gOneDecPos)
      End If
      grid1.Row = 3
      grid1.Text = Format(gOTOffering(I).AcceptedHrs, gOneDecPos)
    Next I
  grid1.Visible = True
  DoEvents
    
   
   gDone = False
   For I = 1 To RecCount
    If gOTOffering(I).AcceptedHrs = -1 Then 'minus 1 indicates nothing accepted or refused for the date
     gDone = False
     EntryOK = True
     Do Until gDone
        Set gEntry = New clsEntry
        gEntry.Title = "Enter Hrs Accepted For " & Format(gOTOffering(I).Actdate, "MM/DD/YY")
        gEntry.EntryType = 0    'Any Displayable Characters
        strHours = UCase(basGetEntry())
        If gEntry.Escaped Then
            Set gEntry = Nothing
            gDone = True
            Exit Sub
        End If
        Set gEntry = Nothing
        If Len(Trim(strHours)) = 0 Then  'user chose not to accept/reject leave blank
            strHours = "    "
            grid1.Row = 2
            grid1.Col = I
            If Len(Trim(grid1.Text)) = 0 Then
                  grid1.Text = strHours
           End If
           gDone = True
       ElseIf Len(strHours) > 4 Then
            ErrBox LoadResString(gLanguageOffset + 92), 2
            ErrorBeep
       ElseIf Not CheckHrs(strHours) Then
                ErrBox LoadResString(gLanguageOffset + 111), 2
                ErrorBeep
       ElseIf CCur(strHours) > gOTOffering(I).OfferedHrs Then
                ErrBox "Accepted Hours greater than Offered Hours", 3
                ErrorBeep
       Else
            grid1.Row = 2
            grid1.Col = I
            grid1.Text = Format(CCur(strHours), "##.0")
            gOTOffering(I).AcceptedHrs = CCur(strHours)
            gDone = True
        End If
     Loop
    End If
   Next I
   
   If Not EntryOK Then
     gDone = False
     ErrBox "All OT for these dates has been accepted or refused", 2, 3
     Do Until gDone
        Set gEntry = New clsEntry
        gEntry.Title = "Press E or ESC to return to menu"
        gEntry.Add "E", LoadResString(gLanguageOffset + 144)
        strOK = basGetEntry()
        If gEntry.Escaped Then
            Set gEntry = Nothing
            gDone = True
        End If
        Set gEntry = Nothing
        Select Case strOK
        Case "E"
            gDone = True
        Case Else
            gDone = False
        End Select
      Loop
   End If
   
'DoEvents
'save data to database
For I = 1 To RecCount
   grid1.Col = I
   grid1.Row = 0
   OTOffereddate = CDate(grid1.Text)
   grid1.Row = 2
   If grid1.Text <> "    " And Len(Trim(grid1.Text)) > 0 Then
       NewHours = CCur(Trim(grid1.Text))
       grid1.Row = 3
       If Len(Trim(grid1.Text)) > 0 Then
            OriginalHrs = CCur(Trim(grid1.Text))
       End If
       If NewHours <> OriginalHrs And OriginalHrs = -1 Then
           Call UpdateAcceptedOT(gEmployee.EmplNo, OTOffereddate, NewHours)
       End If
    End If
  Next I
  
End Sub



Private Sub Form_Activate()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmLaborChk.Form_Activate
'Description      : When form is activated for the first time after the
'                   form is loaded, the initialization process is called.
'Parameters       : None
'Called From      : Activate Event for form
'Procedures Used  : LaborChkProcess
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    'If gDone set, then unload forms all the way back to frmDCMain
    If gDone Then
        Unload Me
    End If
   
    If mintLoaded <> 1 Then
        mintLoaded = 1
        Call OTAcceptanceProcess
        mintLoaded = 0
        Unload Me
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmLaborChk.Form_KeyDown
'Description      : Process the Esc keystroke event.  Scrolling events will be
'                   handled by the grid control automatically.  The escape key
'                   will be set to unload form.  KeyPreview property is set
'                   on the form, so that the form captures keystrokes first.
'Parameters       : KeyCode and Shift indicator
'Called From      : KeyDown Event
'Procedures Used  : None
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    If KeyCode = vbKeyEscape Then
        Unload Me
        mintLoaded = 0
        Exit Sub
    End If

End Sub


