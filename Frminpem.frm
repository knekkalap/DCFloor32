VERSION 5.00
Begin VB.Form frmInpEmplNo 
   Caption         =   "Data Collection - Input Employee No."
   ClientHeight    =   3555
   ClientLeft      =   3120
   ClientTop       =   2670
   ClientWidth     =   8730
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
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
   ScaleHeight     =   3555
   ScaleWidth      =   8730
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   90
      Width           =   7935
   End
   Begin VB.Label LblErrMsg 
      AutoSize        =   -1  'True
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
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   2565
      Width           =   60
   End
End
Attribute VB_Name = "frmInpEmplNo"
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
    If frmInpEmplNo.Tag = 7 Then
        If Valid_Empl_badge(strEntryEmpNo) Then
                  Call GetSysDateTime     'save current time
                  DCOvertime.Show 1
                  gDone = True
                  Unload Me
                  Exit Sub
        Else
           Exit Sub
        End If
    End If
    If Valid_Empl_badge(strEntryEmpNo) Then 'must be a valid employee number
'        If div = stroms Then
'            gEmployee.AttdStatus = 1 'always clocked in
'            Call GetSysDateTime     'save current time
'            gEmployee.AttdStartDate = gSysDateTime
'        End If
        Select Case gEmployee.AttdStatus
            Case 2, 5, 8
                ErrBox LoadResString(gLanguageOffset + 76), 4
            Case 3
                ErrBox LoadResString(gLanguageOffset + 2), 4
            Case Else
                Call GetSysDateTime     'save current time
                If Elapsed_hrs(gEmployee.AttdStartDate, gSysDateTime) >= 24# And _
                   frmInpEmplNo.Tag <> 2 Then
                    ErrBox LoadResString(gLanguageOffset + 5), 4
                Else
                    Select Case frmInpEmplNo.Tag
                       Case 1
                          If div = milw Then
                            If GetMsgsForShopFloor("L") = 1 Then
                               frmShopMessage.Show 1
                               Unload frmShopMessage
                            End If
                          End If
'                          gEntry.ESCDisabled = False 'reset to false

                          frmLabor.Show 1
                          gDone = True
                          Unload Me
                          Exit Sub
                       Case 2
                          frmLaborChk.Show 1
                          gDone = True
                          Unload Me
                          Exit Sub
                       Case 3
                          frmNotepad.Show 1
                          gDone = True
                          Unload Me
                          Exit Sub
                       Case 4
                          frmDCStdsChgs.Show 1
                          gDone = True
                          Unload Me
                          Exit Sub
                       Case 5
                          ' These boolean flags must be set here since this
                          ' transaction is not initiated from the labor menu
                          gbNumbMach = False
                          gbNoLoads = False
                          gbInpSetOrMach = False   'Elapsed time calculation
                         ' gbAnotherRec = True
                          gbMachRecY = True
                          frmAssyCmpl.Show 1
                          gDone = True
                          Unload Me
                          Exit Sub
                       Case 6 'Scrap from main menu .. for Monterrey demo
                          gbNumbMach = False
                          gbNoLoads = False
                          gbInpSetOrMach = False   'Elapsed time calculation
                         ' gbAnotherRec = True
                          gbMachRecY = True
                          frmScrap.Show 1
                          gDone = True
                          Unload Me
                          Exit Sub
                       Case 8
                          gDone = True
                          Unload Me
                          frmDCMassEntry.Show vbModal
                          Exit Sub
                    End Select
                End If   'Elapsed_hrs(gEmployee.AttdStartDate, gSysDateTime) >= 24#
        End Select
    End If  'Valid_Empl_badge(strEntryEmpNo)
    
End Sub

Private Sub Form_Activate()
    Dim strEntrySub As String
    Dim strEntryEmpNo As String

    If gDone Then
        Unload Me
    End If
 '  Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
 '  MyForm.Height = Me.Height ' Remember the current size
 '  MyForm.Width = Me.Width

    Select Case Me.Tag
       Case 1
           Me.Caption = "DC Labor Entry - Input Employee No."
       Case 2
           Me.Caption = "DC Review Labor and Attd - Input Employee No."
       Case 3
           Me.Caption = "DC Notepad Entry - Input Employee No."
       Case 4
           Me.Caption = "DC Standards Changes  - Input Employee No."
       Case 5
           Me.Caption = "DC Assy Complete - Input Employee No."
       Case 6
           Me.Caption = "DC Scrap Entry - Input Employee No."
       Case 7
           Me.Caption = "DC OT Acceptance - Input Employee No."
       Case 8
           Me.Caption = "Upload Labor Spreadsheet - Input Employee No."
       Case Else
           Me.Caption = "Data Collection - Input Employee No."
    End Select
    
           
    
    If mintLoaded <> 1 Then ' if I am not loaded, load me up
        mintLoaded = 1  ' indicate that I am now loaded
        gDone = False   ' preset to not abort entry
        Call basEntry   ' show me the menu
        mintLoaded = 0  ' allow the menu to come up again
        Unload Me       ' unload me to refresh the form next time
    End If
End Sub


