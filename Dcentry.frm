VERSION 5.00
Begin VB.Form frmDCEntry 
   BorderStyle     =   0  'None
   Caption         =   "Data Collection Entry Form"
   ClientHeight    =   1320
   ClientLeft      =   1020
   ClientTop       =   1875
   ClientWidth     =   11130
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1320
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   10860
      Begin VB.ComboBox cmbEntry 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Dcentry.frx":0000
         Left            =   5520
         List            =   "Dcentry.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.TextBox txtEntryInput 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5580
         TabIndex        =   0
         Top             =   135
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label lblErrMsg 
         AutoSize        =   -1  'True
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   75
      End
      Begin VB.Label lblLongDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5640
         TabIndex        =   17
         Top             =   630
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   10305
         TabIndex        =   16
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   9945
         TabIndex        =   15
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   9585
         TabIndex        =   14
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   9225
         TabIndex        =   13
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   8865
         TabIndex        =   12
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   8505
         TabIndex        =   11
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   8145
         TabIndex        =   10
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   7785
         TabIndex        =   9
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   7425
         TabIndex        =   8
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7065
         TabIndex        =   7
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6705
         TabIndex        =   6
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   6345
         TabIndex        =   5
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5985
         TabIndex        =   4
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblMenuSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5670
         TabIndex        =   3
         Top             =   195
         Width           =   300
      End
      Begin VB.Label lblEntry 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   225
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmDCEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim miInMeth As Integer

Const vbInMethNone = 0
Const vbInMethScan = 1
Const vbInMethInput = 2
Const vbInMethCombo = 3
Const vbInMethMenu = 4
Const CB_ShowDropDown = &H14F


Dim mstrEntry As String             'which Entry item selected
Dim miEntryItemCur As Integer 'Which Entry item is highlighted

Const vbColorSelected = &H80000005
Const vbColorNormal = &H8000000F



Private Sub cmbEntry_Click()
      lblErrMsg.Caption = "Press Enter key to Select"
   '   ErrBox "Press Enter key to select", 0
End Sub


Private Sub Form_Activate()
    GetServerTime
    Call basEntry
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDCEntry.KeyDown
'Description      : Keystroke capture for special keys
'Parameters       : None
'Called From      : Keystroke Events
'Procedures Used  : UnloadApp, basEntryMove
'Return Parameters: None
'Tables Updates   : None
'Dependencies     : gEntry global variable (object of type clsEntry)
'==============================================================================
    Dim i As Integer
    Dim CtrlDown As Integer
    Dim AltDown As Integer
    
    Rem Check for Ctrl & Alt & X key combination -> Shutdown App
    Rem only do this if the MainEsc property has been set to 1
    CtrlDown = (Shift And 2) > 0
    AltDown = (Shift And 4) > 0
    If gEntry.MainEsc And KeyCode = vbKeyX Then
        If AltDown And CtrlDown Then
            Call UnloadApp
        End If
    End If
    
'   If gEntry.MainEsc And KeyCode = vbKeyF12 Then
 '          frmDCMain.SetFocus
  ' End If
    
    
    
    If miInMeth = vbInMethMenu Then
        'Menu System
        Select Case KeyCode
        Case vbKeyEscape
            If Not gEntry.MainEsc And gEntry.ESCDisabled = False Then
                gEntry.Response = ""
                gEntry.Escaped = True
                Unload Me
                Exit Sub
            End If
        Case vbKeyReturn
            'Select the item Highlighted
            'mstrEntryResponse = lblMenuSel(miEntryItemCur)
            gEntry.Response = lblMenuSel(miEntryItemCur)
            Unload Me
            Exit Sub
        Case vbKeyLeft, vbKeyDown
            Call basEntryMove(miEntryItemCur - 1)
            KeyCode = 0
        Case vbKeyRight, vbKeyUp
            Call basEntryMove(miEntryItemCur + 1)
            KeyCode = 0
        Case Else
            'Don't let anything else happen to field
            'KeyCode = 0
        End Select
    ElseIf miInMeth = vbInMethCombo Then
        'Combo box System
        Select Case KeyCode
        Case vbKeyEscape
            gEntry.Response = ""
            gEntry.Escaped = True
            Unload Me
            Exit Sub
        Case vbKeyReturn
            gEntry.Response = cmbEntry.Text
            Unload Me
            Exit Sub
        End Select
    ElseIf miInMeth = vbInMethInput Then
        'Text Box System
        Select Case KeyCode
        Case vbKeyEscape
            gEntry.Response = ""
            gEntry.Escaped = True
            Unload Me
            Exit Sub
        Case vbKeyReturn
            gEntry.Response = txtEntryInput.Text
            Unload Me
            Exit Sub
        End Select
    ElseIf miInMeth = vbInMethScan Then
        'In the Process of Scanning
        Select Case KeyCode
        Case vbKeyEscape
            If Not gEntry.MainEsc Then
                gEntry.Response = ""
                gEntry.Escaped = True
                Unload Me
                Exit Sub
            End If
        Case vbKeyReturn
            gEntry.Response = mstrEntry
            Unload Me
            Exit Sub
        End Select
    Else
        'Boop
        Select Case KeyCode
        Case vbKeyEscape
            gEntry.Escaped = True
            Unload Me
        'Case Else
        '    KeyCode = 0
        End Select
    End If

End Sub





Public Function basEntry()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDCEntry.basEntry
'Description      : Main processing for this form.  This form may be a Text
'                   input form, a Menu input form, or it may be a lookup combo
'                   box.  The procedures for setting the type of form, is
'                   determined by getting properties from the global object
'                   gEntry, in the clsEntry class.
'Parameters       : None
'Called From      : DirLaborProcess (after an entry is done)
'Procedures Used  : None
'Return Parameters: None
'Tables Updates   : None
'Dependencies     : gEntry global variable (object of type clsEntry)
'                   clsEntry class for code interface to this form
'                   clsCodeItem class for menu or lookup item storage
'==============================================================================
    Dim x As Variant
    Dim z As Long
    Dim i As Integer
    Const conHwndTopmost = -1
    
    lblEntry.Caption = gEntry.Title
    
    GetServerTime
    If gEntry.Count > 0 Then
        'This is a menu, because we have 1 or more entries in the MenuItems Collection
        For i = 1 To gEntry.Count
            lblMenuSel(i - 1).Caption = gEntry.getCode(i)
        Next i
        miEntryItemCur = 0
        lblMenuSel(0).BackColor = vbColorSelected
        txtEntryInput.Visible = False
        lblLongDesc.Visible = True
        cmbEntry.Visible = False
        lblLongDesc.Caption = gEntry.getDesc(1)
        miInMeth = vbInMethMenu
     ElseIf gEntry.CountLookupItems > 0 Then
        'The number of Items in the LookUpItems Collection, using a combo box
        lblLongDesc.Visible = False
        txtEntryInput.Visible = False
        cmbEntry.Visible = True
        miInMeth = vbInMethCombo
        For i = 1 To gEntry.CountLookupItems
            cmbEntry.AddItem gEntry.getLookupItem(i)
        Next i
        cmbEntry.SetFocus
        z = sendmessage(cmbEntry.hwnd, CB_ShowDropDown, 1, 0&)
'        SendKeys "%{DOWN}"
     Else
        'Free Form Text Box
        lblLongDesc.Visible = False
        txtEntryInput.Visible = True
        cmbEntry.Visible = False
        miInMeth = vbInMethInput
    End If
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDCEntry.KeyPress
'Description      : Keystroke capture for printable characters.  Used to allow
'                   menu selection by typing the appropriate letter.  Also used
'                   for dropping non numerics for numeric field.  Places response
'                   into gEntry.Response
'Parameters       : None
'Called From      : Keystroke Events
'Procedures Used  : None
'Return Parameters: None
'Tables Updates   : None
'Dependencies     : gEntry global variable (object of type clsEntry)
'==============================================================================
' Only process printable characters in this section
' See Form_KeyDown event for special characters
    Dim i As Integer
    'Do something for both input and field? or just input and fix KeyCode?
    If miInMeth = vbInMethMenu Then
        'Snag Scanned Entries from menu processing
        For i = 1 To gEntry.CountScanItems
            If gEntry.getScanItem(i) = UCase(Chr(KeyAscii)) Then
                miInMeth = vbInMethScan
                If gEntry.EntryType = gEntryNumeric Then
                    mstrEntry = ""
                Else
                    mstrEntry = UCase(Chr(KeyAscii))
                End If
                Exit Sub
            End If
        Next i
        Select Case KeyAscii
        Case 32 To 126   'printable characters
            For i = 0 To 13
                If lblMenuSel(i) = UCase(Chr(KeyAscii)) Then
                    gEntry.Response = lblMenuSel(i)
                    Unload Me
                    Exit Sub
                End If
            Next i
        End Select
    ElseIf miInMeth = vbInMethCombo Then
        If gEntry.EntryType = gEntryNumeric Then
            Select Case KeyAscii
            Case 48 To 57, 45, 46   'Numerics, - and .
                'OK
            Case vbKeyBack, vbKeyDelete, vbKeyLeft, vbKeyRight
            Case Else
                'beep and then ignore key
                Beep
                KeyAscii = 0
            End Select
        End If
    ElseIf miInMeth = vbInMethInput Then
        If gEntry.EntryType = gEntryNumeric Then
            Select Case KeyAscii
            Case 48 To 57, 45, 46   'Numerics, - and .
                'OK
            Case vbKeyBack, vbKeyDelete, vbKeyLeft, vbKeyRight
                'OK
            Case Else
                'beep and then ignore key
                Beep
                KeyAscii = 0
            End Select
        End If
    ElseIf miInMeth = vbInMethScan Then
    ' handle scanning during menu operation
        If gEntry.EntryType = gEntryNumeric Then
            Select Case KeyAscii
            Case 48 To 57, 45, 46   'Numerics, - and .
                'OK
            Case Else
                'beep and then ignore key
                Beep
                KeyAscii = 0
            End Select
        End If
        Select Case KeyAscii
        Case 32 To 126      'Printable Characters
            mstrEntry = mstrEntry & UCase(Chr(KeyAscii))
        Case Else
            'beep and then ignore key
            Beep
            KeyAscii = 0
        End Select
    End If

End Sub









Public Sub basEntryMove(iAmt As Integer)
    If miInMeth <> vbInMethMenu Then
        Exit Sub
    End If
    
    If iAmt > -1 And iAmt < gEntry.Count Then
        If miEntryItemCur > -1 Then
            lblMenuSel(miEntryItemCur).BackColor _
                = vbColorNormal
        End If
        miEntryItemCur = iAmt
        lblMenuSel(iAmt).BackColor _
            = vbColorSelected
        lblLongDesc.Caption = gEntry.getDesc(iAmt + 1)
    End If

End Sub






Private Sub Form_Load()
    
''Win32 or Win 31 API Calls to Set the Window Position. Originally
'   used to try to resolve problem with Border Style 0 not getting
'   control back to app after alt-tab.  Must have control box set
'   to true to solve this.
'
'Attempt to force the dcentry form to top of Z order.
'
'#If Win32 Then
'    lrc = SetWindowPos(frmDCEntry.hWnd, frmDCEntry.hWnd, 0, 0, 0, 0, _
'        SWP_NOSIZE Or HWND_TOPMOST Or SWP_NOMOVE)
'#Else
'    Call SetWindowPos(frmDCEntry.hWnd, frmDCEntry.hWnd, 0, 0, 0, 0, _
'        SWP_NOSIZE Or HWND_TOPMOST Or SWP_NOMOVE)
'#End If 'WIN32
'
'Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
'MyForm.Height = Me.Height ' Remember the current size
'MyForm.Width = Me.Width

Me.Top = 385
Me.Left = 50

End Sub



