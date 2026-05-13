VERSION 5.00
Begin VB.Form frmLookup 
   Caption         =   "Data Collection Code Selection Form"
   ClientHeight    =   4350
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   5310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4350
   ScaleWidth      =   5310
   Begin VB.ListBox lstCodes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   240
      Picture         =   "Frmlooku.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   870
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      Caption         =   " "
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
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   45
   End
   Begin VB.Label lblHeader 
      Caption         =   "Code Description"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Use Up, Down arrows, or Page Up, Page Down to view.  Press F1 or Enter to Select."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Data Collection Scrap Code Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    frmLookup.lstCodes.ListIndex = 0
    gSelIndx = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmLookup.Form_KeyDown
'Description      : Passes selected Scrap code to global variable gSelDesc.
'                   Routine that handles keystrokes in the scrap code lookup
'                   selection form.  This form was developed before the code
'                   lookup combo box was added to the clsEntry class.  This may
'                   be reworked if desired.
'Parameters       : KeyCode and Shift indicator from KeyDown Event
'Called From      : KeyDown Event
'Procedures Used  : None
'Forms Loaded     : None
'Return Parameters: gSelCode with code, gSelDesc with Description, and
'                   gSelIndx with the index value of entry in table/collection
'Tables Updates   : None
'==============================================================================
    Dim iAt As Integer
    Dim iPage As Integer
    iPage = 9
    gSelIndx = frmLookup.lstCodes.ListIndex
    
    Select Case KeyCode
    Case vbKeyEscape
        gSelIndx = -1
        gSelDesc = ""
        gSelCode = ""
        Unload Me
        Exit Sub
    Case vbKeyReturn, vbKeyF1
        'Select the current item
        gSelDesc = lstCodes.List(lstCodes.ListIndex)
        iAt = InStr(gSelDesc, " ")
        If iAt > 1 Then
            gSelCode = Left(gSelDesc, InStr(gSelDesc, " ") - 1)
        Else
            gSelCode = "    "
        End If
        Unload Me
        Exit Sub
    Case vbKeyDown
        If gSelIndx < frmLookup.lstCodes.ListCount - 2 Then
            frmLookup.lstCodes.ListIndex = gSelIndx + 1
            gSelIndx = gSelIndx + 1
        Else
            frmLookup.lstCodes.ListIndex = frmLookup.lstCodes.ListCount - 2
            gSelIndx = frmLookup.lstCodes.ListCount - 2
        End If
        KeyCode = 0
    Case vbKeyUp
        If gSelIndx > 0 Then
            frmLookup.lstCodes.ListIndex = gSelIndx - 1
            gSelIndx = gSelIndx - 1
        Else
            frmLookup.lstCodes.ListIndex = 0
            gSelIndx = 0
        End If
        KeyCode = 0
    Case vbKeyPageDown
        If gSelIndx < frmLookup.lstCodes.ListCount - iPage - 1 Then
            frmLookup.lstCodes.ListIndex = gSelIndx + iPage
            gSelIndx = gSelIndx + iPage
        Else
            frmLookup.lstCodes.ListIndex = frmLookup.lstCodes.ListCount - 2
            gSelIndx = frmLookup.lstCodes.ListCount - 2
        End If
        KeyCode = 0
    Case vbKeyPageUp
        If gSelIndx > iPage Then
            frmLookup.lstCodes.ListIndex = gSelIndx - iPage
           gSelIndx = gSelIndx - iPage
        Else
            frmLookup.lstCodes.ListIndex = 0
            gSelIndx = 0
        End If
        KeyCode = 0
    'Case Else
    '    'Don't let anything else happen to field
    '    KeyCode = 0
    End Select
End Sub



