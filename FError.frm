VERSION 5.00
Begin VB.Form FError 
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   435
      Left            =   2100
      TabIndex        =   0
      Top             =   2835
      Width           =   1905
   End
   Begin VB.TextBox txtErrorMsg 
      Height          =   2430
      Left            =   210
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   210
      Width           =   5685
   End
End
Attribute VB_Name = "FError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'========================================================
'Name: Shari M. Fogel
'Date: 03/01/1999
'Comments:
'   Error Form
'
'Properties:
'   ErrorModuleName     = Error Module Name
'   ErrorFunctionName   = Error Function Name
'   ErrorMessage        = Error Message
'========================================================
Private mstrErrorModuleName         As String
Private mstrErrorFunctionName       As String
Private mstrErrorMessage            As String

Public Property Get ErrorModuleName() As String
    ErrorModuleName = mstrErrorModuleName
End Property

Public Property Let ErrorModuleName(ByVal strErrorModuleName As String)
    mstrErrorModuleName = strErrorModuleName
End Property

Public Property Get ErrorFunctionName() As String
    ErrorFunctionName = mstrErrorFunctionName
End Property

Public Property Let ErrorFunctionName(ByVal strErrorFunctionName As String)
    mstrErrorFunctionName = strErrorFunctionName
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = mstrErrorMessage
End Property

Public Property Let ErrorMessage(ByVal strErrorMessage As String)
    mstrErrorMessage = strErrorMessage
End Property

Private Sub cmdOK_Click()
    ' unload form
    Unload Me
End Sub

Private Sub Form_Load()
    
    ' center form
  '  Call CenterForm(Me)
    
    With Me
        .Caption = "Error in " & ErrorModuleName & "." & ErrorFunctionName
        .txtErrorMsg = ErrorMessage
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FError = Nothing
End Sub

