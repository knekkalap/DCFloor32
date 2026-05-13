VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmEntryValidation 
   Caption         =   "Validate Entry Record"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7735.426
   ScaleMode       =   0  'User
   ScaleWidth      =   10782.86
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   485
      Left            =   5588
      TabIndex        =   23
      Top             =   6840
      Width           =   1589
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   485
      Left            =   3548
      TabIndex        =   22
      Top             =   6840
      Width           =   1589
   End
   Begin VB.Frame grpEmployeeInfo 
      Caption         =   "Employee Info"
      Height          =   2775
      Left            =   248
      TabIndex        =   11
      Top             =   1320
      Width           =   5295
      Begin VB.TextBox txtEmplNo 
         Height          =   375
         Left            =   1920
         TabIndex        =   0
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtEmplName 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblEmpNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Employee Number"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label lblEmpName 
         Alignment       =   1  'Right Justify
         Caption         =   "Employee Name"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1020
         Width           =   1335
      End
   End
   Begin VB.Frame grpLaborInfo 
      Caption         =   "Labor Details"
      Height          =   2295
      Left            =   248
      TabIndex        =   13
      Top             =   4200
      Width           =   10215
      Begin VB.ComboBox cmbScrapCd 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7200
         TabIndex        =   9
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtScrapQty 
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtActQty 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtMachHrs 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   953
         Width           =   2655
      End
      Begin VB.ComboBox cmbUnplnReas 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IntegralHeight  =   0   'False
         Left            =   7200
         TabIndex        =   8
         Top             =   487
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtActDate 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   228524033
         CurrentDate     =   43509
         MinDate         =   36161
      End
      Begin VB.Label lblScrapCd 
         Alignment       =   1  'Right Justify
         Caption         =   "Scrap Code"
         Height          =   255
         Left            =   5520
         TabIndex        =   28
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label lblScrapQty 
         Alignment       =   1  'Right Justify
         Caption         =   "Scrap Quantity"
         Height          =   255
         Left            =   5400
         TabIndex        =   27
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblActQty 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantity"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1500
         Width           =   1335
      End
      Begin VB.Label lblMachHrs 
         Alignment       =   1  'Right Justify
         Caption         =   "Machine Hrs"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label lblUnplnReas 
         Alignment       =   1  'Right Justify
         Caption         =   "Unplanned Reason"
         Height          =   255
         Left            =   5520
         TabIndex        =   24
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label lblActDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Activity Date"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   540
         Width           =   1335
      End
   End
   Begin VB.Frame grpOrderInfo 
      Caption         =   "Order Info"
      Height          =   2775
      Left            =   5768
      TabIndex        =   12
      Top             =   1320
      Width           =   4695
      Begin VB.TextBox txtMchStd 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtSeqNumb 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtOrderQty 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtOrderNo 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblMachStd 
         Alignment       =   1  'Right Justify
         Caption         =   "Machine Std"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Label lblSeqNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Seq. Num"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label lblOrderQty 
         Alignment       =   1  'Right Justify
         Caption         =   "Order Qty"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label lblJobNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Job Number"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   540
         Width           =   1095
      End
   End
   Begin VB.Label lblErrorMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   29
      Top             =   240
      Width           =   10215
   End
End
Attribute VB_Name = "frmEntryValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EntryIndex As Integer
Dim Entry As New MassEntry
Dim UpdatedEntry As New MassEntry
Dim EmplNoChanged As Boolean
Dim OrderNoChanged As Boolean
Dim SeqNumbChanged As Boolean
Dim ActDateChanged As Boolean
Dim UnplnReasChanged As Boolean
Dim ScrapQtyChanged As Boolean
Dim ScrapCdChanged As Boolean

Private Sub ResetChangedValues()
    EmplNoChanged = False
    OrderNoChanged = False
    SeqNumbChanged = False
    ActDateChanged = False
    UnplnReasChanged = False
    ScrapCdChanged = False
    ScrapQtyChanged = False
End Sub

Public Sub SetEntry(SelectedIndex As Integer, SelectedRecord As MassEntry)
    EntryIndex = SelectedIndex
    Set Entry = SelectedRecord                  'Set values for the initial record selected
    Call CopyEntry(SelectedRecord, UpdatedEntry)    'Initialize values for the record with user updates
End Sub

Public Sub Localization()
    lblActDate.Caption = LoadResString(gLanguageOffset + 114)
    lblActQty.Caption = LoadResString(gLanguageOffset + 170)
    lblEmpNo.Caption = LoadResString(gLanguageOffset + 136)
    lblEmpName.Caption = LoadResString(gLanguageOffset + 137)
    lblJobNo.Caption = LoadResString(gLanguageOffset + 150)
    lblMachHrs.Caption = LoadResString(gLanguageOffset + 168)
    lblMachStd.Caption = LoadResString(gLanguageOffset + 164)
    lblOrderQty.Caption = LoadResString(gLanguageOffset + 152)
    lblScrapCd.Caption = LoadResString(gLanguageOffset + 180)
    lblScrapQty.Caption = LoadResString(gLanguageOffset + 171)
    lblSeqNum.Caption = LoadResString(gLanguageOffset + 156)
    lblUnplnReas.Caption = LoadResString(gLanguageOffset + 161)
    grpEmployeeInfo.Caption = LoadResString(gLanguageOffset + 307)
    grpOrderInfo.Caption = LoadResString(gLanguageOffset + 308)
    grpLaborInfo.Caption = LoadResString(gLanguageOffset + 309)
    
End Sub

Private Sub BindUnplnReas()
    Dim i As Integer
    Dim UplnCount As Integer
    UplnCount = gcolUnplnCodes.Count
    
    If UplnCount > 0 Then
        For i = 1 To UplnCount
            cmbUnplnReas.AddItem gcolUnplnCodes.getCodeDesc(i)
        Next i
    End If
End Sub


Private Sub BindScrapCd()
    Dim i As Integer
    Dim SCdCount As Integer
    SCdCount = gcolScrapCodes.Count
    
    If SCdCount > 0 Then
        For i = 1 To SCdCount
            cmbScrapCd.AddItem gcolScrapCodes.getCodeDesc(i)
        Next i
    End If
End Sub

Private Sub ValidateUpdatedEntry()
    '//OPEN CONNECTION TO DATABASE IF IT IS NOT OPEN
    
    If ConnectToDB() = vbError Then
        gbConnected = False
        Call MainError("Cannot connect to DB")
        Exit Sub
    Else
        gbConnected = True
    End If
    
    Call frmDCMassEntry.IsValid(UpdatedEntry)
    
    '//CLOSE CONNECTION TO DATABASE
    DisconnectDB
End Sub
Private Sub BindEntryData(ByRef record As MassEntry, Optional DataType As String = "ALL")
    
    If DataType = "EmplNo" Or DataType = "ALL" Then
        '//Employee Info
        txtEmplNo.Text = record.EmplNo
        txtEmplName.Text = record.EmplName
    End If
    
    If DataType = "OrderNo" Or DataType = "ALL" Then
        '//Job Info
        txtOrderNo.Text = record.OrderNo
        txtOrderQty.Text = record.OrderQty
    End If
    
    If DataType = "SeqNumb" Or DataType = "OrderNo" Or DataType = "ALL" Then
        txtSeqNumb.Text = record.SeqNumb
        txtMchStd.Text = record.MchStd
    End If
    
    If DataType = "ALL" Then
        '//Labor Info
        dtActDate.Value = IIf(IsDate(Format$(record.Actdate, "mm/dd/yyyy")), record.Actdate, Now)
        txtActQty.Text = record.ActQty
        txtMachHrs.Text = record.MachHrs
        
        If record.UnplnReasIsValid Then
            If cmbUnplnReas.Text <> gcolUnplnCodes.getCodeDesc(Right("00" & record.UnplnReas, 2)) Then
                cmbUnplnReas.SelText = gcolUnplnCodes.getCodeDesc(Right("00" & record.UnplnReas, 2))
            Else
                'Do Nothing
            End If
        End If
        
        txtScrapQty.Text = record.scrapQty
        
        If record.ScrapCdIsValid Then
            If cmbScrapCd.Text <> gcolScrapCodes.getCodeDesc(Right("00" & record.ScrapCd, 2)) Then
                cmbScrapCd.SelText = gcolScrapCodes.getCodeDesc(Right("00" & record.ScrapCd, 2))
            Else
                'Do Nothing
            End If
        End If
    End If
    
    If Len(record.ErrorMsg) > 0 Then
        lblErrorMsg.Caption = "Errors Found: " & record.ErrorMsg
        lblErrorMsg.FontBold = True
        lblErrorMsg.ForeColor = vbRed
    Else
        lblErrorMsg.Caption = "No Errors Found!"
        lblErrorMsg.FontBold = True
        lblErrorMsg.ForeColor = vbGreen
    End If
    
    Call ResetChangedValues
    
End Sub

Private Sub txtEmplNo_Change()
    EmplNoChanged = True
End Sub

Private Sub txtEmplNo_Validate(Cancel As Boolean)
    
    If EmplNoChanged Then
        UpdatedEntry.EmplNo = Trim(txtEmplNo.Text)
    
        Call ValidateUpdatedEntry
        
        If UpdatedEntry.EmplNoIsValid Then
            Call BindEntryData(UpdatedEntry, "EmplNo")
            Cancel = False
        Else
            Call BindEntryData(UpdatedEntry, "EmplNo")
            Cancel = True
        End If
    End If
    
End Sub

Private Sub txtOrderNo_Change()
    OrderNoChanged = True
End Sub

Private Sub txtOrderNo_Validate(Cancel As Boolean)
    
    If OrderNoChanged Then
        UpdatedEntry.OrderNo = Trim(txtOrderNo.Text)
        
        Call ValidateUpdatedEntry
        
        If UpdatedEntry.OrderNoIsValid Then
            Call BindEntryData(UpdatedEntry, "OrderNo")
            Cancel = False
        Else
            Call BindEntryData(UpdatedEntry, "OrderNo")
            Cancel = True
        End If
    End If
    
End Sub

Private Sub txtScrapQty_Change()
    ScrapQtyChanged = True
End Sub

Private Sub txtScrapQty_Validate(Cancel As Boolean)

    If ScrapQtyChanged Then
        UpdatedEntry.scrapQty = Trim(txtScrapQty.Text)
        
        Call ValidateUpdatedEntry
        
        If Len(UpdatedEntry.ErrorMsg) = 0 Then
            Call BindEntryData(UpdatedEntry, "ALL")
            Cancel = False
        Else
            Call BindEntryData(UpdatedEntry, "ALL")
            Cancel = True
        End If
    End If
    
End Sub

Private Sub txtSeqNumb_Change()
    SeqNumbChanged = True
End Sub

Private Sub txtSeqNumb_Validate(Cancel As Boolean)

    If SeqNumbChanged Then
        UpdatedEntry.SeqNumb = Trim(txtSeqNumb.Text)
        
        Call ValidateUpdatedEntry
        
        If UpdatedEntry.SeqNumbIsValid Then
            Call BindEntryData(UpdatedEntry, "SeqNumb")
            Cancel = False
        Else
            Call BindEntryData(UpdatedEntry, "SeqNumb")
            Cancel = True
        End If
    End If
    
End Sub

Private Sub dtActDate_Change()
    ActDateChanged = True
End Sub

Private Sub dtActDate_Validate(Cancel As Boolean)

    If ActDateChanged Then
        UpdatedEntry.Actdate = dtActDate
        
        Call ValidateUpdatedEntry
        
        If Strings.Left(UpdatedEntry.ErrorMsg, 4) = "Date" Then
            Call BindEntryData(UpdatedEntry, "ALL")
            Cancel = True
        Else
            Call BindEntryData(UpdatedEntry, "ALL")
            Cancel = False
        End If
    End If
    
End Sub

Private Sub cmbUnplnReas_Click()
    UnplnReasChanged = True

    If UnplnReasChanged Then
        UpdatedEntry.UnplnReas = Strings.Left(cmbUnplnReas.Text, 2)
        
        Call ValidateUpdatedEntry
        
        If Not UpdatedEntry.UnplnReasIsValid Then
            Call BindEntryData(UpdatedEntry, "ALL")
            Cancel = True
        Else
            Call BindEntryData(UpdatedEntry, "ALL")
            Cancel = False
        End If
    End If
End Sub

Private Sub cmbScrapCd_Click()
    ScrapCdChanged = True
    
    If ScrapCdChanged Then
        UpdatedEntry.ScrapCd = Strings.Left(cmbScrapCd.Text, 2)
        
        Call ValidateUpdatedEntry
        
        If Not UpdatedEntry.ScrapCdIsValid Then
            Call BindEntryData(UpdatedEntry, "ALL")
            Cancel = True
        Else
            Call BindEntryData(UpdatedEntry, "ALL")
            Cancel = False
        End If
    End If
End Sub

Private Sub cmdSave_Click()

    '//Check if the data is valid
    If Len(UpdatedEntry.ErrorMsg) = 0 Then
        '//Data was valid, go ahead and save changes.
        Call CopyEntry(UpdatedEntry, Entry)
        Call frmDCMassEntry.UpdateGridItem(UpdatedEntry, EntryIndex)
        Unload Me
    Else
        MsgBox "Please correct the remaining issues before saving your changes.", vbOKOnly + vbExclamation, "Errors remain"
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Set UpdatedEntry = Nothing
    frmDCMassEntry.Tag = "EntryValidation"
    Unload Me
End Sub

Private Sub Form_Activate()

    Call Localization
    Call BindUnplnReas
    Call BindScrapCd
    Call BindEntryData(Entry, "ALL")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Entry = Nothing
    Set UpdatedEntry = Nothing
End Sub
