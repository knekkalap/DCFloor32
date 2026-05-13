VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form frmDCMassEntry 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9885
   ScaleMode       =   0  'User
   ScaleWidth      =   18500
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbLoading 
      Height          =   855
      Left            =   4230
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   1508
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame fmLoading 
      Height          =   2055
      Left            =   3975
      TabIndex        =   12
      Top             =   3915
      Width           =   10455
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         Caption         =   "LOADING... Please Wait... 0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   10095
      End
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Delete Record"
      Height          =   475
      Left            =   10767
      TabIndex        =   9
      Top             =   9250
      Width           =   1990
   End
   Begin VB.CommandButton btnProcess 
      Caption         =   "&Process Records"
      Height          =   475
      Left            =   8207
      TabIndex        =   8
      Top             =   9250
      Width           =   1990
   End
   Begin VB.CommandButton btnValidate 
      Caption         =   "&Modify Record"
      Height          =   475
      Left            =   5648
      TabIndex        =   7
      Top             =   9250
      Width           =   1990
   End
   Begin VB.TextBox ssname 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4507
      TabIndex        =   2
      Top             =   1320
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog dlgFileRequest 
      Left            =   13080
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbXLSheetName 
      Height          =   315
      Left            =   11962
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton btnBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   8640
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLaborRecs 
      Bindings        =   "frmDCMassEntry.frx":0000
      Height          =   6899
      Left            =   165
      TabIndex        =   5
      Top             =   2205
      Width           =   18060
      _ExtentX        =   31856
      _ExtentY        =   12171
      _Version        =   393216
      RowSizingMode   =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblErrorCount 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   8895
      TabIndex        =   10
      Top             =   1815
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   6611
      Picture         =   "frmDCMassEntry.frx":0013
      Stretch         =   -1  'True
      Top             =   240
      Width           =   870
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Upload Labor Sheet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7653
      TabIndex        =   6
      Top             =   360
      Width           =   4140
   End
   Begin VB.Label Label2 
      Caption         =   "Excel Sheet"
      Height          =   255
      Left            =   10867
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Excel File"
      Height          =   255
      Left            =   3667
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "frmDCMassEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xl As New Excel.Application
Dim xlSheet As New Excel.Worksheet
Dim xlwbook As New Excel.Workbook
Dim xlPass As String

Dim cnxn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strCmd As String

Dim headerRowExists As Boolean
Dim iLastRow, iRow As Integer
Dim iLastCol, iCol As Integer
Dim prevRow, curRow, selRow As Integer
Dim prevText As String

Dim entries As New Collection
Dim invalidRows As Integer
Dim perComp As Double

Private Sub GetExcelFile()
    '//Open dialog box for the user to select Excel file
    With dlgFileRequest
        .DialogTitle = "Select Excel Spreasheet"
        .InitDir = App.Path
        .Filter = "Excel Spreadsheets (*.xlsx)|*.xlsx|Excel Spreadsheets 97 (*.xls)|*.xls"
        .ShowOpen
    End With
End Sub

Private Sub OpenExcel()
    If xl Is Nothing Then
        Set xl = New Excel.Application
    End If
    
    If xl.Workbooks.Count = 0 And Len(ssname.Text) > 0 Then
        Set xlwbook = xl.Workbooks.Open(FileName:=ssname.Text, Password:=xlPass)
    End If
End Sub

Private Sub CloseExcel()
    If dlgFileRequest.FileName > "" And Not xl.ActiveWorkbook Is Nothing Then
        xl.ActiveWorkbook.Close SaveChanges:=False, FileName:=ssname.Text
        xl.Quit
        Set xlSheet = Nothing
        Set xl = Nothing
    End If
End Sub

Private Function RequestPassword() As String
    RequestPassword = InputBox("Please enter the spreadsheet password.", "Password Entry")
End Function

Private Sub FillGrid()
On Error GoTo Handler

    invalidRows = 0
    Dim i As Integer
    
    Call OpenExcel
    
    
    fmLoading.Visible = True
    pbLoading.Visible = True
    
    '//Update Progress
    lblProgress.Caption = "Connecting to file."
    pbLoading.Value = 0
    
    '//Set the active sheet to the sheet selected by the user
    Set xlSheet = xlwbook.Sheets.Item(cmbXLSheetName.ListIndex + 1)
    
    xlSheet.Activate
    
    '//Update Progress
    lblProgress.Caption = "Opening Work Sheet."
    pbLoading.Value = 5
    
    '//With the sheet activated, get the total number of rows that need to be inserted
    With xlSheet
        iLastCol = 11
        iLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    
    '//Update Progress
    lblProgress.Caption = "Counting Records."
    pbLoading.Value = 10
    
    If xl.Cells(1, 1) = "Date" Then
        'Skip header row
        headerRowExists = True
    End If
    
    '//Size the grid to fit the data in the spreadsheet
    With grdLaborRecs
        .Redraw = False

        If headerRowExists Then
            .Rows = iLastRow
        Else
            .Rows = iLastRow + 1 'Expand size of the grid to account for missing header row
        End If
    End With
    
    '//Update Progress
    lblProgress.Caption = "Reading Records."
    pbLoading.Value = 20
    
    '//Open connection and get data from the spreadsheet
    With cnxn
       .Provider = "Microsoft.ACE.OLEDB.12.0"
       .Properties("Data Source").Value = dlgFileRequest.FileName
       .Properties("Extended Properties").Value = "Excel 12.0 Xml;HDR=YES;IMEX=1"
       .Properties("Jet OLEDB:Database Password").Value = xlPass
       .Open
    End With
    
    strCmd = "SELECT [Date], [Emp No], [Job No], [Order Qty], [Seq No], [Machine Std], [Machine Hrs], [Quantity], [Unplan Reas], " & _
                "[Scrap Quantity], [Scrap Code] FROM [" & cmbXLSheetName.Text & "$]"
    
    rs.Open strCmd, cnxn, adOpenStatic, adLockReadOnly
    
    '//Update Progress
    lblProgress.Caption = "Extracting Records."
    pbLoading.Value = 40
    
    '//Clear previous collection if one exists
    If entries.Count > 0 Then
        Call ClearCollection(entries)
    End If
    
    '//Update Progress
    pbLoading.Value = 50
    
    '//Build a collection of entries from the recordset
    Call BuildCollection(rs, entries)
    
    '//Update Progress
    lblProgress.Caption = "Preparing to Validate Records."
    pbLoading.Value = 95
    
    Call CloseExcel
    
    '//Update Progress
    pbLoading.Value = 100
    
    '//Validate data from spreadsheet
    Call ValidateGrid(grdLaborRecs)
    
    '//Bind data to the grid
    For i = 1 To entries.Count
    
        '//Update progress bar
        perComp = (i / entries.Count * 100)
        lblProgress.Caption = "Importing Records... Please Wait. " & CInt(perComp) & "%"
        lblProgress.Refresh
        pbLoading.Value = perComp
        
        grdLaborRecs.TextMatrix(i, 1) = Format$(entries.Item(i).Actdate, gDateformat)
        grdLaborRecs.TextMatrix(i, 2) = entries.Item(i).EmplNo
        grdLaborRecs.TextMatrix(i, 3) = entries.Item(i).OrderNo
        grdLaborRecs.TextMatrix(i, 4) = entries.Item(i).OrderQty
        grdLaborRecs.TextMatrix(i, 5) = entries.Item(i).SeqNumb
        grdLaborRecs.TextMatrix(i, 6) = entries.Item(i).MchStd
        grdLaborRecs.TextMatrix(i, 7) = entries.Item(i).MachHrs
        grdLaborRecs.TextMatrix(i, 8) = entries.Item(i).ActQty
        grdLaborRecs.TextMatrix(i, 9) = entries.Item(i).UnplnReas
        grdLaborRecs.TextMatrix(i, 10) = entries.Item(i).ScrapQty
        grdLaborRecs.TextMatrix(i, 11) = entries.Item(i).ScrapCd
    Next i
    
    grdLaborRecs.Redraw = True
    
    Call DisplayErrorCount
    
    GoTo Cleanup
    
    Exit Sub
    
Handler:
    If Err.Number > 0 Then
        MsgBox "ERROR: " & Err.Number & ", " & Err.Description
    End If
    
    Call CloseExcel
    
Cleanup:
    rs.Close
    Set rs = Nothing
    cnxn.Close
    Set cnxn = Nothing
    
End Sub


Private Sub RefreshGrid()

    '//Resize the grid to have the right number of columns
    grdLaborRecs.Cols = 12
    
    If Me.Tag = "EntryValidation" And curRow > 0 Then
        selRow = curRow
        Me.Tag = ""
    Else
        selRow = 0
    End If
        
    '//Set column widths and headers
    With grdLaborRecs
    
        .TextMatrix(0, 1) = "Date"                                  'Date
        .TextMatrix(0, 2) = "Emp No"                                'Employee Number
        .TextMatrix(0, 3) = "Job No"                                'Job Number
        .TextMatrix(0, 4) = "Order Qty"                             'Order Qty
        .TextMatrix(0, 5) = "Seq No"                                'Sequence Number
        .TextMatrix(0, 6) = "Machine Std"                           'Machine Std
        .TextMatrix(0, 7) = "Machine Hrs"                           'Machine Hours
        .TextMatrix(0, 8) = "Quantity"                              'Quantity
        .TextMatrix(0, 9) = "Unplan Reas"                          'Unplanned Reason
        .TextMatrix(0, 10) = "Scrap Quantity"                       'Scrap Quantity
        .TextMatrix(0, 11) = "Scrap Code"                           'Scrap Code
        
           
        .ColWidth(0) = 800                              'Selection Column
        .ColWidth(1) = 1200                             'Date
        .ColWidth(2) = 1200                             'Employee Number
        .ColWidth(3) = 1800                             'Job Number
        .ColWidth(4) = 1600                             'Order Qty
        .ColWidth(5) = 1200                             'Sequence Number
        .ColWidth(6) = 1600                             'Machine Std
        .ColWidth(7) = 1600                             'Machine Hours
        .ColWidth(8) = 1200                             'Quantity
        .ColWidth(9) = 1800                            'Unplanned Reason
        .ColWidth(10) = 1800                            'Scrap Quantity
        .ColWidth(11) = 1800                            'Scrap Code
        
        
        .Row = selRow
        .Refresh
        
    End With
End Sub

Public Sub UpdateGridItem(UpdatedEntry As MassEntry, EntryIndex As Integer)
    With grdLaborRecs
        .TextMatrix(EntryIndex, 0) = ""                         'Clear selection
        
        '//Update line item
        .TextMatrix(EntryIndex, 1) = Format$(UpdatedEntry.Actdate, gDateformat)
        .TextMatrix(EntryIndex, 2) = UpdatedEntry.EmplNo
        .TextMatrix(EntryIndex, 3) = UpdatedEntry.OrderNo
        .TextMatrix(EntryIndex, 4) = UpdatedEntry.OrderQty
        .TextMatrix(EntryIndex, 5) = UpdatedEntry.SeqNumb
        .TextMatrix(EntryIndex, 6) = UpdatedEntry.MchStd
        .TextMatrix(EntryIndex, 7) = UpdatedEntry.MachHrs
        .TextMatrix(EntryIndex, 8) = UpdatedEntry.ActQty
        .TextMatrix(EntryIndex, 9) = UpdatedEntry.UnplnReas
        .TextMatrix(EntryIndex, 10) = UpdatedEntry.ScrapQty
        .TextMatrix(EntryIndex, 11) = UpdatedEntry.ScrapCd
        
        '//Check if record was previously marked invalid
        If .CellBackColor = vbRed Then

            '//Clear invalid highlighting
            .Row = EntryIndex
            .Col = .FixedCols
            .RowSel = EntryIndex
            .ColSel = .Cols - 1
            .FillStyle = flexFillRepeat
            .CellBackColor = vbWhite
            .Refresh
            
            '//Reduce error count
            invalidRows = invalidRows - 1
            Call DisplayErrorCount
            lblErrorCount.Refresh
        End If
    End With
End Sub

Public Function IsValid(ByRef Entry As MassEntry) As Boolean
On Error GoTo Cleanup

    With Entry
    
        '//Clear previous ErrorMsg
        .ErrorMsg = ""
        
        '//Check for valid Date format
        If IsDate(Format$(.Actdate, gDateformat)) Then
            '//Check to see if this is a future date
            If CDate(.Actdate) > Now Then
                .ErrorMsg = "Date is in the future."
            End If
        Else
            .ErrorMsg = "Date is not in a valid format."
        End If
        
        '//Check if either Scrap Qty or Code are entered and the other is missing
        If (.ScrapCd = "" And .ScrapQty > "") Or (.ScrapCd > "" And .ScrapQty = "") Then
            If Len(.ErrorMsg) > 0 Then
                .ErrorMsg = .ErrorMsg + " " + LoadResString(gLanguageOffset + 66) & " or " & LoadResString(gLanguageOffset + 67) & "."
            Else
                .ErrorMsg = LoadResString(gLanguageOffset + 66) & " or " & LoadResString(gLanguageOffset + 67) & "."
            End If
            
            .ScrapCdIsValid = False
        End If
        
        '//Validate numeric values are numeric
        If Not IsNumeric(.ActQty) Or Not IsNumeric(.MachHrs) Or (Not IsNumeric(.ScrapQty) And .ScrapQty <> "") Or Not IsNumeric(.SeqNumb) Then
            If Len(.ErrorMsg) > 0 Then
                .ErrorMsg = .ErrorMsg + " " + "Quantity, Machine Hours, Scrap Qty, or Sequence Number are expected to be numeric."
            Else
                .ErrorMsg = "Quantity, Machine Hours, Scrap Qty, or Sequence Number are expected to be numeric."
            End If
        End If
        
        If Len(.UnplnReas) > 0 Then
            If gcolUnplnCodes.getCode(Right("00" & .UnplnReas, 2)) = LoadResString(gLanguageOffset + 4) Then
                If Len(.ErrorMsg) > 0 Then
                    .ErrorMsg = .ErrorMsg + " " + "Unplanned Reason code invalid."
                Else
                    .ErrorMsg = "Unplanned Reason code invalid."
                End If
                
                .UnplnReas = ""
                .UnplnReasIsValid = False
            Else
                .UnplnReasIsValid = True
            End If
        End If
        
        If Len(.ScrapCd) > 0 Then
            If gcolScrapCodes.getCode(Right("00" & .ScrapCd, 2)) = LoadResString(gLanguageOffset + 4) Then
                If Len(.ErrorMsg) > 0 Then
                    .ErrorMsg = .ErrorMsg + " " + "Scrap code invalid."
                Else
                    .ErrorMsg = "Scrap code invalid."
                End If
                
                .ScrapCd = ""
                .ScrapCdIsValid = False
            Else
                .ScrapCdIsValid = True
            End If
        End If
        
    End With
        
    '//Validate all other info
    Call ValidateEntry(Entry)
    
    If Len(Entry.ErrorMsg) > 0 Then
        GoTo Invalid
    End If
    
Valid:
    IsValid = True
    Exit Function

Invalid:
    IsValid = False
    Exit Function

Cleanup:
    Call CloseExcel
    
End Function


Public Sub ValidateGrid(LaborRecs As MSHFlexGrid)
On Error GoTo Cleanup
    
    Dim i As Integer
    
    '//OPEN CONNECTION TO DATABASE IF IT IS NOT OPEN
    
    If ConnectToDB() = vbError Then
        gbConnected = False
        Call MainError("Cannot connect to DB")
        Exit Sub
    Else
        gbConnected = True
    End If
    
    '//Start validating items in the grid
    For i = 1 To entries.Count
        
        '//Update progress bar
        perComp = (i / entries.Count * 100)
        lblProgress.Caption = "Validating Records... Please Wait. " & CInt(perComp) & "%"
        lblProgress.Refresh
        pbLoading.Value = perComp
    
        '//Validate records
        If Not IsValid(entries.Item(i)) Then
            With LaborRecs
                '//Highlight invalid rows in red
                    .Row = i
                    .Col = .FixedCols
                    .RowSel = i
                    .ColSel = .Cols - 1
                    .FillStyle = flexFillRepeat
                    .CellBackColor = vbRed
            End With
            invalidRows = invalidRows + 1
        End If
    Next
    
    '//CLOSE DATABASE CONNECTION
    DisconnectDB

    Exit Sub

Cleanup:
    Call CloseExcel
    DisconnectDB
End Sub

Private Sub DisplayErrorCount(Optional Processed As Boolean = False)
    If Not Processed Then
        If invalidRows > 0 Then
            '//No cookie for you. You have bad data.
            lblErrorCount = "There are " & invalidRows & " errors found. See red items below. " & entries.Count & " records were found."
            lblErrorCount.Left = (Me.Width / 2) - (lblErrorCount.Width / 2)
            lblErrorCount.BackColor = vbRed
            lblErrorCount.Visible = True
        Else
            '//Cookies for everyone!
            lblErrorCount = "There are " & invalidRows & " known errors in the " & entries.Count & " records found." & _
                            "You may now process the data."
            lblErrorCount.BackColor = vbGreen
            lblErrorCount.Visible = True
        End If
    Else
        If entries.Count > 0 Then
            lblErrorCount.Caption = entries.Count & " records may be duplicates. Please review these records and enter them manually if needed."
            lblErrorCount.BackColor = RGB(255, 127, 0)
        Else
            lblErrorCount.Caption = "All records have been processed!"
            lblErrorCount.BackColor = vbGreen
        End If
    End If
    lblErrorCount.Left = (Me.Width / 2) - (lblErrorCount.Width / 2)
End Sub

Private Sub btnProcess_Click()
    If invalidRows > 0 Then
        '//Nope - fix your data
        MsgBox "You cannot process these records due to invalid entries. Please correct the items marked in red and try again.", vbExclamation + vbOKOnly, _
                    "Cannot Process Entries"
        
        curRow = grdLaborRecs.RowSel
        If curRow > 0 Then
            Call grdLaborRecs_Click
        End If
    Else
        Dim resp As VbMsgBoxResult
        
        resp = MsgBox("Processing these records will send them to Data Collection and clear all records from the grid." & vbCrLf & vbCrLf & _
                "Click OK to continue processing the records.", _
                vbOKCancel + vbExclamation, "Attention!")
        
        If resp = vbCancel Then
            Exit Sub
        End If
        
        '//Code for handling data processing
        Dim i As Integer
    
        '//OPEN CONNECTION TO DATABASE IF IT IS NOT OPEN
        
        If ConnectToDB() = vbError Then
            gbConnected = False
            Call MainError("Cannot connect to DB")
            Exit Sub
        Else
            gbConnected = True
        End If
        
        For i = entries.Count To 1 Step -1
            If ProcessEntry(entries.Item(i)) Then
            
                If grdLaborRecs.Rows > 2 Then
                    grdLaborRecs.RemoveItem (i)
                Else
                    grdLaborRecs.Clear
                    Call RefreshGrid
                End If
                
                entries.Remove (i)
            End If
        Next i
        
        Call DisplayErrorCount(True)
        
        '//CLOSE DATABASE CONNECTION
        DisconnectDB
    End If
End Sub

Private Sub btnBrowse_Click()
    dlgFileRequest.CancelError = True
On Error GoTo ERROR
    
    ssname.Text = ""
    ssname.Refresh
    
    '//Prompt for password and open dialog box for user to supply Excel file
    xlPass = RequestPassword()
    Call GetExcelFile
    
    lblErrorCount.Caption = ""
    lblErrorCount.Visible = False
    lblErrorCount.Refresh
    
    ssname.Text = dlgFileRequest.FileName
    ssname.Refresh
    
    '//Update the file name box and get the count of sheets in the workbook
    Dim nos As Integer
    If ssname.Text <> "" Then
    
        grdLaborRecs.Clear
        grdLaborRecs.Rows = 2
        Call RefreshGrid
        
        If xl.Workbooks.Open(ssname.Text, Password:=xlPass).HasPassword Then
            
            Set xlwbook = xl.Workbooks.Open(ssname.Text, Password:=xlPass)
            
            nos = xlwbook.Sheets.Count
            
            Dim n As Integer
            '//Refresh the combo box
            cmbXLSheetName.Clear
            cmbXLSheetName.Refresh
            
            '//Add sheet list to the drop-down box so the user can select which one has the data
            For n = 1 To nos
                cmbXLSheetName.AddItem xlwbook.Sheets(n).Name
            Next
            '//Automatically select the sheet  for processing if only one exists
            If nos = 1 Then
                cmbXLSheetName.ListIndex = 0
                cmbXLSheetName.Refresh
            End If
        Else
            MsgBox "Incorrect file format." & vbCrLf & "Expecting password protected file."
        End If
    End If
    
    Exit Sub
    
ERROR:
    If Err.Description = "Cancel was selected." Then
        '//Do Nothing
    Else
        MsgBox Err.Description, vbOKOnly, "Unable to open workbook"
    End If
End Sub

Private Sub btnDelete_Click()
    
    Dim confirm As String
    
    If grdLaborRecs.Row > 0 Then
        confirm = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Confirm Delete?")
            
        If confirm = vbYes And grdLaborRecs.Row > 0 Then
            If grdLaborRecs.Rows > grdLaborRecs.Row And grdLaborRecs.Rows > 2 Then
            
                '//Remove the item from the collection
                If grdLaborRecs.CellBackColor = vbRed Then
                    invalidRows = invalidRows - 1
                End If
                
                entries.Remove (grdLaborRecs.Row)
                '//Remove the item from the grid
                grdLaborRecs.RemoveItem (grdLaborRecs.Row)
                grdLaborRecs.Row = 0
                Call DisplayErrorCount
            Else
                MsgBox "Cannot delete this record. The table must contain at least one row.", vbOKOnly, "Warning!"
            End If
        End If
    Else
        MsgBox "You must first select a row to delete.", vbOKOnly, "Warning!"
    End If
    
End Sub

Private Sub btnValidate_Click()
    If Not grdLaborRecs.Row = 0 Then
        Load frmEntryValidation
        Call frmEntryValidation.SetEntry(grdLaborRecs.Row, entries.Item(grdLaborRecs.Row))
        frmEntryValidation.Show vbModal
    Else
        MsgBox "You must first select a record to validate.", vbExclamation + vbOKOnly, "No Record Selected"
    End If
End Sub

Private Sub grdLaborRecs_Click()
    
    curRow = grdLaborRecs.RowSel
    prevText = grdLaborRecs.TextMatrix(curRow, 0)
    
    If curRow = 0 Then
        MsgBox "You cannot select this row.", vbOKOnly, "Warning!"
    Else
        If curRow <> prevRow And prevRow > 0 Then
            grdLaborRecs.TextMatrix(curRow, 0) = grdLaborRecs.TextMatrix(curRow, 0) & "    >>"
            If prevRow < grdLaborRecs.Rows Then
                grdLaborRecs.TextMatrix(prevRow, 0) = ""
            End If
        Else
            grdLaborRecs.TextMatrix(curRow, 0) = grdLaborRecs.TextMatrix(curRow, 0) & "    >>"
        End If
    End If
        
    prevRow = curRow
    
End Sub

Private Sub cmbXLSheetName_Click()
On Error GoTo Cleanup

    '//Clear and hide error message before refreshing data
    lblErrorCount.Caption = ""
    lblErrorCount.Visible = False
    lblErrorCount.Refresh
    
    Call FillGrid
    
Cleanup:

    fmLoading.Visible = False
    pbLoading.Visible = False
    
    Call CloseExcel

End Sub

Private Sub Form_Activate()
    
    Call RefreshGrid
    pbLoading.Value = 0
    fmLoading.Visible = False
    pbLoading.Visible = False
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    fmLoading.Visible = False
    pbLoading.Visible = False
    If Not xl Is Nothing Then
        Call CloseExcel
    End If
End Sub

