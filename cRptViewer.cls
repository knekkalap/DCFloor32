VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mrptname As String
Private mNumbOfParams As Long
Private mFromDateParam As Date
Private mToDateParam As Date
Private mRptRunOption As Long

'EmplNo used with numbofparams = 5
Private mEmplNo As Long
'mstartshift and mendshift used with numbofparams = 4
Private mStartShift As Integer
Private mEndShift As Integer
'mlaborrecapemplno and mlaborrecaphomedpt used with numbofparams = 6
Private mLaborRecapEmplNo As Long
Private mLaborRecapHomeDpt As String
Private mDatabase As String
Private mSQLServer As String

Const SQLLOGIN = "DataCollection"
Const SQLPASSWD = "Cherrygarcia"
Const NO_PARAMS = 0
Const FROM_DATE_ONLY = 1
Const FROM_TO_DATES = 2
Const FROM_TO_DATES_W_ROPT = 3
Const SPECL_LABOR = 4
Const EMPL_ATTD_HIST = 5
Const LABOR_RECAP = 6
Const MTD_TESTS = 7
Const SHOP_MESSAGES = 8


Public Property Get Database() As String
Database = Trim(mDatabase)
End Property
Public Property Let Database(sDatabase As String)
mDatabase = sDatabase
End Property
Public Property Get SQLServer() As String
SQLServer = Trim(mSQLServer)
End Property
Public Property Let SQLServer(sSQLServer As String)
mSQLServer = sSQLServer
End Property

Public Property Get FromDateParam() As Date
   FromDateParam = mFromDateParam
End Property
Public Property Let FromDateParam(sFromDateParam As Date)
mFromDateParam = sFromDateParam
End Property
Public Property Get ToDateParam() As Date
   ToDateParam = mToDateParam
End Property
Public Property Let ToDateParam(sToDateParam As Date)
mToDateParam = sToDateParam
End Property
Public Property Get rptRunOption() As Long
   rptRunOption = mRptRunOption
End Property
Public Property Let rptRunOption(lrptRunOption As Long)
mRptRunOption = lrptRunOption
End Property
Public Property Get LaborRecapEmplNo() As Long
   LaborRecapEmplNo = mLaborRecapEmplNo
End Property
Public Property Let LaborRecapEmplNo(lLaborRecapEmplNo As Long)
mLaborRecapEmplNo = lLaborRecapEmplNo
End Property
Public Property Get LaborRecapHomeDpt() As String
   LaborRecapHomeDpt = mLaborRecapHomeDpt
End Property
Public Property Let LaborRecapHomeDpt(sLaborRecapHomeDpt As String)
mLaborRecapHomeDpt = sLaborRecapHomeDpt
End Property
Public Property Get EmplNo() As Long
   EmplNo = mEmplNo
End Property
Public Property Let EmplNo(lEmplNo As Long)
mEmplNo = lEmplNo
End Property
Public Property Get startshift() As Integer
   startshift = mStartShift
End Property
Public Property Let startshift(lstartshift As Integer)
mStartShift = lstartshift
End Property
Public Property Get Endshift() As Integer
   Endshift = mEndShift
End Property
Public Property Let Endshift(lEndshift As Integer)
mEndShift = lEndshift
End Property
Public Property Get rptname() As String
   rptname = mrptname
End Property
Public Property Let rptname(sRptname As String)
mrptname = sRptname
End Property
Public Property Get numbOfParams() As Long
   numbOfParams = mNumbOfParams
End Property
Public Property Let numbOfParams(lnumbOfParams As Long)
mNumbOfParams = lnumbOfParams
End Property
Public Function ShowReport()
Dim i As Long
Dim j As Long
Dim frmrptviewer As frmCrystalReports

Dim MyApp As New CRAXDRT.Application
Dim MyRpt As New CRAXDRT.Report
Dim crxTables As CRAXDRT.DatabaseTables
Dim crxTable As CRAXDRT.DatabaseTable
Dim crxSubreportObject As CRAXDRT.SubreportObject
Dim crxSubReport As CRAXDRT.Report
Dim crxSections As CRAXDRT.Sections
Dim crxSection As CRAXDRT.Section


Set frmrptviewer = New frmCrystalReports
Set MyRpt = MyApp.OpenReport(rptname, 1)
Screen.MousePointer = vbHourglass
MyRpt.DiscardSavedData

'Set the connection for the report.
'SetLogOnInfo is a deprecated method in CR 9, _
    but still works.
'The suggested method for CR9 is to use _
    the ConnectionProperty object.
MyRpt.Database.Tables(1).SetLogOnInfo SQLServer, Database, SQLLOGIN, SQLPASSWD
'This removes the schema from the Database Table's Location property.
Set crxTables = MyRpt.Database.Tables
For Each crxTable In crxTables
    With crxTable
         .Location = .Name
    End With
Next

Set crxSections = MyRpt.Sections

For i = 1 To crxSections.Count
    Set crxSection = crxSections(i)
    
    For j = 1 To crxSection.ReportObjects.Count
    
        If crxSection.ReportObjects(j).Kind = crSubreportObject Then
            Set crxSubreportObject = crxSection.ReportObjects(j)
            
            'Open the subreport, and treat like any other report
            Set crxSubReport = crxSubreportObject.OpenSubreport
            Set crxTables = crxSubReport.Database.Tables
            
            For Each crxTable In crxTables
                With crxTable
                    .SetLogOnInfo SQLServer, Database, SQLLOGIN, SQLPASSWD
                    .Location = .Name
                End With
            Next
        End If
    Next j
Next i

Select Case numbOfParams
  Case NO_PARAMS
  Case FROM_DATE_ONLY
    MyRpt.ParameterFields(1).ClearCurrentValueAndRange
    MyRpt.ParameterFields(1).AddCurrentValue FromDateParam
  Case FROM_TO_DATES
    MyRpt.ParameterFields(1).ClearCurrentValueAndRange
    MyRpt.ParameterFields(1).AddCurrentValue FromDateParam
    MyRpt.ParameterFields(2).ClearCurrentValueAndRange
    MyRpt.ParameterFields(2).AddCurrentValue ToDateParam
  Case FROM_TO_DATES_W_ROPT
    MyRpt.ParameterFields(1).ClearCurrentValueAndRange
    MyRpt.ParameterFields(1).AddCurrentValue rptRunOption
    MyRpt.ParameterFields(2).ClearCurrentValueAndRange
    MyRpt.ParameterFields(2).AddCurrentValue FromDateParam
    MyRpt.ParameterFields(3).ClearCurrentValueAndRange
    MyRpt.ParameterFields(3).AddCurrentValue ToDateParam
  Case SPECL_LABOR
    MyRpt.ParameterFields(1).ClearCurrentValueAndRange
    MyRpt.ParameterFields(1).AddCurrentValue FromDateParam
    MyRpt.ParameterFields(2).ClearCurrentValueAndRange
    MyRpt.ParameterFields(2).AddCurrentValue ToDateParam
    MyRpt.ParameterFields(3).ClearCurrentValueAndRange
    MyRpt.ParameterFields(3).AddCurrentValue startshift
    MyRpt.ParameterFields(4).ClearCurrentValueAndRange
    MyRpt.ParameterFields(4).AddCurrentValue Endshift
  Case EMPL_ATTD_HIST
    MyRpt.ParameterFields(1).ClearCurrentValueAndRange
    MyRpt.ParameterFields(1).AddCurrentValue EmplNo
    MyRpt.ParameterFields(2).ClearCurrentValueAndRange
    MyRpt.ParameterFields(2).AddCurrentValue FromDateParam
    MyRpt.ParameterFields(3).ClearCurrentValueAndRange
    MyRpt.ParameterFields(3).AddCurrentValue ToDateParam
  Case LABOR_RECAP
    MyRpt.ParameterFields(1).ClearCurrentValueAndRange
    MyRpt.ParameterFields(1).AddCurrentValue FromDateParam
    MyRpt.ParameterFields(2).ClearCurrentValueAndRange
    MyRpt.ParameterFields(2).AddCurrentValue ToDateParam
    MyRpt.ParameterFields(3).ClearCurrentValueAndRange
    MyRpt.ParameterFields(3).AddCurrentValue rptRunOption
    MyRpt.ParameterFields(4).ClearCurrentValueAndRange
    MyRpt.ParameterFields(4).AddCurrentValue LaborRecapEmplNo
    MyRpt.ParameterFields(5).ClearCurrentValueAndRange
    MyRpt.ParameterFields(5).AddCurrentValue LaborRecapHomeDpt
  Case MTD_TESTS
    MyRpt.ParameterFields(1).ClearCurrentValueAndRange
    MyRpt.ParameterFields(1).AddCurrentValue rptRunOption
    MyRpt.ParameterFields(2).ClearCurrentValueAndRange
    MyRpt.ParameterFields(2).AddCurrentValue FromDateParam
  Case SHOP_MESSAGES
    MyRpt.ParameterFields(1).ClearCurrentValueAndRange
    MyRpt.ParameterFields(1).AddCurrentValue FromDateParam
    MyRpt.ParameterFields(2).ClearCurrentValueAndRange
    MyRpt.ParameterFields(2).AddCurrentValue EmplNo
  
  End Select

'MyRpt.Database.Verify     'required!!! but causes unwanted external prompting from CR.

frmrptviewer.CRViewer1.EnableCloseButton = False
frmrptviewer.CRViewer1.ReportSource = MyRpt
frmrptviewer.Show

frmrptviewer.CRViewer1.ViewReport
While frmrptviewer.CRViewer1.IsBusy
   DoEvents
Wend
Screen.MousePointer = vbNormal
Set MyRpt = Nothing
Set MyApp = Nothing

End Function

Private Sub Class_Initialize()
'Database = ReadIniDatabase()
'SQLServer = ReadIniServer()
End Sub
'Public Function ReadIniServer() As String
'Const strFunction = "ReadIniServer"
'Dim strErrorMsg         As String   'Var to Contain the Error Message
'ReadIniServer = ""
'ReadIniServer = GetIniSetting("General", "Server")
'End Function
'Public Function ReadIniDatabase() As String
'Const strFunction = "ReadIniDatabase"
'Dim strErrorMsg         As String   'Var to Contain the Error Message
'ReadIniDatabase = ""
'ReadIniDatabase = GetIniSetting("General", "Database")
'End Function

