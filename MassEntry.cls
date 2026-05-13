VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MassEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarActDate As String 'local copy
'local variable(s) to hold property value(s)
Private mvarEmplNo As String 'local copy
Private mvarEmplNoIsValid As Boolean 'local copy
Private mvarEmplName As String 'local copy
Private mvarOrderNo As String 'local copy
Private mvarOrderNoIsValid As Boolean 'local copy
Private mvarOrderQty As String 'local copy
Private mvarSeqNumb As String 'local copy
Private mvarSeqNumbIsValid As Boolean 'local copy
Private mvarCCID As String 'local copy
Private mvarWCID As String 'local copy
Private mvarSetStd As String 'local copy
Private mvarMchStd As String 'local copy
Private mvarMachHrs As String 'local copy
Private mvarActQty As String 'local copy
Private mvarUnplnReas As String 'local copy
Private mvarUnplnReasIsValid As Boolean 'local copy
Private mvarScrapQty As String 'local copy
Private mvarScrapCd As String 'local copy
Private mvarScrapCdIsValid As Boolean 'local copy
Private mvarErrorMsg As String 'local copy
Private mvarIsDuplicate As Boolean 'local copy
Private mvarForceProcess As Boolean 'local copy

Public Property Let CCID(ByVal vData As String)
    mvarCCID = vData
End Property

Public Property Get CCID() As String
    CCID = mvarCCID
End Property


Public Property Let WCID(ByVal vData As String)
    mvarWCID = vData
End Property

Public Property Get WCID() As String
    WCID = mvarWCID
End Property


Public Property Let SetStd(ByVal vData As String)
    mvarSetStd = vData
End Property

Public Property Get SetStd() As String
    SetStd = mvarSetStd
End Property


Public Property Let ErrorMsg(ByVal vData As String)
    mvarErrorMsg = vData
End Property

Public Property Get ErrorMsg() As String
    ErrorMsg = mvarErrorMsg
End Property


Public Property Let EmplName(ByVal vData As String)
    mvarEmplName = vData
End Property

Public Property Get EmplName() As String
    EmplName = mvarEmplName
End Property


Public Property Let ScrapCd(ByVal vData As String)
    mvarScrapCd = vData
End Property

Public Property Get ScrapCd() As String
    ScrapCd = mvarScrapCd
End Property


Public Property Let ScrapCdIsValid(ByVal vData As Boolean)
    mvarScrapCdIsValid = vData
End Property

Public Property Get ScrapCdIsValid() As Boolean
    ScrapCdIsValid = mvarScrapCdIsValid
End Property


Public Property Let ScrapQty(ByVal vData As String)
    mvarScrapQty = vData
End Property

Public Property Get ScrapQty() As String
    ScrapQty = mvarScrapQty
End Property


Public Property Let UnplnReas(ByVal vData As String)
    mvarUnplnReas = vData
End Property

Public Property Get UnplnReas() As String
    UnplnReas = mvarUnplnReas
End Property


Public Property Let UnplnReasIsValid(ByVal vData As Boolean)
    mvarUnplnReasIsValid = vData
End Property

Public Property Get UnplnReasIsValid() As Boolean
    UnplnReasIsValid = mvarUnplnReasIsValid
End Property


Public Property Let ActQty(ByVal vData As String)
    mvarActQty = vData
End Property

Public Property Get ActQty() As String
    ActQty = mvarActQty
End Property


Public Property Let MachHrs(ByVal vData As String)
    mvarMachHrs = vData
End Property


Public Property Get MachHrs() As String
    MachHrs = mvarMachHrs
End Property


Public Property Let MchStd(ByVal vData As String)
    mvarMchStd = vData
End Property

Public Property Get MchStd() As String
    MchStd = mvarMchStd
End Property


Public Property Let SeqNumb(ByVal vData As String)
    mvarSeqNumb = vData
End Property

Public Property Get SeqNumb() As String
    SeqNumb = mvarSeqNumb
End Property


Public Property Let SeqNumbIsValid(ByVal vData As Boolean)
    mvarSeqNumbIsValid = vData
End Property

Public Property Get SeqNumbIsValid() As Boolean
    SeqNumbIsValid = mvarSeqNumbIsValid
End Property


Public Property Let OrderQty(ByVal vData As String)
    mvarOrderQty = vData
End Property

Public Property Get OrderQty() As String
    OrderQty = mvarOrderQty
End Property


Public Property Let OrderNoIsValid(ByVal vData As Boolean)
    mvarOrderNoIsValid = vData
End Property

Public Property Get OrderNoIsValid() As Boolean
    OrderNoIsValid = mvarOrderNoIsValid
End Property


Public Property Let OrderNo(ByVal vData As String)
    mvarOrderNo = vData
End Property

Public Property Get OrderNo() As String
    OrderNo = mvarOrderNo
End Property


Public Property Let EmplNo(ByVal vData As String)
    mvarEmplNo = vData
End Property

Public Property Get EmplNo() As String
    EmplNo = mvarEmplNo
End Property


Public Property Let ForceProcess(ByVal vData As Boolean)
    mvarForceProcess = vData
End Property

Public Property Get ForceProcess() As Boolean
    ForceProcess = mvarForceProcess
End Property


Public Property Let IsDuplicate(ByVal vData As Boolean)
    mvarIsDuplicate = vData
End Property

Public Property Get IsDuplicate() As Boolean
    IsDuplicate = mvarIsDuplicate
End Property


Public Property Let EmplNoIsValid(ByVal vData As Boolean)
    mvarEmplNoIsValid = vData
End Property

Public Property Get EmplNoIsValid() As Boolean
    EmplNoIsValid = mvarEmplNoIsValid
End Property


Public Property Let Actdate(ByVal vData As String)
    mvarActDate = vData
End Property

Public Property Get Actdate() As String
    Actdate = mvarActDate
End Property
