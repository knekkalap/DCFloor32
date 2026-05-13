VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIIItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Class Name       : clsIIItem
'Description      : The class that is an object to represent an Indirect
'                   Incentives Item.
'Properties       : CCID, Code, CodeDesc, Desc, IncenType, MachNumb, WrkStd
'Called From      : clsAllIIItems.AddIIItem, clsAllIIItem.GetIIItem,
'                   clsCCIIItems.AddIIItem, clsCCIIItem.GetIIItem,
'                   DBCommon.LdIndirIncn, Inputs,getIndirIncn,
'Dependencies     :
'==============================================================================

' The Private Properties of a Indirect Incentive Item
Private strCode As String
Private strDesc As String
Private strCodeDesc As String
Private strCCID As String
Private iType As Integer
Private iNumb As Integer
Private cWrkStd As Currency
Property Get Code() As String
    'Expose the Code property of the Indirect Incentive Item.
    Code = strCode
End Property

Property Let Code(strDesc As String)
    'Expose the Code property of the Indirect Incentive Item.
    strCode = strDesc
End Property

Property Get CodeDesc() As String
    'Expose the Code and Desc properties of the Indirect Incentive Item.
    CodeDesc = strCodeDesc
End Property

Property Let CodeDesc(strCodeDescIn As String)
    'Expose the Code and Desc properties of the Indirect Incentive Item.
    strCodeDesc = strCodeDescIn
End Property

Property Get Desc() As String
    'Expose the Desc property of the Indirect Incentive Item.
    Desc = strDesc
End Property

Property Let Desc(strDescIn As String)
    'Expose the Desc property of the Indirect Incentive Item.
    strDesc = strDescIn
End Property

Public Property Get IncenType() As Integer
    'Expose the Incentive type property of the Indirect Incentive Item.
    IncenType = iType
End Property

Public Property Let IncenType(iTypeIn As Integer)
    'Expose the Incentive type property of the Indirect Incentive Item.
    iType = iTypeIn
End Property

Public Property Get MachNumb() As Integer
    'Expose the Machine No property of the Indirect Incentive Item.
    MachNumb = iNumb
End Property

Public Property Let MachNumb(iNumbIn As Integer)
    'Expose the Machine No property of the Indirect Incentive Item.
    iNumb = iNumbIn
End Property

Public Property Get CCID() As String
    'Expose the CCID property of the Indirect Incentive Item.
    CCID = strCCID
End Property

Public Property Let CCID(strCCIDIn As String)
    'Expose the CCID property of the Indirect Incentive Item.
    strCCID = strCCIDIn
End Property

Public Property Get WrkStd() As Currency
    'Expose the Work Standard property of the Indirect Incentive Item.
    WrkStd = cWrkStd
End Property

Public Property Let WrkStd(cWrkStdIn As Currency)
    'Expose the Work Standard property of the Indirect Incentive Item.
    cWrkStd = cWrkStdIn
End Property
