VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBIItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Class Name       : clsBIItem
'Description      : Class to contain one Batch Incentive code.
'Properties       : Code, CodeDesc, Desc, IncMin, Method, WCID
'Called From      : clsAllBITable.AddBIItem, clsAllBITable.GetBIItem,
'                   clsWCBITable.AddBIItem, clsWCBITable.GetBIItem,
'                   DBCommon.LdBatchIncn
'Dependencies     : None
'==============================================================================

' The Private Properties of a Batch Incentive Item
Private strCode As String
Private strDesc As String
Private strCodeDesc As String
Private strWCID As String
Private siMethod As Byte
Private siIncMin As Byte


Property Get Code() As String
    'Expose the Code property of the Batch Incentive Item.
    Code = strCode
End Property


Property Let Code(strDesc As String)
    'Expose the Code property of the Batch Incentive Item.
    strCode = strDesc
End Property

Property Get CodeDesc() As String
    'Expose the CodeDesc property of the Batch Incentive Item.
    'Has both the Code and description returned as a string
    CodeDesc = strCodeDesc
End Property



Property Let CodeDesc(strCodeDescIn As String)
    'Expose the CodeDesc property of the Batch Incentive Item.
    'Sets both the Code and description
    strCodeDesc = strCodeDescIn
End Property



Property Get Desc() As String
    'Expose the Desc property of the Batch Incentive Item.
    Desc = strDesc
End Property


Property Let Desc(strDescIn As String)
    'Expose the Desc property of the Batch Incentive Item.
    strDesc = strDescIn
End Property



Public Property Get WCID() As String
    'Expose the WDID property of the Batch Incentive Item.
    WCID = strWCID
End Property

Public Property Let WCID(strWCIDIn As String)
    'Expose the WCID property of the Batch Incentive Item.
    strWCID = strWCIDIn
End Property

Public Property Get Method() As Byte
    'Expose the Method property of the Batch Incentive Item.
    Method = siMethod
End Property

Public Property Let Method(siMethodIn As Byte)
    'Expose the Method property of the Batch Incentive Item.
    siMethod = siMethodIn
End Property

Public Property Get IncMin() As Byte
    'Expose the IncMin property of the Batch Incentive Item.
    IncMin = siIncMin
End Property

Public Property Let IncMin(siIncMinIn As Byte)
    'Expose the IncMin property of the Batch Incentive Item.
    siIncMin = siIncMinIn
End Property

