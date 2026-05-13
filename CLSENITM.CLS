VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCodeItem"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Class Name       : clsCodeItem
'Description      : The class that is an object to represent a Code used
'                   in the Menu Entry or Combo Box Lookup sections of the
'                   frmDCEntry form.
'Properties       : Code, CodeDesc, Desc, Value
'Called From      : gcolIndirCodes
'Dependencies     : None
'==============================================================================

'The Private Properties of a Code Item
Private strCode As String   'Just the Code
Private strDesc As String   'Just the Description
Private strCodeDesc As String   'The Code and Description
Private vValue As Variant   'An assigned value to an item
Property Get Code() As String
    'Expose the Code property of the Batch Incentive Item.
    Code = strCode
End Property

Property Get Desc() As String
    'Expose the Desc property of the Batch Incentive Item.
    Desc = strDesc
End Property


Property Get CodeDesc() As String
    'Expose the CodeDesc property of the Batch Incentive Item.
    'Has both the Code and description returned as a string
    CodeDesc = strCodeDesc
End Property


Property Let Code(strDesc As String)
    'Expose the Code property of the Batch Incentive Item.
    strCode = strDesc
End Property

Property Let Desc(strDescIn As String)
    'Expose the Desc property of the Batch Incentive Item.
    strDesc = strDescIn
End Property

Property Let CodeDesc(strCodeDescIn As String)
    'Expose the CodeDesc property of the Batch Incentive Item.
    'Sets both the Code and description
    strCodeDesc = strCodeDescIn
End Property


Private Sub Class_Initialize()
    strCode = ""
    strDesc = ""
    strCodeDesc = ""
    vValue = Null
End Sub



Public Property Get Value() As Variant
    'Expose the Value property of the Batch Incentive Item.
    Value = vValue
End Property

Public Property Let Value(vValueIn As Variant)
    'Expose the Value property of the Batch Incentive Item.
    vValue = vValueIn
End Property

