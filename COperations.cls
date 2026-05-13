VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COperations"
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Attribute VB_Description = "Class of Work Order Operations"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarSeq_No  As String 'local copy
Private mvarWC_ID   As String 'local copy
Private mvarCC_ID   As String 'local copy
Private mvarOper_Cd As String 'local copy
Public Property Let Oper_Cd(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Oper_Cd = 5
    mvarOper_Cd = vData
End Property


Public Property Get Oper_Cd() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Oper_Cd
    Oper_Cd = mvarOper_Cd
End Property



Public Property Let CC_ID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CC_ID = 5
    mvarCC_ID = vData
End Property


Public Property Get CC_ID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CC_ID
    CC_ID = mvarCC_ID
End Property



Public Property Let WC_ID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WC_ID = 5
    mvarWC_ID = vData
End Property


Public Property Get WC_ID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WC_ID
    WC_ID = mvarWC_ID
End Property



Public Property Let Seq_No(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Seq_No = 5
    mvarSeq_No = vData
End Property


Public Property Get Seq_No() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Seq_No
    Seq_No = mvarSeq_No
End Property



