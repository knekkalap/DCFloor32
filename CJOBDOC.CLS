VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CJObDoc"
Attribute VB_Creatable = True
Attribute VB_Exposed = False


'Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const mstrModuleName = "CJobDoc"
'==============================================================================
'Name : John Canny
'Date : 06/01/98
'Comments : Class for Maintaining the Job Ticket Data

'==============================================================================

'local variable(s) to hold property value(s)
Private mvarintSeq_Num      As Single       'Job Sequence Number
Private mvarstrOrderNum     As String       'Job Work Order Number
Private mvarstrPartNum      As String       'Job Part Number
Private mvarstrWrkCtr       As String       'Job Work Center
Private mvarstrOperation    As String       'Job Operation
Private mvarlngRunStd       As Long         'Run Standard, Plan Qty
Private mvarcurSetStd       As Currency     'Set Standard
Private mvarstrEorSstd      As String       'E or S Standard
Private mvarstrPartNumber   As String       'Part Number



Public Property Let strPartNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strPartNumber = 5
    mvarstrPartNumber = vData
End Property


Public Property Get strPartNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strPartNumber
    strPartNumber = mvarstrPartNumber
End Property



Public Property Let strEorSstd(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strEorSstd = 5
    mvarstrEorSstd = vData
End Property


Public Property Get strEorSstd() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strEorSstd
    strEorSstd = mvarstrEorSstd
End Property



Public Property Let curSetStd(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.curSetStd = 5
    mvarcurSetStd = vData
End Property


Public Property Get curSetStd() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.curSetStd
    curSetStd = mvarcurSetStd
End Property



Public Property Let lngRunStd(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngRunStd = 5
    mvarlngRunStd = vData
End Property


Public Property Get lngRunStd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngRunStd
    lngRunStd = mvarlngRunStd
End Property



Public Property Let strOperation(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strOperation = 5
    mvarstrOperation = vData
End Property


Public Property Get strOperation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strOperation
    strOperation = mvarstrOperation
End Property



Public Property Let strWrkCtr(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strWrkCtr = 5
    mvarstrWrkCtr = vData
End Property


Public Property Get strWrkCtr() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strWrkCtr
    strWrkCtr = mvarstrWrkCtr
End Property



Public Property Let strPartNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strPartNum = 5
    mvarstrPartNum = vData
End Property


Public Property Get strPartNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strPartNum
    strPartNum = mvarstrPartNum
End Property



Public Property Let strOrderNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strOrderNum = 5
    mvarstrOrderNum = vData
End Property


Public Property Get strOrderNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strOrderNum
    strOrderNum = mvarstrOrderNum
End Property



Public Property Let intSeq_Num(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.intSeq_Num = 5
    mvarintSeq_Num = vData
End Property


Public Property Get intSeq_Num() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.intSeq_Num
    intSeq_Num = mvarintSeq_Num
End Property





