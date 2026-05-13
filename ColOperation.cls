VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColOperation"
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public Function Add(Seq_No As String, WC_ID As String, CC_ID As String, Oper_Cd As String, Optional sKey As Variant) As COperations
'create a new object
Dim objNewMember As COperations
Set objNewMember = New COperations


    'set the properties passed into the method
    objNewMember.Seq_No = Seq_No
    objNewMember.WC_ID = WC_ID
    objNewMember.CC_ID = CC_ID
    objNewMember.Oper_Cd = Oper_Cd

    If IsMissing(sKey) = True Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, Str$(sKey)
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As COperations

    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


#If Win32 Then
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property
#End If

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

