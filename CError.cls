VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrErrNum As String
Private mstrErrDesc As String
Private mstrMyErrDesc As String
Private mbooErrUnknown As Boolean
Private mstrModuleName As String
Private mstrFunctionName As String
Private mstrUserID As String
Private mstrComputerName As String
Private ErrorForm As FError

' Error constants
Const ERR_NULL_VALUE = 94
Const ERR_CRW_RPT_MISSING = 20504
Const ERR_LOGIN_FAILED = 32003
Const ERR_PRIMARY_KEY_EXISTS = -2147217873
Const ERR_OBJECT_UNLOADED = 364
Const ERR_ZERO = 0
Const ERR_RESUME = 20
Const ERR_TYPE_MISMATCH = 13
Const ERR_CANT_FIND_SQLSRVR = 20599
Const ERR_SQLSRVR = 20527
Const ERR_NOT_ENOUGH_MEM = 20500
Const ERR_INVALID_CALL = 5
Const ERR_CONNECT_STRING = -2147467259
Const ERR_DB_NOT_FOUND = -2147217843
Const ERR_PROVIDER = 3706

Public Property Get ErrNum() As String
    ErrNum = mstrErrNum
End Property

Public Property Let ErrNum(ByVal strErrNum As String)
    mstrErrNum = strErrNum
End Property

Public Property Get ErrDesc() As String
    ErrDesc = mstrErrDesc
End Property

Public Property Let ErrDesc(ByVal strErrDesc As String)
    mstrErrDesc = strErrDesc
End Property

Public Property Get MyErrDesc() As String
    MyErrDesc = mstrMyErrDesc
End Property

Public Property Let MyErrDesc(ByVal strMyErrDesc As String)
    mstrMyErrDesc = strMyErrDesc
End Property

Public Property Get ErrUnknown() As Boolean
    ErrUnknown = mbooErrUnknown
End Property

Public Property Let ErrUnknown(ByVal booErrUnknown As Boolean)
    mbooErrUnknown = booErrUnknown
End Property

Public Property Get ModuleName() As String
    ModuleName = mstrModuleName
End Property

Public Property Let ModuleName(ByVal strModuleName As String)
    mstrModuleName = strModuleName
End Property

Public Property Get FunctionName() As String
    FunctionName = mstrFunctionName
End Property

Public Property Let FunctionName(ByVal strfunctionname As String)
    mstrFunctionName = strfunctionname
End Property

Public Property Get UserID() As String
    UserID = mstrUserID
End Property

Public Property Let UserID(ByVal strUserID As String)
    mstrUserID = strUserID
End Property

Public Property Get ComputerName() As String
    ComputerName = mstrComputerName
End Property

Public Property Let ComputerName(ByVal strComputerName As String)
    mstrComputerName = strComputerName
End Property

Public Function ErrorRoutine() As Integer
 
    Dim strErrFile As String
    Dim strErrWrite As String
    Dim strDate As String
    Dim fso As Scripting.FileSystemObject
    Dim txtfile As TextStream
    
 Select Case ErrNum
    Case ERR_NULL_VALUE  'Null Value encountered
        MyErrDesc = "ERROR: NULL value encountered"
        ErrUnknown = False
        ErrorRoutine = CONTINUE
        
    Case ERR_CRW_RPT_MISSING 'Crystal RPT file is missing
        MyErrDesc = "ERROR: Crystal Report file is missing. Please contact the system administrator."
        ErrUnknown = False
        ErrorRoutine = ABORT
        
    Case ERR_LOGIN_FAILED 'MAPI login failed -- Outlook isn't running
        MyErrDesc = "ERROR: Error while attempting to send mail notifications.  Outlook must be started for mail to work.  Please start Outlook, and then retry your transaction."
        ErrUnknown = False
        ErrorRoutine = ABORT
        
    Case ERR_PRIMARY_KEY_EXISTS ' ID (primary key) already exists
        MyErrDesc = "ERROR: Error while attempting to save data.  The Primary Key Field already exisits.  Please retry your save."
        ErrUnknown = False
        ErrorRoutine = ABORT
        
    Case ERR_OBJECT_UNLOADED ' object was unloaded
        MyErrDesc = "ERROR: An error prevented the form from loading properly."
        ErrUnknown = False
        ErrorRoutine = CONTINUE
    
    Case ERR_ZERO ' form loading without showing
        MyErrDesc = "Resume without error."
        ErrUnknown = False   ' don't know why this error happens
        ErrorRoutine = CONTINUE
        
    Case ERR_RESUME ' resume without error
        MyErrDesc = "Resume without error."
        ErrUnknown = False  ' This error happens in response to ERR_ZERO
        ErrorRoutine = CONTINUE
        
    Case ERR_TYPE_MISMATCH ' type mismatch
        MyErrDesc = "ERROR: An invalid value type was passed to the server.  Please contact the MIS Help Desk at (860) 409 - 4690."
        ErrUnknown = False
        ErrorRoutine = ABORT
        
    Case ERR_CANT_FIND_SQLSRVR ' cannot find the SQL server (usually when printing)
        MyErrDesc = "ERROR: Cannot find SQL Server.  Please contact the MIS Help Desk at (860) 409 - 4690."
        ErrUnknown = False
        ErrorRoutine = ABORT
        
    Case ERR_SQLSRVR ' SQL server error
        MyErrDesc = "ERROR: An unspecified SQL Server error occurred.  Please contact the MIS Help Desk at (860) 409 - 4690."
        ErrUnknown = False
        ErrorRoutine = ABORT
        
    Case ERR_NOT_ENOUGH_MEM ' not enough memory (usually when printing)
        MyErrDesc = "ERROR: Not enough memory for operation.  Please try again.  If problems persist, please contact the MIS Help Desk at (860) 409 - 4690."
        ErrUnknown = False
        ErrorRoutine = ABORT
        
    Case ERR_INVALID_CALL ' invalid procedure call
        MyErrDesc = "ERROR: Couldn't set focus on control."
        ErrUnknown = False
        ErrorRoutine = CONTINUE
    
    Case ERR_CONNECT_STRING ' something wrong with ADO connection string
        MyErrDesc = "ERROR: There is a problem connecting to the SQL database.  First check the .ini file to see if connection parameters are correct."
        ErrUnknown = False
        ErrorRoutine = ABORT
        
    Case ERR_DB_NOT_FOUND ' database not found
        MyErrDesc = "ERROR: There is a problem connecting to the SQL database.  First check the .ini file to see if connection parameters are correct."
        ErrUnknown = False
        ErrorRoutine = ABORT
        
    Case ERR_PROVIDER ' problem with OLEDB provider
        MyErrDesc = "ERROR: There is a problem connecting to the SQL database.  First check the .ini file to see if connection parameters are correct."
        ErrUnknown = False
        ErrorRoutine = ABORT
        
    Case Else
        ErrDesc = "ERROR: " & Err.Description
        MyErrDesc = ""
        ErrUnknown = True
        ErrorRoutine = ABORT
        
  End Select
  
If ErrorRoutine = ABORT Then

    'Write Error to file
    strDate = Format(Now(), "mm/dd/yyyy - hh:nn AM/PM")
    strErrWrite = "Module: " & ModuleName & vbCrLf
    strErrWrite = strErrWrite & "Function: " & FunctionName & vbCrLf
    strErrWrite = strErrWrite & "User: " & UserID & vbCrLf
    strErrWrite = strErrWrite & "Computer: " & ComputerName & vbCrLf
    strErrWrite = strErrWrite & "Date: " & strDate & vbCrLf
    strErrWrite = strErrWrite & "Error #: " & ErrNum & vbCrLf
    strErrWrite = strErrWrite & "Error Desc.: " & ErrDesc & vbCrLf & vbCrLf
    strErrWrite = strErrWrite & "My Error Desc.: " & MyErrDesc & vbCrLf & vbCrLf
    Set fso = New Scripting.FileSystemObject
    ' check to make sure that drive exists
    If fso.DriveExists(Left(gErrLocation, 2)) Then
        strErrFile = gErrLocation
    Else
       strErrFile = "C:\" & App.EXEName & ".ERR"
    End If
    If Not fso.FileExists(strErrFile) Then
        fso.CreateTextFile (strErrFile)
    End If
    Set txtfile = fso.OpenTextFile(strErrFile, ForAppending, True)
    txtfile.WriteLine strErrWrite
    txtfile.Close
    Set fso = Nothing
    'MsgBox "                  **** A critical error has occurred ****                " & Chr(13) & _
    '       "Please notify your supervisor or call the MIS help desk at (860) 409-4690" & Chr(13) & _
    '       "   This pc should be restarted before attempting another DC transaction  " & Chr(13) & Chr(13) & _
    '       "                              Press Enter                                ", vbCritical + vbOKOnly, "Critical Error"
    'Call UnloadApp
 End If

End Function
Public Sub Clear()

    ErrNum = ""
    ErrDesc = ""
    MyErrDesc = ""
    ErrUnknown = False
    ModuleName = ""
    FunctionName = ""
    
End Sub

Private Sub Class_Initialize()
UserID = ap_GetUserName
ComputerName = ap_GetComputerName
End Sub
Function ap_GetUserName() As Variant
    Dim strUsername As String
    Dim lngLength As Long
    Dim lngResult As Long

    '-----Set up the Buffer
    strUsername = String$(255, 0)
    lngLength = 255
    '-----make the Call
    lngResult = wu_GetUserName(strUsername, lngLength)
    '------clean up and Assign the value
    'ap_GetUserName = strUsername
    ap_GetUserName = Left(strUsername, InStr(1, strUsername, _
    Chr(0)) - 1)
End Function
Function ap_GetComputerName() As Variant
    Dim strComputerName As String
    Dim lngLength As Long
    Dim lngResult As Long


    '----- Set up Buffer
    strComputerName = String$(255, 0)
    lngLength = 255


    '-----Make the Call
    lngResult = wu_GetComputerName(strComputerName, lngLength)


    '----Clean up and assign the value
    ap_GetComputerName = Left(strComputerName, InStr(1, strComputerName, _
    Chr(0)) - 1)
End Function

Public Sub ShowMsg(Optional ByVal bClear As Boolean = True)
'========================================================
'Name: Shari M. Fogel
'Date: 03/01/1999
'Comments:
'   Displays a message box notifying the user of the error
'
'Parameters:
'   bClear  =   optional parameter
'           =   says whether to clear error info after displaying message
'           =   default is true -- meaning clear the info
'
'========================================================
    
    Set ErrorForm = New FError
    
    ErrorForm.ErrorModuleName = ModuleName
    ErrorForm.ErrorFunctionName = FunctionName
    ErrorForm.ErrorMessage = ErrNum & vbCrLf & ErrDesc & vbCrLf & MyErrDesc
    ErrorForm.Show vbModal
    
    If bClear Then
        'reset error info
        Clear
    End If
'    Select Case gError.ErrorRoutine
'        Case CONTINUE
'            gError.Clear
'            Resume Next
'        Case ABORT
'            gError.ShowMsg (True)
'            GoTo Cleanup
'    End Select
'
   
End Sub

