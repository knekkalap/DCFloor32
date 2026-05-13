Attribute VB_Name = "ERROR"
Option Explicit

'Const ErrFile = "ERROR.LOG"
Public gstrLastOper As String

Public Sub ErrBox(inp_msg As String, msgtype As Integer, Optional msgduration As Integer)
   Dim start As Variant
   Dim dispduration As Integer
   
   On Error GoTo ErrBoxError
   
 Select Case msgtype
    Case 1
           Screen.ActiveForm.lblerrmsg.ForeColor = &H4040&
           Screen.ActiveForm.lblerrmsg.FontSize = 14
   Case 2
           Screen.ActiveForm.lblerrmsg.ForeColor = &HFF&
           Screen.ActiveForm.lblerrmsg.FontSize = 14
   Case 3
           Screen.ActiveForm.lblerrmsg.ForeColor = &HFF&
           Screen.ActiveForm.lblerrmsg.FontSize = 28
   Case 4
           Screen.ActiveForm.lblerrmsg.ForeColor = &HFF&
           Screen.ActiveForm.lblerrmsg.FontSize = 28
   Case Else
           Screen.ActiveForm.lblerrmsg.ForeColor = &H4040&
           Screen.ActiveForm.lblerrmsg.FontSize = 14
   End Select
   Screen.ActiveForm.lblerrmsg.Caption = inp_msg
   
   If msgduration > 0 Then
      dispduration = msgduration
   Else
      dispduration = gErrBoxDuration
   End If

   start = Timer
   Do While Timer <= start + dispduration
'  Beep
   Loop
   
   Screen.ActiveForm.lblerrmsg.Caption = ""
   Exit Sub
   
'Section to handle forms without lblErrMsg label.
ErrBoxError:
    MsgBox LoadResString(gLanguageOffset + 46) & inp_msg
End Sub

Public Sub WriteAttdLog(ByVal strUser As String, ByVal strEmpNo As String, ByVal strAttdSel As String)
'Write strMsg to Error Log
Dim strAttdLog     As String
Dim fso As Scripting.FileSystemObject
Dim txtfile As TextStream
  
 strAttdLog = "Attendance transaction: " & strAttdSel & "  Employee No: " & strEmpNo & _
                       "  DC station: " & gError.ComputerName & "  DateTime: " & Now

   Set fso = New Scripting.FileSystemObject
    ' check to make sure that drive exists
   If fso.FileExists(gAttdLogFile) Then
        Set txtfile = fso.OpenTextFile(gAttdLogFile, ForAppending, True)
        txtfile.WriteLine strAttdLog
        txtfile.Close
   End If
   Set fso = Nothing

'Write string to local file also, in case of network problems
    Set fso = New Scripting.FileSystemObject
    If Not fso.FolderExists("C:\AttdLog\") Then
          fso.CreateFolder ("C:\AttdLog\")
    End If
    
    Set txtfile = fso.OpenTextFile(gAttdLogLocal, ForAppending, True)
    txtfile.WriteLine strAttdLog
    txtfile.Close
    Set fso = Nothing

End Sub





