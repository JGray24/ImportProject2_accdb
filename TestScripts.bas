Attribute VB_Name = "TestScripts"
Option Compare Database
Option Explicit

Function GexpTst()
Call KillFile("G:\My Drive\Joel's Files\Work\ImportProject2x\")
Call Gexp
End Function

Function EtlTst()
Dim Q:  Q = "G:\My Drive\Joel's Files\Work\Testing2Tabs.xlsx"
Dim r:  r = "G:\My Drive\Joel's Files\Work\Testing2Tabs.xlsx.config"
Dim EtlErrMsg As String

''''GblForceOnPrintDebugLog = True

Q = "G:\My Drive\Joel's Files\Work\ActualImportTest.xlsx"
Q = "G:\My Drive\Joel's Files\Work\NewTest.xlsx"
Q = "G:\My Drive\Joel's Files\Work\LogErrors\Testing2Tabs.xlsx"
Q = "G:\My Drive\Joel's Files\Work\LogErrors\ActTest.xlsm"
Q = "G:\My Drive\Joel's Files\Work\LogErrors\TestAtsMatchedToOne.xlsm"


'EtlErrMsg = Etl_123(q, , "Activity,Table")
'EtlErrMsg = Etl_123(q, , "Activity*,    Category,  ", "Testing2Tabs.xlsx/Category,Testing2Tabs.xlsx/NoHeadings")
'EtlErrMsg = Etl_123(q, "CreateConfig")
'EtlErrMsg = Etl_123(q, , "Activity2")

'EtlErrMsg = Etl_123(q, "None", , "*", , False, True)
'EtlErrMsg = Etl_123(Q, GetTesting2WkShtsConfig, , "*", , True, True)
'Call Etl_123(Q, GetActTestConfig, , , , False, True)
DelTbl ("ETL123_table_errors")

Call KillFile(Q & ".Config")
Call Etl_123(Q, , , , False, False, True)

Debug.Print ("Done")



End Function




Private Sub DelTbl(tblName As String)
Dim Errmsg    As String
Dim Response  '  Yes=6  No=7  Retry=4  OK=1  Cancel=2
Dim crlf:  crlf = vbCrLf & Chr(10)
   
On Error GoTo Error_Handler
   
Application.CurrentDb.TableDefs.Delete tblName
On Error GoTo 0
Exit Sub
   
Error_Handler:
  Select Case err.Number
  Case 3211
    Errmsg = "Error number: " & Str(err.Number) & vbNewLine & _
             "Source: " & err.source & vbNewLine & _
             "Description: " & err.Description
    Response = MsgBox(Errmsg, vbRetryCancel)
    Debug.Print (vbCrLf & "****" & Errmsg & vbCrLf)
    If Response = 4 Then Resume
    End
  Case Else
    Resume Next
    Resume  ' Extra Resume
  End Select
  
End Sub

Private Function KillFile(ByVal strFile As String) As Boolean

Dim ErrorHasOccured:  ErrorHasOccured = False
Dim aCMD As String
Dim SetAttrCmdFlag  As Boolean:  SetAttrCmdFlag = False
On Error GoTo Error_Handler

KillFile = True   ' Default to Successful Delete...
If Not FileOrDirExists(strFile) Then Exit Function

If Len(Dir$(strFile)) > 0 Then
    SetAttrCmdFlag = True
    SetAttr strFile, vbNormal   ' Attempt to set attribute to normal.
    SetAttrCmdFlag = False
    
    Kill strFile
    Exit Function
End If

GoTo EndOfRoutine   '  Error Handling should always be the last thing in a routine.
Dim Errmsg, errResponse ' Call Err.Raise(????) The range 513 - 65535 is available for user errors.
Error_Handler:
  ErrorHasOccured = True
  Errmsg = "Error number: " & Str(err.Number) & vbNewLine & _
           "Source: " & err.source & vbNewLine & _
           "Description: " & err.Description & vbCrLf & vbCrLf
  Select Case err.Number
  Case 75, 70
    On Error GoTo 0
    aCMD = "cmd /c erase """ & strFile & """"
    Shell aCMD, vbNormalFocus
    On Error GoTo Error_Handler
    Call WaitFor(3) ' Wait for 3 seconds.....
    Resume Next  ' To continue...
  Case 53
    If SetAttrCmdFlag Then Resume Next
    KillFile = False
    Exit Function
  Case Else
    Errmsg = Errmsg & "No specific Handling.. " & vbCrLf & vbCrLf & _
                      "Abort will launch standard error handling. (Use to Debug)" & vbCrLf & _
                      "Retry will Try Again." & vbCrLf & _
                      "Ignore will END the process/program."
    errResponse = MsgBox(Errmsg, vbAbortRetryIgnore)
    ' 3 Abort, 4 Retry, 5 Ignore
    Debug.Print (vbCrLf & "****" & Errmsg & vbCrLf)
    If errResponse = 3 Then
      On Error GoTo 0 ' Turn off error trap.
      Resume  ' Ignore To continue...
    End If
    If errResponse = 4 Then Resume  ' Retry....
    If errResponse = 5 Then
      Debug.Print ("Process Aborted by User")
      End     ' Ignore will end the process.
    End If
    Debug.Print ("Process Aborted by User")
    End
  End Select
Resume  ' Extra Resume for debug.  Locate source of the error..
EndOfRoutine:

End Function

Private Function FileOrDirExists(strDest As String) As Boolean
  Dim intLen As Long
  Dim fReturn As Boolean

  fReturn = False

  If strDest <> vbNullString Then
    On Error Resume Next
    intLen = Len(Dir$(strDest, vbDirectory + vbNormal))
    On Error GoTo PROC_ERR
    fReturn = (Not err And intLen > 0)
  End If

PROC_EXIT:
  FileOrDirExists = fReturn
  Exit Function

PROC_ERR:
  MsgBox "Error: " & err.Number & ". " & err.Description, , "FileOrDirExists"
  Resume PROC_EXIT
End Function


Private Sub WaitFor(ByVal Seconds As Long)
'  Wait for number of seconds to expire before returning.
Dim II, JJ
Dim aTimeValue   As String
Dim waitTill As Date

Dim Hrs As Long, Min As Long, sec As Long, xMod As Long

Hrs = Int(Seconds / 3600)
xMod = Seconds Mod 3600
Min = Int(xMod / 60)
sec = xMod Mod 60

Dim HH As String, MM As String, SS As String
HH = Hrs:  HH = Lpad(HH, 2)
MM = Min:  MM = Lpad(MM, 2)
SS = sec:  SS = Lpad(SS, 2)
Dim Interval   As String:  Interval = HH & ":" & MM & ":" & SS

If Seconds > 86399 Then
  Call MsgBox("WaitFor(" & Seconds & ") is more than max 24 hours." & vbCrLf & vbCrLf & _
              "Program will ABORT")
  End
End If

waitTill = Now() + TimeValue(Interval)  ' Wait Proper num of seconds for program to finish....
While Now() < waitTill
  DoEvents
Wend

End Sub

Private Function Lpad(ByVal strInput As String, ByVal NewLen As Long, Optional ByVal PadChar As String = "0") As String
If Len(PadChar) > 1 Then PadChar = Left(PadChar, 1)
If PadChar = "" Then PadChar = "0"
Lpad = strInput
Do
  If Len(Lpad) >= NewLen Then Exit Function
  Lpad = PadChar & Lpad
Loop
End Function

