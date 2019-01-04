Attribute VB_Name = "GitSupport"
Option Compare Database
Option Explicit

Dim ThisModuleName   As String

Dim GblLastFilePickerPath  As String
Private Const GitHubRemoteAdd As String = "https://github.com/JGray24/REPOSITORY.git"
Private Const GitHubPush As String = "git push -u origin master"
'https://github.com/JGray24/ImportProject2.git
'git remote add origin2 https://github.com/JGray24/ImportProject2.git
'https://github.com/USERNAME/REPOSITORY.git


'''''''''''''''''''''''''''
' Error Number Constants
'''''''''''''''''''''''''''
Private Const C_ERR_NO_ERROR = 0&
Private Const C_ERR_SUBSCRIPT_OUT_OF_RANGE = 9&
Private Const C_ERR_ARRAY_IS_FIXED_OR_LOCKED = 10&


Public Function Gexp(Optional ByVal SaveReminder As Boolean = True)
'  Git Export
Dim I  As Long

'  First locate the GIT folder that is located in the same folder as the Office project.
Dim GitProjectFolderName  As String

Dim msg, Response
If SaveReminder Then
  msg = "Did you manually SAVE your MS/Office project?" & vbCrLf & vbCrLf & "  (Required for good results)"
  Response = MsgBox(msg, vbYesNo, "Save Reminder Question?")
  If Response = vbNo Then
    Call MsgBox("After SAVE is complete, please try again.", , "Save Reminder Question?")
    End
  End If
End If

GitProjectFolderName = Application.CurrentProject.Path & "\" & GitProjectName & "\"

If Not FileOrDirExists(GitProjectFolderName) Then _
  Call EnsureProjectFolderExists(GitProjectName) ' Make sure that project is either initialized or cloned.

Call ExportSourceFiles(GitProjectFolderName)  ' Push VBA Source Files into Git Project Folder

Call InitGitIgnore(GitProjectFolderName)  ' Initialize GitIgnore file if not already existing.
Call InitReadMe(GitProjectFolderName)     ' Initialize ReadMe file if not already existing.
Call InitGitCommit(GitProjectFolderName)  ' Initialize GitCommit batch file if not already existing.
Call InitGitPush(GitProjectFolderName)    ' Initialize GitPush batch file if not already existing.

'Call CreateGitBatFile(GitProjectFolderName)  ' Create and launch the Git "1 Gexp-VBA-Git.bat" file to execute the export process.
Call MsgBox("Git VBA Export to: " & GitProjectFolderName & vbCrLf & _
            "  is COMPLETE..." & vbCrLf & vbCrLf & "GitCommit.bat and GitPush.bat can now be run." & vbCrLf & vbCrLf & _
            "Look for instructions in the Immediate window...")
            
Debug.Print ("Gexp (Git Export) has completed capturing all of the individual VBA files " & vbCrLf & "   to the """ & Application.CurrentProject.Name & """ project folder." & vbCrLf & vbCrLf & _
             "To run the GIT process to Commit files to the Repo, copy and paste (including quotes)" & vbCrLf & "   the respective batch files to run from Windows CMD prompt:" & vbCrLf & _
             """" & GitProjectFolderName & ".GitCommit.bat""  --> Commit to Local Repo." & vbCrLf & _
             """" & GitProjectFolderName & ".GitPush.bat""    --> Push to the Remote Repo.")


End Function

Public Function Gimp()   '  Git Import all modules from the project directory.

ThisModuleName = "GitSupport"  ' Name of this module....

'  First locate the GIT folder that is located in the same folder as the Office project.
Dim GitProjectFolderName  As String
Dim I  As Long
GitProjectFolderName = Application.CurrentProject.Name
I = InStrRev(GitProjectFolderName, ".")
If I > 1 Then GitProjectFolderName = Left(GitProjectFolderName, I - 1)
GitProjectFolderName = Application.CurrentProject.Path & "\" & GitProjectFolderName & "\"
  
Call ImportSourceFiles
Call MsgBox("Git VBA Import to: " & vbCrLf & vbCrLf & GitProjectFolderName & vbCrLf & vbCrLf & "   is COMPLETE...")
End Function


Private Sub ImportSourceFiles()

Dim I   As Long
Dim ModuleName   As String

'  First locate the GIT folder that is located in the same folder as the Office project.
Dim GitProjectFolderName  As String
GitProjectFolderName = Application.CurrentProject.Name
I = InStrRev(GitProjectFolderName, ".")
If I > 1 Then GitProjectFolderName = Left(GitProjectFolderName, I - 1)
GitProjectFolderName = Application.CurrentProject.Path & "\" & GitProjectFolderName & "\"

Dim file As String
file = Dir(GitProjectFolderName)
While (file <> vbNullString)
  Call RemoveAModule(file)
  I = InStrRev(file, ".")
  If I > 1 Then ModuleName = Left(file, I - 1)
  If ThisModuleName <> ModuleName Then _
    Application.VBE.ActiveVBProject.VBComponents.Import GitProjectFolderName & file
  file = Dir
Wend

End Sub

'Private Sub RemoveAllModules()
'Dim project As VBProject
'Set project = Application.VBE.ActiveVBProject
'
'Dim comp As VBComponent
'For Each comp In project.VBComponents
'  If Not comp.Name = GitProjectFolderName And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
'    project.VBComponents.Remove comp
'  End If
'Next
'End Sub


Private Sub RemoveAModule(ByVal ModuleName As String)
Dim project As VBProject
Dim I   As Long

I = InStrRev(ModuleName, ".")
If I > 1 Then ModuleName = Left(ModuleName, I - 1)

If ModuleName = ThisModuleName Then Exit Sub  '  Don't remove this module.

Set project = Application.VBE.ActiveVBProject
 
Dim comp As VBComponent
For Each comp In project.VBComponents
  If comp.Name = ModuleName And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
    project.VBComponents.Remove comp
    Exit Sub
  End If
Next
End Sub


Public Sub ExportSourceFiles(destPath As String)
 
Dim component As VBComponent
Dim KillFileAndPath  As String

For Each component In Application.VBE.ActiveVBProject.VBComponents
  If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
    KillFileAndPath = destPath & component.Name & ToFileExtension(component.Type)
    If Len(Dir$(KillFileAndPath)) > 0 Then Call Kill(KillFileAndPath)
    component.Export destPath & component.Name & ToFileExtension(component.Type)
  End If
Next

End Sub
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
Select Case vbeComponentType
Case vbext_ComponentType.vbext_ct_ClassModule
ToFileExtension = ".cls"
Case vbext_ComponentType.vbext_ct_StdModule
ToFileExtension = ".bas"
Case vbext_ComponentType.vbext_ct_MSForm
ToFileExtension = ".frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner
Case vbext_ComponentType.vbext_ct_Document
Case Else
ToFileExtension = vbNullString
End Select
 
End Function


Private Function FileSave(ByVal OutpFileData As String, _
                  ByVal FilePath As String, _
                  Optional ByVal BoxTitle As String = "", _
                  Optional ByVal AppendFlag As Boolean = False) As String
                  
' This routine will set up a new file or update/replace data to an existing file.
' OutpFileData will contain the actual data that will be written to the file.
' FilePath will contain the fully qualified path and file (New or Pre-Existing).
'     Ex: "G:\My Drive\Joel's Files\Work\JohnsTestData.txt"
' FilePath can also supply a pattern for a new file name that will be used to validate user picked name.
'     Ex: "G:\My Drive\Joel's Files\Work\*.csv, *.txt, *.config"

' 1) If the file already exists

' If the file/path does not exist, and the file name is not valid, then the routine will attempt to
'   set up the new file.
'   If the new file name needs to be set by the user, the file name will be a list of patterns that will
'   be used to allocate new file, and validate that the new name matches one of the patterns.
'
'

If BoxTitle <> "" Then BoxTitle = BoxTitle & " / "
Dim MsgBoxTitle:  MsgBoxTitle = BoxTitle & "FileSave Routine"
On Error GoTo Error_Handler

Dim fso As FileSystemObject
Set fso = New FileSystemObject
Dim FileReadOnly
Dim FileStream As TextStream
Dim OrigData As String, II
Dim Err52Ctr As Long:  Err52Ctr = 0
Dim Err62FileEmpty As Boolean
Dim Err5FileNameInvalid As Boolean
Dim Err76FolderPathInvalid As Boolean
Dim TrimChr  As String:  TrimChr = Chr(9) & Chr(10) & Chr(11) & Chr(12) & Chr(13) & Chr(32)
Dim aOld, aNew, iFileNumber As Integer

If Not AppendFlag And FileOrDirExists(FilePath) Then
  aOld = TrimChars(GetEntireFile(FilePath), TrimChr)
  aNew = TrimChars(OutpFileData, TrimChr)
  If aOld = aNew Then GoTo FileSaveExit   ' Data is not changed.
  Call KillFile(FilePath)
End If

If FileOrDirExists(FilePath) Then
  iFileNumber = FreeFile                   ' Get unused file number
  Open FilePath For Append As #iFileNumber    ' Connect to the file
  Print #iFileNumber, OutpFileData            ' Append our string
  Close #iFileNumber                       ' Close the file
  GoTo FileSaveExit
End If

' Here the actual file is created and opened for write access
Err76FolderPathInvalid = False
Err5FileNameInvalid = False
Dim FileAndPathAreGood As Boolean:  FileAndPathAreGood = False
Set FileStream = fso.CreateTextFile(FilePath)
If Not Err76FolderPathInvalid And Not Err5FileNameInvalid Then FileAndPathAreGood = True

' If there is any problem using the new file/path then the user will be prompted to select a name.
If Not FileAndPathAreGood Then
    FilePath = DiaglogSaveAs(FilePath, False, MsgBoxTitle)
    If FilePath <> "" Then Set FileStream = fso.CreateTextFile(FilePath)
    If FilePath = "" Then
      Call MsgBox("File was not saved..." & vbCrLf & vbCrLf & "Process will END", vbOKCancel, MsgBoxTitle)
      End
    End If
End If


' Write Data to the file
FileStream.WriteLine OutpFileData

' Close it, so it is not locked anymore
FileStream.Close

FileSaveExit:
  Set FileStream = Nothing
  Set fso = Nothing
  FileSave = FilePath

GoTo EndOfRoutine   '  Error Handling should always be the last thing in a routine.
Dim Errmsg, errResponse ' Call Err.Raise(????) The range 513 - 65535 is available for user errors.
Error_Handler:
  Errmsg = "Error number: " & Str(err.Number) & vbNewLine & _
           "Source: " & err.source & vbNewLine & _
           "Description: " & err.Description & vbCrLf & vbCrLf
  Select Case err.Number
  
  Case 62
    Err62FileEmpty = True
    Resume Next
  Case 5   ' Invalid file name used.  need to invoke SaveAs process to select a new file.
    Err5FileNameInvalid = True
    Resume Next
  Case 76   ' Folder Name is Invalid. Need to cleanse the path and invoke SaveAs process to select a new file.
    Err76FolderPathInvalid = True
    Resume Next
  Case 52   ' Folder Name is Invalid. Need to cleanse the path and invoke SaveAs process to select a new file.
    Err52Ctr = Err52Ctr + 1  '  Count the number of times we had this occur.
    If Err52Ctr > 5 Then GoTo GeneralErrResponse
    Resume
  Case 53
    Errmsg = Errmsg & "My custom message...   Cancel will continue with next statement.."
    errResponse = MsgBox(Errmsg, vbRetryCancel, MsgBoxTitle)
    Debug.Print (vbCrLf & "****" & Errmsg & vbCrLf)
    If errResponse = 4 Then
      On Error GoTo 0  ' Turn off error handling.
      Resume  ' Retry....
    End If
    Resume Next  ' To continue...
  Case Else
    GoTo GeneralErrResponse
  End Select

GeneralErrResponse:
Errmsg = Errmsg & "No specific Handling.. " & vbCrLf & vbCrLf & _
                  "Abort will launch standard error handling. (Use to Debug)" & vbCrLf & _
                  "Retry will Try Again." & vbCrLf & _
                  "Ignore will END the process/program."
errResponse = MsgBox(Errmsg, vbAbortRetryIgnore, MsgBoxTitle)
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


Private Function GetEntireFile(ByVal FileName As String, Optional ByVal MsgBoxTitle = "GetEntireFile Error Handling", _
                      Optional ByVal ErrorMsgReturn As String = "") As String

Dim textData As String, fileNo As Integer
'fileName = "G:\My Drive\Joel's Files\VA Services Design\VA Services Data Repository\Normalized Data Tables\WorkingOnFiduciaryAccess\ImportSpecTest.txt"
fileNo = FreeFile 'Get first free file number

On Error GoTo Error_Handler
 
Open FileName For Input As #fileNo
GetEntireFile = Input$(LOF(fileNo), fileNo)
Close #fileNo
Exit Function

Dim Errmsg, Response
Error_Handler:
  Errmsg = "Error number: " & Str(err.Number) & vbNewLine & _
           "Source: " & err.source & vbNewLine & _
           "Description: " & err.Description & vbCrLf & vbCrLf
  Select Case err.Number
  Case 53
    Errmsg = Errmsg & FileName
    ErrorMsgReturn = Errmsg
    GetEntireFile = ""
    Exit Function
  Case Else
    Resume Next
    Resume ' Extra Resume
  End Select

End Function


Private Function TrimChars(ByVal XX, _
                   Optional ByVal chars As String = " ", _
                   Optional ByVal LTrimX As Boolean = True, _
                   Optional ByVal RTrimX As Boolean = True) As String

' Similar operation to Trim() function, but can also be used to
' trim any characters (in addition to blanks) from the input string.
Dim I, xStart, xEnd

If chars = "" Or Len(XX) = 0 Then
  TrimChars = XX
  If LTrimX Then TrimChars = LTrim(TrimChars)
  If RTrimX Then TrimChars = RTrim(TrimChars)
  Exit Function
End If
chars = chars & " "  ' Always trim blanks....

xStart = 1
If LTrimX Then
  For I = 1 To Len(XX)
    xStart = I
    If InStr(1, chars, Mid(XX, I, 1)) = 0 Then Exit For
  Next I
End If

xEnd = Len(XX)
If RTrimX Then
  For I = Len(XX) To 1 Step (-1)
    xEnd = I
    If InStr(1, chars, Mid(XX, I, 1)) = 0 Then Exit For
  Next I
End If
If xEnd < xStart Then
  TrimChars = ""
  Else
    TrimChars = Mid(XX, xStart, (xEnd - xStart) + 1)
End If

If Len(TrimChars) = 1 And InStr(1, chars, TrimChars) <> 0 Then TrimChars = ""

End Function

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


Private Function DiaglogSaveAs(ByVal OutpFileName As String, _
                       Optional ByVal AddCreateDateTimeStamp As Boolean = False, _
                       Optional ByVal BoxTitle) As String

Dim MsgBoxTitle:  MsgBoxTitle = BoxTitle & " / DiaglogSaveAs Routine"
On Error GoTo Error_Handler
 
' Filter  can be multiples comma separated.  Example:  "*.xlsx, *.csv, *.txt"

DiaglogSaveAs = ""

' The CleansePath will make sure that the for saving the file is good.  If directory and/or
' sub-directories need to be allocated, then this is handled in the function.
OutpFileName = CleansePath(OutpFileName, "DiaglogSaveAs ", True)
                      
Dim myObj As FileDialog
Dim HoldShow, II

Dim StartingFolderPath, FilePatternsFilters
Dim HoldPath, HoldPatterns As String
Dim Patterns()  As String
II = InStrRev(OutpFileName, "\")
If II <> 0 Then
  HoldPath = Mid(OutpFileName, 1, II)
  HoldPatterns = TrimChars(Mid(OutpFileName, II + 1), ",")
End If

If OutpFileName <> "" Then Patterns = Split(HoldPatterns, ",")
StartingFolderPath = HoldPath & "*.*"  ' Default
If IsArrayAllocated(Patterns) Then
  Dim ValidPatterns:  ValidPatterns = ""
  For II = LBound(Patterns) To UBound(Patterns)
    If ValidPatterns <> "" Then ValidPatterns = ValidPatterns & "  OR  "
    ValidPatterns = ValidPatterns & Patterns(II)
  Next II
  MsgBoxTitle = MsgBoxTitle & " / Valid Pattern(s): " & ValidPatterns
  If Len(MsgBoxTitle) > 127 Then ' Make sure title is not too long...
    MsgBoxTitle = Right(MsgBoxTitle, 124)
    MsgBoxTitle = "..." & MsgBoxTitle
  End If
End If
If IsArrayAllocated(Patterns) Then StartingFolderPath = HoldPath & Patterns(0)

If GblLastFilePickerPath = "" Then GblLastFilePickerPath = CurrentProject.Path
If StartingFolderPath = "" Then StartingFolderPath = GblLastFilePickerPath & "\"

''''Set myObj = Application.FileDialog(msoFileDialogSaveAs)
    Set myObj = Application.FileDialog(2)
TryAgain:
    myObj.InitialFileName = StartingFolderPath
    myObj.AllowMultiSelect = False
    myObj.Title = MsgBoxTitle
       
    HoldShow = myObj.Show
    ' -1 when file selected.
    ' 0 when cancel is pressed.
    If HoldShow = 0 Then
      DiaglogSaveAs = ""
      GoTo ExitSubroutine
    End If
    If HoldShow = -1 Then DiaglogSaveAs = myObj.SelectedItems(1)

Dim HoldOnlyFileName:  HoldOnlyFileName = ""
II = InStrRev(DiaglogSaveAs, "\")
If II <> 0 Then HoldOnlyFileName = Mid(DiaglogSaveAs, II + 1)

' Match to all existing patterns for a good match.
Dim AtLeastOneGoodPatternMatch As Boolean:  AtLeastOneGoodPatternMatch = False
If IsArrayAllocated(Patterns) Then
  For II = LBound(Patterns) To UBound(Patterns)
    If HoldOnlyFileName Like Patterns(II) Then AtLeastOneGoodPatternMatch = True
  Next II
  Dim Response, ErrMessage
  If Not AtLeastOneGoodPatternMatch Then
    ErrMessage = "File Name Picked - """ & HoldOnlyFileName & """ does not match any of the patterns:" & vbCrLf & vbCrLf
    For II = LBound(Patterns) To UBound(Patterns)
      ErrMessage = ErrMessage & Patterns(II) & ",   "
    Next II
    ErrMessage = TrimChars(ErrMessage, "," & vbCrLf) & vbCrLf & vbCrLf & _
             "File name must match the pattern(s) to be valid." & vbCrLf & _
             "To save the file, please try again."
    Response = MsgBox(ErrMessage, vbOKCancel, MsgBoxTitle)
    If Response = vbCancel Then End
    GoTo TryAgain
  End If
End If

' Save File Path for next time...
If DiaglogSaveAs <> "" Then
  II = InStrRev(DiaglogSaveAs, "\")
  If II <> 0 Then GblLastFilePickerPath = Left(DiaglogSaveAs, II)
End If

' Right Here add date and time stamp...
If Not FileOrDirExists(DiaglogSaveAs) And AddCreateDateTimeStamp Then _
  DiaglogSaveAs = AddDateTimeToFileName(DiaglogSaveAs)

ExitSubroutine:
  Set myObj = Nothing

GoTo EndOfRoutine   '  Error Handling should always be the last thing in a routine.
Dim Errmsg, errResponse ' Call Err.Raise(????) The range 513 - 65535 is available for user errors.
Error_Handler:
  Errmsg = "Error number: " & Str(err.Number) & vbNewLine & _
           "Source: " & err.source & vbNewLine & _
           "Description: " & err.Description & vbCrLf & vbCrLf
  Select Case err.Number
  Case 53
    Errmsg = Errmsg & "My custom message...   Cancel will continue with next statement.."
    errResponse = MsgBox(Errmsg, vbRetryCancel, MsgBoxTitle)
    Debug.Print (vbCrLf & "****" & Errmsg & vbCrLf)
    If errResponse = 4 Then
      On Error GoTo 0  ' Turn off error handling.
      Resume  ' Retry....
    End If
    Resume Next  ' To continue...
  Case Else
    Errmsg = Errmsg & "No specific Handling.. " & vbCrLf & vbCrLf & _
                      "Abort will launch standard error handling. (Use to Debug)" & vbCrLf & _
                      "Retry will Try Again." & vbCrLf & _
                      "Ignore will END the process/program."
    errResponse = MsgBox(Errmsg, vbAbortRetryIgnore, MsgBoxTitle)
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

Private Function CleansePath(ByVal DirectoryAndFile As String, _
                     Optional ByVal BoxTitle As String = "", _
                     Optional ByVal WhenMissingCreateFolders As Boolean = False, _
                     Optional ByVal WhenMissingRemoveFoldersFromPath As Boolean = False) As String

Dim MsgBoxTitle:  MsgBoxTitle = TrimChars(BoxTitle & " / CleansePath", "/")
Dim ErrorHasOccured:  ErrorHasOccured = False
On Error GoTo Error_Handler
Dim Errmsg: Errmsg = ""
Dim ErrResp

Dim II: II = 0
Dim JJ: JJ = 0

' This routine will first validate the path as valid.  If the path is valid then the file/path
' string is returned unchanged.

' When a bad path is detected, a MsgBox is displayed that will inform the user about the folders
' in the path that are not found.  User options will be:
'
'
' Path & File: "G:\My Drive\Joel's Files\Work\MyTest\MySubFolder\Johns.txt" has invalid Path.
'
' Missing Folders found in the path: "...\MyTest\MySubFolder\"
'
' Do you want these missing folders created automatically?
'
' Yes    - Missing folders will be created.  Input Path/File will be returned unchanged.
' No     - Missing folders will removed from the Input Path/File and returned.
' Cancel - Missing folders will NOT be created, and Input Path/File will be returned unchanged.
'
' a - shorten the path by removing the invalid parts to the right.
' b - set up the additional folder and subfolders that will be needed so that the original path be valid.
    '  BankPath = "cmd /c mkdir """ & HoldBankAcctFolderPath & """"
     ' Shell (BankPath) ' Set up Bank Account Folder.....
     
         
Dim HoldPath As String, HoldFileName As String

II = InStrRev(DirectoryAndFile, "\")
If II = 0 Then
  Call err.Raise(603, "CleansePath", """\"" is missing from a qualified Path/File:" & vbCrLf & DirectoryAndFile)
  Debug.Print ("Process Aborted by User")
  End  ' Half Baked..
End If

HoldPath = Left(DirectoryAndFile, II)
HoldFileName = Mid(DirectoryAndFile, II + 1)

Dim aFolders()  As String
aFolders = Split(TrimChars(HoldPath, "\"), "\")

Dim MissingFolders As String, PossibleValidPath As String, ValidPath As String, ThisFolder As String

'  Find folder names that are missing in the directory file system....
JJ = -1
If Not FileOrDirExists(HoldPath) Then
  For II = LBound(aFolders) To UBound(aFolders)
    ThisFolder = aFolders(II)
    PossibleValidPath = PossibleValidPath & ThisFolder & "\"
    If FileOrDirExists(PossibleValidPath) Then
      ValidPath = PossibleValidPath
      GoTo NextII
    End If
    MissingFolders = MissingFolders & "\" & ThisFolder
    If JJ = -1 Then JJ = II
NextII:
  Next II
End If

' Optional ByVal WhenMissingCreateFolders As Boolean,
' Optional ByVal WhenMissingRemoveFoldersFromPath As Boolean = False) As String
' When
If WhenMissingCreateFolders = WhenMissingRemoveFoldersFromPath Then
  Errmsg = "Path & File:  """ & DirectoryAndFile & """ has invalid Path." & vbCrLf & vbCrLf & _
           "Missing Folders found in the path: ""..." & MissingFolders & "\""" & vbCrLf & vbCrLf & _
           "Do you want these missing folders created AUTOMATICALLY?" & vbCrLf & vbCrLf & _
           "YES   - Missing folders WILL be created. Orig Input Path/File returned." & vbCrLf & _
           "NO    - Missing folders NOT created. CLEANSED Input Path/File returned." & vbCrLf & _
           "CANCEL - Missing folders NOT created. Original Input Path/File returned."
  ErrResp = MsgBox(Errmsg, vbYesNoCancel, MsgBoxTitle)
  WhenMissingCreateFolders = False
  WhenMissingRemoveFoldersFromPath = False
  If ErrResp = vbYes Then WhenMissingCreateFolders = True
  If ErrResp = vbNo Then WhenMissingRemoveFoldersFromPath = True
End If

If Not WhenMissingCreateFolders And Not WhenMissingRemoveFoldersFromPath Then
  ' Cancel - Missing folders will NOT be created, and Input Path/File will be returned unchanged.
  CleansePath = DirectoryAndFile
  Exit Function
End If

If WhenMissingRemoveFoldersFromPath Then
  ' No     - Missing folders will removed from the Input Path/File and returned.
  CleansePath = ValidPath & HoldFileName
  Exit Function
End If

' Yes    - Missing folders will be created.  Input Path/File will be returned unchanged.
'MkDir DirectoryPath

For II = JJ To UBound(aFolders)
  ValidPath = ValidPath & aFolders(II) & "\"
  MkDir ValidPath
Next II
CleansePath = DirectoryAndFile


GoTo EndOfRoutine   '  Error Handling should always be the last thing in a routine.
Dim errResponse ' Call Err.Raise(????) The range 513 - 65535 is available for user errors.
Error_Handler:
  ErrorHasOccured = True
  Errmsg = "Error number: " & Str(err.Number) & vbNewLine & _
           "Source: " & err.source & vbNewLine & _
           "Description: " & err.Description & vbCrLf & vbCrLf
  Select Case err.Number
  Case 603
    Errmsg = Errmsg & "My custom message...   Cancel will continue with next statement.."
    errResponse = MsgBox(Errmsg, vbOK, MsgBoxTitle)
    Debug.Print (vbCrLf & "****" & Errmsg & vbCrLf)
    On Error GoTo 0  ' Turn off error handling.
    Resume  ' Retry....
  Case 53
    Errmsg = Errmsg & "My custom message...   Cancel will continue with next statement.."
    errResponse = MsgBox(Errmsg, vbRetryCancel, MsgBoxTitle)
    Debug.Print (vbCrLf & "****" & Errmsg & vbCrLf)
    If errResponse = 4 Then
      On Error GoTo 0  ' Turn off error handling.
      Resume  ' Retry....
    End If
    Resume Next  ' To continue...
  Case Else
    Errmsg = Errmsg & "No specific Handling.. " & vbCrLf & vbCrLf & _
                      "Abort will launch standard error handling. (Use to Debug)" & vbCrLf & _
                      "Retry will Try Again." & vbCrLf & _
                      "Ignore will END the process/program."
    errResponse = MsgBox(Errmsg, vbAbortRetryIgnore, MsgBoxTitle)
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

Private Function IsArrayAllocated(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayAllocated
' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
' allocated.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is just the reverse of IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim N As Long
On Error Resume Next

' if Arr is not an array, return FALSE and get out.
If IsArray(Arr) = False Then
    IsArrayAllocated = False
    Exit Function
End If

' Attempt to get the UBound of the array. If the array has not been allocated,
' an error will occur. Test Err.Number to see if an error occurred.
N = UBound(Arr, 1)
If (err.Number = 0) Then
    ''''''''''''''''''''''''''''''''''''''
    ' Under some circumstances, if an array
    ' is not allocated, Err.Number will be
    ' 0. To acccomodate this case, we test
    ' whether LBound <= Ubound. If this
    ' is True, the array is allocated. Otherwise,
    ' the array is not allocated.
    '''''''''''''''''''''''''''''''''''''''
    If LBound(Arr) <= UBound(Arr) Then
        ' no error. array has been allocated.
        IsArrayAllocated = True
    Else
        IsArrayAllocated = False
    End If
Else
    ' error. unallocated array
    IsArrayAllocated = False
End If

End Function


Private Function AddDateTimeToFileName(ByVal FileName, _
                               Optional ByVal UseThisDateTime As Date = 0) As String

If UseThisDateTime = 0 Then UseThisDateTime = Now()
Dim II, DT:  DT = "_" & Format(UseThisDateTime, "yyyymmdd") & "-" & Format(UseThisDateTime, "hhmmss")
II = InStrRev(FileName, ".")
If II = 0 Then
  AddDateTimeToFileName = FileName & DT
  Exit Function
End If
AddDateTimeToFileName = Left(FileName, II - 1) & DT & Mid(FileName, II)

End Function


Private Function Lpad(ByVal strInput As String, ByVal NewLen As Long, Optional ByVal PadChar As String = "0") As String
If Len(PadChar) > 1 Then PadChar = Left(PadChar, 1)
If PadChar = "" Then PadChar = "0"
Lpad = strInput
Do
  If Len(Lpad) >= NewLen Then Exit Function
  Lpad = PadChar & Lpad
Loop
End Function

Function RunFile(strFile As String, strWndStyle As String)
On Error GoTo Error_Handler
 
    Shell "cmd /k """ & strFile & """", strWndStyle
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    err.Number & vbCrLf & "Error Source: RunFile" & vbCrLf & "Error Description: " & _
    err.Description, vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
    Resume ' Extra Resume
End Function

Private Function InsertNewElementIntoArray(ByRef InputArray As Variant, _
                                          ByVal Value As Variant, _
                                          Optional ByVal Unique As Boolean = False, _
                                          Optional ByVal Sorted As Boolean = False, _
                                          Optional ByVal LowerBound As Long = 0 _
                                                                                      ) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' InsertNewElementIntoArray - JGray 4/10/2018
' The input/output array must be dynamic so that the size can be expanded properly.
' This function will insert new values into an array.  When Unique is True, duplicate values will be culled
' from the table.  If Sorted is True, the table returned will always be sorted.  If Sorted is False, the input
' sequence of the values will be maintained in the Array.
' Returns True or False indicating success.
'    False also indicates that a "Non-Unique" value was NOT added to array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim I  As Integer
InsertNewElementIntoArray = True

''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    InsertNewElementIntoArray = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''
' Ensure InputArray is a dynamic array.
'''''''''''''''''''''''''''''''''''
If IsArrayDynamic(InputArray) = False Then
    InsertNewElementIntoArray = False
    Exit Function
End If

If IsArrayEmpty(InputArray) Then
  ReDim InputArray(LowerBound To LowerBound)
  InputArray(LowerBound) = Value
  Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure InputArray is a one-dimensional
' array.
''''''''''''''''''''''''''''''''''
Dim XX  As Long: XX = NumberOfArrayDimensions(InputArray)
If NumberOfArrayDimensions(InputArray) <> 1 Then
    InsertNewElementIntoArray = False
    Exit Function
End If

For I = LBound(InputArray) To UBound(InputArray)
  If InputArray(I) = Value And Unique Then
    InsertNewElementIntoArray = False
    Exit Function
  End If
  If InputArray(I) > Value And Sorted Then Exit For
Next I

Call InsertElementIntoArray(InputArray, UBound(InputArray) + 1, Value)

If Not Sorted Then Exit Function
If Not IsArraySorted(InputArray) Then
  InsertNewElementIntoArray = QuickSort(InputArray, LBound(InputArray), UBound(InputArray))
End If

End Function

Private Function IsArrayDynamic(ByRef Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayDynamic
' This function returns TRUE or FALSE indicating whether Arr is a dynamic array.
' Note that if you attempt to ReDim a static array in the same procedure in which it is
' declared, you'll get a compiler error and your code won't run at all.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim LUBound As Long

' If we weren't passed an array, get out now with a FALSE result
If IsArray(Arr) = False Then
    IsArrayDynamic = False
    Exit Function
End If

' If the array is empty, it hasn't been allocated yet, so we know
' it must be a dynamic array.
If IsArrayEmpty(Arr:=Arr) = True Then
    IsArrayDynamic = True
    Exit Function
End If

' Save the UBound of Arr.
' This value will be used to restore the original UBound if Arr
' is a single-dimensional dynamic array. Unused if Arr is multi-dimensional,
' or if Arr is a static array.
LUBound = UBound(Arr)

On Error Resume Next
err.Clear

' Attempt to increase the UBound of Arr and test the value of Err.Number.
' If Arr is a static array, either single- or multi-dimensional, we'll get a
' C_ERR_ARRAY_IS_FIXED_OR_LOCKED error. In this case, return FALSE.
'
' If Arr is a single-dimensional dynamic array, we'll get C_ERR_NO_ERROR error.
'
' If Arr is a multi-dimensional dynamic array, we'll get a
' C_ERR_SUBSCRIPT_OUT_OF_RANGE error.
'
' For either C_NO_ERROR or C_ERR_SUBSCRIPT_OUT_OF_RANGE, return TRUE.
' For C_ERR_ARRAY_IS_FIXED_OR_LOCKED, return FALSE.

ReDim Preserve Arr(LBound(Arr) To LUBound + 1)

Select Case err.Number
    Case C_ERR_NO_ERROR
        ' We successfully increased the UBound of Arr.
        ' Do a ReDim Preserve to restore the original UBound.
        ReDim Preserve Arr(LBound(Arr) To LUBound)
        IsArrayDynamic = True
    Case C_ERR_SUBSCRIPT_OUT_OF_RANGE
        ' Arr is a multi-dimensional dynamic array.
        ' Return True.
        IsArrayDynamic = True
    Case C_ERR_ARRAY_IS_FIXED_OR_LOCKED
        ' Arr is a static single- or multi-dimensional array.
        ' Return False
        IsArrayDynamic = False
    Case Else
        ' We should never get here.
        ' Some unexpected error occurred. Be safe and return False.
        IsArrayDynamic = False
End Select

End Function


Private Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayEmpty
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is really the reverse of IsArrayAllocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim LB As Long
Dim UB As Long

err.Clear
On Error Resume Next
If IsArray(Arr) = False Then
    ' we weren't passed an array, return True
    IsArrayEmpty = True
End If

' Attempt to get the UBound of the array. If the array is
' unallocated, an error will occur.
UB = UBound(Arr, 1)
If (err.Number <> 0) Then
    IsArrayEmpty = True
Else
    ''''''''''''''''''''''''''''''''''''''''''
    ' On rare occassion, under circumstances I
    ' cannot reliably replictate, Err.Number
    ' will be 0 for an unallocated, empty array.
    ' On these occassions, LBound is 0 and
    ' UBoung is -1.
    ' To accomodate the weird behavior, test to
    ' see if LB > UB. If so, the array is not
    ' allocated.
    ''''''''''''''''''''''''''''''''''''''''''
    err.Clear
    LB = LBound(Arr)
    If LB > UB Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
End If

End Function

Private Function NumberOfArrayDimensions(Arr As Variant) As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Integer
Dim Res As Integer
On Error Resume Next
' Loop, increasing the dimension index Ndx, until an error occurs.
' An error will occur when Ndx exceeds the number of dimension
' in the array. Return Ndx - 1.
Do
    Ndx = Ndx + 1
    Res = UBound(Arr, Ndx)
Loop Until err.Number <> 0

NumberOfArrayDimensions = Ndx - 1

End Function


Private Function InsertElementIntoArray(InputArray As Variant, Index As Long, _
    Value As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' InsertElementIntoArray
' This function inserts an element with a value of Value into InputArray at locatation Index.
' InputArray must be a dynamic array. The Value is stored in location Index, and everything
' to the right of Index is shifted to the right. The array is resized to make room for
' the new element. The value of Index must be greater than or equal to the LBound of
' InputArray and less than or equal to UBound+1. If Index is UBound+1, the Value is
' placed at the end of the array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long

'''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''
InsertElementIntoArray = False

''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Ensure InputArray is dynamic.
''''''''''''''''''''''''''''''''
If IsArrayDynamic(Arr:=InputArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''
' Ensure InputArray is allocated.
'''''''''''''''''''''''''''''''''
If IsArrayAllocated(Arr:=InputArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''
' Ensure InputArray is a single
' dimensional array.
'''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=InputArray) <> 1 Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Ensure Index is a valid element index.
' We allow Index to be equal to
' UBound + 1 to facilitate inserting
' a value at the end of the array. E.g.,
' InsertElementIntoArray(Arr,UBound(Arr)+1,123)
' will insert 123 at the end of the array.
'''''''''''''''''''''''''''''''''''''''''
If (Index < LBound(InputArray)) Or (Index > UBound(InputArray) + 1) Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' Resize the array
'''''''''''''''''''''''''''''''''''''''''''''
ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) + 1)
'''''''''''''''''''''''''''''''''''''''''''''
' First, we set the newly created last element
' of InputArray to Value. This is done to trap
' an error 13, type mismatch. This last entry
' will be overwritten when we shift elements
' to the right, and the Value will be inserted
' at Index.
'''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
err.Clear
InputArray(UBound(InputArray)) = Value
If err.Number <> 0 Then
    ''''''''''''''''''''''''''''''''''''''
    ' An error occurred, most likely
    ' an error 13, type mismatch.
    ' Redim the array back to its original
    ' size and exit the function.
    '''''''''''''''''''''''''''''''''''''''
    ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) - 1)
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''''
' Shift everything to the right.
'''''''''''''''''''''''''''''''''''''''''''''
For Ndx = UBound(InputArray) To Index + 1 Step -1
    InputArray(Ndx) = InputArray(Ndx - 1)
Next Ndx

'''''''''''''''''''''''''''''''''''''''''''''
' Insert Value at Index
'''''''''''''''''''''''''''''''''''''''''''''
InputArray(Index) = Value

   
InsertElementIntoArray = True


End Function

Private Function IsArraySorted(TestArray As Variant, _
    Optional Descending As Boolean = False) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArraySorted
' This function determines whether a single-dimensional array is sorted. Because
' sorting is an expensive operation, especially so on large array of Variants,
' you may want to determine if an array is already in sorted order prior to
' doing an actual sort.
' This function returns True if an array is in sorted order (either ascending or
' descending order, depending on the value of the Descending parameter -- default
' is false = Ascending). The decision to do a string comparison (with StrComp) or
' a numeric comparison (with < or >) is based on the data type of the first
' element of the array.
' If TestArray is not an array, is an unallocated dynamic array, or has more than
' one dimension, or the VarType of TestArray is not compatible, the function
' returns NULL.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim StrCompResultFail As Long
Dim NumericResultFail As Boolean
Dim Ndx As Long
Dim NumCompareResult As Boolean
Dim StrCompResult As Long

Dim IsString As Boolean
Dim VType As VbVarType

''''''''''''''''''''''''''''''''''
' Ensure TestArray is an array.
''''''''''''''''''''''''''''''''''
If IsArray(TestArray) = False Then
    IsArraySorted = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a single dimensional array.
''''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=TestArray) <> 1 Then
    IsArraySorted = Null
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' The following code sets the values of
' comparison that will indicate that the
' array is unsorted. It the result of
' StrComp (for strings) or ">=" (for
' numerics) equals the value specified
' below, we know that the array is
' unsorted.
'''''''''''''''''''''''''''''''''''''''''''''
If Descending = True Then
    StrCompResultFail = -1
    NumericResultFail = False
Else
    StrCompResultFail = 1
    NumericResultFail = True
End If

''''''''''''''''''''''''''''''''''''''''''''''
' Determine whether we are going to do a string
' comparison or a numeric comparison.
''''''''''''''''''''''''''''''''''''''''''''''
VType = VarType(TestArray(LBound(TestArray)))
Select Case VType
    Case vbArray, vbDataObject, vbEmpty, vbError, vbNull, vbObject, vbUserDefinedType
    '''''''''''''''''''''''''''''''''
    ' Unsupported types. Reutrn Null.
    '''''''''''''''''''''''''''''''''
        IsArraySorted = Null
        Exit Function
    Case vbString, vbVariant
    '''''''''''''''''''''''''''''''''
    ' Compare as string
    '''''''''''''''''''''''''''''''''
        IsString = True
    Case Else
    '''''''''''''''''''''''''''''''''
    ' Compare as numeric
    '''''''''''''''''''''''''''''''''
        IsString = False
End Select

For Ndx = LBound(TestArray) To UBound(TestArray) - 1
    If IsString = True Then
        StrCompResult = StrComp(TestArray(Ndx), TestArray(Ndx + 1))
        If StrCompResult = StrCompResultFail Then
            IsArraySorted = False
            Exit Function
        End If
    Else
        NumCompareResult = (TestArray(Ndx) >= TestArray(Ndx + 1))
        If NumCompareResult = NumericResultFail Then
            IsArraySorted = False
            Exit Function
        End If
    End If
Next Ndx


''''''''''''''''''''''''''''
' If we made it out of  the
' loop, then the array is
' in sorted order. Return
' True.
''''''''''''''''''''''''''''
IsArraySorted = True

End Function

Private Function QuickSort(vArray As Variant, Optional ByVal inLow As Long = -1, Optional ByVal inHi As Long = -1) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' QuickSort - JGray 4/10/2018
' This subroutine will do an Ascending sort on a single Dimension array.
' Returns True or False indicating success.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim XinLow   As Long:  XinLow = inLow
Dim XinHi    As Long:  XinHi = inHi
QuickSort = True

''''''''''''''''''''''''''''''
' Ensure vArray is an array.
''''''''''''''''''''''''''''''
If IsArray(vArray) = False Then
    QuickSort = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure vArray is a one-dimensional
' array.
''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(vArray) <> 1 Then
    QuickSort = False
    Exit Function
End If

If XinLow < 0 Then XinLow = LBound(vArray)
If XinHi < 0 Then XinHi = UBound(vArray)
 
''''''''''''''''''''''''''''''''''''''''''
' Ensure Hi is greater than Low
''''''''''''''''''''''''''''''''''''''''''
If XinLow >= XinHi Then
    QuickSort = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Call private (Recurrsive) subroutine to handle sorting.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call QuickSortx(vArray, XinLow, XinHi)
End Function


Private Sub QuickSortx(vArray As Variant, inLow As Long, inHi As Long)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' QuickSortx - JGray 4/10/2018  (Called by QuickSort)
' This subroutine will do an Ascending sort on a single Dimension array.
' Returns True or False indicating success.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSortx vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSortx vArray, tmpLow, inHi

End Sub

Private Sub InitGitIgnore(ByVal GitProjectFolderName As String)

Dim AccessAccdbName  As String
AccessAccdbName = Application.CurrentProject.Name
Dim GitIgnoreName  As String:  GitIgnoreName = ".gitignore"

If FileOrDirExists(GitProjectFolderName & GitIgnoreName) Then Exit Sub

Dim GitIgnoreText  As String


GitIgnoreText = "#" & vbCrLf & _
                "# >>> GitIgnore list of file patterns in the project folder that Git will ignore......" & vbCrLf & _
                "# >>> GitIgnore Rules can be found in:   https://git-scm.com/book/en/v2/Git-Basics-Recording-Changes-to-the-Repository#Ignoring-Files" & vbCrLf & _
                "#" & vbCrLf & vbCrLf & vbCrLf & _
                "# Windows BAT that is generated by Gexp" & vbCrLf & _
                ".Git*.bat" & vbCrLf & _
                ".RunBth*.bat" & vbCrLf & _
                ".RunBth*.txt" & vbCrLf & _
                "*.git" & vbCrLf
                
Call FileSave(GitIgnoreText, GitProjectFolderName & GitIgnoreName)

End Sub


Private Sub InitReadMe(ByVal GitProjectFolderName As String)

Dim AccessAccdbName  As String
AccessAccdbName = Application.CurrentProject.Name
Dim ReadMeTxtName  As String:  ReadMeTxtName = "ReadMe.md"

If FileOrDirExists(GitProjectFolderName & ReadMeTxtName) Then Exit Sub

Dim ReadMeText  As String

'GitProjectName
ReadMeText = "# " & GitProjectName & " " & vbCrLf & _
             "This is a readme text for " & AccessAccdbName & ":  To be altered to suit your needs...."
                
Call FileSave(ReadMeText, GitProjectFolderName & ReadMeTxtName)

End Sub


Private Sub CreateGitBatFile(ByVal GitProjectFolderName As String)

Dim AccessAccdbName  As String, II, JJ
Dim aDriveLetter, aPathName
AccessAccdbName = Application.CurrentProject.Name
Dim GexpBatFile As String: GexpBatFile = "1 Gexp-VBA-Git.bat"

Dim VbaProjectFrom As String, VbaProjectTo As String
If Right(GitProjectFolderName, 1) <> "\" Then GitProjectFolderName = GitProjectFolderName & "\"
VbaProjectTo = GitProjectFolderName & AccessAccdbName
II = InStrRev(VbaProjectTo, ".")
If II <> 0 Then VbaProjectTo = Left(VbaProjectTo, II - 1) & "-Git" & Mid(VbaProjectTo, II)

VbaProjectFrom = GitProjectFolderName
II = InStrRev(VbaProjectFrom, "\", Len(VbaProjectFrom) - 1)
If II <> 0 Then VbaProjectFrom = Left(VbaProjectFrom, II) & AccessAccdbName

II = InStr(GitProjectFolderName, ":")
If II <> 0 Then
  aDriveLetter = Left(GitProjectFolderName, II)
  aPathName = Mid(GitProjectFolderName, II + 1)
End If

Dim Response
Response = MsgBox("Do you want to include the entire project """ & AccessAccdbName & """ as part of the Git Export?", vbYesNo)

Dim XX As String, YY As String, ZZ As String, AA As String
XX = "": YY = "": ZZ = ""

XX = "c:" & vbCrLf & _
     "" & vbCrLf & _
     "Rem -- This bat file was generated by Gexp function to commit VBA project changes.." & vbCrLf & _
     "@echo off" & vbCrLf & _
     "set /p gititle=" & AccessAccdbName & " -  Enter Title/Description for Gexp export: " & vbCrLf & _
     "Echo ""This bat file was generated by Gexp function to commit VBA project changes for - %gititle%""" & vbCrLf & _
     "" & vbCrLf

Dim aDriveLetter2, aPathFileName2
If Response = vbYes Then
  YY = ""
  If Not FileOrDirExists(VbaProjectTo) Then
    II = InStr(VbaProjectTo, ":")
    If II <> 0 Then
      aDriveLetter2 = Left(VbaProjectTo, II)
      aPathFileName2 = Mid(VbaProjectTo, II + 1)
    End If
    YY = "Rem -- Create an empty/NUL version of the project file so that XCOPY will avoid un-needed propmts." & vbCrLf & _
         aDriveLetter2 & vbCrLf & _
         "type NUL > """ & aPathFileName2 & """" & vbCrLf
  End If

  YY = YY & _
       ": TryCopy" & vbCrLf & _
       "echo on" & vbCrLf & _
       "xcopy """ & VbaProjectFrom & """ """ & VbaProjectTo & """ /Y" & vbCrLf & _
       "Echo off" & vbCrLf & _
       "if ""%errorlevel%"" == ""0"" (" & vbCrLf & _
       "  GoTo Continue" & vbCrLf & _
       "  )" & vbCrLf & _
       "" & vbCrLf & _
       "    Echo ""xCopy failed due with errorlevel=%errorlevel%""" & vbCrLf & _
       "    Echo ""Ensure that Office has closed the " & AccessAccdbName & "  project before pressing any key to try again.""" & vbCrLf & _
       "  pause" & vbCrLf & _
       "  GoTo TryCopy" & vbCrLf & _
       ": Continue" & vbCrLf & vbCrLf & _
       "Echo """ & AccessAccdbName & " has been successfully copied..  Git Commit process will continue.""" & vbCrLf
End If

ZZ = aDriveLetter & vbCrLf & _
     "cd """ & aPathName & """" & vbCrLf & _
     "echo on" & vbCrLf & _
     "git add --all" & vbCrLf & _
     "git commit -am """ & Now() & "  VBA Gexp - %gititle%""" & vbCrLf & _
     "git ls-tree -r master" & vbCrLf & _
     "Echo off" & vbCrLf & _
     "Echo ""Ended...........""" & vbCrLf & _
     "pause" & vbCrLf
     
Dim ShutMsAppBeforeRunningBTH   As String
ShutMsAppBeforeRunningBTH = ""
If Response = vbYes Then _
  ShutMsAppBeforeRunningBTH = Application.CurrentProject.Name & _
          " will be copied to the Git project folder and committed to Git Repo."

Call MakeAndRunBatchFile(XX & YY & ZZ, GitProjectFolderName, ShutMsAppBeforeRunningBTH)

Debug.Print ("Gexp - Process has ended...")


End Sub


Private Sub MakeAndRunBatchFile(ByVal BatchFileText As String, _
                                ByVal HoldFilePath As String, _
                                Optional ByVal ShutMsAppBeforeRunningBTH As String = "", _
                                Optional ByVal TimeOutValue As Long = 120, _
                                Optional ByVal FinishedMarker As String = "C:\BthFinished.txt")
                                
                                
' First add Batch CMDs to the end of the script that will create a "Finished" marker.
Dim II, JJ, KK, XX, MarkerFile
Dim TimeNow: TimeNow = Now()
Dim BatchFileNameAndPath As String

Dim winDriveLetter
II = InStr(1, HoldFilePath, ":")
If II <= 1 Then
  MsgBox ("MakeAndRunBatchFile - HoldFilePath=" & HoldFilePath & vbCrLf & _
         "      does not begin with a driver letter.  Process will abort.")
  End
End If
winDriveLetter = Left(HoldFilePath, II - 1)
XX = Left(HoldFilePath, II - 1) & ":" & vbCrLf & "cd """
If Mid(HoldFilePath, II + 1, 1) <> "\" Then II = II + 1
XX = XX & Mid(HoldFilePath, II + 1) & vbCrLf & "type NUL > "

'  Now make sure that the file path does NOT end in '\'
If Right(HoldFilePath, 1) = "\" Then HoldFilePath = Left(HoldFilePath, Len(HoldFilePath) - 1)

MarkerFile = AddDateTimeToFileName(".RunBthMarkerFile.txt", TimeNow)
BatchFileNameAndPath = HoldFilePath & "\" & AddDateTimeToFileName(".RunBthFile.bat", TimeNow)

' Remove a final "pause" if it exists in the script...
BatchFileText = TrimChars(BatchFileText, " " & vbCrLf)
If Right(BatchFileText, 5) = "pause" Then BatchFileText = Left(BatchFileText, Len(BatchFileText) - 5)

Dim DelBthFileName   As String
II = InStrRev(BatchFileNameAndPath, "\")
If II = 0 Then II = InStrRev(BatchFileNameAndPath, ":")
If (II <> 0 And II < Len(BatchFileNameAndPath)) Then _
  DelBthFileName = "Echo off" & vbCrLf & _
                   "Echo ""Script had Ended...........  """ & vbCrLf & _
                   "Echo ""Enter EXIT to close the CMD window....  (The batch file cannot be found.) is normal and can be ignored. """ & vbCrLf & _
                   "del """ & Mid(BatchFileNameAndPath, II + 1) & """" & vbCrLf

XX = vbCrLf & "pause" & vbCrLf & XX & """" & MarkerFile & """" & vbCrLf & "exit"
' Don't add text to create marker file if project must be closed.
If ShutMsAppBeforeRunningBTH <> "" Then XX = vbCrLf & "pause" & vbCrLf & DelBthFileName
BatchFileText = BatchFileText & XX

Call FileSave(BatchFileText, BatchFileNameAndPath)
Call RunFile(BatchFileNameAndPath, vbNormalFocus) '  vbHide can also be used...

If ShutMsAppBeforeRunningBTH = "" Then
    MarkerFile = HoldFilePath & "\" & MarkerFile
    Do Until (Dir(MarkerFile) <> "")  ' Wait until the file shows up in the folder.
      Call WaitFor(1) ' Wait 1 second before looking for file again....
    Loop
    Call Kill(MarkerFile)
    Call Kill(BatchFileNameAndPath)
    Exit Sub
End If

' Access, Excel or Word must be shut down....
ShutMsAppBeforeRunningBTH = ShutMsAppBeforeRunningBTH & vbCrLf & vbCrLf & _
    "Program will end. " & Application.CurrentProject.Name & " project will need to be manually ended."
Call MsgBox(ShutMsAppBeforeRunningBTH)
End

End Sub


Private Sub InitGitCommit(ByVal GitProjectFolderName As String)

Dim AccessAccdbName  As String, II, JJ, aDriveLetter, aPathName
AccessAccdbName = Application.CurrentProject.Name
Dim GitCommitName  As String:  GitCommitName = ".GitCommit.bat"

'If FileOrDirExists(GitProjectFolderName & GitCommitName) Then Exit Sub

Dim GitCommitText  As String

Dim VbaProjectFrom As String, VbaProjectTo As String
If Right(GitProjectFolderName, 1) <> "\" Then GitProjectFolderName = GitProjectFolderName & "\"
VbaProjectTo = GitProjectFolderName & AccessAccdbName
II = InStrRev(VbaProjectTo, ".")
If II <> 0 Then VbaProjectTo = Left(VbaProjectTo, II - 1) & "-Git" & Mid(VbaProjectTo, II)

VbaProjectFrom = GitProjectFolderName
II = InStrRev(VbaProjectFrom, "\", Len(VbaProjectFrom) - 1)
If II <> 0 Then VbaProjectFrom = Left(VbaProjectFrom, II) & AccessAccdbName

II = InStr(GitProjectFolderName, ":")
If II <> 0 Then
  aDriveLetter = Left(GitProjectFolderName, II)
  aPathName = Mid(GitProjectFolderName, II + 1)
End If

GitCommitText = "Rem -- This bat file was generated by Gexp function to commit VBA project changes.." & vbCrLf & _
  "c:" & vbCrLf & _
  "@echo off" & vbCrLf & _
  "set /p gititle=" & Application.CurrentProject.Name & " -  Enter Title/Description for Gexp export: " & vbCrLf & _
  "Echo This bat file was generated by Gexp function to commit VBA project changes for - ""%gititle%""" & vbCrLf & _
  "Echo ." & vbCrLf & _
  "" & vbCrLf & _
  ": AskQ" & vbCrLf & _
  "set /p gitInc=" & Application.CurrentProject.Name & " -  Should project be included in the Git Commit? (Y or N) " & vbCrLf & _
  "if ""%gitInc%"" == ""Y"" (GoTo TryCopy)" & vbCrLf & _
  "if ""%gitInc%"" == ""y"" (GoTo TryCopy)" & vbCrLf & _
  "if ""%gitInc%"" == ""N"" (GoTo ContinueScript)" & vbCrLf & _
  "if ""%gitInc%"" == ""n"" (GoTo ContinueScript)" & vbCrLf & _
  "    " & vbCrLf & _
  "Echo Response ""%gitInc%"" was invalid.  Must be Y or N.  Please try again." & vbCrLf & _
  "GoTo AskQ" & vbCrLf
  
GitCommitText = GitCommitText & _
  "" & vbCrLf & _
  ": TryCopy" & vbCrLf & _
  "echo on" & vbCrLf & _
  "xcopy """ & VbaProjectFrom & """ """ & VbaProjectTo & """ /Y" & vbCrLf & _
  "Echo off" & vbCrLf & _
  "if ""%errorlevel%"" == ""0"" (GoTo Continue)" & vbCrLf & _
  "    Echo xCopy failed due with errorlevel=%errorlevel%" & vbCrLf & _
  "    Echo Ensure that Office has closed the " & Application.CurrentProject.Name & "  project before pressing any key to try again." & vbCrLf & _
  "  pause" & vbCrLf & _
  "  GoTo TryCopy" & vbCrLf & _
  ": Continue" & vbCrLf & _
  "Echo " & Application.CurrentProject.Name & " has been successfully copied..  Git Commit process will continue." & vbCrLf & _
  ": ContinueScript" & vbCrLf & _
  aDriveLetter & vbCrLf & _
  "cd """ & aPathName & """" & vbCrLf & _
  "echo on" & vbCrLf & _
  "git add --all " & vbCrLf & _
  "git commit -am ""%date% %time%   VBA Gexp - %gititle%""" & vbCrLf & _
  "git ls-tree -r master" & vbCrLf & _
  "Echo off" & vbCrLf & _
  "Echo GitCommit Script Ended..........." & vbCrLf & _
  "pause"
                
Call FileSave(GitCommitText, GitProjectFolderName & GitCommitName)

End Sub


Private Sub InitGitPush(ByVal GitProjectFolderName As String)

Dim AccessAccdbName  As String, II, Response, XX As String, ZZ As String
AccessAccdbName = Application.CurrentProject.Name

Dim aDriveLetter, aPathName
II = InStr(GitProjectFolderName, ":")
If II <> 0 Then
  aDriveLetter = Left(GitProjectFolderName, II)
  aPathName = Mid(GitProjectFolderName, II + 1)
End If

Dim GitRepoName   As String
XX = Left(GitProjectFolderName, Len(GitProjectFolderName) - 1)  ' Trim trailing backslash
II = InStrRev(XX, "\")
GitRepoName = Mid(XX, II + 1)
'  the ""ImportProject2_accdb"" repo
'  the """ & GitRepoName & """ repo


AccessAccdbName = Application.CurrentProject.Name
Dim GitPushName  As String:  GitPushName = ".GitPush.bat"

'If FileOrDirExists(GitProjectFolderName & GitPushName) Then Exit Sub

Dim GitPushText  As String    'GitHubRemoteUrl
  
GitPushText = ""  ' Start fresh...

GitPushText = GitPushText & _
   "Rem -- This bat file was generated by Gexp function to Push VBA project changes.." & vbCrLf & _
   aDriveLetter & vbCrLf & _
   "Echo off" & vbCrLf & _
   "cd """ & GitProjectFolderName & """" & vbCrLf & _
   "" & vbCrLf & _
   ":: First establish a remote for GitHub" & vbCrLf & _
   "echo on" & vbCrLf & _
   "git remote remove origin" & vbCrLf & _
   "git remote add origin " & GitHubRemoteUrl & vbCrLf & _
   "Echo off" & vbCrLf & _
   "" & vbCrLf & _
   ": TryAgain" & vbCrLf & _
   "" & vbCrLf & _
   "echo on" & vbCrLf & _
   "git remote -v" & vbCrLf & _
   "git push -u origin master" & vbCrLf & _
   "Echo off" & vbCrLf & _
   "" & vbCrLf

   
GitPushText = GitPushText & _
   "if ""%errorlevel%"" == ""0"" (GoTo PushSuccessful)" & vbCrLf & _
   "if ""%errorlevel%"" == """" (GoTo PushSuccessful)" & vbCrLf & _
   "" & vbCrLf & _
   "if ""%errorlevel%"" == ""1"" (GoTo MustPullFirst)" & vbCrLf & _
   "" & vbCrLf & _
   "GoTo AskAgain" & vbCrLf & _
   "" & vbCrLf & _
   ": MustPullFirst" & vbCrLf & _
   "echo ." & vbCrLf & _
   "set /p ExecPull=Ok to first do a pull (with -f to force) from """ & GitRepoName & """ and try the push again? (Y or N) " & vbCrLf & _
   "if ""%ExecPull%"" == ""Y"" (GoTo Pull)" & vbCrLf & _
   "if ""%ExecPull%"" == ""y"" (GoTo Pull)" & vbCrLf

GitPushText = GitPushText & _
   "if ""%ExecPull%"" == ""N"" (GoTo TakeTheExit)" & vbCrLf & _
   "if ""%ExecPull%"" == ""n"" (GoTo TakeTheExit)" & vbCrLf & _
   "GoTo MustPullFirst" & vbCrLf & _
   "echo on" & vbCrLf & _
   ": Pull" & vbCrLf & _
   "git pull -f origin master" & vbCrLf & _
   "Rem -- errorlevel = ""%errorlevel%""" & vbCrLf & _
   "pause" & vbCrLf & _
   "GoTo TryAgain" & vbCrLf & _
   "" & vbCrLf & _
   ": AskAgain" & vbCrLf & _
   "echo errorlevel = ""%errorlevel%""   128 indicates that GitHub cannot find the Repo with this " & GitLogIn & " logon" & vbCrLf & _
   "echo  ""git push"" was not successful.  Make sure that the """ & GitRepoName & """ repo exists in GitHub." & vbCrLf & _
   "echo After resolving the error, try the ""git push"" again." & vbCrLf & _
   "echo ." & vbCrLf & _
   "set /p TryAgain=Are you ready to try to ""git push"" again? (Y or N) " & vbCrLf & _
   "if ""%TryAgain%"" == ""Y"" (GoTo TryAgain)" & vbCrLf & _
   "if ""%TryAgain%"" == ""y"" (GoTo TryAgain)" & vbCrLf & _
   "if ""%TryAgain%"" == ""N"" (GoTo TakeTheExit)" & vbCrLf & _
   "if ""%TryAgain%"" == ""n"" (GoTo TakeTheExit)" & vbCrLf & _
   "GoTo AskAgain" & vbCrLf & _
   "" & vbCrLf

GitPushText = GitPushText & _
   "pause" & vbCrLf & _
   "exit" & vbCrLf & _
   "" & vbCrLf & _
   ": PushSuccessful" & vbCrLf & _
   "echo ""git push"" appears to be successful." & vbCrLf & _
   "pause" & vbCrLf & _
   "exit" & vbCrLf & _
   "" & vbCrLf & _
   ": TakeTheExit" & vbCrLf & _
   "echo ""git push"" NOT successful." & vbCrLf & _
   "pause" & vbCrLf & _
   "exit" & vbCrLf
                
Call FileSave(GitPushText, GitProjectFolderName & GitPushName)

End Sub

Public Function GitLogIn() As String
Dim XX, II, JJ
II = InStrRev(GitHubRemoteAdd, "/")
GitLogIn = ""
If II = 0 Then Exit Function
JJ = InStrRev(GitHubRemoteAdd, "/", II - 1)
If JJ = 0 Then Exit Function
GitLogIn = Mid(GitHubRemoteAdd, JJ + 1, (II - JJ) - 1)


End Function

Public Function GitHubRemoteUrl() As String
Dim XX, II, JJ
GitHubRemoteUrl = GitHubRemoteAdd

GitHubRemoteUrl = Replace(GitHubRemoteUrl, "REPOSITORY", GitProjectName)
If GitHubRemoteUrl <> GitHubRemoteAdd Then Exit Function

' Repository was not located in the GitHubRemoteAdd variable constant....
If Right(GitHubRemoteUrl, 4) <> ".git" Then Exit Function

II = InStrRev(GitHubRemoteUrl, "/")
If II = 0 Then Exit Function
GitHubRemoteUrl = Left(GitHubRemoteUrl, II) & GitProjectName & ".git"

End Function

Public Function GitProjectName() As String

Dim I
GitProjectName = Application.CurrentProject.Name
I = InStrRev(GitProjectName, ".")
If I > 1 Then GitProjectName = Left(GitProjectName, I - 1) & "_" & Mid(GitProjectName, I + 1)

End Function



Private Sub EnsureProjectFolderExists(ByVal GitRepoName As String)

Dim II, Response, XX As String

Dim aDriveLetter, aPathName, GitProjectFolderName
GitProjectFolderName = Application.CurrentProject.Path

II = InStr(GitProjectFolderName, ":")
If II <> 0 Then
  aDriveLetter = Left(GitProjectFolderName, II)
  aPathName = Mid(GitProjectFolderName, II + 1)
End If

XX = "A new Git Repo being initialized for: " & Application.CurrentProject.Name & vbCrLf & vbCrLf

Dim GitHubRemoteAddIsValid  As Boolean
If GitHubRemoteAdd <> "" And GitLogIn <> "" And GitLogIn <> "USERNAME" Then GitHubRemoteAddIsValid = True

If Not GitHubRemoteAddIsValid Then
  Response = MsgBox(XX & "GitHubRemoteAdd = """" so no Remote Repo can be cloned or attached. " & vbCrLf & vbCrLf & _
                    "Continue?  No will abort the run so that ""GitHubRemoteAdd"" variable can be set properly.", vbYesNo)
 Else
  Response = MsgBox(XX & "GitHubRemoteAdd = """ & GitHubRemoteUrl & """ will be used to set the Remote Repo. " & vbCrLf & vbCrLf & _
                    "Continue?  No will abort the run so that ""GitHubRemoteAdd"" variable can be set properly.", vbYesNo)
End If
If Response = vbNo Then End

XX = ""
XX = XX & _
     aDriveLetter & vbCrLf & _
     "cd " & aPathName & vbCrLf & _
     "Rem -- This bat file was generated by Gexp function to Init VBA project changes.." & vbCrLf & _
     "set GitRemoteName=" & GitHubRemoteUrl & vbCrLf & _
     "Set GitProjectName=" & GitRepoName & vbCrLf & _
     "" & vbCrLf & _
     "Rem First attempt to clone from Remote Repo" & vbCrLf & _
     "" & vbCrLf & _
     "git clone %GitRemoteName%" & vbCrLf & _
     "if ""%errorlevel%"" == ""0"" (" & vbCrLf & _
     "  :: Change Directory and wrap up with Git Status" & vbCrLf & _
     "  cd %GitProjectName%" & vbCrLf & _
     "  git status" & vbCrLf & _
     "  GoTo ScriptEnds" & vbCrLf & _
     "  )" & vbCrLf & _
     "" & vbCrLf

XX = XX & _
     "Rem errorlevel=%errorlevel%   was encountered by ""git clone""." & vbCrLf & _
     "mkdir %GitProjectName%" & vbCrLf & _
     "cd %GitProjectName%" & vbCrLf & _
     "git init" & vbCrLf & _
     "" & vbCrLf & _
     ": ScriptEnds" & vbCrLf & _
     "git remote -v " & vbCrLf & _
     "Echo off" & vbCrLf & _
     "Echo ""Ended...........""" & vbCrLf & _
     "pause" & vbCrLf
     
Call MakeAndRunBatchFile(XX, Application.CurrentProject.Path & "\")

Exit_Sub:
End Sub

