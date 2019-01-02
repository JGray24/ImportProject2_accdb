Attribute VB_Name = "ETL123_Import"
Option Compare Database
Option Explicit

Dim GblTogglePrintDebugLog As Boolean
Dim GblForceOnPrintDebugLog As Boolean
Dim GblForceOffPrintDebugLog As Boolean
Dim GblLastFilePickerPath  As String
Dim GblLastFolderPickerPath  As String

Dim aPrintDebugLog As Boolean ' If this is True, then the Debug log is produced.

' Run Time File Config.....
Dim XslFileAndPath
Dim RT_FileID
Dim RT_FileName
Dim RT_WorkSheetName As String
Dim RT_OutputTableName As String
Dim RT_UniqueKey As Boolean
Dim RT_MarkActiveFlag As String  ' Valid values: "Changed", "New", "Imported"
Dim RT_ActiveFlagField As String
Dim RT_KeyField() As Boolean
Dim RT_ID() As String
Dim RT_ExcelColumnLetter() As String
Dim RT_ExcelColumnLetterOrig() As String
Dim RT_ExcelHeadingText() As String
Dim RT_FieldNameOutput() As String
Dim RT_AcceptChanges()
Dim RT_AllowChange2Blank()
Dim RT_RejectErrFile()
Dim RT_RejectErrRows()
Dim RT_SaveDateChanged()
Dim RT_SaveDateChangedLineLvl()

Dim Config(), ConfigID(), ConfigFileName()  ' These Arrays will hold all Import Configurations found in Import.config file.

Dim strSql         As String
Dim ErrorMsg       As String
Dim HoldSelItem    As String    ' Combined path and file name.
Dim HoldFileName   As String    ' Only the file name.  Path is not included.
Dim HoldFilePath   As String    ' Only the path name.

'  These Following array values are populated by GetImportSpecifications and includes only Active Items
Dim numOfSpecs   As Long
Dim Add_Date_Changed_to_Rows  As Boolean
Dim aActive_flag_name  As String
Dim aMark_active_flag  As String

Dim aOutput_Table_Name() As String
Dim aExcel_Column_Number() As String
Dim aExcel_Heading_Text() As String
Dim aField_Name_Output() As String
Dim aData_Type() As String
Dim aData_Type_Enum() As Variant
Dim aXcelFieldHasLongText() As Boolean
Dim aDup_Key_Field() As Boolean
Dim aReject_Err_File() As Boolean
Dim aReject_Err_Rows() As Boolean
Dim aAccept_Changes() As Boolean
Dim aAllowChange2Blank() As Boolean
Dim aSave_Date_Changed() As Boolean

Dim aold_Val() As String      ' Hold the old value from the target table.
Dim afield_err() As String    ' Any error that is found with this field.
Dim holdMergedData() As String
Dim holdKeyNames() As String

Dim aNoImportErrorsWereFound As Boolean  ' This will be false if there were errors found.
Dim aWorkSheetErrorFound As Boolean      ' This flag will be used to flag an error in a worksheet.

Dim GblKeepExcelWkShtHeadings() As String
Dim ElapsedTimerName()  As String
Dim ElapsedTimerBegin() As Date

'''''''''''''''''''''''''''
' Error Number Constants
'''''''''''''''''''''''''''
Private Const C_ERR_NO_ERROR = 0&
Private Const C_ERR_SUBSCRIPT_OUT_OF_RANGE = 9&
Private Const C_ERR_ARRAY_IS_FIXED_OR_LOCKED = 10&










Public Function Etl_123(Optional ByVal InputPathAndFile As String = "", _
                        Optional ByVal EtlConfigInput As String = "", _
                        Optional ByVal SelectWkShtElemts As String = "", _
                        Optional ByVal SelectConfigIDs As String = "", _
                        Optional ByVal IgnoreWarnings As Boolean = False, _
                        Optional ByVal CompileOnly As Boolean = True, _
                        Optional ByVal AddDateTimeToRept As Boolean = False, _
                        Optional ByVal AllowUnmatchedWarnings As Boolean = True) As String

' InputPathAndFile - This field will contain a fully qualified Path and File to be imported.
'                    Example: C:\MyDirectory\Customers.xlsx
'
'                    If it is not a valid file, then User will be prompted to pick the file and
'                        use a pattern match:
'                    Examples:
'                       C:\MyDirectory\*.xlsx, *.xls, *.csv
'                       C:\MyDirectory\Customer*.xlsx, Customer*.xls, Invoice*.xlsx, Invoice*.xlsx
'
' EtlConfigInput - This is the Config to be used for this Import.
'     Options:    1) It is a valid file name with "*.config" file extension.
'                 2) It is an actual Config source code passed by the caller.  This allows the caller to
'                    have the configuration to come from another source, either internal or external.
'                 3) If this optional parm is not supplied, then the import config is read from the
'                    *.config version of the selected InputPathAndFile.
'                 4) If *.config file is not located, then user is prompted to pick the *.config file.
'                 5) "CreateConfig" will create a new Config Template for the Input File.
'                 6) "None" - This will cause the program to create a temporary Config using parms
'                             from the Input File and use it to import the file.  Useful in testing.
'
'
' Matching Process will match each WkSht in the picked Excel file with matching config elements.
' This match is done using a combination of the Excel headings and/or the WkSht names.
' The config file WkSht= parameter is matched to the WkShts found in the selected Input file.
' Pattern matching is supported for all field comparisons (Both WkShts and headings).

' SelectWkShtElemts - This will be a list of Excel WkShts that should be present in the Excel input file to be
'                   processed according to rules in the matching Config ID.  Pattern matching is supported.
'                   "*" will be interpreted as selecting ALL Excel WkShts.  All WkShts in this list will be
'                   required to have a matched Config ID for processing.
'                   (Ignored for a *.csv input file and forced to be "*")

' SelectConfigIDs - This will be a list of Config IDs that should be processed.  All IDs in this list will
'                   be required to match to at least one Excel WkSht.  Pattern matching is supported. "*" will be
'                   interpreted as selecting ALL ConfigIDs.  (Ignored for a *.csv input file)
'
' IgnoreWarnings - This Boolean if true will cause warnings to be ignored and not stop the Import Process.
'

Dim HoldConfigData As String
Dim II, JJ, KK, LL, Errmsg: Errmsg = ""
Dim ConfigSource:  ConfigSource = ""
Dim EtlConfig As String:  EtlConfig = EtlConfigInput
Dim Response
Dim NewConfigCreated As Boolean: NewConfigCreated = False
Dim ConfigPassedFromCaller  As Boolean:  ConfigPassedFromCaller = False

Call ElapsedTimeSince("Debug.Print", , True) ' Reset the Debug Timer....
Call ElapsedTimeSince("Etl123", , True) ' Reset the Etl123 Timer to start over.....

'1)  Validate the input file and pick the file to be processed if needed.
XslFileAndPath = FilePrePicker(InputPathAndFile, "*.xlsx, *.xlsm, *.xls, *.csv")
DebugPrintOn ("Processing: " & XslFileAndPath)


'2) Get the Config script using EtlConfig.  Sources can be to read from a script file or to have the
'   script text passed in by the caller.

'   a) It is a valid file name with "*.config" file extension?  Yes, then use this script.
    If EtlConfig Like "*.config" Then
      ConfigSource = "Config File Name was Passed from the Caller"
      If Not FileOrDirExists(EtlConfig) Then GoTo UserPickedConfig
      GoTo ConfigFileRead
    End If

'   b) If this optional parm is "CreateConfig" then use the InputPathAndFile to create a *.config template
'      that can process the file.
      If EtlConfig = "CreateConfig" Or EtlConfig = "None" Then GoTo CreateNewConfig

'   c) It is an actual Config source code passed by the caller.  This allows the caller to
'      have the configuration to come from another source, either internal or external.
      If EtlConfig <> "" Then
        ConfigPassedFromCaller = True
        ConfigSource = "Config Data was Passed from the Caller.   *.config file was NOT used"
        
        II = InStr(1, EtlConfig, "Config string generated by function")
        If II <> 0 Then
          ConfigSource = Mid(EtlConfig, II, 200)
          JJ = InStr(1, ConfigSource, "*/")
          If JJ <> 0 Then ConfigSource = Left(ConfigSource, JJ)
          ConfigSource = ConfigSource & "  Passed from the Caller.   *.config file was NOT used"
        End If
        
        HoldConfigData = EtlConfig
        EtlConfig = ""
        GoTo ContinueProcess
      End If

'   d) If this optional parm is not supplied, then the import config is read from the
'      *.config version of the selected InputPathAndFile, from the same path as the *.xlsx file.
      ConfigSource = "Config File Name derivied from the Input File"
      EtlConfig = XslFileAndPath & ".config"
      If FileOrDirExists(EtlConfig) Then GoTo ConfigFileRead
      
'   e) Config file is not located.  Ask the User if they want to create a *.config file.
      Response = MsgBox("*.config file has not been located for """ & EtlConfig & """. " & vbCrLf & vbCrLf _
                      & "Do you want the system to create a new *.config template?" & vbCrLf & vbCrLf _
                      & "Yes - Will create (and save) new *.config template for manual editing." & vbCrLf _
                      & "No - Will prompt the user for an existing *.config file." & vbCrLf _
                      & "Cancel - Will ABORT the process.", vbYesNoCancel, "Etl_123 - Create a new *.config template")
      If Response = vbYes Then GoTo CreateNewConfig
      If Response = vbNo Then GoTo UserPickedConfig
      End
      
'   f) If the file is not located, then User Should Pick the file.
UserPickedConfig:
      ConfigSource = ConfigSource & " was NOT found.  Config file picked by the user"
      EtlConfig = FilePrePicker(EtlConfig, "*.config")


ConfigFileRead:
  DebugPrintOn ("ConfigFileRead")
  HoldConfigData = GetEntireFile(EtlConfig)
  GoTo ContinueProcess
  
CreateNewConfig:
Dim NewConfigFileName As String, ReturnSelWkShtNames As Variant, SaveNewConfig As Boolean
If EtlConfig = "None" Then SaveNewConfig = False Else SaveNewConfig = True
HoldConfigData = CreateNewImportConfig(XslFileAndPath, NewConfigFileName, ReturnSelWkShtNames, SaveNewConfig)
            
NewConfigCreated = True   ' Set flag to indicate that this is a new Config file.
''''CompileOnly = True  ' New Config file should only be Compiled.
SelectConfigIDs = ""
IgnoreWarnings = True
SelectWkShtElemts = ""  ' Select each WkSht found to be evaluated.
If IsArrayAllocated(ReturnSelWkShtNames) Then
  For II = LBound(ReturnSelWkShtNames) To UBound(ReturnSelWkShtNames)
    If SelectWkShtElemts <> "" Then SelectWkShtElemts = SelectWkShtElemts & ", "
    SelectWkShtElemts = SelectWkShtElemts & ReturnSelWkShtNames(II)
  Next II
End If

ContinueProcess:

' Start the Report File....
Dim ReptFileName:  ReptFileName = XslFileAndPath & ".Etl_123Rept.txt"

Dim ReptData As String
Dim ReptDateTime As Date:  ReptDateTime = Format(Now(), "mm-dd-yyyy  hh:mm:ss")
ReptData = ReptDateTime & _
           "      Etl_123 Processing Results            User-" & Environ("USERNAME") & vbCrLf & vbCrLf & _
           "Run Time Parameters passed in to Etl_123:" & vbCrLf
Call FileSave(ReptData, ReptFileName) ' Start new report file.
           
ReptData = "Input Excel File (Supplied by the Caller) - """ & XslFileAndPath & """" & vbCrLf

If XslFileAndPath <> InputPathAndFile Then
  ReptData = "Picked (by the User) Excel File - """ & XslFileAndPath & """"
  ReptData = ReptData & vbCrLf & "        (Original Caller Input) - """ & InputPathAndFile & """"
End If
ReptData = "Parm #1 - InputPathAndFile - " & vbCrLf & "   " & TrimChars(ReptData, vbCrLf)
Call FileSave(ReptData, ReptFileName, , True): ReptData = "" ' Append to the report....
  
If EtlConfig = "" Then
  ReptData = ReptData & "   Config data for: """ & XslFileAndPath & """" _
           & vbCrLf & "                is: " & ConfigSource & "."
 Else
  ReptData = ReptData & "   Config data for: """ & XslFileAndPath & """" _
           & vbCrLf & "                is: """ & EtlConfig & """" _
           & vbCrLf & "   " & ConfigSource & "."
End If
   
ReptData = "Parm #2 - EtlConfigInput - " & vbCrLf & ReptData
Call FileSave(ReptData, ReptFileName, , True): ReptData = "" ' Append to the report....
    
Dim ParmMsg As String:  ParmMsg = ""
If RemoveWhiteSpace(SelectWkShtElemts) = "" And RemoveWhiteSpace(SelectConfigIDs) = "" Then
  SelectWkShtElemts = "All Matched Pairs"
  ParmMsg = "  (Both Parm #3 and #4 were missing.  All Matched Pairs will be Selected.)"
End If

ReptData = "Parm #3 - SelectWkShtElemts - """ & RemoveWhiteSpace(SelectWkShtElemts) & """" & ParmMsg
Call FileSave(ReptData, ReptFileName, , True): ReptData = "" ' Append to the report....
    
ReptData = "Parm #4 - SelectConfigIDs - """ & RemoveWhiteSpace(SelectConfigIDs) & """"
Call FileSave(ReptData, ReptFileName, , True): ReptData = "" ' Append to the report....

ReptData = "Parm #5 - IgnoreWarnings - " & IgnoreWarnings
ReptData = Rpad(ReptData, 36)
If IgnoreWarnings Then ReptData = ReptData & "(User will NOT be alerted that warnings exist in the Etl_123 report. Processing will continue with Warnings.)" _
                  Else ReptData = ReptData & "(User will be alerted that warnings exist in the Etl_123 report.)"
Call FileSave(ReptData, ReptFileName, , True): ReptData = "" ' Append to the report....

ReptData = "Parm #6 - CompileOnly - " & CompileOnly
ReptData = Rpad(ReptData, 36)
If CompileOnly Then ReptData = ReptData & "(Config will be Compiled and Etl_123 report produced.  Actually running of the Import will be skipped.)" _
               Else ReptData = ReptData & "(Config will be compiled and RUN if NO errors.)"
Call FileSave(ReptData, ReptFileName, , True): ReptData = "" ' Append to the report....

ReptData = "Parm #7 - AddDateTimeToRept - " & AddDateTimeToRept
ReptData = Rpad(ReptData, 36)
If AddDateTimeToRept Then ReptData = ReptData & "(Etl_123 report name will contain a Date/Time stamp.  This avoids replacing the report from previous run.)" _
                     Else ReptData = ReptData & "(Etl_123 report name will NOT contain a Date/Time stamp.  The report from previous run will be replaced.)"
If AddDateTimeToRept Then _
  ReptFileName = RenameAddDateTime(ReptFileName)
Call FileSave(ReptData, ReptFileName, , True): ReptData = "" ' Append to the report....
    
ReptData = "Parm #8 - AllowUnmatchedWarnings - " & AllowUnmatchedWarnings & "  "
ReptData = Rpad(ReptData, 36)
If AllowUnmatchedWarnings Then ReptData = ReptData & "(Warnings for UnMatched Config IDs and eXcel WkShts will be allowed to process the matched plans, and bypass those UnMatched.)" _
                          Else ReptData = ReptData & "(Warnings for UnMatched Config IDs and eXcel WkShts will be treated as Severe Errors and Abort the Import.)"
Call FileSave(ReptData, ReptFileName, , True): ReptData = "" ' Append to the report....
    
ReptData = vbCrLf & "Selected Configuration for this Import.........................................................."
Call FileSave(ReptData & vbCrLf & HoldConfigData, ReptFileName, , True): ReptData = "" ' Append to the report....

Dim SelectWkShts() As String, SelectIDs() As String
If Right(XslFileAndPath, 4) = ".csv" Then  ' Force the proper selections for a *.csv
  SelectWkShtElemts = "*"
  SelectConfigIDs = ""
End If
SelectIDs = Split(TrimChars(SelectConfigIDs, ","), ",")
SelectWkShts = Split(TrimChars(SelectWkShtElemts, ","), ",")
For II = LBound(SelectIDs) To UBound(SelectIDs)
  SelectIDs(II) = Trim(SelectIDs(II))
Next II
For II = LBound(SelectWkShts) To UBound(SelectWkShts)
  SelectWkShts(II) = Trim(SelectWkShts(II))
Next II

Dim HoldConfigExcel() As String
Dim HoldConfig() As String
Dim HoldConfigErrs() As String
Dim HoldConfigIDs() As String
Dim ConfigMatched() As String, ConfigSelectedID() As Boolean  ' This will be our selected list.

Dim HoldConfigWkShts() As String, HoldXcelWkShts() As String
Dim HoldConfigHeadings() As String, HoldXcelHeadings() As String
Dim HoldOutputTableName() As String
Dim XcelMatched() As String, XcelMatchedVia() As String, XcelSelectedWkSht() As Boolean  ' This will be our selected list.

' Populate Config array data that Matches the selected file.
Call GetAllConfigs(HoldConfigData, HoldConfigIDs, HoldConfigExcel, _
                   HoldConfigWkShts, HoldConfigHeadings, HoldConfigErrs, HoldConfig, _
                   HoldOutputTableName, XslFileAndPath)
                   
Dim ConfigErrsPresent  As Boolean: ConfigErrsPresent = False
Dim SevereConfigErrsEncountered:   SevereConfigErrsEncountered = False
Dim WarningConfigErrsEncountered:  WarningConfigErrsEncountered = False
Dim AbortProcessFlag:  AbortProcessFlag = False
For II = LBound(HoldConfigErrs) To UBound(HoldConfigErrs)
  If HoldConfigErrs(II) <> "" Then ConfigErrsPresent = True
  If InStr(1, HoldConfigErrs(II), vbCrLf & "***") <> 0 Then SevereConfigErrsEncountered = True
  If InStr(1, HoldConfigErrs(II), vbCrLf & "Warning") <> 0 Then WarningConfigErrsEncountered = True
Next II

Dim ConfigFunc  As String
If Not ConfigPassedFromCaller And Not SevereConfigErrsEncountered Then
  ConfigFunc = vbCrLf & vbCrLf _
           & "*************  The BELOW Function can be copy/pasted into your ETL program to hardcode the Etl123 Config into your process..." _
           & vbCrLf & CreateConfigFunction(HoldConfigData, XslFileAndPath) _
           & vbCrLf & "*************  END of Function to copy/paste config into ETL123 program..."
  Call FileSave(ConfigFunc, ReptFileName, , True)
End If

Dim ConfigReport:  ConfigReport = "Config Elements have no errors or warnings reported."
If ConfigErrsPresent Then
  ConfigReport = vbCrLf & vbCrLf & "Config Elements - Beginning Error Report...................................." & vbCrLf
  For II = LBound(HoldConfigErrs) To UBound(HoldConfigErrs)
    If HoldConfigErrs(II) <> "" Then _
      ConfigReport = ConfigReport & HoldConfigErrs(II) & vbCrLf
  Next II
  ConfigReport = ConfigReport & "Config Elements - End of Configuration Error Report...................................."
End If
Call FileSave(ConfigReport & vbCrLf, ReptFileName, , True) ' Append to the report....
                 
' Populate eXcel array data pulled from the selected eXcel file.
Call GetAllXcel(HoldXcelWkShts, HoldXcelHeadings, XslFileAndPath)

' Check for Duplicate WkSht names.  Names are EXPECTED to be Unique.
Dim DupWkshtName As String
If FindFirstDupValue(HoldXcelWkShts, DupWkshtName) Then
  Call FileSave("*** WkSheet - """ & DupWkshtName & """ has been found multiple times.  WkSheet names must be unique.", ReptFileName, , True) ' Append to the report....
  SevereConfigErrsEncountered = True
End If

If SevereConfigErrsEncountered Then
  Call FileSave("Config errors present.  Process ABORT...", ReptFileName, , True) ' Append to the report....
  MsgBox ("Config errors present.  Etl_123Rept.txt will have details.  Process ABORT...")
  DebugPrintOn ("Config errors present.  Etl_123Rept.txt will have details.  Process ABORT...")
  GoTo EndOfEtl_123
End If

If WarningConfigErrsEncountered And Not IgnoreWarnings Then
  If Not CompileOnly Then
    Response = MsgBox("Config Warnings present.  Etl_123Rept.txt will have details.  " & vbCrLf _
                    & vbCrLf & "Do you want to ignore Warnings and continue?", vbYesNo)
    If Response = vbNo Then GoTo EndOfEtl_123
  End If
  If CompileOnly Then _
    Response = MsgBox("Config Warnings present.  Etl_123Rept.txt will have details.  ", vbOKOnly)
End If
                  
' Allocate new Arrays to select and match items.
For II = LBound(HoldConfigIDs) To UBound(HoldConfigIDs)
    Call InsertNewElementIntoArray(ConfigMatched, "")
    Call InsertNewElementIntoArray(ConfigSelectedID, False)
Next II
' Allocate new Arrays to select and match items.
For II = LBound(HoldXcelWkShts) To UBound(HoldXcelWkShts)
    Call InsertNewElementIntoArray(XcelMatched, "")
    Call InsertNewElementIntoArray(XcelMatchedVia, "")
    Call InsertNewElementIntoArray(XcelSelectedWkSht, False)
Next II

Call GetMatchedPairs(HoldConfigIDs, _
                     HoldConfig, _
                     HoldOutputTableName, _
                     HoldXcelWkShts, _
                     HoldConfigHeadings, _
                     HoldXcelHeadings, _
                     HoldConfigWkShts, _
                     XcelMatched, _
                     XcelMatchedVia, _
                     ConfigMatched, _
                     ReptFileName, _
                     IgnoreWarnings, _
                     AllowUnmatchedWarnings)

' Now Process the selected Config IDs and Excel WkShts....
' Result will be a list of Selected Matched Pairs for import processing.

' Edits....
' 1) Any selected Config should also be matched/paired.
' 2) Any selected WkSht should also be matched/paired.
'   2a) If *.csv then or SelectWkShts(??) = "*" then provide a warning of configs not processed.
' 3) Any selected Config should match to at least one Confid ID.
' 4) Any selected WkSht Element should match to at least one WkSht element.
' 5) Select All matched/pairs when SelectWkShtElemts  SelectConfigIDs are both = ""
'   5a) In this case, provide a warning of WkShts not processed, and configs unaccounted for.
' 6) The selected Excel WkSht should have one and only one Config ID match.


' Mark HoldXcelWkShts as selected based on Caller.

Dim Ptrs() As String, P As Integer
ReptData = ""
Dim Warnings:  Warnings = ""
' 1) Any selected Config should also be matched/paired.
For II = LBound(HoldConfigIDs) To UBound(HoldConfigIDs)
  For JJ = LBound(SelectIDs) To UBound(SelectIDs)
    If ConfigMatched(II) <> "" And _
       HoldConfigIDs(II) Like SelectIDs(JJ) Then ConfigSelectedID(II) = True
  Next JJ
  If ConfigSelectedID(II) Then
    Erase Ptrs
    Ptrs = Split(TrimChars(ConfigMatched(II), "/"), "/")
    If IsArrayAllocated(Ptrs) Then
      For JJ = LBound(Ptrs) To UBound(Ptrs)
        P = Ptrs(JJ)
        XcelSelectedWkSht(P) = True
      Next JJ
    End If
  End If
Next II

' 2) Any selected WkSht should also be matched/paired.
For II = LBound(HoldXcelWkShts) To UBound(HoldXcelWkShts)
  For JJ = LBound(SelectWkShts) To UBound(SelectWkShts)
    If XcelMatched(II) <> "" And _
       HoldXcelWkShts(II) Like SelectWkShts(JJ) Then XcelSelectedWkSht(II) = True
    If XcelMatched(II) <> "" And _
       SelectWkShts(JJ) = "All Matched Pairs" Then XcelSelectedWkSht(II) = True
  Next JJ
  If XcelSelectedWkSht(II) Then
    Erase Ptrs
    Ptrs = Split(TrimChars(XcelMatched(II), "/"), "/")
    If IsArrayAllocated(Ptrs) Then
      For JJ = LBound(Ptrs) To UBound(Ptrs)
        P = Ptrs(JJ)
        ConfigSelectedID(P) = True
      Next JJ
    End If
  End If
Next II

'   2a) If *.csv then or SelectWkShts(??) = "*" then provide a warning of configs not processed.
For II = LBound(SelectWkShts) To UBound(SelectWkShts)
  If SelectWkShts(II) = "*" Then
    For JJ = LBound(ConfigSelectedID) To UBound(ConfigSelectedID)
      If Not ConfigSelectedID(JJ) Then _
        Warnings = Warnings & "Warning - Config ID-""" & HoldConfigIDs(JJ) & """ was not processed. " _
                            & "Was not matched to any Excel WkSht." & vbCrLf
    Next JJ
  End If
Next II

' 3) Any selected Config should match to at least one Confid ID.
For II = LBound(SelectIDs) To UBound(SelectIDs)
  KK = 0
  For JJ = LBound(HoldConfigIDs) To UBound(HoldConfigIDs)
    If HoldConfigIDs(JJ) Like SelectIDs(II) Then
      KK = KK + 1
      ' 4) The selected Config ID should have at least one Excel WkSht match.
      If ConfigMatched(II) = "" Then _
        ReptData = ReptData & "*** SelectConfigIDs-""" & SelectIDs(II) & """ ID is not paired with any WkShts in the eXcel file." & vbCrLf
    End If
  Next JJ
  If KK = 0 Then
    ReptData = ReptData & "*** SelectConfigIDs-""" & SelectIDs(II) & """ ID was not found in the Config." & vbCrLf
  End If
Next II


' 5) Any selected Excel WkSht should match to at least one WkSht element.
Dim ID_List As String
For II = LBound(SelectWkShts) To UBound(SelectWkShts)
  If SelectWkShts(II) = "All Matched Pairs" Then Exit For
  KK = 0
  For JJ = LBound(HoldXcelWkShts) To UBound(HoldXcelWkShts)
    If HoldXcelWkShts(JJ) Like SelectWkShts(II) Then
      KK = KK + 1
      ' 6) The selected Excel WkSht should have one and only one Config ID match.
      If XcelMatched(JJ) = "" Then _
        ReptData = ReptData & "*** SelectWkShtElemts-""" & SelectWkShts(II) & """ WkSht is not paired with any IDs from the Config file." & vbCrLf
      If XcelMatched(JJ) <> "" Then
        Ptrs = Split(TrimChars(XcelMatched(JJ), "/"), "/")
        ' 6) The selected Excel WkSht should have only one Config ID match.
        If UBound(Ptrs) > 0 Then
          P = Ptrs(0)
          
      'Dim ListOfMatchedConfigs As String
      'ListOfMatchedConfigs = GetListOfMatchedConfigs(XcelMatched(JJ), HoldConfigIDs)
          Warnings = Warnings & "Warning - SelectWkShtElemts-""" & SelectWkShts(II) & """ WkSht is paired with Multiple IDs from the Config file.  ID-" & GetListOfMatchedConfigs(XcelMatched(JJ), HoldConfigIDs) & " will be used to process Import." & vbCrLf
        End If
      End If
    End If
  Next JJ
  If KK = 0 Then
    ReptData = ReptData & "*** SelectWkShtElemts-""" & SelectWkShts(II) & """ WkSht was not found in the eXcel file." & vbCrLf
  End If

Next II

Dim AllElmtSelWithAts As Boolean

' 7) Select "*" All WkShts.  Expect to find a matching pair.
AllElmtSelWithAts = False
For II = LBound(SelectWkShts) To UBound(SelectWkShts)
  If SelectWkShts(II) = "*" Then AllElmtSelWithAts = True
Next II
If AllElmtSelWithAts Then
  For II = LBound(HoldXcelWkShts) To UBound(HoldXcelWkShts)
    If XcelMatched(II) = "" Then _
      Warnings = Warnings & "Warning - SelectWkShtElemts-""*"" and WkSht-""" & HoldXcelWkShts(II) & """  Expected matching Config not found.  This WkSht will not be processed." & vbCrLf
  Next II
End If

' 8) Select "*" All Configs.  Expect to find a matching pair.
AllElmtSelWithAts = False
For II = LBound(SelectIDs) To UBound(SelectIDs)
  If SelectIDs(II) = "*" Then AllElmtSelWithAts = True
Next II
If AllElmtSelWithAts Then
  For II = LBound(HoldConfigIDs) To UBound(HoldConfigIDs)
    If ConfigMatched(II) = "" Then _
      Warnings = Warnings & "Warning - SelectConfigIDs-""*"" and ConfigID-""" & HoldConfigIDs(II) & """  Expected matching WkSht not found.  This ConfigID  will not be processed." & vbCrLf
  Next II
End If

'  Build the processing list for next Import Steps..........................
Dim ImportList() As String, aListItem As String  ' List of selected
Dim SelectionRept: SelectionRept = ""
' aListItem structure:  eXcelFilePath~ConfigID~eXcelWkSht#~eXcelWkShtName

' Loop thru all Selected WkShts and build an Import List for processing.....
Dim MatchedVia() As String
For II = LBound(XcelSelectedWkSht) To UBound(XcelSelectedWkSht)
  If XcelSelectedWkSht(II) Then
    Ptrs = Split(TrimChars(XcelMatched(II), "/"), "/")
    MatchedVia = Split(TrimChars(XcelMatchedVia(II), "/"), "/")
    For JJ = LBound(Ptrs) To UBound(Ptrs)
      P = Ptrs(JJ)  ' Use the first matched Config for processing...
      aListItem = XslFileAndPath & "~" & II + 1 & "~" & HoldXcelWkShts(II) & "~" & MatchedVia(JJ) & "~" & HoldConfigIDs(P) & "~" & HoldConfig(P)
      Call InsertNewElementIntoArray(ImportList, aListItem, True)
    Next JJ
  End If
  
Next II
If Not IsArrayAllocated(ImportList) Then _
    ReptData = ReptData & "*** No Files were selected for """ & XslFileAndPath & """." & vbCrLf



'ReptData = ReptData & "*** My error..." & vbCrLf

If ReptData = "" And Warnings = "" Then
    SelectionRept = SelectionRept & vbCrLf & "No Selection errors were found related to: (Parm #3-SelectWkShtElemts and Parm #4-SelectConfigIDs) " & vbCrLf
   ' SelectionRept = vbCrLf & "Beginning Analysis Report on Selected Matched Pairs for: " & XslFileAndPath & ".............................................................." & vbCrLf _
   '             & SelectionRept _
   '             & "End of Analysis Report on Selected Matched Pairs................................................................." & vbCrLf
    Call FileSave(SelectionRept, ReptFileName, , True) ' Append to the report....
End If

If ReptData <> "" Or Warnings <> "" Then
  If ReptData <> "" Then _
    ReptData = ReptData & "*** Severe Selection errors present. Process is ABORTed." & vbCrLf
  
  SelectionRept = vbCrLf & "Beginning Report on Selected Matched Pairs.............................................................." & vbCrLf _
                & Warnings & ReptData _
                & "End of Report on Selected Matched Pairs................................................................." & vbCrLf
             
  Call FileSave(SelectionRept, ReptFileName, , True) ' Append to the report....
  If ReptData <> "" Then
    MsgBox ("Severe Selection errors present.  Etl_123Rept.txt will have details.  Process Will ABORT...")
    GoTo EndOfEtl_123
  End If
  If Warnings <> "" And Not IgnoreWarnings Then
    Response = MsgBox("Selection Warnings present.  Etl_123Rept.txt will have details.  " & vbCrLf & vbCrLf & _
                      Warnings & vbCrLf & _
                      "Do you want to ignore Warnings and continue?", vbYesNo)
    If Response = vbNo Then GoTo EndOfEtl_123
  End If
End If
 
' Now process the items in the Import List one at a time.
Dim ImpListElmts() As String, ImportResult As String
'Call DelTbl("imported_table_errors")
For II = LBound(ImportList) To UBound(ImportList)
  ImpListElmts = Split(ImportList(II), "~")
  ImportResult = ProcessOneWkShtInput(II, UBound(ImportList), _
                                      ImpListElmts(0), _
                                      ImpListElmts(1), _
                                      ImpListElmts(2), _
                                      ImpListElmts(3), _
                                      ImpListElmts(4), _
                                      ImpListElmts(5), _
                                      ReptFileName, _
                                      CompileOnly)
                                    

Next II

EndOfEtl_123:

If NewConfigCreated Then _
  Call MsgBox("Create NEW Config is COMPLETE." & vbCrLf & vbCrLf & _
              "Check Etl_123Rept.txt report for results of the Config compile.")
              
Dim BeginningTime  As Date, ElapsedEtl123Time
ElapsedEtl123Time = ElapsedTimeSince("Etl123", , True, , BeginningTime) ' Reset the Etl123 Timer....
              
JJ = vbCrLf & ReptDateTime & "  Etl123 process is complete.   Total Elapsed time (mm:ss = " & ElapsedEtl123Time & ")"
Call FileSave(JJ, ReptFileName, , True) ' Append to the report....
             

End Function

Private Function GetListOfMatchedConfigs(ByVal Matched As String, _
                                         ByVal HoldConfigIDs) As String
Dim Ptrs() As String, P As Long
Ptrs = Split(TrimChars(Matched, "/"), "/")
Dim II, JJ, XX
'ID-"Activity1" will be used to process Import.

XX = ""
For II = LBound(Ptrs) To UBound(Ptrs)
  P = Ptrs(II)
  XX = XX & """" & HoldConfigIDs(P) & """, "
  If UBound(Ptrs) > 0 And II = UBound(Ptrs) - 1 Then XX = Left(XX, Len(XX) - 2) & " and "
Next II
If Right(XX, 2) = ", " Then XX = Left(XX, Len(XX) - 2)
If UBound(Ptrs) > 0 Then XX = "(" & XX & ")"
GetListOfMatchedConfigs = XX

End Function



Private Sub GetAllXcel(ByRef HoldXcelWkShts() As String, _
                       ByRef HoldHeadings() As String, _
                       ByVal XslFileAndPath As String)
                         
Dim aConfig As String

Dim II, JJ, KK
Dim XX As String
Dim ZZ As String, YY As String
Dim HeadingArr As Variant
Dim HeadingStr As String, HoldWkShtNumber As String


HoldXcelWkShts = GetExcelWkShtNames(XslFileAndPath, True)

For II = LBound(HoldXcelWkShts) To UBound(HoldXcelWkShts)
  HoldWkShtNumber = II + 1
  HeadingArr = GetExcelHeadingText(XslFileAndPath, II + 1, True)
  'Clean Up Headings (Remove parms to right of : delimiter) and build HeadingStr
  HeadingStr = ""
  If IsArrayAllocated(HeadingArr) Then
    For KK = LBound(HeadingArr) To UBound(HeadingArr)
      HeadingArr(KK) = RemoveWhiteSpace(HeadingArr(KK))
      HeadingStr = HeadingStr & HeadingArr(KK) & "~"
    Next KK
  End If
  
  HeadingStr = TrimChars(HeadingStr, "~")
  
  Call InsertNewElementIntoArray(HoldHeadings, HeadingStr)
Next II

                          
End Sub

Private Sub GetAllConfigs(ByVal CompleteConfig As String, _
                          ByRef HoldConfigIDs() As String, _
                          ByRef HoldConfigExcel() As String, _
                          ByRef HoldConfigWkSht() As String, _
                          ByRef HoldHeadings() As String, _
                          ByRef HoldConfigErrs() As String, _
                          ByRef HoldConfig() As String, _
                          ByRef HoldOutputTableName() As String, _
                          ByVal MatchToThisFileName)
                          
                         
                         
Dim aConfig As String

Dim II, JJ, KK
Dim XX As String: XX = CompleteConfig
Dim ZZ As String, YY As String

JJ = InStrRev(MatchToThisFileName, "\")
If JJ <> 0 Then MatchToThisFileName = Mid(MatchToThisFileName, JJ + 1)

XX = Check4ValidKeyWordChars(XX)  ' Remove invalid Parentheses

'Get all of the FileName/WkSht configs and populate HoldConfig array.
Do
  ZZ = ExtractDelimitedStr(XX, "FileName=(")  ' Each section of config elements begins with this string.
  If ZZ = "" Then Exit Do
  
  ' Should this Config be captured?
  YY = ZZ
  Dim HoldExcelFilePattern
  HoldExcelFilePattern = ExtractKeyWord(YY, "Excel=")
  
  If MatchToThisFileName <> "" And MatchToThisFileName <> "*" And _
        Not (MatchToThisFileName Like HoldExcelFilePattern) Then GoTo ContinueDoLoop  ' Don't Capture
  
  Call InsertNewElementIntoArray(HoldConfig, ZZ)
  Dim ParseImportConfigErrs: ParseImportConfigErrs = ParseImportConfig(ZZ)
  'RT_FileID
  Dim FoundDup:  FoundDup = False
  If IsArrayAllocated(HoldConfigIDs) Then
    For II = LBound(HoldConfigIDs) To UBound(HoldConfigIDs)
      If HoldConfigIDs(II) = RT_FileID Then FoundDup = True
    Next II
  End If
  DebugPrintOn (ParseImportConfigErrs)
  If FoundDup Then ParseImportConfigErrs = ParseImportConfigErrs & vbCrLf & _
    "*** FileName=(ID=""" & RT_FileID & """) is a duplicate of another FileName Element." & vbCrLf
  
  Call InsertNewElementIntoArray(HoldConfigErrs, ParseImportConfigErrs)
  
  'II = UBound(HoldConfigErrs)
  'If HoldConfigErrs(II) <> "" Then
  '  DebugPrintOn (II & " --Top-------------------------------------------------------------------------------")
  '  DebugPrintOn ("Config=" & HoldConfig(II))
  '  DebugPrintOn ("Errors=" & HoldConfigErrs(II))
  '  DebugPrintOn (II & " --Bottom-------------------------------------------------------------------------------")
  'End If
  
  YY = ""
  If IsArrayAllocated(RT_ExcelHeadingText) Then
    For II = LBound(RT_ExcelHeadingText) To UBound(RT_ExcelHeadingText)
      YY = YY & RT_ExcelHeadingText(II) & "~"
    Next II
  End If
  Call InsertNewElementIntoArray(HoldHeadings, YY)
  
  Call InsertNewElementIntoArray(HoldConfigIDs, RT_FileID)
  Call InsertNewElementIntoArray(HoldConfigExcel, RT_FileName)
  Call InsertNewElementIntoArray(HoldConfigWkSht, RT_WorkSheetName)
  Call InsertNewElementIntoArray(HoldOutputTableName, RT_OutputTableName)
ContinueDoLoop:
Loop

                          
End Sub




Private Function EquivalentHeadings(ByVal HdrsFrmConfig As String, _
                                    ByVal HdrsFrmFileWkSht As String) As String

Dim Config() As String, FileWkSht() As String
Config = Split(TrimChars(HdrsFrmConfig, "~"), "~")  ' Trim Nulls
FileWkSht = Split(TrimChars(HdrsFrmFileWkSht, "~"), "~")


' Every Config Header must be accounted for in the File Headings for headings to be Equilivalent
' Pattern matchs are checked after all Equal matchs to ensure the most exact match.

Dim II, JJ, Matched

'Remove Whitespace from Headings...
For II = LBound(Config) To UBound(Config)
  Config(II) = RemoveWhiteSpace(Config(II))
  If Config(II) = "" Then Config(II) = "??Empty??"
Next II
For II = LBound(FileWkSht) To UBound(FileWkSht)
  FileWkSht(II) = RemoveWhiteSpace(FileWkSht(II))
  If FileWkSht(II) = "" Then FileWkSht(II) = "??Empty??"
Next II

' Every item in Config must be found in FileWkSht to be equivalent (Unless it is empty)
For II = LBound(Config) To UBound(Config)
  ' First look for ALL Equal Matched Pairs...
    For JJ = LBound(FileWkSht) To UBound(FileWkSht)
    If Config(II) = FileWkSht(JJ) And Config(II) <> "??Empty??" And Config(II) <> "??Matched??" Then
      FileWkSht(JJ) = "??Matched??"
      Config(II) = "??Matched??"
      Exit For
    End If
  Next JJ
Next II

' Now go and look for the pattern matches.
For II = LBound(Config) To UBound(Config)
  For JJ = LBound(FileWkSht) To UBound(FileWkSht)
    If FileWkSht(JJ) Like Config(II) And Config(II) <> "??Empty??" And Config(II) <> "??Matched??" Then
      FileWkSht(JJ) = "??Matched??"
      Config(II) = "??Matched??"
      Exit For
    End If
  Next JJ
Next II

EquivalentHeadings = ""
For II = LBound(Config) To UBound(Config)
  If Config(II) <> "??Empty??" And Config(II) <> "??Matched??" Then
    EquivalentHeadings = EquivalentHeadings & "Heading=""" & Config(II) & """ and "
  End If
Next II
If EquivalentHeadings <> "" Then EquivalentHeadings = Mid(EquivalentHeadings, 1, Len(EquivalentHeadings) - 4) & "are unaccounted for."

End Function

Private Sub GetMatchedPairs(ByRef HoldConfigIDs() As String, _
                           ByRef HoldConfig() As String, _
                           ByRef HoldOutputTableName() As String, _
                           ByRef HoldXcelWkShts() As String, _
                           ByRef HoldConfigHeadings() As String, _
                           ByRef HoldXcelHeadings() As String, _
                           ByRef HoldConfigWkShts() As String, _
                           ByRef XcelMatched() As String, _
                           ByRef XcelMatchedMethod() As String, _
                           ByRef ConfigMatched() As String, _
                           ByVal ReptFileName As String, _
                           ByVal IgnoreWarnings As Boolean, _
                           ByVal AllowUnmatchedWarnings As Boolean)

Dim II, JJ, Response
'Dim MatchedOn As String
Dim MatchedReport: MatchedReport = ""

Dim MatchedWkShtCount:  MatchedWkShtCount = 0
Dim WkshtPlans As String:  WkshtPlans = ""

Dim MatchedHdrCount:  MatchedHdrCount = 0
Dim HdrPlans As String:  HdrPlans = ""

Dim UnMatchedConfigIDs: UnMatchedConfigIDs = 0
Dim ConfigUnMatched:  ConfigUnMatched = ""

Dim UnMatchedXcelWkShts:  UnMatchedXcelWkShts = 0
Dim WkShtUnMatched:  WkShtUnMatched = ""

Dim MatchedOnWkSht  As Boolean: MatchedOnWkSht = False
Dim MatchedOnHdrs  As Boolean: MatchedOnHdrs = False

Dim ConfigWkShts() As String, aConfig
' Extract WkSht keywords
For II = LBound(HoldConfigIDs) To UBound(HoldConfigIDs)
  aConfig = HoldConfig(II)
  Call InsertNewElementIntoArray(ConfigWkShts, ExtractKeyWord(aConfig, "WkSht=", "WS"))
Next II

Dim EquHdrs As String
' Match up Xcel WkShts from the Picked/Selected file to the Configurations
For II = LBound(HoldConfigIDs) To UBound(HoldConfigIDs)
  For JJ = LBound(HoldXcelWkShts) To UBound(HoldXcelWkShts)
  
    If EquivalentHeadings(HoldConfigHeadings(II), HoldXcelHeadings(JJ)) = "" And Trim(HoldConfigWkShts(II)) = "" Then _
      MatchedOnHdrs = True Else MatchedOnHdrs = False
    If HoldXcelWkShts(JJ) Like HoldConfigWkShts(II) Then _
      MatchedOnWkSht = True Else MatchedOnWkSht = False
    If MatchedOnWkSht Then MatchedOnHdrs = False  ' Ensure that both are not True...
      
    If MatchedOnWkSht Or MatchedOnHdrs Then
      If MatchedOnHdrs Then _
        XcelMatchedMethod(JJ) = XcelMatchedMethod(JJ) & "Headers" & "/"
      If MatchedOnWkSht Then _
        XcelMatchedMethod(JJ) = XcelMatchedMethod(JJ) & "WorkSheet Name" & "/"
      XcelMatched(JJ) = XcelMatched(JJ) & II & "/"
      ConfigMatched(II) = ConfigMatched(II) & JJ & "/"
    End If
    If MatchedOnWkSht Then
      MatchedWkShtCount = MatchedWkShtCount + 1
      WkshtPlans = FormatMatchedPairs(WkshtPlans, HoldXcelWkShts(JJ), _
                                      HoldConfigIDs(II), HoldOutputTableName(II))
    End If
    If MatchedOnHdrs Then
      MatchedHdrCount = MatchedHdrCount + 1
      HdrPlans = FormatMatchedPairs(HdrPlans, HoldXcelWkShts(JJ), _
                                    HoldConfigIDs(II), HoldOutputTableName(II))
    End If
  Next JJ
Next II

' Now report on elements that were not matched....
ConfigUnMatched = ""
For II = LBound(ConfigMatched) To UBound(ConfigMatched)
  If ConfigMatched(II) = "" Then
    JJ = II
    If ConfigWkShts(JJ) = "" Then ConfigWkShts(JJ) = "Null"
    ConfigUnMatched = FormatMatchedPairs(ConfigUnMatched, ConfigWkShts(JJ), _
                                      HoldConfigIDs(JJ), HoldOutputTableName(JJ))
    UnMatchedConfigIDs = UnMatchedConfigIDs + 1
  End If
Next II
ConfigUnMatched = "Warning - " & ConfigUnMatched & " not matched to ANY WorkSheets."


WkShtUnMatched = ""
For II = LBound(XcelMatched) To UBound(XcelMatched)
  If XcelMatched(II) = "" Then
    WkShtUnMatched = WkShtUnMatched & """" & HoldXcelWkShts(II) & """, "
    UnMatchedXcelWkShts = UnMatchedXcelWkShts + 1
  End If
Next II
If Len(WkShtUnMatched) > 0 Then
  WkShtUnMatched = Left(WkShtUnMatched, Len(WkShtUnMatched) - 2)
  II = InStrRev(WkShtUnMatched, ", ")
  If II > 0 Then WkShtUnMatched = _
    Mid(WkShtUnMatched, 1, II - 1) & " and " & Mid(WkShtUnMatched, II + 2)
  If UnMatchedXcelWkShts > 1 Then WkShtUnMatched = "(" & WkShtUnMatched & ")"
  WkShtUnMatched = "Warning - " & WkShtUnMatched & " not matched to ANY Config IDs."
End If

Dim UnMatchedConfigMsg, UnMatchedWkShtMsg

If MatchedWkShtCount > 0 Then _
  MatchedReport = MatchedReport & Rpad("WorkSheets Matched on WorkSheet Name ", 38, "-") & " " & MatchedWkShtCount & "   " & WkshtPlans & vbCrLf
If MatchedHdrCount > 0 Then _
  MatchedReport = MatchedReport & Rpad("WorkSheets Matched on Headings ", 38, "-") & " " & MatchedHdrCount & "   " & HdrPlans & vbCrLf
If UnMatchedConfigIDs > 0 Then
  UnMatchedConfigMsg = Rpad("UnMatched Config IDs ", 38, "-") & " " & UnMatchedConfigIDs & "   " & ConfigUnMatched & vbCrLf
  MatchedReport = MatchedReport & UnMatchedConfigMsg
End If
If UnMatchedXcelWkShts > 0 Then
  UnMatchedWkShtMsg = Rpad("UnMatched eXcel WkShts ", 38, "-") & " " & UnMatchedXcelWkShts & "   " & WkShtUnMatched & vbCrLf
  MatchedReport = MatchedReport & UnMatchedWkShtMsg
End If

Call FileSave(vbCrLf & "WkSht-->Config-ID Matched Summary ----:    (WkSheet-->""Config-ID""-->DB Table)", ReptFileName, , True) ' Append to the report....
MatchedReport = TrimChars(MatchedReport, vbCrLf)
Call FileSave(MatchedReport, ReptFileName, , True) ' Append to the report....
'Call FileSave("Ending of Analysis for Matched Pairs.", ReptFileName, , True) ' Append to the report....
If (MatchedHdrCount + MatchedWkShtCount) = 0 Then
  MatchedReport = MatchedReport & "No Matched Pairs of WkSht/Configs were found.  Process will ABORT."
  Call FileSave(MatchedReport, ReptFileName, , True) ' Append to the report....
  MsgBox ("WkSht Matching errors present.  Etl_123Rept.txt will have details.  Process Will ABORT...")
  End
End If

If (UnMatchedConfigIDs + UnMatchedXcelWkShts) > 0 Then
  MatchedReport = ""
  If UnMatchedConfigIDs > 0 Then MatchedReport = MatchedReport & UnMatchedConfigMsg & vbCrLf
  If UnMatchedXcelWkShts > 0 Then MatchedReport = MatchedReport & UnMatchedWkShtMsg & vbCrLf
  If Not AllowUnmatchedWarnings Then
    Response = MsgBox("UnMatched Wksht-->Config  Warnings are found:  " & vbCrLf & vbCrLf & _
                      "Parm #8 - AllowUnmatchedWarnings=" & AllowUnmatchedWarnings & _
                      "  so these will be treated as Severe issues:" & _
                      vbCrLf & vbCrLf & MatchedReport & _
                      "Etl_123Rept.txt will have details.  Process Will ABORT...")
    End
  End If
  
  If Not IgnoreWarnings Then
    Response = MsgBox("UnMatched Wksht-->Config  Warnings are found:" & vbCrLf & vbCrLf & MatchedReport & _
                      "Parm #8 - AllowUnmatchedWarnings=" & AllowUnmatchedWarnings & _
                      "  will allow you to continue." & vbCrLf & vbCrLf & _
                      "Do you want to want to continue?  ""No"" will abort entire process.", vbYesNo)
    If Response = vbNo Then End
  End If
End If

End Sub

Private Function FormatMatchedPairs(ByVal FormatedPairs As String, _
                            ByVal WkshtTab As String, _
                            ByVal ConfigID As String, _
                            ByVal DbTable As String) As String
                            
'Number of Tab Name Matched Pairs - 2  (Activity-->"Config-1"-->Activity)   (Activity2-->"Config-2"-->Activity)
'Number of Headings Matched Pairs - 2
 Dim XX
 
 WkshtTab = Trim(WkshtTab)
 ConfigID = Trim(ConfigID)
 DbTable = Trim(DbTable)
 
 If Contains(WkshtTab, """'- ") <> 0 Then WkshtTab = """" & WkshtTab & """"
 If Contains(ConfigID, """'- ") <> 0 Then ConfigID = """" & ConfigID & """"
 If Contains(DbTable, """'- ") <> 0 Then DbTable = """" & DbTable & """"
 XX = "(" & WkshtTab & "-->" & ConfigID & "-->" & DbTable & ")"
 FormatMatchedPairs = FormatedPairs & XX & "   "
                            
End Function



Private Function Check4ValidKeyWordChars(ByVal XX As String) As String

' This subroutine is responsible for verifying that certain keywords do not contain reserved characters that
' will confuse the parsing of FileName and/or Column components.

Dim II, JJ, KK, LL, MM, ZZ, Invalids() As String, KWrds() As String
Dim Response
II = "WkSht,WS,Excel,Heading,H": KWrds = Split(II, ",")

ZZ = XX    ' Save original contents of config string.
Check4ValidKeyWordChars = ZZ  ' Initialize the return.

For JJ = LBound(KWrds) To UBound(KWrds)
  Do
    KK = ExtractKeyWord(XX, KWrds(JJ))
    If KK = "" Then Exit Do
    II = InStr(1, KK, ")")
    If II = 0 Then II = InStr(1, KK, "(")
'    If II = 0 Then II = InStr(1, KK, """")
    If II <> 0 Then
      MM = KWrds(JJ) & "=""" & KK & """"
      If InStr(1, ZZ, MM) <> 0 Then Call InsertNewElementIntoArray(Invalids, MM)
      MM = KWrds(JJ) & "=" & KK
      If InStr(1, ZZ, MM) <> 0 Then Call InsertNewElementIntoArray(Invalids, MM)
    End If
  Loop
Next JJ

' Now clean up this config string.
If Not IsArrayAllocated(Invalids) Then GoTo EndOfFunction ' No invalids were found.

Dim msg: msg = ""
For II = LBound(Invalids) To UBound(Invalids)
  msg = msg & Invalids(II) & vbCrLf
Next II
msg = "Config contains the following KeyWord values that have reserved characters "" ( ) "".  " & vbCrLf & vbCrLf _
    & msg & vbCrLf & vbCrLf & "ETL123 will automatically change these to ""*"" if you continue." & vbCrLf _
    & "Do you want to Continue?  No will ABORT the process."
Response = MsgBox(msg, vbYesNo)
If Response = vbNo Then
  MsgBox ("Process is ABORTED at user request.")
  End
End If
  
For II = LBound(Invalids) To UBound(Invalids)
  Do
    JJ = InStr(1, ZZ, Invalids(II))
    If JJ = 0 Then Exit Do
    LL = Len(Invalids(II))
    ZZ = Left(ZZ, JJ - 1) & Translate(Mid(ZZ, JJ, LL), "**", "()") & Mid(ZZ, JJ + LL)
  Loop
Next II

EndOfFunction:
  Check4ValidKeyWordChars = ZZ  ' Initialize the return.
 ' Debug.Print (ZZ)
End Function

Private Sub DebugPrintOn(msg As String)
  Call DebugPrint(msg, True)
End Sub


Private Sub DebugPrint(msg As String, Optional PrintOn As Boolean = False)

If PrintOn Then
  Debug.Print (ElapsedTimeSince("Debug.Print", , True) & " " & msg)
  Exit Sub
End If

If GblForceOffPrintDebugLog Then
  GblTogglePrintDebugLog = False
  GblForceOnPrintDebugLog = False
  GblForceOffPrintDebugLog = False
  aPrintDebugLog = False
  DebugPrintOn ("GblForceOffPrintDebugLog is true.  Debug.Print is now turned OFF.")
End If

If GblForceOnPrintDebugLog Then
  GblTogglePrintDebugLog = False
  GblForceOnPrintDebugLog = False
  GblForceOffPrintDebugLog = False
  aPrintDebugLog = True
  DebugPrintOn ("GblForceOnPrintDebugLog is true.  Debug.Print is now turned ON.")
End If

If GblTogglePrintDebugLog Then
  GblTogglePrintDebugLog = False
  GblForceOnPrintDebugLog = False
  GblForceOffPrintDebugLog = False
  If aPrintDebugLog Then
    aPrintDebugLog = False
    DebugPrintOn ("GblTogglePrintDebugLog is true.  Debug.Print is now turned OFF.")
  Else
    aPrintDebugLog = True
    DebugPrintOn ("GblTogglePrintDebugLog is true.  Debug.Print is now turned ON.")
  End If
End If
 
If aPrintDebugLog Then
  Debug.Print (ElapsedTimeSince("Debug.Print", , True) & " " & msg)
  Exit Sub
End If

End Sub


Private Function ProcessOneWkShtInput(ByVal OneOf As Long, _
                                   ByVal TotProcesses As Long, _
                                   ByVal ImpFile As String, _
                                   ByVal WkShtNum As Long, _
                                   ByVal WkShtName As String, _
                                   ByVal MatchedVia As String, _
                                   ByVal ConfigID As String, _
                                   ByVal ConfigText As String, _
                                   ByVal OutPutReptName As String, _
                                   ByVal CompileOnly As Boolean) As String
Dim II, JJ, KK, ZZ, Response
'**** Import Processing Plan (#1 of 3) for  (Activity-->"Config-1"-->Activity)  (Matched Via WorkSheet Name)    "G:\My Drive\Joel

'**** Import Processing (#2 of 3) for "G:\My Drive\Joel's Files\Work\LogErrors\ActTest.xlsm"

'Step 1-2-3 Processing Plan for (2:Activity2-->"Config-2"-->Activity) --  Matched Via the WorkSheet Name:
'Step 1-2-3 Processing Plan for eXcel Worksheet-"2:Activity2" and DB Table-"Activity" data import:

KK = "(#" & OneOf + 1 & " of " & TotProcesses + 1 & ")"
                                   
Dim result As String
result = ParseImportConfig(ConfigText)
If InStr(1, result, "***") = 0 Then result = ""  ' Look for any severe errors in the messages.
If result <> "" Then  ' There should be no Severe errors expected here.
  MsgBox ("ProcessOneWkShtInput - Unexpected " & vbCrLf & "Result=""" & result & """  from ParseImportConfig")
  End
End If

Dim ProcessPlanErrs:  ProcessPlanErrs = ""
Dim Rept:
Rept = vbCrLf & "------------------------------------------------------------------------------------------------" & _
                "------------------------------------------------------------------------------------------------" & _
       vbCrLf & "**** Import Processing Plan " & KK & " for " _
     & " ConfigID=""" & ConfigID & """   """ & ImpFile & """" & vbCrLf
Rept = vbCrLf & "------------------------------------------------------------------------------------------------" & _
                "------------------------------------------------------------------------------------------------" & _
       vbCrLf & "**** Import Processing Plan " & KK & " for " _
     & """" & ImpFile & """" & vbCrLf
     
Call FileSave(Rept, OutPutReptName, , True) ' Append to the report....
'Import List:  WkSht-"1:Category"  ConfigID-"Testing2Tabs.xlsx/Category"  "G:\My Drive\Joel's Files\Work\Testing2Tabs.xlsx"
'  SelectionRept = vbCrLf & "Beginning Report on Selected Matched Pairs.............................................................." & vbCrLf _
'                & Warnings & ReptData _
'                & "End of Report on Selected Matched Pairs................................................................." & vbCrLf
'
'  Call FileSave(SelectionRept, ReptFileName, , True) ' Append to the report....


' Step #1 is to assign column numbers to all RT_ExcelColumnLetter variables to match the column headings
'         found in the spreadsheet.  In the process, ensure that ALL WkSht Headings match as expected.
Dim WkShtHeadings() As String:
WkShtHeadings = GetExcelHeadingText(ImpFile, WkShtNum, True)
' Step #1a Fill in any blank Headings if user has supplied column letter.
For II = LBound(RT_ExcelHeadingText) To UBound(RT_ExcelHeadingText)
  If Trim(RT_ExcelHeadingText(II)) = "" And Trim(RT_ExcelColumnLetter(II)) <> "" Then
    JJ = xlColNum(RT_ExcelColumnLetter(II)) - 1
    If JJ <= UBound(WkShtHeadings) Then _
      RT_ExcelHeadingText(II) = WkShtHeadings(JJ)
  End If
Next II

' Step #1b - Verify that those config columns that have specific columns match properly to the heading text.
For II = LBound(RT_ExcelHeadingText) To UBound(RT_ExcelHeadingText)
  JJ = 0
  If RT_ExcelColumnLetter(II) <> "" Then
    KK = xlColNum(RT_ExcelColumnLetter(II)) - 1
    If Not (WkShtHeadings(KK) Like RT_ExcelHeadingText(II)) And Trim(RT_ExcelHeadingText(II)) <> "" Then
      ProcessPlanErrs = ProcessPlanErrs _
         & "*** Error - Col=" & RT_ExcelColumnLetter(II) & " and Heading=""" & WkShtHeadings(KK) & """ does not match the expected Config Heading=""" & RT_ExcelHeadingText(II) & """" & vbCrLf
    End If
  End If
Next II
RT_ExcelColumnLetterOrig = RT_ExcelColumnLetter
If ProcessPlanErrs <> "" Then GoTo SkipStep1 ' Go report this error and skip the rest.

' *** Error - Worksheet contains 750 columns. Maximum number supported is 700.
If UBound(WkShtHeadings) + 1 > 700 Then _
  ProcessPlanErrs = ProcessPlanErrs _
    & "*** Error - Worksheet contains " & UBound(WkShtHeadings) + 1 & " columns. Maximum number supported is 700." & vbCrLf
If ProcessPlanErrs <> "" Then GoTo SkipStep1 ' Go report this error and skip the rest.

' *** Error - ("C:Post Date" and "G:Post Date") were found in the eXcel input file.  "ID-C" for "Post Date" must identify Col= to identify the correct one.
For II = LBound(RT_ExcelHeadingText) To UBound(RT_ExcelHeadingText)
  KK = 0:  ZZ = ""
    For JJ = LBound(WkShtHeadings) To UBound(WkShtHeadings)
      If WkShtHeadings(JJ) = RT_ExcelHeadingText(II) Then
        KK = KK + 1
        ZZ = ZZ & """" & xlColAlfa(JJ + 1) & ":" & WkShtHeadings(JJ) & """ and "
      End If
    Next JJ
  
  If Right(ZZ, 5) = " and " Then ZZ = Left(ZZ, Len(ZZ) - 5)
  ZZ = "(" & ZZ & ")"

  If KK > 1 And Trim(RT_ExcelColumnLetter(II)) = "" Then
    ProcessPlanErrs = ProcessPlanErrs _
          & "*** Error - " & ZZ & " were found in the eXcel input file.  """ & RT_ID(II) & """ for """ & RT_ExcelHeadingText(II) & """ must identify Col= to map the correct one." & vbCrLf
  End If
Next II
If ProcessPlanErrs <> "" Then GoTo SkipStep1 ' Go report this error and skip the rest.

' Step #1c - Assign proper Col numbers to all RT_ExcelColumnLetter that do not already have them.
Dim Col
RT_ExcelColumnLetterOrig = RT_ExcelColumnLetter
For II = LBound(RT_ExcelHeadingText) To UBound(RT_ExcelHeadingText)
  JJ = 0
  If RT_ExcelColumnLetter(II) = "" Then
    For KK = LBound(WkShtHeadings) To UBound(WkShtHeadings)
      Col = xlColAlfa(KK + 1)
      If WkShtHeadings(KK) Like RT_ExcelHeadingText(II) And FindPosition(RT_ExcelColumnLetter, Col) < 0 Then
        RT_ExcelColumnLetter(II) = Col
        Exit For
      End If
    Next KK
  End If
Next II



' Step #1d - Confirm that all config columns match properly with the heading text and that are no
'            duplicate columns in the RT_ExcelColumnLetter list.
For II = LBound(RT_ExcelColumnLetter) To UBound(RT_ExcelColumnLetter)
  If RT_ExcelColumnLetter(II) = "" Then _
      ProcessPlanErrs = ProcessPlanErrs _
         & "*** Error - Config ID=""" & RT_ID(II) & """ Heading=""" & RT_ExcelHeadingText(II) & """ does not match any Excel Headings. Column Letter is unresolved. Possible duplicate headings present." & vbCrLf
  
  KK = FindNonUniquePos(RT_ExcelColumnLetter, RT_ExcelColumnLetter(II))
  If RT_ExcelColumnLetter(II) <> "" And KK >= 0 Then _
      ProcessPlanErrs = ProcessPlanErrs _
         & "*** Error - Config ID=""" & RT_ID(II) & """ Heading=""" & RT_ExcelHeadingText(II) & """ " _
                   & "Config ID=""" & RT_ID(KK) & """ Heading=""" & RT_ExcelHeadingText(KK) _
                   & """ have duplicated column letter """ & RT_ExcelColumnLetter(II) & "1""."
         
Next II

' Step #1e - Verify that ALL config columns that have specific columns match properly to the heading text.
For II = LBound(RT_ExcelHeadingText) To UBound(RT_ExcelHeadingText)
  JJ = 0
  If RT_ExcelColumnLetter(II) <> "" Then
    KK = xlColNum(RT_ExcelColumnLetter(II)) - 1
    If Not (WkShtHeadings(KK) Like RT_ExcelHeadingText(II)) And Trim(RT_ExcelHeadingText(II)) <> "" Then
      ProcessPlanErrs = ProcessPlanErrs _
         & "*** Error - Col=" & RT_ExcelColumnLetter(II) & " and Heading=""" & WkShtHeadings(KK) & """ does not match the expected Config Heading=""" & RT_ExcelHeadingText(II) & """" & vbCrLf
    End If
  End If
Next II

SkipStep1:

'  Articulate the Processing Plan...........

Dim Step1Plan As String
JJ = ""
For II = LBound(RT_KeyField) To UBound(RT_KeyField)
  ZZ = RT_ExcelHeadingText(II)
  If RT_ExcelColumnLetterOrig(II) <> "" Then ZZ = RT_ExcelColumnLetterOrig(II) & ":" & ZZ
  If RT_KeyField(II) Then
    'If Left(JJ, 5) = " and " Then JJ = ", " & Mid(JJ, 5)
    JJ = JJ & """" & ZZ & """, "
  End If

Next II
Dim KeyFieldsArePresent As Boolean
If JJ = "" Then
  KeyFieldsArePresent = False
  Step1Plan = "NO Key fields have been defined.  All rows will be APPENDed to the table." & vbCrLf
 Else
  KeyFieldsArePresent = True
  If Right(JJ, 2) = ", " Then JJ = Left(JJ, Len(JJ) - 2)
  KK = InStrRev(JJ, """, ")
  If KK <> 0 Then JJ = Left(JJ, KK) & " and" & Mid(JJ, KK + 2)
  If RT_UniqueKey Then Step1Plan = "A Unique Key for " & JJ Else Step1Plan = "A NON-Unique Key for " & JJ
  Step1Plan = Step1Plan & " is defined." & vbCrLf
End If

Dim ImpPlan As String

ImpPlan = "Step 1-2-3 Processing Plan for (" & WkShtNum & ":" & FrameQuotes(WkShtName) & "-->" & FrameQuotes(ConfigID) & _
          "-->" & FrameQuotes(RT_OutputTableName) & ") --  Matched Via the """ & MatchedVia & """:" & vbCrLf
        
ImpPlan = ImpPlan & "  Step 1 ----  (Match Key Plan)" _
        & vbCrLf & "     " & Step1Plan

ImpPlan = ImpPlan & "  Step 2 ----  (Update Fields Plan for existing Rows)" _
          & vbCrLf & GetUpdatePlan(KeyFieldsArePresent)

ImpPlan = ImpPlan & "  Step 3 ----  (Error Handling Plan)" _
          & vbCrLf & GetErrorHandlePlan

ImpPlan = ImpPlan & "  Mapping Plan -----------" _
          & vbCrLf & GetMappingPlan(ImpFile, WkShtName)
          
ImpPlan = ImpPlan & "  Save the Date Changed Plan -----------" _
          & vbCrLf & GetSaveDatePlan(ImpFile, WkShtName)
          
If CompileOnly Then ImpPlan = ImpPlan & "CompileOnly=True,  Actual Import will be SKIPPED."
          
If ProcessPlanErrs = "" Then
  ImpPlan = ImpPlan & vbCrLf & "No Processing Plan Errors Reported.." & vbCrLf
  Call FileSave(ImpPlan, OutPutReptName, , True)
 Else
  ProcessPlanErrs = "Processing Plan Errors: " & vbCrLf & ProcessPlanErrs & vbCrLf
  ImpPlan = ImpPlan & vbCrLf & ProcessPlanErrs & vbCrLf
  Call FileSave(ImpPlan, OutPutReptName, , True)
  Response = MsgBox("ProcessOneWkShtInput - Unexpected Plan error for Worksheet """ & WkShtName & """" & vbCrLf & vbCrLf & _
          ProcessPlanErrs & vbCrLf & vbCrLf & _
          "Worksheet """ & WkShtName & """ will be skipped and not processed." & vbCrLf & _
          "CANCEL will ABORT whole process..", vbOKCancel)
  If Response = 1 Then ' Yes...
    Call FileSave("Worksheet skipped due to Plan Errors.", OutPutReptName, , True)
    Exit Function
  End If
  
  ' Cancel.....
   Call FileSave("User chose to ABORT entire process due to Plan Errors.", OutPutReptName, , True)
   DebugPrintOn ("Process Aborted by User")
   End
End If
ProcessPlanErrs = "" ' Append to the report....
          
Dim EtlImportResult As Boolean:  EtlImportResult = True
If Not CompileOnly Then EtlImportResult = EtlImport(ImpFile, WkShtName, WkShtNum, OutPutReptName, True)

If Not EtlImportResult Then MsgBox ("EtlImport=" & EtlImportResult & "  Process Paused.....")

End Function

Private Function GetUpdatePlan(ByVal KeyFieldsArePresent As Boolean) As String

' This routine will analyse and articulate the Update plan.
If Not KeyFieldsArePresent Then
  GetUpdatePlan = "     No Update Plan, since all rows will be APPENDed. (No KeyFields were set)" & vbCrLf
  Exit Function
End If

'RT_AcceptChanges
'RT_AllowChange2Blank
Dim AcceptChanges4All As Boolean:  AcceptChanges4All = True
Dim AllowChange2Blank4All As Boolean:  AllowChange2Blank4All = True

Dim II, JJ, KK, XX, ZZ
For II = LBound(RT_AcceptChanges) To UBound(RT_AcceptChanges)
  If Not RT_AcceptChanges(II) Then
    AcceptChanges4All = False
  End If
  If Not RT_AllowChange2Blank(II) Then
    'AllowChange2Blank4All
  End If
Next II

Dim UpdCntYY, UpdCntYN, UpdCntNN  ' Count of combinations of AcceptChanges and AllowChange2Blank
Dim UpdFldYY As String, UpdFldYN As String, UpdFldNN As String  ' Lists of fields.
UpdCntYY = 0: UpdCntYN = 0: UpdCntNN = 0 ' Initial Values
Dim UpdPlan() As String  ' Mark each item with update plan.


For II = LBound(RT_AcceptChanges) To UBound(RT_AcceptChanges)
  XX = "N~N"  ' Default field to N and N
  If RT_AcceptChanges(II) Then
    If RT_AllowChange2Blank(II) Then XX = "Y~Y" Else XX = "Y~N"
  End If
  Call InsertNewElementIntoArray(UpdPlan, XX) ' Populate Update Plans for each field.
  If XX = "N~N" Then UpdCntNN = UpdCntNN + 1
  If XX = "Y~Y" Then UpdCntYY = UpdCntYY + 1
  If XX = "Y~N" Then UpdCntYN = UpdCntYN + 1
  
  ZZ = RT_ExcelHeadingText(II)
  If RT_ExcelColumnLetterOrig(II) <> "" Then ZZ = RT_ExcelColumnLetterOrig(II) & ":" & ZZ
  If XX = "N~N" Then UpdFldNN = UpdFldNN & """" & ZZ & """, "
  If XX = "Y~Y" Then UpdFldYY = UpdFldYY & """" & ZZ & """, "
  If XX = "Y~N" Then UpdFldYN = UpdFldYN & """" & ZZ & """, "
Next II
If UpdFldNN <> "" Then
  ZZ = UpdFldNN
  ZZ = Left(ZZ, Len(ZZ) - 2)
  XX = InStrRev(ZZ, ", ")
  If XX > 0 Then ZZ = "(" & Left(ZZ, XX - 1) & " and " & Mid(ZZ, XX + 2) & ")"
  UpdFldNN = ZZ
End If

If UpdFldYY <> "" Then
  ZZ = UpdFldYY
  ZZ = Left(ZZ, Len(ZZ) - 2)
  XX = InStrRev(ZZ, ", ")
  If XX > 0 Then ZZ = "(" & Left(ZZ, XX - 1) & " and " & Mid(ZZ, XX + 2) & ")"
  UpdFldYY = ZZ
End If

If UpdFldYN <> "" Then
  ZZ = UpdFldYN
  ZZ = Left(ZZ, Len(ZZ) - 2)
  XX = InStrRev(ZZ, ", ")
  If XX > 0 Then ZZ = "(" & Left(ZZ, XX - 1) & " and " & Mid(ZZ, XX + 2) & ")"
  UpdFldYN = ZZ
End If

Dim Upd() As String
XX = Lpad(UpdCntNN, 9) & "~N~N~" & UpdFldNN
Call InsertNewElementIntoArray(Upd, XX, , True) ' Store and sort low to high..
XX = Lpad(UpdCntYY, 9) & "~Y~Y~" & UpdFldYY
Call InsertNewElementIntoArray(Upd, XX, , True) ' Store and sort low to high..
XX = Lpad(UpdCntYN, 9) & "~Y~N~" & UpdFldYN
Call InsertNewElementIntoArray(Upd, XX, , True) ' Store and sort low to high..

Dim UpdParms(1 To 3, 1 To 4) As Variant
Dim hParms() As String, L As Long
hParms = Split(Upd(2), "~")
L = hParms(0)
UpdParms(1, 1) = L: UpdParms(1, 2) = hParms(1): UpdParms(1, 3) = hParms(2): UpdParms(1, 4) = hParms(3)

hParms = Split(Upd(1), "~")
L = hParms(0)
UpdParms(2, 1) = L: UpdParms(2, 2) = hParms(1): UpdParms(2, 3) = hParms(2): UpdParms(2, 4) = hParms(3)

hParms = Split(Upd(0), "~")
L = hParms(0)
UpdParms(3, 1) = L: UpdParms(3, 2) = hParms(1): UpdParms(3, 3) = hParms(2): UpdParms(3, 4) = hParms(3)

GetUpdatePlan = ""
For II = LBound(UpdParms, 1) To UBound(UpdParms, 1)
  If UpdParms(II, 2) = "Y" And UpdParms(II, 1) > 0 Then ' Update Process will Accept Changes for 3 fields.
    GetUpdatePlan = GetUpdatePlan & "     Update Process will Accept Changes for " & UpdParms(II, 1) & " fields. "
    If UpdParms(II, 3) = "Y" Then _
      GetUpdatePlan = GetUpdatePlan & "(And Allow BLANK/Empty cells to clear a database field) " _
      Else _
      GetUpdatePlan = GetUpdatePlan & "(BLANK/Empty cells will not change a database field) "
    GetUpdatePlan = GetUpdatePlan & " Fields " & UpdParms(II, 4) & vbCrLf
  End If
   
  If UpdParms(II, 2) = "N" And UpdParms(II, 1) > 0 Then ' Update will not be done (No Changes Accepted) for 3 fields.
    GetUpdatePlan = GetUpdatePlan & "     Update will not be done (No Changes Accepted) for " & UpdParms(II, 1) & " field(s). "
    GetUpdatePlan = GetUpdatePlan & " Field(s) " & UpdParms(II, 4) & vbCrLf
  End If
   
Next II

End Function



Private Function GetMappingPlan(ByVal ImpFile As String, ByVal WkShtName As String) As String

Dim Headings, II, JJ, KK, ZZ, YY
Headings = GetExcelHeadingText(ImpFile, WkShtName, True)

Dim Mapped(), UnMapped(), Ignored()

'  Find all Un-Mapped and Mapped fields
For II = LBound(RT_ExcelHeadingText) To UBound(RT_ExcelHeadingText)
  ZZ = RT_ExcelHeadingText(II)
  If RT_ExcelColumnLetterOrig(II) <> "" Then ZZ = RT_ExcelColumnLetterOrig(II) & ":" & ZZ
  If RT_FieldNameOutput(II) = "" Then
    Call InsertNewElementIntoArray(UnMapped, ZZ)
  Else
    Call InsertNewElementIntoArray(Mapped, ZZ & "-->" & RT_FieldNameOutput(II))
  End If
Next II

'  Find all Ignored fields........
'Dim HoldRT_ExcelHeadingText() As String:  HoldRT_ExcelHeadingText = RT_ExcelHeadingText
For II = LBound(Headings) To UBound(Headings)
  ZZ = xlColAlfa(II + 1) & ":" & Headings(II)    '''  Add Excel Column to heading...
  KK = -1
  For JJ = LBound(RT_ExcelHeadingText) To UBound(RT_ExcelHeadingText)
    YY = RT_ExcelColumnLetter(JJ) & ":" & RT_ExcelHeadingText(JJ)    '''  Add Excel Column to heading...
    If ZZ = YY Then
      KK = II
      Exit For
    End If
  Next JJ
  
  If KK = -1 Then Call InsertNewElementIntoArray(Ignored, ZZ)
Next II

Dim MappedFields As String, AllMapped As String:  AllMapped = ""
If Not IsArrayAllocated(UnMapped) And Not IsArrayAllocated(Ignored) Then AllMapped = "ALL "

If Not IsArrayAllocated(Mapped) Then
  MappedFields = "No Fields were mapped to the DB Table." & vbCrLf
 Else
  If UBound(Mapped) = 0 Then
    MappedFields = "Mapping for " & AllMapped & UBound(Mapped) + 1 & " field.  (""" & Mapped(0) & """) "
   Else
    MappedFields = "Mapping for " & AllMapped & UBound(Mapped) + 1 & " fields. ("
    For II = LBound(Mapped) To UBound(Mapped)
      MappedFields = MappedFields & """" & Mapped(II) & """, "
      If II = UBound(Mapped) - 1 Then MappedFields = Left(MappedFields, Len(MappedFields) - 2) & " and "
    Next II
    MappedFields = Left(MappedFields, Len(MappedFields) - 2) & ")"
  End If
End If
MappedFields = "     " & MappedFields

'Dim Mapped(), UnMapped(), Ignored()
Dim UnMappedFields As String
If Not IsArrayAllocated(UnMapped) Then GoTo SkipUnMapped
If Not IsArrayAllocated(Mapped) Then
  UnMappedFields = "ALL Fields were UnMapped to the DB Table." & vbCrLf
 Else
  If UBound(UnMapped) = 0 Then
    UnMappedFields = "Un-Mapped for " & UBound(UnMapped) + 1 & " field.  (""" & UnMapped(0) & """) "
   Else
    UnMappedFields = "Un-Mapped for " & UBound(UnMapped) + 1 & " fields. ("
    For II = LBound(UnMapped) To UBound(UnMapped)
      UnMappedFields = UnMappedFields & """" & UnMapped(II) & """, "
      If II = UBound(UnMapped) - 1 Then UnMappedFields = Left(UnMappedFields, Len(UnMappedFields) - 2) & " and "
    Next II
    UnMappedFields = Left(UnMappedFields, Len(UnMappedFields) - 2) & ")"
  End If
End If
UnMappedFields = "     " & UnMappedFields
SkipUnMapped:

'Found 4 fields in that are Ignored in Config.  Fields ("Trans Date", "Amount" and "Balance")
Dim IgnoredFields As String
If IsArrayAllocated(Ignored) Then
  If UBound(Ignored) = 0 Then
    IgnoredFields = " Found 1 UN-Mapped field in """ & WkShtName & """ that is IGNORED in Config.  (""" & Ignored(0) & """) "
   Else
    IgnoredFields = " Found " & UBound(Ignored) + 1 & " UN-Mapped fields in """ & WkShtName & """ that are Ignored in Config.  Fields ("
    For II = LBound(Ignored) To UBound(Ignored)
      IgnoredFields = IgnoredFields & """" & Ignored(II) & """, "
      If II = UBound(Ignored) - 1 Then IgnoredFields = Left(IgnoredFields, Len(IgnoredFields) - 2) & " and "
    Next II
    IgnoredFields = Left(IgnoredFields, Len(IgnoredFields) - 2) & ")"
  End If
  IgnoredFields = IgnoredFields & vbCrLf & _
                  "           (Data contained in Un-Mapped columns will not be captured in the """ & WkShtName & """ table.)"
End If
IgnoredFields = "    " & IgnoredFields

If MappedFields <> "" Then GetMappingPlan = GetMappingPlan & MappedFields & vbCrLf
If UnMappedFields <> "" Then GetMappingPlan = GetMappingPlan & UnMappedFields & vbCrLf
If IgnoredFields <> "" Then GetMappingPlan = GetMappingPlan & IgnoredFields & vbCrLf

End Function



Private Function GetErrorHandlePlan() As String

'  Step 3 ----  (Error Handling Plan)
'     Error Handling will reject field level data for 3 fields. Fields ("Trans Date", "Amount" and "Balance")
'     Error Handling will reject ENTIRE Row if these 2 fields have an error.  Fields ("Amount" and "Balance")
'     Error Handling will reject entire FILE if these 2 fields have an error.  Fields ("Amount" and "Balance")

'  Need to Add warning if rowlevel and filelevel are both specified.  (Done)

Dim FldLevel(), RowLevel(), FileLevel()
Dim II, JJ, ZZ

'Global RT_RejectErrFile()
'Global RT_RejectErrRows()

For II = LBound(RT_ExcelHeadingText) To UBound(RT_ExcelHeadingText)
  
  ZZ = RT_ExcelHeadingText(II)
  If RT_ExcelColumnLetterOrig(II) <> "" Then ZZ = RT_ExcelColumnLetterOrig(II) & ":" & ZZ
  
  If Not RT_RejectErrFile(II) And Not RT_RejectErrRows(II) Then _
    Call InsertNewElementIntoArray(FldLevel, ZZ)

  If RT_RejectErrFile(II) And RT_RejectErrRows(II) Then _
    Call InsertNewElementIntoArray(FileLevel, ZZ)

  If Not RT_RejectErrFile(II) And RT_RejectErrRows(II) Then _
    Call InsertNewElementIntoArray(RowLevel, ZZ)

  If RT_RejectErrFile(II) And Not RT_RejectErrRows(II) Then _
    Call InsertNewElementIntoArray(FileLevel, ZZ)

Next II

Dim FldLevelFields As String
If Not IsArrayAllocated(FldLevel) Then GoTo SkipFieldLevel
If Not IsArrayAllocated(RowLevel) And _
   Not IsArrayAllocated(FileLevel) Then
  FldLevelFields = "All Error Handling will be FIELD level.  Any error will result in ONLY the BAD field data being rejected." & vbCrLf
 Else
  If UBound(FldLevel) = 0 Then
    FldLevelFields = "Error Handling will reject FIELD level data for 1 field.  """ & FldLevel(0) & """) "
   Else
    FldLevelFields = "Error Handling will reject FIELD level data for " & UBound(FldLevel) + 1 & " fields.  A field error will result in BAD field data being rejected. ("
    For II = LBound(FldLevel) To UBound(FldLevel)
      FldLevelFields = FldLevelFields & """" & FldLevel(II) & """, "
      If II = UBound(FldLevel) - 1 Then FldLevelFields = Left(FldLevelFields, Len(FldLevelFields) - 2) & " and "
    Next II
    FldLevelFields = Left(FldLevelFields, Len(FldLevelFields) - 2) & ")"
  End If
End If
FldLevelFields = "     " & FldLevelFields
SkipFieldLevel:

Dim RowLevelFields As String
If Not IsArrayAllocated(RowLevel) Then GoTo SkipRowLevel
If Not IsArrayAllocated(FldLevel) And _
   Not IsArrayAllocated(FileLevel) Then
  RowLevelFields = "All Error Handling will be ROW level.  Any DATA errors will result in entire row being rejected." & vbCrLf
 Else
  If UBound(RowLevel) = 0 Then
    RowLevelFields = "Data Handling errors will reject ROW level data for 1 field.  """ & RowLevel(0) & """) "
   Else
    RowLevelFields = "Data Handling errors will reject ROW level data for " & UBound(RowLevel) + 1 & " fields. ("
    For II = LBound(RowLevel) To UBound(RowLevel)
      RowLevelFields = RowLevelFields & """" & RowLevel(II) & """, "
      If II = UBound(RowLevel) - 1 Then RowLevelFields = Left(RowLevelFields, Len(RowLevelFields) - 2) & " and "
    Next II
    RowLevelFields = Left(RowLevelFields, Len(RowLevelFields) - 2) & ")"
  End If
End If
RowLevelFields = "     " & RowLevelFields
SkipRowLevel:


Dim FileLevelFields As String
If Not IsArrayAllocated(FileLevel) Then GoTo SkipFileLevel
If Not IsArrayAllocated(FldLevel) And _
   Not IsArrayAllocated(RowLevel) Then
  FileLevelFields = "All Error Handling will be FILE level.  Any error in any field will result in the entire file being rejected." & vbCrLf
 Else
  If UBound(FileLevel) = 0 Then
    FileLevelFields = "Error Handling will reject FILE level data for 1 field.  (""" & FileLevel(0) & """) "
   Else
    FileLevelFields = "Error Handling will reject FILE level data for " & UBound(FileLevel) + 1 & " fields. ("
    For II = LBound(FileLevel) To UBound(FileLevel)
      FileLevelFields = FileLevelFields & """" & FileLevel(II) & """, "
      If II = UBound(FileLevel) - 1 Then FileLevelFields = Left(FileLevelFields, Len(FileLevelFields) - 2) & " and "
    Next II
    FileLevelFields = Left(FileLevelFields, Len(FileLevelFields) - 2) & ")"
  End If
End If
FileLevelFields = "     " & FileLevelFields
SkipFileLevel:

If FldLevelFields <> "" Then GetErrorHandlePlan = GetErrorHandlePlan & FldLevelFields & vbCrLf
If RowLevelFields <> "" Then GetErrorHandlePlan = GetErrorHandlePlan & RowLevelFields & vbCrLf
If FileLevelFields <> "" Then GetErrorHandlePlan = GetErrorHandlePlan & FileLevelFields & vbCrLf

End Function



Private Function GetSaveDatePlan(ByVal ImpFile As String, ByVal WkShtName As String) As String

Dim Headings, II, JJ, KK, ZZ
Headings = GetExcelHeadingText(ImpFile, WkShtName, True)

Dim SaveDate(), NoSaveDate()

'  Find all Un-SaveDate and SaveDate fields
For II = LBound(RT_ExcelHeadingText) To UBound(RT_ExcelHeadingText)
  If RT_FieldNameOutput(II) <> "" Then ' Only concerned with Mapped Fields....
    ZZ = RT_ExcelHeadingText(II)
    If RT_ExcelColumnLetterOrig(II) <> "" Then ZZ = RT_ExcelColumnLetterOrig(II) & ":" & ZZ
    If RT_SaveDateChanged(II) Then
      Call InsertNewElementIntoArray(SaveDate, ZZ & "-->" & RT_FieldNameOutput(II))
     Else
      Call InsertNewElementIntoArray(NoSaveDate, ZZ & "-->" & RT_FieldNameOutput(II))
    End If
  End If
Next II

Dim SaveDateFields As String, AllSaveDate As String:  AllSaveDate = ""
If Not IsArrayAllocated(NoSaveDate) Then AllSaveDate = "ALL "

If Not IsArrayAllocated(SaveDate) Then
  SaveDateFields = "No Fields were selected to Save Date/Time Stamps to the DB Table." & vbCrLf
 Else
  If UBound(SaveDate) = 0 Then
    SaveDateFields = "Save Date/Time Changed for " & AllSaveDate & UBound(SaveDate) + 1 & " field.  (""" & SaveDate(0) & """) "
   Else
    SaveDateFields = "Save Date/Time Changed for " & AllSaveDate & UBound(SaveDate) + 1 & " fields. ("
    For II = LBound(SaveDate) To UBound(SaveDate)
      SaveDateFields = SaveDateFields & """" & SaveDate(II) & """, "
      If II = UBound(SaveDate) - 1 Then SaveDateFields = Left(SaveDateFields, Len(SaveDateFields) - 2) & " and "
    Next II
    SaveDateFields = Left(SaveDateFields, Len(SaveDateFields) - 2) & ")"
  End If
End If
SaveDateFields = "     " & SaveDateFields

'Dim SaveDate(), NoSaveDate()
Dim NoSaveDateFields As String
If Not IsArrayAllocated(NoSaveDate) Then GoTo SkipNoSaveDate
If IsArrayAllocated(SaveDate) Then
  If UBound(NoSaveDate) = 0 Then
    NoSaveDateFields = "Save Date/Time Changed NOT selected for " & UBound(NoSaveDate) + 1 & " field.  (""" & NoSaveDate(0) & """) "
   Else
    NoSaveDateFields = "Save Date/Time Changed NOT selected for " & UBound(NoSaveDate) + 1 & " fields. ("
    For II = LBound(NoSaveDate) To UBound(NoSaveDate)
      NoSaveDateFields = NoSaveDateFields & """" & NoSaveDate(II) & """, "
      If II = UBound(NoSaveDate) - 1 Then NoSaveDateFields = Left(NoSaveDateFields, Len(NoSaveDateFields) - 2) & " and "
    Next II
    NoSaveDateFields = Left(NoSaveDateFields, Len(NoSaveDateFields) - 2) & ")"
  End If
End If
NoSaveDateFields = "     " & NoSaveDateFields
SkipNoSaveDate:

If SaveDateFields <> "" Then GetSaveDatePlan = GetSaveDatePlan & SaveDateFields & vbCrLf
If NoSaveDateFields <> "" Then GetSaveDatePlan = GetSaveDatePlan & NoSaveDateFields & vbCrLf

End Function



Private Function VerifyImportConfig() As String

Dim Errmsg:  Errmsg = ""
Dim ColErrMsg:  ColErrMsg = ""
Dim II, JJ, KK, ZZ

'DBTable - Required, Must be a valid table.
'
' Column Validation....
' At least one column is required.
' RT_ID ID=2 is required and must be unique
' RT_ExcelHeadingText Heading='Rept ID' and RT_ExcelColumnLetter Col=B are not required individually, but one of them must be present.
' RT_ExcelColumnLetter Col=B if not empty, must be unique.
' RT_ExcelColumnLetter Col=B must be a valid Excel Column ID
' RT_ExcelHeadingText Heading='Rept ID' if not unique, must have a RT_ExcelColumnLetter
' RT_FieldNameOutput DB=field is provided, then either RT_ExcelColumnLetter or RT_ExcelHeadingText is required.
' RT_FieldNameOutput DB=field not required individually, but there must be at least one found present.
' RT_FieldNameOutput DB=field if found, must be unique.
' RT_FieldNameOutput DB=field if not missing, must be a valid field name in RT_OutputTableName DBTable='x_table'
' RT_KeyField() As Boolean  at least one key field must be present.  ***Done
' RT_KeyField() As Boolean  must have a DB=field filled in to be a valid Key field.  Done
' RT_KeyField() There cannot be more than 12 key fields specified.  ***Done
' check defaults against old Import routine.
' RT_ExcelHeadingText Warning if heading is present but no DB mapping provided. Not Done

' RT_ActiveFlag Field Require Boolean for field??? Needs to be checked.
'aMark_active_flag <> "Changed") And (aMark_active_flag <> "New") And (aMark_active_flag <> "Imported")

' Active Flag Field does not exist yet.
'Global RT_MarkActiveFlag  ' Valid values: "Changed", "New", "Imported"
'Global RT_ActiveFlagField

' Warnings.....
' RT_KeyField Key=Y - Non specified.  Complete file will be appended.  done

' Check for these severe errors before proceeding....
If Not TableExists(RT_OutputTableName) Then _
  Errmsg = Errmsg & "*** DBTable=""" & RT_OutputTableName & """ table name is NOT found in database." & vbCrLf
' At least one column is required.
If Not IsArrayAllocated(RT_ID) Then _
  Errmsg = Errmsg & "*** DBTable=""" & RT_OutputTableName & """ no columns defined for this table." & vbCrLf
VerifyImportConfig = Errmsg
If Errmsg <> "" Then Exit Function  '  If the output table does not exist or no columns are mapped, then Exit.


' Assign proper Col numbers to all RT_ExcelColumnLetter that do not already have them.
Dim Col
Dim AllColumnLetters: AllColumnLetters = RT_ExcelColumnLetter

Dim WkShtHeadings
Dim rtWorkSheetNameExists As Boolean

Dim WkShtNames As Variant, WkShtFound As Boolean, WkShtCount As Long, HoldWkShtName As String
WkShtNames = GetExcelWkShtNames(XslFileAndPath, True)
If IsArrayAllocated(WkShtNames) Then
  WkShtCount = 0
  For II = LBound(WkShtNames) To UBound(WkShtNames)
    If WkShtNames(II) Like RT_WorkSheetName Then
      HoldWkShtName = WkShtNames(II)
      WkShtCount = WkShtCount + 1
    End If
  Next II
End If

If HoldWkShtName <> "" Then _
  WkShtHeadings = GetExcelHeadingText(XslFileAndPath, HoldWkShtName, True)
If IsArrayAllocated(WkShtHeadings) Then WkShtFound = True Else WkShtFound = False

If WkShtFound Then
  For II = LBound(RT_ExcelHeadingText) To UBound(RT_ExcelHeadingText)
    JJ = 0
    If RT_ExcelColumnLetter(II) = "" Then
      For KK = LBound(WkShtHeadings) To UBound(WkShtHeadings)
        Col = xlColAlfa(KK + 1)
        If WkShtHeadings(KK) Like RT_ExcelHeadingText(II) And FindPosition(RT_ExcelColumnLetter, Col) < 0 Then
          AllColumnLetters(II) = Col
          Exit For
        End If
      Next KK
    End If
  Next II
End If

If (RT_ActiveFlagField <> "" Or RT_MarkActiveFlag <> "") And _
   Not fieldexists(RT_ActiveFlagField, RT_OutputTableName) Then _
  Errmsg = Errmsg & "*** ActiveFlagField=""" & RT_ActiveFlagField & """ field name is NOT found in database." & vbCrLf
  
If (RT_ActiveFlagField <> "" Or RT_MarkActiveFlag <> "") And _
   RT_MarkActiveFlag <> "Imported" And RT_MarkActiveFlag <> "Changed" And RT_MarkActiveFlag <> "New" Then _
  Errmsg = Errmsg & "*** MarkActiveFlag=""" & RT_MarkActiveFlag & """ value is NOT valid.  Must be ""Imported"" or ""Changed"" or ""New""" & vbCrLf
  
Dim NumOfDBFields:  NumOfDBFields = 0
Dim NumOfKeyFields:   NumOfKeyFields = 0
Dim NumOfColLettersUsed:  NumOfColLettersUsed = 0
Dim ColLetterList:  ColLetterList = ""
Dim NumOfUnmappedHeadings: NumOfUnmappedHeadings = 0

' Column Validation....
For II = LBound(RT_ID) To UBound(RT_ID)
  ' RT_ID ID=2 is required and must be unique
  If IsEmpty(RT_ID(II)) Or RT_ID(II) = "" Then _
    ColErrMsg = ColErrMsg & "*** Column Col=" & RT_ExcelColumnLetter(II) & " Heading=""" & RT_ExcelHeadingText(II) & """," & _
      "Does not contain a required ID= keyword." & vbCrLf
      
  JJ = FindNonUniquePos(RT_ID, RT_ID(II))
  If Not IsEmpty(RT_ID(II)) And RT_ID(II) <> "" And JJ <> -1 Then _
    ColErrMsg = ColErrMsg & "*** Column ID=""" & RT_ID(II) & """ " & _
      "Must be unique. Duplicate was found- Heading=""" & RT_ExcelHeadingText(II) & """," & _
      "Col=" & RT_ExcelColumnLetter(II) & vbCrLf
  
  ' RT_ExcelHeadingText Heading='Rept ID' and RT_ExcelColumnLetter Col=B are not required individually, but one of them must be present.
  If (IsEmpty(RT_ExcelHeadingText(II)) Or RT_ExcelHeadingText(II) = "") And _
     (IsEmpty(RT_ExcelColumnLetter(II)) Or RT_ExcelColumnLetter(II) = "") Then _
    ColErrMsg = ColErrMsg & "*** Column ID=""" & RT_ID(II) & """ " & _
      "Must contain either Col= or Heading=, Both ARE missing." & vbCrLf
  
  ' RT_ExcelColumnLetter Col=B if not empty, must be unique.
  JJ = FindNonUniquePos(RT_ExcelColumnLetter, RT_ExcelColumnLetter(II))
  If Not IsEmpty(RT_ExcelColumnLetter(II)) And RT_ExcelColumnLetter(II) <> "" And JJ <> -1 Then _
    ColErrMsg = ColErrMsg & "*** Column ID=""" & RT_ID(II) & """   " & _
        "Col=" & RT_ExcelColumnLetter(II) & " must be unique. " & _
        "Duplicate found. " & vbCrLf
        
  ' RT_ExcelColumnLetter Col=B must be a valid Excel Column ID
  If Not IsEmpty(RT_ExcelColumnLetter(II)) And RT_ExcelColumnLetter(II) <> "" And _
     Not ValidExcelColID(RT_ExcelColumnLetter(II)) Then _
    ColErrMsg = ColErrMsg & "*** Column ID=""" & RT_ID(II) & """   " & _
        "Col=" & RT_ExcelColumnLetter(II) & " must be unique. " & _
        "Duplicate found. " & vbCrLf
    
  ' RT_ExcelHeadingText Heading='Rept ID' if not unique, must have a RT_ExcelColumnLetter
  JJ = FindNonUniquePos(RT_ExcelHeadingText, RT_ExcelHeadingText(II))
  If RT_FieldNameOutput(II) <> "" And _
     Not IsEmpty(RT_ExcelHeadingText(II)) And RT_ExcelHeadingText(II) <> "" And JJ <> -1 And _
    (IsEmpty(RT_ExcelColumnLetter(II)) Or RT_ExcelColumnLetter(II) = "") Then _
    ColErrMsg = ColErrMsg & "*** Column ID=""" & RT_ID(II) & """   " & _
        "Heading=""" & RT_ExcelHeadingText(II) & _
        """ is NOT Unique.  Col=?? must be explicitly defined for Duplicate Headings." & vbCrLf
  
  If Not IsEmpty(RT_FieldNameOutput(II)) And RT_FieldNameOutput(II) <> "" Then _
    NumOfDBFields = NumOfDBFields + 1
    
  If RT_KeyField(II) Then NumOfKeyFields = NumOfKeyFields + 1

  ' RT_KeyField() As Boolean  must have a DB=field filled in to be a valid Key field.
  If RT_KeyField(II) And Not (Not IsEmpty(RT_FieldNameOutput(II)) And RT_FieldNameOutput(II) <> "") Then _
    ColErrMsg = ColErrMsg & "*** Column ID=""" & RT_ID(II) & """ " & _
        "is marked as a KeyField with KF=Y.  The DB="""" field name is blank. Field Name is required for a KeyField.  DB Table-""" & _
        RT_OutputTableName & """. " & vbCrLf

  ' RT_SaveDateChangedLineLvl() As Boolean  must have a DB=field filled in to be a valid Key field.
  If RT_SaveDateChangedLineLvl(II) And Not (Not IsEmpty(RT_FieldNameOutput(II)) And RT_FieldNameOutput(II) <> "") Then _
    ColErrMsg = ColErrMsg & "Warning - Column ID=""" & RT_ID(II) & """ " & _
        "is marked as SavDtChange=Yes.  The DB="""" field name is blank. Field Name is needed for SDC=Y to process.  DB Table-""" & _
        RT_OutputTableName & """. " & vbCrLf

  ' Attempt to map "C:Seq" to AutoNumber field "ID" is not allowed."
  If (Not IsEmpty(RT_FieldNameOutput(II)) And Field_Type(RT_OutputTableName, RT_FieldNameOutput(II)) = "AutoNumber") Then
    JJ = AllColumnLetters(II) & ":" & RT_ExcelHeadingText(II)
    If Left(JJ, 1) = ":" Then JJ = Mid(JJ, 2)
    If Right(JJ, 1) = ":" Then JJ = Left(JJ, Len(JJ) - 1)
    ColErrMsg = ColErrMsg & "*** Column ID=""" & RT_ID(II) & """ " & _
        " is attempting to Illegally map Excel column-""" & JJ & """ to an AutoNumber field """ & RT_FieldNameOutput(II) & """ found in Table-" & _
        RT_OutputTableName & """. " & vbCrLf
  End If

  ' RT_ExcelHeadingText Warning if heading is present but no DB mapping provided.
  If Not IsEmpty(RT_ExcelHeadingText(II)) And RT_ExcelHeadingText(II) <> "" And _
        Not (Not IsEmpty(RT_FieldNameOutput(II)) And RT_FieldNameOutput(II) <> "") Then
    NumOfUnmappedHeadings = NumOfUnmappedHeadings + 1
    ColErrMsg = ColErrMsg & "Warning - Column ID=""" & RT_ID(II) & """   " & _
        "Heading=""" & RT_ExcelHeadingText(II) & """ is not mapped to any database field with DB= keyword. " & vbCrLf
  End If
  
  ' RT_FieldNameOutput DB=field is provided, then either RT_ExcelColumnLetter or RT_ExcelHeadingText is required.
  If Not IsEmpty(RT_FieldNameOutput(II)) And RT_FieldNameOutput(II) <> "" Then
    If (IsEmpty(RT_ExcelColumnLetter(II)) Or RT_ExcelColumnLetter(II) = "") And _
       (IsEmpty(RT_ExcelHeadingText(II)) Or RT_ExcelHeadingText(II) = "") Then _
    ColErrMsg = ColErrMsg & "*** Column ID=""" & RT_ID(II) & """ " & _
        "DB=""" & RT_FieldNameOutput(II) & """ is not mapped to an eXcel column.  Either Col= or Heading= is required. " & _
        "Duplicate found. " & vbCrLf
  End If
  
  If Not IsEmpty(RT_ExcelColumnLetter(II)) And RT_ExcelColumnLetter(II) <> "" Then
    NumOfColLettersUsed = NumOfColLettersUsed + 1
    KK = xlColNum(RT_ExcelColumnLetter(II)) - 1
    If RT_ExcelHeadingText(II) = "" Then KK = WkShtHeadings(KK) Else KK = RT_ExcelHeadingText(II)
    ColLetterList = ColLetterList & """" & RT_ExcelColumnLetter(II) & ":" & KK & """, "
  End If
  
  ' RT_FieldNameOutput DB=field if found, must be unique.
  JJ = FindNonUniquePos(RT_FieldNameOutput, RT_FieldNameOutput(II))
  If Not IsEmpty(RT_FieldNameOutput(II)) And RT_FieldNameOutput(II) <> "" And JJ <> -1 Then _
    ColErrMsg = ColErrMsg & "*** Column ID=""" & RT_ID(II) & """   " & _
        "DB=""" & RT_FieldNameOutput(II) & """ must be unique. " & _
        "Duplicate found. " & vbCrLf
  
  ' RT_FieldNameOutput DB=field if not missing, must be a valid field name in RT_OutputTableName DBTable='x_table'
  If (Not IsEmpty(RT_FieldNameOutput(II)) And RT_FieldNameOutput(II) <> "") And _
      Not fieldexists(RT_FieldNameOutput(II), RT_OutputTableName) Then _
    ColErrMsg = ColErrMsg & "*** Column ID=""" & RT_ID(II) & """   " & _
        "DB=""" & RT_FieldNameOutput(II) & """ is not a valid field in table-""" & _
        RT_OutputTableName & """. " & vbCrLf
        
  ' Warning: AcceptChanges=Y not allowed on a KeyField AcceptChanges will be ignored.
  If RT_KeyField(II) And RT_AcceptChanges(II) And RT_FieldNameOutput(II) <> "" Then
    ColErrMsg = ColErrMsg & "Warning - Column ID=""" & RT_ID(II) & """   " & _
        "DB=""" & RT_FieldNameOutput(II) & """ is KeyField=Y has AcceptChanges=Y. Changes to this field will be ignored for table-""" & _
        RT_OutputTableName & """. " & vbCrLf
    RT_AcceptChanges(II) = False ' Force AcceptChanges=N
  End If
  
  ' Warning: AllowChange2Blank not allowed unless AcceptChanges=Y. AllowChange2Blank will be ignored.
  If Not RT_AcceptChanges(II) And RT_AllowChange2Blank(II) And _
    RT_FieldNameOutput(II) <> "" And Not RT_KeyField(II) Then
    ColErrMsg = ColErrMsg & "Warning - Column ID=""" & RT_ID(II) & """   " & _
        "DB=""" & RT_FieldNameOutput(II) & """ cannot Accept Changes. AllowChange2Blank has no meaning for table-""" & _
        RT_OutputTableName & """. " & vbCrLf
    RT_AllowChange2Blank(II) = False
  End If
  
  ' Warning: RejErrRow=Y and RejErrFile=Y are ambiguous.  ONLY RejErrFile=Y will be honored.
  If RT_RejectErrFile(II) And RT_RejectErrRows(II) Then
    ColErrMsg = ColErrMsg & "Warning - Column ID=""" & RT_ID(II) & """   " & _
        "DB=""" & RT_FieldNameOutput(II) & """ have RejErrRow=Y and RejErrFile=Y.  Both cannot be Yes.  Only RejErrFile=Y will be honored." & vbCrLf
    RT_RejectErrRows(II) = False
  End If
      
ProcessNext:
Next II
  
' Warning - If Col="" is specified, then Excel columns cannot be shifted without causing issues.
If ColLetterList <> "" Then
  If Right(ColLetterList, 2) = ", " Then ColLetterList = Left(ColLetterList, Len(ColLetterList) - 2)
  II = InStrRev(ColLetterList, """, ")
  If II <> 0 Then ColLetterList = Mid(ColLetterList, 1, II) & " and " & Mid(ColLetterList, II + 3)
  ColLetterList = " (" & ColLetterList & ")"
End If

If NumOfColLettersUsed = 1 Then _
  Errmsg = Errmsg & "Warning - Found " & NumOfColLettersUsed & " column" & ColLetterList & " mapped with eXcel column letters. " & _
  "NOT recommended (unless there are duplicate heading titles or other reasons)." & vbCrLf & "     Using specific Col= letters will limit capability to shift eXcel Columns." & _
  " Config-""" & RT_FileID & """." & vbCrLf
If NumOfColLettersUsed > 1 Then _
  Errmsg = Errmsg & "Warning - Found " & NumOfColLettersUsed & " columns" & ColLetterList & " mapped with eXcel column letters. " & _
  "NOT recommended (unless there are duplicate heading titles or other reasons)." & vbCrLf & "     Using specific Col= letters will limit capability to shift eXcel Columns." & _
  " Config-""" & RT_FileID & """." & vbCrLf
  
' RT_FieldNameOutput DB=field not required individually, but there must be at least one found present.
If NumOfDBFields = 0 Then _
  Errmsg = Errmsg & "*** DBTable=""" & RT_OutputTableName & _
  """   ALL DB= were blank.  " & vbCrLf & "  At least 1 Column must define a DBField for the Configuration to work. " & vbCrLf
  
If NumOfUnmappedHeadings = 1 Then _
  Errmsg = Errmsg & "Warning - There is " & NumOfUnmappedHeadings & " un-mapped Heading found in Config-""" & _
      RT_FileID & """." & vbCrLf
If NumOfUnmappedHeadings > 1 Then _
  Errmsg = Errmsg & "Warning - There are " & NumOfUnmappedHeadings & " un-mapped Headings found in Config-""" & _
      RT_FileID & """." & vbCrLf

If NumOfKeyFields = 0 Then _
  Errmsg = Errmsg & "Warning - DBTable=""" & RT_OutputTableName _
  & """ had NO key fields marked with KF=Y.  This is NOT Normal for majority of ETL steps.  " & vbCrLf _
    & "Warning - If no key fields are used, then ALL records in the spreadsheet will be APPENDed to the table. " & vbCrLf
  
If NumOfKeyFields > 12 Then _
  Errmsg = Errmsg & "*** DBTable=""" & RT_OutputTableName _
  & """ had too many key fields marked with KF=Y.  Max number of Key Fields is 12.  " & vbCrLf
  
VerifyImportConfig = Errmsg & ColErrMsg

End Function

Private Function ValidExcelColID(ByVal XX As String) As Boolean

If Len(XX) < 1 Then Exit Function
If Verify(XX, "abcdefghijklmnopqrstuvwxyz") = 0 Then ValidExcelColID = True
End Function



Private Function ParseImportConfig(ByVal Config) As String

Dim HoldFileID
Dim HoldFileName
Dim HoldWorkSheetName As String
Dim HoldOutputTableName As String
Dim HoldActiveFlagField As String
Dim HoldMarkActiveFlag As String
Dim HoldUniqueKey
Dim FN_HoldAcceptChanges
Dim FN_HoldAllowChange2Blank
Dim FN_HoldRejectErrFile
Dim FN_HoldRejectErrRows
Dim FN_HoldSaveDateChanged

Dim HoldKeyField
Dim HoldID As String
Dim HoldExcelColumnLetter As String
Dim HoldExcelHeadingText As String
Dim HoldFieldNameOutput As String
Dim HoldAcceptChanges
Dim HoldAllowChange2Blank
Dim HoldRejectErrFile
Dim HoldRejectErrRows
Dim HoldSaveDateChanged
Dim HoldSaveDateChangedLineLvl
Dim Errmsg:  Errmsg = ""

' Initialize the Global File Configuration....
RT_FileID = ""
RT_FileName = ""
RT_WorkSheetName = ""
RT_OutputTableName = ""
RT_UniqueKey = True
RT_MarkActiveFlag = ""
RT_ActiveFlagField = ""
Erase RT_KeyField
Erase RT_ID
Erase RT_ExcelColumnLetter
Erase RT_ExcelHeadingText
Erase RT_FieldNameOutput
Erase RT_AcceptChanges
Erase RT_AllowChange2Blank
Erase RT_RejectErrFile
Erase RT_RejectErrRows
Erase RT_SaveDateChanged
Erase RT_SaveDateChangedLineLvl

Dim HoldFileNamex, aFileName, HoldColumn(), aColumn
If IsArrayAllocated(HoldColumn) Then Erase HoldColumn

' Parse out the 1 FileName= keywords, and then all of the Column= keywords...
HoldFileNamex = ExtractKeyWord(Config, "FileName", "FN")
Do
  aColumn = ExtractKeyWord(Config, "Column", "Col")
  If aColumn = "" Then Exit Do
  Call InsertNewElementIntoArray(HoldColumn, aColumn)
Loop

' Verify that high Keyword= level parameters are present.
If HoldFileNamex = "" Then _
  Errmsg = Errmsg & "*** FileName=  Keyword definition is missing from the Configuration...." & vbCrLf
If Not IsArrayAllocated(HoldColumn) Then _
  Errmsg = Errmsg & "*** Column=  Keyword definition is missing from the Configuration...." & vbCrLf
If Verify(Config, ", ") <> 0 Then _
  Errmsg = Errmsg & "*** Invalid/Unknown syntax:" & vbCrLf & """" & Config & """" & vbCrLf
  
If Errmsg <> "" Then
  ParseImportConfig = "ParseImportConfig detected invalid syntax:" & vbCrLf & vbCrLf & Errmsg
  DebugPrintOn (ParseImportConfig)
  Exit Function
End If

' First process the HoldFileNamex parms....
aFileName = HoldFileNamex
HoldFileNamex = "FileName=(" & HoldFileNamex & ")"
HoldFileID = ExtractKeyWord(aFileName, "ID=")
HoldFileName = ExtractKeyWord(aFileName, "Excel=")

HoldWorkSheetName = ExtractKeyWord(aFileName, "WkSht=", "WS")
HoldOutputTableName = ExtractKeyWord(aFileName, "DBTable=", "DB=")
HoldActiveFlagField = ExtractKeyWord(aFileName, "ActiveFlagField", "AFF")
HoldMarkActiveFlag = ExtractKeyWord(aFileName, "MarkActiveFlag", "MAF")

Call ExtractYnKeyword(aFileName, Errmsg, "", "UniqueKey=", HoldUniqueKey, "UK")
Call ExtractYnKeyword(aFileName, Errmsg, "", "AcceptChanges=", FN_HoldAcceptChanges, "AC")
Call ExtractYnKeyword(aFileName, Errmsg, "", "Change2Blank=", FN_HoldAllowChange2Blank, "C2B")
Call ExtractYnKeyword(aFileName, Errmsg, "", "RejErrFile=", FN_HoldRejectErrFile, "REF")
Call ExtractYnKeyword(aFileName, Errmsg, "", "RejErrRow=", FN_HoldRejectErrRows, "RER")
Call ExtractYnKeyword(aFileName, Errmsg, "", "SavDtChange=", FN_HoldSaveDateChanged, "SDC")

aFileName = TrimChars(aFileName, ",")  ' Trim commas and blanks....
If Verify(aFileName, ", ") <> 0 Then
  Errmsg = Errmsg & "*** """ & aFileName & """ is Invalid or Duplicated Keyword syntax." & vbCrLf
End If

If Errmsg <> "" Then
  ParseImportConfig = "ParseImportConfig detected invalid/unknown syntax in:" & vbCrLf & _
          HoldFileNamex & vbCrLf & Errmsg
  DebugPrintOn (ParseImportConfig)
  Exit Function
End If

'  Save parsed File Level values in the Global Variables
RT_FileID = HoldFileID
RT_FileName = HoldFileName
RT_WorkSheetName = HoldWorkSheetName
RT_OutputTableName = HoldOutputTableName
RT_ActiveFlagField = HoldActiveFlagField
RT_MarkActiveFlag = HoldMarkActiveFlag
RT_UniqueKey = HoldUniqueKey

Dim II: II = 0
Dim HoldLetterAndHeading As String
For II = LBound(HoldColumn) To UBound(HoldColumn)
  aColumn = HoldColumn(II)

  HoldID = ExtractKeyWord(aColumn, "ID=")
  HoldExcelColumnLetter = ExtractKeyWord(aColumn, "Col=", "C=")
  HoldExcelHeadingText = ExtractKeyWord(aColumn, "Heading=", "H=")
  HoldFieldNameOutput = ExtractKeyWord(aColumn, "DB=")
  
  HoldLetterAndHeading = HoldExcelHeadingText
  If HoldExcelHeadingText <> "" And HoldExcelColumnLetter <> "" Then _
    HoldLetterAndHeading = ":" & HoldLetterAndHeading
  HoldLetterAndHeading = HoldExcelColumnLetter & HoldLetterAndHeading
  
  HoldKeyField = Empty
  Call ExtractYnKeyword(aColumn, Errmsg, HoldID, "KeyField=", HoldKeyField, "KF")
  If IsEmpty(HoldKeyField) Then HoldKeyField = False ' Default Value.
  
  HoldAcceptChanges = Empty
  Call ExtractYnKeyword(aColumn, Errmsg, HoldID, "AcceptChanges=", HoldAcceptChanges, "AC")
  If IsEmpty(HoldAcceptChanges) Then HoldAcceptChanges = FN_HoldAcceptChanges
  If IsEmpty(HoldAcceptChanges) Then
    If HoldKeyField Then HoldAcceptChanges = False Else HoldAcceptChanges = True  ' Initial Value
    HoldAcceptChanges = True ' Default Value.
    Errmsg = Errmsg & "Default Value - AcceptChanges=" & HoldAcceptChanges & _
      "  for ID=" & HoldID & "  Heading=""" & HoldLetterAndHeading & """" & vbCrLf
  End If
  
  HoldAllowChange2Blank = Empty
  Call ExtractYnKeyword(aColumn, Errmsg, HoldID, "Change2Blank=", HoldAllowChange2Blank, "C2B")
  If IsEmpty(HoldAllowChange2Blank) Then HoldAllowChange2Blank = FN_HoldAllowChange2Blank
  If IsEmpty(HoldAllowChange2Blank) Then
    HoldAllowChange2Blank = False ' Default Value is always False.
    Errmsg = Errmsg & "Default Value - AllowChange2Blank=" & HoldAllowChange2Blank & _
      "  for ID=" & HoldID & "  Heading=""" & HoldLetterAndHeading & """" & vbCrLf
  End If
  
  HoldRejectErrFile = Empty
  Call ExtractYnKeyword(aColumn, Errmsg, HoldID, "RejErrFile=", HoldRejectErrFile, "REF")
  If IsEmpty(HoldRejectErrFile) Then HoldRejectErrFile = FN_HoldRejectErrFile
  If IsEmpty(HoldRejectErrFile) Then
    HoldRejectErrFile = False ' Default Value.
    Errmsg = Errmsg & "Default Value - RejectErrFile=" & HoldRejectErrFile & _
      "  for ID=" & HoldID & "  Heading=""" & HoldLetterAndHeading & """" & vbCrLf
  End If
  
  HoldRejectErrRows = Empty
  Call ExtractYnKeyword(aColumn, Errmsg, HoldID, "RejErrRow=", HoldRejectErrRows, "RER")
  If IsEmpty(HoldRejectErrRows) Then HoldRejectErrRows = FN_HoldRejectErrRows
  If IsEmpty(HoldRejectErrRows) Then
    HoldRejectErrRows = False ' Default Value.
    Errmsg = Errmsg & "Default Value - RejectErrRows=" & HoldRejectErrRows & _
      "  for ID=" & HoldID & "  Heading=""" & HoldLetterAndHeading & """" & vbCrLf
  End If
  
  HoldSaveDateChanged = Empty
  HoldSaveDateChangedLineLvl = Empty
  Call ExtractYnKeyword(aColumn, Errmsg, HoldID, "SavDtChange=", HoldSaveDateChangedLineLvl, "SDC")
  HoldSaveDateChanged = HoldSaveDateChangedLineLvl
  If IsEmpty(HoldSaveDateChanged) Then HoldSaveDateChanged = FN_HoldSaveDateChanged
  If IsEmpty(HoldSaveDateChanged) Then
    HoldSaveDateChanged = False ' Default Value.
    Errmsg = Errmsg & "Default Value - SaveDateChanged=" & HoldSaveDateChanged & _
      "  for ID=" & HoldID & "  Heading=""" & HoldLetterAndHeading & """" & vbCrLf
  End If
  If IsEmpty(HoldSaveDateChangedLineLvl) Then HoldSaveDateChangedLineLvl = False

  ' Save individual parm values in field level array(s)
  Call InsertNewElementIntoArray(RT_KeyField, HoldKeyField)
  Call InsertNewElementIntoArray(RT_ID, HoldID)
  Call InsertNewElementIntoArray(RT_ExcelColumnLetter, HoldExcelColumnLetter)
  Call InsertNewElementIntoArray(RT_ExcelHeadingText, HoldExcelHeadingText)
  Call InsertNewElementIntoArray(RT_FieldNameOutput, HoldFieldNameOutput)
  Call InsertNewElementIntoArray(RT_AcceptChanges, HoldAcceptChanges)
  Call InsertNewElementIntoArray(RT_AllowChange2Blank, HoldAllowChange2Blank)
  Call InsertNewElementIntoArray(RT_RejectErrFile, HoldRejectErrFile)
  Call InsertNewElementIntoArray(RT_RejectErrRows, HoldRejectErrRows)
  Call InsertNewElementIntoArray(RT_SaveDateChanged, HoldSaveDateChanged)
  Call InsertNewElementIntoArray(RT_SaveDateChangedLineLvl, HoldSaveDateChangedLineLvl)
  
  aColumn = TrimChars(aColumn, ",")  ' Trim commas and blanks....
  If Verify(aColumn, ", ") <> 0 Then
    Errmsg = Errmsg & "*** Column ID=" & HoldID & "  """ & aColumn & """ is Invalid or Duplicated Keyword syntax." & vbCrLf
  End If

Next II


Errmsg = Errmsg & VerifyImportConfig
'ParseImportConfig detected errors or warnings for FileName ID="Config-1"  invalid/unknown syntax in:
'ParseImportConfig detected errors/warnings or invalid/unknown syntax for FileName ID="Config-1"  --------
If Errmsg <> "" Then
  ParseImportConfig = "ParseImportConfig detected errors/warnings or invalid syntax for FileName ID=""" & RT_FileID & """  --------" & _
          vbCrLf & Errmsg
  DebugPrintOn (ParseImportConfig)
  Exit Function
End If


End Function


Private Function ExtractYnKeyword(ByRef ParmStr, _
                          ByRef Errmsg, _
                          ByVal ColID As String, _
                          ByVal KeyW, _
                          ByRef ReturnNormalizedValue, _
                          Optional ByVal AbreviatedKeyW As String = "")
                          
' This function will handle extracting the KeyWord Y/N valued from the ParmStr and then validate
' and then validate by checking for the values of Yes, No, True, False, Y, N, T, F, E or Empty.
'
' The value will be normalized to True or False in the ReturnTrueFalse parm.  It will be left empty if there's
'     an error.
' The Function return value will be an error message that can be used to report a bad value.

Dim YesNoParm
ReturnNormalizedValue = Empty
YesNoParm = ExtractKeyWord(ParmStr, KeyW, AbreviatedKeyW)

If YesNoParm = "Empty" Or YesNoParm = "E" Or YesNoParm = "" Then Exit Function

If YesNoParm = "Yes" Or YesNoParm = "Y" Or YesNoParm = "True" Or YesNoParm = "T" Then
  ReturnNormalizedValue = True
  Exit Function
End If

If YesNoParm = "No" Or YesNoParm = "N" Or YesNoParm = "False" Or YesNoParm = "F" Then
  ReturnNormalizedValue = False
  Exit Function
End If

If IsEmpty(Errmsg) Then Errmsg = ""
If ColID <> "" Then
  Errmsg = Errmsg & "*** Column ID=" & ColID & "  """ & KeyW & YesNoParm & """ is not valid value. (Can only be Yes/No)" & vbCrLf
 Else
  Errmsg = Errmsg & "*** """ & KeyW & YesNoParm & """ is not valid value. (Can only be Yes/No)" & vbCrLf
End If

End Function





Private Function CreateNewImportConfig(Optional ByVal aPrePickedFile As String = "", _
                                      Optional ByRef NewFileName As String, _
                                      Optional ByRef ReturnWkShtNames As Variant, _
                                      Optional ByVal SaveNewConfig As Boolean = True) As String

' 1) Pick file name using FilePrePicker()
' 2) get WkSht names in the file.  GetExcelWkShtNames()
' 3) get headings array   GetExcelHeadingText()
' 4) determine which WkShts should be selected for processing:
'   a) Does the WkSht contain a ; after the WkSht name?
'   b) Do the headings have any parameters specified?
'   c) if a or b are true that select this WkSht.
'   d) if no WkShts are selected in the worksheet, then assume that all should be selected.
'   e)
' 5) Take the list of selected WkShts, and build the import config.
' 6) Save the import config using FileSave()

' Pick the file.
Dim XslFileAndPath: XslFileAndPath = FilePrePicker(aPrePickedFile, "*.xlsx, *.xlsm, *.xls, *.csv")
DebugPrintOn ("Processing: " & XslFileAndPath)

Dim WkShtNames, Headings, II, JJ, WkShtShouldBeSelected As Boolean
Dim SelectedWkShtNames(), NonSelectedWkShtNames()
'2) get WkSht names in the file.  GetExcelWkShtNames()
WkShtNames = GetExcelWkShtNames(XslFileAndPath)

' 3) get headings array   GetExcelHeadingText()
' 4) determine which WkShts should be selected for processing:
'   a) Does the WkSht contain a ; after the WkSht name?
'   b) Do the headings have any parameters specified?
'   c) if a or b are true that select this WkSht.
If IsArrayAllocated(WkShtNames) Then
  For II = LBound(WkShtNames) To UBound(WkShtNames)
    WkShtShouldBeSelected = False
    If InStr(WkShtNames(II), ":") <> 0 Or InStr(WkShtNames(II), ";") <> 0 Then WkShtShouldBeSelected = True
    Headings = GetExcelHeadingText(XslFileAndPath, WkShtNames(II))
    If IsArrayAllocated(Headings) Then
      For JJ = LBound(Headings) To UBound(Headings)
        If InStr(Headings(JJ), ":") <> 0 Or InStr(Headings(JJ), ";") <> 0 Then WkShtShouldBeSelected = True
      Next JJ
    End If
    If WkShtShouldBeSelected Then
      Call InsertNewElementIntoArray(SelectedWkShtNames, WkShtNames(II))
     Else
      Call InsertNewElementIntoArray(NonSelectedWkShtNames, WkShtNames(II))
    End If
  Next II
End If

'   d) if no WkShts are selected in the worksheet, then assume that all should be selected.
If Not IsArrayAllocated(SelectedWkShtNames) Then
  Erase NonSelectedWkShtNames
  For II = LBound(WkShtNames) To UBound(WkShtNames)
    Call InsertNewElementIntoArray(SelectedWkShtNames, WkShtNames(II))
  Next II
End If

' 5) Take the list of selected WkShts, and build the import config.
If Not IsArrayAllocated(SelectedWkShtNames) Then
  MsgBox ("No Excel WkShts were selected.")
  Exit Function
End If

Dim ImportConfigFileData As String:  ImportConfigFileData = ""
For II = LBound(SelectedWkShtNames) To UBound(SelectedWkShtNames)
  DebugPrintOn ("Processing-" & SelectedWkShtNames(II))
  Headings = GetExcelHeadingText(XslFileAndPath, SelectedWkShtNames(II))
  Call GenImportConfigContent(ImportConfigFileData, XslFileAndPath, SelectedWkShtNames(II), Headings, II + 1)
Next II
If IsArrayAllocated(NonSelectedWkShtNames) Then
  ImportConfigFileData = ImportConfigFileData & vbCrLf & vbCrLf
  For II = LBound(NonSelectedWkShtNames) To UBound(NonSelectedWkShtNames)
    ImportConfigFileData = ImportConfigFileData & _
        "/**** WkSht NOT processed """ & NonSelectedWkShtNames(II) & """  due to not selected with "";""  */" & vbCrLf
Next II
End If
CreateNewImportConfig = ImportConfigFileData

' 6) Save the import config using FileSave()
Dim ImportConfig: ImportConfig = XslFileAndPath & ".config"
If ImportConfigFileData <> "" And SaveNewConfig Then
  NewFileName = FileSave(ImportConfigFileData, ImportConfig, "CreateNewImportConfig")
  Call MsgBox("New Config file has been created and placed in directory:" & vbCrLf & vbCrLf _
            & """" & ImportConfig & """" & vbCrLf & vbCrLf _
            & "This NEW Config Template can be altered/updated using your Text Editor before being used." & vbCrLf & vbCrLf _
            & "A Compile will be done on the NEW config to check for errors.", _
            vbOKOnly, "New Config File Created:  Etl_123")
End If
ReturnWkShtNames = SelectedWkShtNames
' Clean up WkSht names and remove any key word parameters.
If IsArrayAllocated(ReturnWkShtNames) Then
  For II = LBound(ReturnWkShtNames) To UBound(ReturnWkShtNames)
    JJ = InStr(1, ReturnWkShtNames(II), ";")
    If JJ = 0 Then JJ = InStr(1, ReturnWkShtNames(II), ":")
    If JJ > 1 Then ReturnWkShtNames(II) = Left(ReturnWkShtNames(II), JJ - 1)
  Next II
End If


DebugPrintOn ("CreateNewImportConfig is Done....")
End Function

Private Sub GenImportConfigContent(ByRef ImportConfigFileData As String, _
                                 ByVal InpFileAndPath As String, _
                                 ByVal aWkShtInput As String, _
                                 ByVal Headings As Variant, _
                                 ByVal ConfigNum As Variant)
                                  
                                 
Dim XX: XX = ImportConfigFileData
Dim II, JJ, crlf: crlf = vbCrLf & Chr(10)
Dim TableParms: TableParms = ""
Dim FileNM, HoldFileNM, excelname, Q
Dim Original_aWkShtInput:  Original_aWkShtInput = aWkShtInput
Dim Original_Headings:  Original_Headings = Headings

' Work areas for the FileName parm....
Dim aExcel, aExcel2, aWkSht, aDBTable, aUniqueKey  ' WkSht Level only
Dim aID, aAcceptChanges, aChange2Blank, aRejErrRow, aRejErrFile, aSavDtChange  ' Shared..
Dim aHeading, aCol, aDB, aKeyField  ' Column Level only
Dim Warnings()
'Erase aID: Erase aAcceptChanges: Erase aChange2Blank:  Erase aRejErrRow: Erase aRejErrFile: Erase aSavDtChange

If XX = "" Then XX = "/*  """ & InpFileAndPath & """  " & Format(Now(), "mm-dd-yyyy hh:mm:ss") & "  */"

JJ = InStrRev(InpFileAndPath, "\")
aExcel = "Default"
If JJ <> 0 Then aExcel = Mid(InpFileAndPath, JJ + 1)

JJ = InStr(aWkShtInput, ";")
If JJ <> 0 Then
  TableParms = Trim(Mid(aWkShtInput, JJ + 1))
  aWkShtInput = Trim(Left(aWkShtInput, JJ - 1))
End If

aExcel2 = ExtractKeyWord(TableParms, "Excel")
aDBTable = ExtractKeyWord(TableParms, "DBTable", "DB")
aUniqueKey = ExtractKeyWord(TableParms, "UniqueKey", "UK")
aID = ExtractKeyWord(TableParms, "ID")
aAcceptChanges = NormalizedYN(ExtractKeyWord(TableParms, "AcceptChanges", "AC"))
aChange2Blank = NormalizedYN(ExtractKeyWord(TableParms, "Change2Blank", "C2B"))
aRejErrRow = NormalizedYN(ExtractKeyWord(TableParms, "RejErrRow", "RER"))
aRejErrFile = NormalizedYN(ExtractKeyWord(TableParms, "RejErrFile", "REF"))
aSavDtChange = NormalizedYN(ExtractKeyWord(TableParms, "SavDtChange", "SDC"))

If aExcel2 <> "" Then _
  Call InsertNewElementIntoArray(Warnings, "Warning1 - Excel=" & aExcel2 & " removed.  Cannot originate from renamed Excel WkSht.")

If aUniqueKey = "" Then
  Call InsertNewElementIntoArray(Warnings, "Warning3 - UniqueKey=" & aUniqueKey & " is Required. Must be filled in (Yes/No).")
  aUniqueKey = "**Default**"
End If

If aDBTable = "" Then
  Call InsertNewElementIntoArray(Warnings, "Warning4 - DBTable=" & aDBTable & " is Required. Must be filled in.")
  aDBTable = "**Default**"
End If

If aID = "" Then
  'aID = aExcel & "/" & aWkShtInput
  'aID = aExcel & "/" & ConfigNum
  aID = "Config-" & ConfigNum
  Call InsertNewElementIntoArray(Warnings, "Warning5 - ID=" & aID & " is the DEFAULT.")
End If

If TableParms <> "" Then _
  Call InsertNewElementIntoArray(Warnings, "Warning6 - Unrecognized Text """ & TableParms & _
          """ Found in """ & aWkSht & """ Excel WkSht Name")

If aAcceptChanges = "" Then aAcceptChanges = "Yes"
If aChange2Blank = "" Then aChange2Blank = "Yes"
If aRejErrRow = "" Then aRejErrRow = "Yes"
If aRejErrFile = "" Then aRejErrFile = "No"
If aSavDtChange = "" Then aSavDtChange = "No"

If TableParms <> "" Then TableParms = "," & TableParms

' Now Build out the FileName=(...)
FileNM = crlf & crlf & crlf & "/**** Original Excel WkSht Name -  """ & Original_aWkShtInput & """"

If IsArrayAllocated(Warnings) Then
  For II = LBound(Warnings) To UBound(Warnings)
    FileNM = FileNM & " / " & Warnings(II)
  Next II
  Erase Warnings
End If
FileNM = FileNM & "  */" & crlf

FileNM = FileNM _
  & "FileName=(" _
  & "ID=""" & aID & """, " _
  & "Excel=""" & Translate(aExcel, "***", """()") & """, " _
  & "WkSht=""" & Translate(aWkShtInput, "***", """()") & """, " _
  & "DBTable=""" & aDBTable & """, " _
  & "UniqueKey=" & aUniqueKey & ", " & crlf & "                         " _
  & "AcceptChanges=" & aAcceptChanges & ", " _
  & "Change2Blank=" & aChange2Blank & ", " _
  & "RejErrRow=" & aRejErrRow & ", " _
  & "RejErrFile=" & aRejErrFile & ", " _
  & "SavDtChange=" & aSavDtChange _
  & TableParms _
  & ") " & crlf

XX = XX & FileNM

'  Split the Headings into 2 components. Headings and ColParms
Dim ColParms(), ColParmsFound(), DupsYN()
' First allocate the 2 arrays.
For II = LBound(Headings) To UBound(Headings)
  Call InsertNewElementIntoArray(ColParms, "")  ' Allocate the array.
  Call InsertNewElementIntoArray(ColParmsFound, False)  ' Allocate the array.
  Call InsertNewElementIntoArray(DupsYN, False) ' Allocate the array.
Next II
' Next Split the components
For II = LBound(Headings) To UBound(Headings)
  JJ = InStr(Headings(II), ";")
  If JJ = 0 Then JJ = InStr(Headings(II), ":")
  If JJ <> 0 Then
    ColParms(II) = Trim(Mid(Headings(II), JJ + 1))
    ColParmsFound(II) = True
    Headings(II) = Trim(Left(Headings(II), JJ - 1))
  End If
Next II
' Mark the duplicate headings...
For II = LBound(Headings) To UBound(Headings)
  For JJ = LBound(Headings) To UBound(Headings)
    If JJ <> II And Headings(II) = Headings(JJ) Then DupsYN(II) = True
  Next JJ
Next II

Dim HighWaterSize: HighWaterSize = 50

For II = LBound(Headings) To UBound(Headings)
  Erase Warnings
  aID = ExtractKeyWord(ColParms(II), "ID")
  If aID = "" Then aID = "ID-" & TrimChars(GetColumnHeadingName(II + 1), "1")
  
  aAcceptChanges = NormalizedYN(ExtractKeyWord(ColParms(II), "AcceptChanges", "AC"))
  aChange2Blank = NormalizedYN(ExtractKeyWord(ColParms(II), "Change2Blank", "C2B"))
  aRejErrRow = NormalizedYN(ExtractKeyWord(ColParms(II), "RejErrRow", "RER"))
  aRejErrFile = NormalizedYN(ExtractKeyWord(ColParms(II), "RejErrFile", "REF"))
  aSavDtChange = NormalizedYN(ExtractKeyWord(ColParms(II), "SavDtChange", "SDC"))
  
  'Dim aHeading, aCol, aDB, aKeyField  ' Column Level only
  
  aDB = ExtractKeyWord(ColParms(II), "DB")
  aKeyField = NormalizedYN(ExtractKeyWord(ColParms(II), "KeyField", "KF"))
  
  aCol = ""
  If DupsYN(II) And aDB <> "" Then _
    aCol = TrimChars(GetColumnHeadingName(II + 1), "1") ' Trim off the row number leaving only the column letter.
  
  If ColParms(II) <> "" Then _
    Call InsertNewElementIntoArray(Warnings, "Warning A - Unrecognized Text """ & ColParms(II) & """")

  If aDB = "" Then _
    Call InsertNewElementIntoArray(Warnings, "Warning B - This field is not mapped to database.")
    
  FileNM = "Column=(" _
    & "ID=""" & aID & """, " _
    & "Heading=""" & Translate(Headings(II), "***", """()") & """, "
  If aCol <> "" Then _
    FileNM = FileNM & "Col=" & aCol & ", "
  If aDB <> "" Then _
    FileNM = FileNM & "DB=""" & aDB & """, "
  If aKeyField <> "" Then _
    FileNM = FileNM & "KeyField=" & aKeyField & ", "
  If aAcceptChanges <> "" Then _
    FileNM = FileNM & "AcceptChanges=" & aAcceptChanges & ", "
  If aChange2Blank <> "" Then _
    FileNM = FileNM & "Change2Blank=" & aChange2Blank & ", "
  If aRejErrRow <> "" Then _
    FileNM = FileNM & "RejErrRow=" & aRejErrRow & ", "
  If aRejErrFile <> "" Then _
    FileNM = FileNM & "RejErrFile=" & aRejErrFile & ", "
  If aSavDtChange <> "" Then _
    FileNM = FileNM & "SavDtChange=" & aSavDtChange & ", "
    
  FileNM = TrimChars(FileNM, ",") & ")"
  If HighWaterSize < Len(FileNM) Then HighWaterSize = Len(FileNM)
  FileNM = Rpad(FileNM, HighWaterSize) & _
           " /* Orig Heading-""" & RemoveWhiteSpace(Original_Headings(II)) & """ "
  
  If IsArrayAllocated(Warnings) Then
    FileNM = FileNM & "/"
    For JJ = LBound(Warnings) To UBound(Warnings)
      FileNM = FileNM & "  " & Warnings(JJ)
    Next JJ
  End If
  FileNM = FileNM & " */" & crlf
  If ColParmsFound(II) Then XX = XX & FileNM
Next II

GoTo ExitSub
ExitSub:
  ImportConfigFileData = XX ' Send the result back to caller....

End Sub
Private Function NormalizedYN(XX) As String
  ' Yes, No, True, False, Y, N, T, F, E or Empty

NormalizedYN = XX
If XX = "Yes" Or XX = "Y" Or XX = "T" Or XX = "True" Then NormalizedYN = "Yes"
If XX = "No" Or XX = "N" Or XX = "False" Or XX = "F" Then NormalizedYN = "No"
If XX = "Empty" Or XX = "E" Then NormalizedYN = "Empty"

End Function




Private Function GetColumnHeadingName(ByVal ColNum As Long) As String

Dim ColumnNames(702)

'  Initialize the ColumnNames string.....
Dim alpha  As String
Dim I As Long, J As Long, JJ As Long
alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
         
For I = 0 To Len(alpha)
   For J = 1 To Len(alpha)
      If JJ = UBound(ColumnNames) Then GoTo ExitFunction
      JJ = JJ + 1
      If I <> 0 Then ColumnNames(JJ) = Mid(alpha, I, 1) & Mid(alpha, J, 1)
      If I = 0 Then ColumnNames(JJ) = Mid(alpha, J, 1)
   Next J
Next I
ExitFunction:
  If ColNum > UBound(ColumnNames) Then GetColumnHeadingName = "AAA" Else GetColumnHeadingName = ColumnNames(ColNum)
  GetColumnHeadingName = GetColumnHeadingName & "1"

End Function


Private Function CreateConfigFunction(ByVal ConfigString As String, _
                                     ByVal FilePath As String) As String

Dim HoldConfigText2  As String
Dim II As Long, JJ As Long, aChar As String, x As String
Dim NL_Delimiter    As String

'ConfigString = GetEntireFile(FilePath)
If InStr(1, ConfigString, vbCr) Then NL_Delimiter = vbCr Else NL_Delimiter = vbLf

For II = 1 To Len(ConfigString)
  aChar = Mid(ConfigString, II, 1)
  If aChar = """" Then aChar = """"""  ' Turn all double quotes into a pair of double quotes.
  If NL_Delimiter = vbCr Then
    If aChar = vbLf Then aChar = ""   ' Remove all LF characters leaving only the CR to delimit lines.
   Else
    If aChar = vbCr Then aChar = ""   ' Remove all CR characters leaving only the LF to delimit lines.
  End If
  HoldConfigText2 = HoldConfigText2 & aChar
Next II

Dim HoldFunctionName  As String:   HoldFunctionName = ""
II = InStrRev(FilePath, "\")
If II > 0 Then HoldFunctionName = Mid(FilePath, II + 1)

II = InStrRev(HoldFunctionName, ".xl")
If II = 0 Then II = InStrRev(HoldFunctionName, ".csv")
If II = 0 Then II = InStrRev(HoldFunctionName, ".config")
If II < 1 Then
  Call MsgBox("CreateConfigFunction is UNABLE to create Function.  Unexpected file name:" & vbCrLf & vbCrLf & _
              """" & HoldFunctionName & """ is not a valid config file name")
  Exit Function
End If
HoldFunctionName = "Get" & Left(HoldFunctionName, II - 1) & "Config"

Dim NewConfigFuncFilePath   As String: NewConfigFuncFilePath = FilePath & "-Function.txt"
Dim ConfigLines() As String
ConfigLines = Split(HoldConfigText2, NL_Delimiter)

Dim NewFuncText As String

NewFuncText = "Private Function " & HoldFunctionName & "() As String" & vbCrLf & vbCrLf & _
              "Dim T As String: T = ""/* Etl123 Config string generated by function """"" & HoldFunctionName & "()""""  *************/""  & vbCrLf" & vbCrLf & vbCrLf
 
JJ = -1
For II = UBound(ConfigLines) To LBound(ConfigLines) Step -1
  If Trim(ConfigLines(II)) <> "" Then
    JJ = II
    Exit For
  End If
Next II
If JJ >= 0 Then ReDim Preserve ConfigLines(JJ)
 
For II = LBound(ConfigLines) To UBound(ConfigLines)
  If InStr(1, ConfigLines(II), "/* Etl123 Config generated by") = 0 Then _
    NewFuncText = NewFuncText & "T = T & """ & ConfigLines(II) & """ & vbCrLf" & vbCrLf
Next II
              
NewFuncText = NewFuncText & vbCrLf & HoldFunctionName & " = T" & vbCrLf & "End Function"
'DebugPrintOn (NewFuncText)
CreateConfigFunction = NewFuncText

End Function

Private Function EtlImport(ByVal preSelectedInput As String, _
                          ByVal WkShtName As String, _
                          ByVal WkShtNum As Long, _
                          ByVal ReptFileName As String, _
                          Optional ByVal printDebugLog As Boolean = True) As Boolean
                          
' This function will return True if import has NO errors.  False if import finished with errors.
   EtlImport = True  ' Set initial value, assuming no errors will be found.
   
Dim I As Long
Dim strSql  As String

Dim WkShtNameFull As String:  WkShtNameFull = GetFullWkShtName(preSelectedInput, WkShtName, WkShtNum)
   
aPrintDebugLog = printDebugLog ' set the Public flag....
If Not aPrintDebugLog Then DebugPrint ("Debug.Print log for ""EtlImport"" is turned off.....")

HoldSelItem = preSelectedInput   ' Combined path and file name.
I = InStrRev(HoldSelItem, "\")
If I = 0 Then I = InStrRev(HoldSelItem, ":")
HoldFileName = Mid(HoldSelItem, I + 1)
HoldFilePath = Left(HoldSelItem, I - 1)

Call GetImportSpecifications  ' Fill in Specifications....
   
aNoImportErrorsWereFound = True ' Initial value of this flag
   
' Now process the selected file along with all Worksheets found in the Import Specification...
Dim JJ As Long
'DebugPrint (HoldFileName & " List of valid Excel WkShts for " & HoldFileName & "................")

aWorkSheetErrorFound = False '  Clear this flag before processing a worksheet.

If RT_WorkSheetName <> WkShtName Then
  Call MsgBox("Something is Wrong:  EtlImport has mis-match on WkSht names. " & vbCrLf & _
              "RT_WorkSheetName=""" & RT_WorkSheetName & """" & vbCrLf & _
              "WkShtName=""" & WkShtName & """  WkShtNum=" & WkShtNum & vbCrLf & vbCrLf & _
              "These should always match. There is a program bug." & vbCrLf & _
              "Process will abort.")
  End
End If


Call Process_A_Worksheet(RT_WorkSheetName, WkShtNameFull, WkShtNum, RT_OutputTableName, ReptFileName)

DelTblS ("ETL123_table_" & RT_WorkSheetName & "*") ' Clean up previous table.
If Not aWorkSheetErrorFound Then
  DelTbl ("ETL123_table")
 Else
   DoCmd.Rename "ETL123_table_" & RT_WorkSheetName & Format(Now(), "_" & "yyyymmddhhmmss"), _
                                                                            acTable, "ETL123_table"
End If
      
EtlImport = aNoImportErrorsWereFound '  Set the return value...
      
If preSelectedInput = "" Then MsgBox ("HoldFileName=" & HoldFileName)
Call DelTbl("ETL123_table_count")
Call Remove_Table_Field("etl123_row_number", RT_OutputTableName)

End Function

Private Function GetFullWkShtName(preSelectedInput As String, _
                                WkShtName As String, _
                                WkShtNum As Long) As String

Dim FullWkShtNames As Variant, II As Long
FullWkShtNames = GetExcelWkShtNames(preSelectedInput)

GetFullWkShtName = FullWkShtNames(WkShtNum - 1)

End Function

Private Sub GetImportSpecifications()
                       
Dim I   As Long

numOfSpecs = UBound(RT_ExcelColumnLetter) + 1  ' Initial value / Clear the output and re-allocate to proper dimentions.
ReDim aOutput_Table_Name(1 To numOfSpecs)
ReDim aExcel_Column_Number(1 To numOfSpecs)
ReDim aExcel_Heading_Text(1 To numOfSpecs)
ReDim aField_Name_Output(1 To numOfSpecs)
ReDim aData_Type(1 To numOfSpecs)
ReDim aData_Type_Enum(1 To numOfSpecs)
ReDim aXcelFieldHasLongText(1 To numOfSpecs)
ReDim aDup_Key_Field(1 To numOfSpecs)
ReDim aReject_Err_File(1 To numOfSpecs)
ReDim aReject_Err_Rows(1 To numOfSpecs)
ReDim aAccept_Changes(1 To numOfSpecs)
ReDim aAllowChange2Blank(1 To numOfSpecs)
ReDim aSave_Date_Changed(1 To numOfSpecs)

ReDim aold_Val(1 To numOfSpecs)       ' Hold the old value from the target table.
ReDim afield_err(1 To numOfSpecs)     ' Any error that is found with this field.
ReDim holdMergedData(1 To numOfSpecs)
ReDim holdKeyNames(1 To numOfSpecs)

For I = LBound(aOutput_Table_Name) To UBound(aOutput_Table_Name)
  aOutput_Table_Name(I) = RT_OutputTableName
  aExcel_Column_Number(I) = RT_ExcelColumnLetter(I - 1)
  aExcel_Heading_Text(I) = RT_ExcelHeadingText(I - 1)
  aField_Name_Output(I) = RT_FieldNameOutput(I - 1)
      
  aData_Type(I) = "Short Text/Text"  ' Default value....
  If aOutput_Table_Name(I) <> "" Then _
    aData_Type(I) = Field_Type(aOutput_Table_Name(I), aField_Name_Output(I), aData_Type_Enum(I))
   
  aDup_Key_Field(I) = RT_KeyField(I - 1)
      
  aReject_Err_File(I) = RT_RejectErrFile(I - 1)
  aReject_Err_Rows(I) = RT_RejectErrRows(I - 1)
  aAccept_Changes(I) = RT_AcceptChanges(I - 1)
  aAllowChange2Blank(I) = RT_AllowChange2Blank(I - 1)
      
  aSave_Date_Changed(I) = RT_SaveDateChanged(I - 1)
  If aSave_Date_Changed(I) Then Add_Date_Changed_to_Rows = True
Next I


End Sub


Private Sub Init_Excel_Names(ColumnNames() As String)
                       
Dim alpha  As String
Dim I As Long, J As Long, JJ As Long
alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
         
Dim U   As Long
On Error Resume Next
ReDim ColumnNames(676)
On Error GoTo 0
U = UBound(ColumnNames)
If U > 702 Then U = 702

For I = 0 To Len(alpha)
   For J = 1 To Len(alpha)
      If JJ = UBound(ColumnNames) Then Exit Sub
      JJ = JJ + 1
      If I <> 0 Then ColumnNames(JJ) = Mid(alpha, I, 1) & Mid(alpha, J, 1)
      If I = 0 Then ColumnNames(JJ) = Mid(alpha, J, 1)
   Next J
Next I
End Sub

Private Function AddTableFields(Output_Table As String)
                       
Dim tdf       As TableDef
Dim target    As TableDef
Dim db As DAO.Database, Rst As Recordset
Dim fld As DAO.Field
Dim prop As DAO.Property
Dim I As Long, J As Long, myDataType As Long

Set db = CurrentDb
Set tdf = db.TableDefs("ETL123_table")
Set target = db.TableDefs(Output_Table)

'https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/datatypeenum-enumeration-dao
'https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/datatypeenum?view=sql-server-2017
'http://allenbrowne.com/ser-49.html
' Create the new fields in the work table....
For I = 1 To numOfSpecs   '  Add new fields to the Work Table  "ETL123_table"
   If aField_Name_Output(I) <> "" Then
     myDataType = aData_Type_Enum(I)
     If myDataType = 10 And aXcelFieldHasLongText(I) Then myDataType = 12
     tdf.Fields.Append tdf.CreateField(aField_Name_Output(I), myDataType)
     tdf.Fields.Append tdf.CreateField(aField_Name_Output(I) & "_err", 10)
   End If
Next I
tdf.Fields.Append tdf.CreateField("ErrorText", 10)     ' dbText
tdf.Fields.Append tdf.CreateField("matched_target_table", 10)     ' dbText
tdf.Fields.Append tdf.CreateField("row_contains_error", 10)       ' dbText
tdf.Fields.Append tdf.CreateField("reject_err_row", 10)       ' dbText
tdf.Fields.Append tdf.CreateField("row_has_changed", 10)       ' dbText

If Add_Date_Changed_to_Rows Then
  tdf.Fields.Append tdf.CreateField("create_date_time", 8)     ' dbDate     Now()
  tdf.Fields.Append tdf.CreateField("create_user", 10)     ' dbText         HoldSelItem
  tdf.Fields.Append tdf.CreateField("create_program", 10)     ' dbText      "EtlImport"
  tdf.Fields.Append tdf.CreateField("create_file", 10)     ' dbText         HoldSelItem
  
  tdf.Fields.Append tdf.CreateField("update_date_time", 8)     ' dbDate     Now()
  tdf.Fields.Append tdf.CreateField("update_user", 10)     ' dbText         HoldSelItem
  tdf.Fields.Append tdf.CreateField("update_program", 10)     ' dbText      "EtlImport"
  tdf.Fields.Append tdf.CreateField("update_file", 10)     ' dbText         HoldSelItem
  
  If Not fieldexists("create_date_time", Output_Table) Then target.Fields.Append target.CreateField("create_date_time", 8)  ' dbDate     Now()
  If Not fieldexists("create_user", Output_Table) Then target.Fields.Append target.CreateField("create_user", 10)     ' dbText         HoldSelItem
  If Not fieldexists("create_program", Output_Table) Then target.Fields.Append target.CreateField("create_program", 10)     ' dbText      "EtlImport"
  If Not fieldexists("create_file", Output_Table) Then target.Fields.Append target.CreateField("create_file", 10)     ' dbText         HoldSelItem
  
  If Not fieldexists("update_date_time", Output_Table) Then target.Fields.Append target.CreateField("update_date_time", 8)  ' dbDate     Now()
  If Not fieldexists("update_user", Output_Table) Then target.Fields.Append target.CreateField("update_user", 10)     ' dbText         HoldSelItem
  If Not fieldexists("update_program", Output_Table) Then target.Fields.Append target.CreateField("update_program", 10)     ' dbText      "EtlImport"
  If Not fieldexists("update_file", Output_Table) Then target.Fields.Append target.CreateField("update_file", 10)     ' dbText         HoldSelItem
End If

' Now add any Temporary fields to the target table.
If fieldexists("etl123_row_number", Output_Table) Then target.Fields.Delete ("etl123_row_number")
target.Fields.Append target.CreateField("etl123_row_number", 7)     '  dbDouble

Call ClearLongTextFormatProp("ETL123_table")  ' Make sure "Long Text" fields are set properly.

Dim srcField   As String   ' Source Field string......
strSql = ""
For I = 1 To numOfSpecs '  Update data fields in the ETL123_table....
  If aField_Name_Output(I) <> "" Then
   If srcField = "" Then strSql = "UPDATE ETL123_table SET "
   srcField = "[ETL123_table].[F" & xlColNum(aExcel_Column_Number(I)) & "]"
   If aData_Type(I) = "Date/Time" Then srcField = "CVDate(" & srcField & ")"
   srcField = "ETL123_table.[" & aField_Name_Output(I) & "] = " & srcField & ", "
   strSql = strSql + srcField
  End If
Next I
' Include Update Stamp
If Add_Date_Changed_to_Rows Then
  strSql = strSql & "ETL123_table.create_date_time = Now(), " & _
                    "ETL123_table.create_user = " & Scrub(Environ("USERNAME")) & ", " & _
                    "ETL123_table.create_program = 'EtlImport', " & _
                    "ETL123_table.create_file = " & Scrub(HoldSelItem) & ", "
  strSql = strSql & "ETL123_table.update_date_time = Now(), " & _
                    "ETL123_table.update_user = " & Scrub(Environ("USERNAME")) & ", " & _
                    "ETL123_table.update_program = 'EtlImport', " & _
                    "ETL123_table.update_file = " & Scrub(HoldSelItem) & ", "
End If

If Len(strSql) > 0 Then strSql = Left(strSql, Len(strSql) - 2) & ";"

DebugPrint ("Step #2 'AddTableFields' - " & strSql)
DoCmd.SetWarnings False
DoCmd.RunSQL (strSql)
DoCmd.SetWarnings True

Set fld = Nothing
Set tdf = Nothing
Set target = Nothing
Set db = Nothing
Set prop = Nothing

End Function

Private Sub Create_ETL123_table_Errors_SQL()

Dim strSql  As String
Dim Errmsg As String, errResponse As Long

If TableExists("ETL123_table_errors") Then GoTo Exit_Function

On Error GoTo Error_Handler
DoCmd.SetWarnings False
strSql = "CREATE TABLE ETL123_table_errors (" _
       & "ID AUTOINCREMENT PRIMARY KEY, " _
       & "file_name                       TEXT(255), " _
       & "ETL123_rept_name                TEXT(255), " _
       & "excel_work_sheet_name           TEXT(255), " _
       & "source_file_primary_key_heading TEXT(255), " _
       & "source_file_primary_key_value   TEXT(255), " _
       & "error_field_cell_column         TEXT(255), " _
       & "error_field_cell_row            INTEGER,   " _
       & "error_field_cell_number         TEXT(255), " _
       & "error_field_heading             TEXT(255), " _
       & "error_mapped_field_name         TEXT(255), " _
       & "error_mapped_field_data_type    TEXT(255), " _
       & "error_field_value               MEMO,      " _
       & "error_message                   TEXT(255), " _
       & "file_name_full                  TEXT(255)); "
DoCmd.RunSQL (strSql)
DoCmd.SetWarnings True

Exit_Function:
  aWorkSheetErrorFound = True       ' Set this flag to be used later.
  aNoImportErrorsWereFound = False ' Set this flag to be used later to indicate an error occured.

Exit Sub

Error_Handler:
  
  Select Case err.Number
  Case 3010
    Resume Next
  Case 3211
    Errmsg = "Error number: " & Str(err.Number) & vbNewLine & _
             "Source: " & err.source & vbNewLine & _
             "Description: " & err.Description
    errResponse = MsgBox(Errmsg, vbRetryCancel)
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
    If errResponse = 4 Then Resume
    
    End
  Case Else
    Errmsg = "Error number: " & Str(err.Number) & vbNewLine & _
             "Source: " & err.source & vbNewLine & _
             "Description: " & err.Description
    MsgBox (Errmsg)
    DebugPrintOn (vbCrLf & "****" & Errmsg & vbCrLf)
    End
    Resume
  End Select
End Sub
  

Private Sub Process_Excel_Data_Error_Check(ByVal fldPtr As Long, _
                                           ByVal WorkSheetName As String, _
                                           ByVal ReptFileName As String, _
                                           ByRef NumberReportedErrs As Variant)
                       
Dim I As Long, J As Long, JJ
Dim Rst As DAO.Recordset
Dim strSql As String
Dim HOLD_KEY_VALUES     As String
Dim Entire_Worksheet_Should_be_Rejected  As Boolean
Entire_Worksheet_Should_be_Rejected = False
Dim aLen

If aField_Name_Output(fldPtr) = "" Then Exit Sub   ' This is not an output field.
If aData_Type(fldPtr) = "Short Text/Text" Then Exit Sub

 'Build the select SQL string
 strSql = "SELECT"
 
 'Build the SELECT for all rows with this error....
 Dim Hold_Key_Heading   As String
 Dim Hold_Key_Num       As Long
 Dim ErrorMsg   As String
 For I = 1 To numOfSpecs   ' Get the key field values....
    If aDup_Key_Field(I) Then
       Hold_Key_Num = Hold_Key_Num + 1
       '  Need to give error msgbox if more than 12
       If Hold_Key_Num > 12 Then
          ErrorMsg = "Too many key fields specified for a table.  Cannot be more than 12 fields specified for " & WorkSheetName
          MsgBox (ErrorMsg)
          DebugPrintOn (vbCrLf & "****" & ErrorMsg & vbCrLf)
          End
       End If
       If Hold_Key_Num > 1 Then Hold_Key_Heading = Hold_Key_Heading & " / "
       Hold_Key_Heading = Hold_Key_Heading & aExcel_Heading_Text(I)
       If Hold_Key_Num > 1 Then strSql = strSql & ","
       strSql = strSql & " ETL123_table.F" & xlColNum(aExcel_Column_Number(I)) & " AS key" & Hold_Key_Num
       strSql = strSql & ", ETL123_table." & Br(aField_Name_Output(I)) & " AS akey" & Hold_Key_Num
    End If
 Next I
 strSql = strSql & ", ETL123_table.row_contains_error"
 strSql = strSql & ", ETL123_table.reject_err_row"
 
 strSql = strSql & ", ETL123_table.F" & xlColNum(aExcel_Column_Number(fldPtr)) & " AS bad_data"
 strSql = strSql & ", ETL123_table." & Br(aField_Name_Output(fldPtr)) & " AS abad_data"
 strSql = strSql & ", ETL123_table.[" & aField_Name_Output(fldPtr) & "_err] AS abad_data_err"
 
 strSql = strSql & ", ETL123_table.F" & xlColNum(aExcel_Column_Number(numOfSpecs)) & " AS row_number"
 
 strSql = strSql & " FROM ETL123_table WHERE ((Not (ETL123_table.F" & _
             xlColNum(aExcel_Column_Number(fldPtr)) & ")='blanks') AND ((ETL123_table." & _
             Br(aField_Name_Output(fldPtr)) & ") Is Null));"
 DebugPrint ("Process_Excel_Data_Error_Check=" & strSql)
 
 Set Rst = Application.CurrentDb.OpenRecordset(strSql)  ' Open recordset with intent to edit.
 If Rst.RecordCount = 0 Then
    Rst.Close
    Set Rst = Nothing
    Exit Sub
 End If
 'DebugPrint (strSql)
 Call Create_ETL123_table_Errors_SQL  ' Create table to hold found errors.
 
'   Top of the READ loop for Import Table with field errors....
 Do
    HOLD_KEY_VALUES = ""
    If Hold_Key_Num > 0 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & Rst!key1
    If Hold_Key_Num > 1 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & Rst!key2
    If Hold_Key_Num > 2 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & Rst!key3
    If Hold_Key_Num > 3 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & Rst!key4
    If Hold_Key_Num > 4 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & Rst!key5
    If Hold_Key_Num > 5 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & Rst!key6
    If Hold_Key_Num > 6 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & Rst!key7
    If Hold_Key_Num > 7 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & Rst!key8
    If Hold_Key_Num > 8 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & Rst!key9
    If Hold_Key_Num > 9 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & Rst!key10
    If Hold_Key_Num > 10 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & Rst!key11
    If Hold_Key_Num > 11 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & Rst!key12
 
    Dim ETL123_table_errors As DAO.Recordset
    Dim HoldErrMsg      As String
    Set ETL123_table_errors = CurrentDb.OpenRecordset("SELECT * FROM [ETL123_table_errors]")  ' Intend to edit
    ETL123_table_errors.AddNew
    ETL123_table_errors![file_name] = HoldFileName
    ETL123_table_errors![file_name_full] = HoldSelItem
    ETL123_table_errors![excel_work_sheet_name] = WorkSheetName
    ETL123_table_errors![source_file_primary_key_heading] = Hold_Key_Heading
    ETL123_table_errors![error_mapped_field_name] = aField_Name_Output(fldPtr)
    
    ETL123_table_errors![source_file_primary_key_value] = HOLD_KEY_VALUES
    ETL123_table_errors![error_field_cell_column] = aExcel_Column_Number(fldPtr)
    ETL123_table_errors![error_field_cell_row] = Rst!row_number
    ETL123_table_errors![error_field_cell_number] = aExcel_Column_Number(fldPtr) & Rst!row_number
    ETL123_table_errors![error_field_heading] = aExcel_Heading_Text(fldPtr)
    ETL123_table_errors![error_mapped_field_data_type] = aData_Type(fldPtr)
    If Len(Rst!bad_data) > 255 Then aLen = 255 Else aLen = Len(Rst!bad_data)
    ETL123_table_errors![error_field_value] = Left(Rst!bad_data, aLen)
    
    HoldErrMsg = "Data is invalid for """ & aData_Type(fldPtr) & """ format."
    If aReject_Err_Rows(fldPtr) Then HoldErrMsg = _
             "Entire Row is Rejected. Data is invalid for """ & aData_Type(fldPtr) & """ format."
    If aReject_Err_File(fldPtr) Then HoldErrMsg = _
            "Entire Worksheet-" & HoldFileName & "/" & WorkSheetName & " is Rejected. Data is invalid for """ & aData_Type(fldPtr) & """ format."
    ETL123_table_errors![error_message] = HoldErrMsg
    
    Dim HoldReptName   As String:  HoldReptName = ""
    ETL123_table_errors![ETL123_rept_name] = ""  ' Move empty string to eliminate Null value.
    Call FormatException4Print(ETL123_table_errors![error_field_cell_column], _
                               ETL123_table_errors![error_field_heading], _
                               ETL123_table_errors![error_field_cell_row], _
                               ETL123_table_errors![error_field_value], _
                               ETL123_table_errors![error_mapped_field_name], _
                               ETL123_table_errors![error_mapped_field_data_type], _
                               ETL123_table_errors![error_message], _
                               HoldReptName, _
                               ReptFileName, NumberReportedErrs)
    ETL123_table_errors![ETL123_rept_name] = HoldReptName
    ETL123_table_errors.Update
    ETL123_table_errors.Close
    Set ETL123_table_errors = Nothing
    Rst.Edit
    Rst!row_contains_error = "Y"
    Rst!abad_data_err = HoldErrMsg
    Dim HOLD_aReject_Err_Rows As Boolean
    HOLD_aReject_Err_Rows = aReject_Err_Rows(fldPtr)
    If aReject_Err_Rows(fldPtr) Then Rst!reject_err_row = "Y"
   ' ***** rst!reject_err_row = "Y"  '  Always reject entire row if there is an error.
    If aReject_Err_File(fldPtr) Then Entire_Worksheet_Should_be_Rejected = True

    Rst.Update
    Rst.MoveNext

    If Rst.EOF Then GoTo Finished_Do_Loop
Loop
Finished_Do_Loop:
Rst.Close
Set Rst = Nothing

If Entire_Worksheet_Should_be_Rejected Then
  strSql = "UPDATE ETL123_table SET ETL123_table.reject_err_row = 'Y';"
  DoCmd.RunSQL (strSql) ' Mark all rows to be rejected due to error (When spec says entire worksheet should be rejected)
End If

'        M A P P I N G                    Cell/Data       Error Message.......
'(C:Post Date-->[post_date] as "Date/Time")   (C7="abc")   Entire Row is Rejected. Data is invalid date format.
'"C:Post Date-->post_date as Date/Time"   C8="abc"   Entire Row is Rejected. Data is invalid date format.
'"C:Post Date-->post_date as Date/Time"   C9="abc"   Entire Row is Rejected. Data is invalid date format.
'"C:Post Date-->post_date as Date/Time"   C10="abc"  Entire Row is Rejected. Data is invalid date format.
'"C:Post Date-->post_date as Date/Time"   C11="abc"  Entire Row is Rejected. Data is invalid date format.

End Sub

Private Function FormatException4Print(ByVal XcelColumn As String, _
                                       ByVal XcelHeading As String, _
                                       ByVal XcelRow As String, _
                                       ByVal CellData As String, _
                                       ByVal MappedField As String, _
                                       ByVal DataType As String, _
                                       ByVal ErrMessage As String, _
                                       ByRef aReptFileName As String, _
                                       ByVal ReptFileName As String, _
                                       ByRef NumberReportedErrs As Variant) As String
                                       
'(C:Post Date-->[post_date] as "Date/Time")   (C7="abc")   Entire Row is Rejected. Data is invalid date format.

Dim MappedFld, CellDat

Dim JJ, MaxNumToList As Variant:  MaxNumToList = 20  ' Max number of errors to list in the Report.
  
' Get the file name...
JJ = InStrRev(ReptFileName, "\")
If JJ = 0 Then JJ = InStrRev(ReptFileName, ":")
If JJ = 0 Then aReptFileName = ReptFileName Else aReptFileName = Mid(ReptFileName, JJ + 1)

If Len(CellData) > 20 Then CellData = Left(CellData, 20) & "...."

If NumberReportedErrs = 0 Then _
  Call FileSave(vbCrLf & "Import Data Errors:" & vbCrLf, ReptFileName, , True)
NumberReportedErrs = NumberReportedErrs + 1

MappedFld = "(" & XcelColumn & ":" & XcelHeading & "-->" & "[" & MappedField & "] as " & """" & DataType & """)"
CellDat = "(" & XcelColumn & XcelRow & "=""" & CellData & """)"

FormatException4Print = Rpad(MappedFld, 25) & "  " & Rpad(CellDat, 20) & "  " & ErrMessage

If NumberReportedErrs = (MaxNumToList + 1) Then
  Call FileSave("Max number of " & MaxNumToList & " listed errors encountered.  " & _
                "ALL errors can be found in [ETL123_table_errors] table. ", _
                ReptFileName, , True)
End If
If NumberReportedErrs <= MaxNumToList Then _
  Call FileSave(FormatException4Print, ReptFileName, , True)
                                       
End Function

Private Sub Process_Field_Update_Data_Error_Check(Rst As Recordset, _
                                                  ByVal fldPtr As Long, _
                                                  ByVal target As Variant, _
                                                  ByVal source As Variant, _
                                                  ByVal err As String, _
                                                  ByVal WorkSheetName As String, _
                                                  ByVal ReptName As String, _
                                                  ByRef NumberReportedErrs As Variant)
Dim I As Long, J As Long, JJ
Dim ETL123_table_errors As DAO.Recordset
Dim strSql As String

Dim eXcelColAlfa   As String
eXcelColAlfa = xlColAlfa(fldPtr)
J = -1
For I = LBound(aExcel_Column_Number) To UBound(aExcel_Column_Number)
  If eXcelColAlfa = aExcel_Column_Number(I) Then
    J = I
    Exit For
  End If
Next I
If J < 0 Then Exit Sub ' This column was not mapped.
If aField_Name_Output(J) = "" Then Exit Sub   ' This is not an output field.

Dim xSource: xSource = source
If aData_Type(J) = "Yes/No" Then
  If source = "Y" Or source = "Yes" Or source = "T" Or source = "True" Then xSource = True
  If source = "N" Or source = "No" Or source = "F" Or source = "False" Then xSource = False
  If xSource = target Then Exit Sub
End If
If target = StrNormalize(xSource) Then Exit Sub                   ' The fields are equal, don't report as error.
If err <> "" Then Exit Sub                         ' If a previous error has been reported, then return.

Call Create_ETL123_table_Errors_SQL  ' Create table to hold found errors.
    
Set ETL123_table_errors = CurrentDb.OpenRecordset("SELECT * FROM [ETL123_table_errors]") ' Intend to edit.
ETL123_table_errors.AddNew
ETL123_table_errors![file_name] = HoldFileName
ETL123_table_errors![file_name_full] = HoldSelItem
ETL123_table_errors![excel_work_sheet_name] = WorkSheetName
ETL123_table_errors![source_file_primary_key_heading] = aExcel_Heading_Text(J)
ETL123_table_errors![error_mapped_field_name] = aField_Name_Output(J)
ETL123_table_errors![source_file_primary_key_value] = Rst!key_fld
ETL123_table_errors![error_field_cell_column] = aExcel_Column_Number(J)
ETL123_table_errors![error_field_cell_row] = Rst!excel_row
ETL123_table_errors![error_field_cell_number] = aExcel_Column_Number(J) & Rst!excel_row
ETL123_table_errors![error_field_heading] = aExcel_Heading_Text(J)
ETL123_table_errors![error_mapped_field_data_type] = aData_Type(J)
ETL123_table_errors![error_field_value] = StrNormalize(source)
ETL123_table_errors![error_message] = "Data VALUE in Excel Row has been rejected due to type mismatch or KEY/LOCK database Violations."
    
Dim HoldReptName   As String:  HoldReptName = ""
ETL123_table_errors![ETL123_rept_name] = ""  ' Move empty string to eliminate Null value.
Call FormatException4Print(ETL123_table_errors![error_field_cell_column], _
                           ETL123_table_errors![error_field_heading], _
                           ETL123_table_errors![error_field_cell_row], _
                           ETL123_table_errors![error_field_value], _
                           ETL123_table_errors![error_mapped_field_name], _
                           ETL123_table_errors![error_mapped_field_data_type], _
                           ETL123_table_errors![error_message], _
                           HoldReptName, _
                           ReptName, NumberReportedErrs)
ETL123_table_errors![ETL123_rept_name] = HoldReptName
    
ETL123_table_errors.Update
ETL123_table_errors.Close
Set ETL123_table_errors = Nothing

strSql = "UPDATE ETL123_table SET ETL123_table.row_has_changed = 'N' WHERE ("
strSql = strSql & "ETL123_table.etl123_row_number = " & Rst!excel_row & ");"

DebugPrint ("Step 'row_has_changed=N' - " & strSql)
DoCmd.RunSQL (strSql)   '  Update matching rows...


End Sub



Private Sub Setup_Row_Numbers_Definitions()
                       
'  This routine will:
'   1)  Locate the Row_Number Column
'   2)  Add field definition to field specifications for Row Number...
'   3)  Delete ???xxxxx??? rows from the ETL123_table
'   4)  Add the row_count and excel_row_num fields. Then populate excel_row_num with the contents of the last column imported.

Dim I As Long, J As Long, II
Dim last_column   As String  ' Field that represents last column before row numbers.
Dim rowNumbersColumnName   As String
Dim holdRowFieldNumber   As Long
Dim strSql    As String

II = 0
Do
  II = II + 1
  If Not fieldexists("F" & II, "ETL123_table") Then
    last_column = "F" & (II - 2)
    holdRowFieldNumber = II - 1
    rowNumbersColumnName = holdRowFieldNumber
    Exit Do
  End If
Loop

DebugPrint ("rowNumbersColumnName=" & rowNumbersColumnName & "  " & "holdRowFieldNumber=" & holdRowFieldNumber & _
            "last_column=" & last_column)

If rowNumbersColumnName = "" Then
  MsgBox ("Error in Setup_Row_Numbers subroutine.  Row Number fields were not found.  Aborted import.")
  DebugPrintOn ("Error in Setup_Row_Numbers subroutine.  Row Number fields were not found.  Aborted import.")
  End
End If

'  Add field definition to field specifications for Row Number...
numOfSpecs = numOfSpecs + 1
J = numOfSpecs

Call InsertNewElementIntoArray(aOutput_Table_Name, "")
If J > 1 Then aOutput_Table_Name(J) = aOutput_Table_Name(J - 1)
Call InsertNewElementIntoArray(aExcel_Column_Number, xlColAlfa(holdRowFieldNumber))
Call InsertNewElementIntoArray(aExcel_Heading_Text, "Row_Number")
Call InsertNewElementIntoArray(aField_Name_Output, "")
Call InsertNewElementIntoArray(aData_Type, "Number/Long Integer")
Call InsertNewElementIntoArray(aData_Type_Enum, 4)

Call InsertNewElementIntoArray(aXcelFieldHasLongText, False)
Call InsertNewElementIntoArray(aDup_Key_Field, False)
Call InsertNewElementIntoArray(aReject_Err_File, False)
Call InsertNewElementIntoArray(aReject_Err_Rows, False)
Call InsertNewElementIntoArray(aAccept_Changes, True)
Call InsertNewElementIntoArray(aAllowChange2Blank, False)
Call InsertNewElementIntoArray(aSave_Date_Changed, False)

' Delete ???xxxxx??? rows from the ETL123_table
strSql = "DELETE ETL123_table.* FROM ETL123_table WHERE (((ETL123_table." & last_column & ")='???xxxxx???'));"
DebugPrint ("Step #3 'Setup_Row_Numbers' - " & strSql)
DoCmd.RunSQL (strSql)

'  Delete Empty Rows from the ETL123_table
Dim XX   As String
strSql = "DELETE ETL123_table.* FROM ETL123_table WHERE ("
For I = 1 To numOfSpecs - 1
  XX = "F" & xlColNum(aExcel_Column_Number(I))
  strSql = strSql & "((ETL123_table." & XX & ") Is Null Or (ETL123_table." & XX & ")='') AND "
Next I
strSql = Left(strSql, Len(strSql) - 5) & ");"
DebugPrint ("Step #3 'Setup_Row_Numbers' - " & strSql)
DoCmd.RunSQL (strSql)

Dim tdf       As TableDef
Dim db As DAO.Database

Dim seqColumnNum     As String
seqColumnNum = "F" & xlColNum(aExcel_Column_Number(numOfSpecs))

' Add the row_count and excel_row_num fields. Then populate excel_row_num with the contents of the last column imported.
Set db = CurrentDb
Set tdf = db.TableDefs("ETL123_table")
On Error Resume Next
tdf.Fields.Append tdf.CreateField("etl123_row_number", 7)     ' dbDouble
tdf.Fields.Append tdf.CreateField("row_count", 7)     ' dbDouble
On Error GoTo 0
Set tdf = Nothing
Set db = Nothing
strSql = "UPDATE " & "ETL123_table" & " SET " & "ETL123_table" & ".etl123_row_number = [" & _
            "ETL123_table" & "].[" & seqColumnNum & "];"
DebugPrint ("Step #4 'Setup_Row_Numbers' - " & strSql)
DoCmd.SetWarnings False
DoCmd.RunSQL (strSql)
DoCmd.SetWarnings True


End Sub

Private Sub Process_A_Worksheet(Worksheet_Name As String, _
                                Worksheet_Name_Full As String, _
                                WkShtNum As Long, _
                                Output_Table As String, _
                                ReptFileName As String)

Dim I As Long, J As Long, k As Long, II, JJ, ZZ, XX
Dim strSql     As String
Dim db As DAO.Database
Set db = CurrentDb
Dim Rst As DAO.Recordset
Dim fName   As String
Dim Response, FieldTypeEnum

Call GetImportSpecifications
If numOfSpecs = 0 Then Exit Sub
If Not Specification_Is_Valid(Output_Table) Then
  MsgBox ("Import for " & Worksheet_Name & "/" & Output_Table & " is terminated due to errors in the import specification.")
  DebugPrintOn ("Import for " & Worksheet_Name & "/" & Output_Table & " is terminated due to errors in the import specification.")
  End
End If

Dim prepedFile   As String
prepedFile = PrepXL(HoldSelItem, WkShtNum)

DebugPrintOn (" ImportTables Process the Worksheet-" & Worksheet_Name)
DelTbl ("ETL123_table")

Dim holdRowNumberColumn       As Long
holdRowNumberColumn = Find_Row_Number_Column(prepedFile, WkShtNum)
If holdRowNumberColumn = 0 Then
  ErrorMsg = "Import for " & Worksheet_Name & "/" & Output_Table & " is terminated due to problem in the Prep_XL routine." & _
          vbCrLf & vbCrLf & "'Row_Number' column was not found to have been added to Excel spreadsheet." & _
          vbCrLf & vbCrLf & "Process will terminate"
  MsgBox (ErrorMsg)
  DebugPrintOn (vbCrLf & "****" & ErrorMsg & vbCrLf)
  End
End If

' Make sure that any "Long Text" fields in the target output table don't have a "Format Property" specified.
' This property has been shown to improperly truncate data longer than 255 characters.
Call ClearLongTextFormatProp(Output_Table)

'Step #1 - Transfer the data from the spreadsheet to imported table
'DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "ETL123_table", _
'        prepedFile, False, Worksheet_Name_Full & "!"
DoCmd.TransferSpreadsheet acImport, 10, "ETL123_table", _
        prepedFile, False, Worksheet_Name_Full & "!"
Kill (prepedFile)


' Set aXcelFieldHasLongText flag so that warning can be issued that text > 255 will be truncated.
For I = LBound(aExcel_Column_Number) To UBound(aExcel_Column_Number)
  J = xlColNum(aExcel_Column_Number(I))
  If J <> 0 And Field_Type("ETL123_table", "F" & J, FieldTypeEnum) = "Long Text/Memo" Then _
      aXcelFieldHasLongText(I) = True
Next I

' Check to see if any fields have Long Text requirements, > 255
For I = LBound(aXcelFieldHasLongText) To UBound(aXcelFieldHasLongText)
  If aXcelFieldHasLongText(I) And aData_Type(I) <> "Long Text/Memo" Then
    II = aData_Type(I)
    ZZ = """" & aExcel_Column_Number(I) & ":" & aExcel_Heading_Text(I) & """"
    XX = """" & Output_Table & "!" & aField_Name_Output(I) & """"
    JJ = "Some data WILL be lost." & vbCrLf & vbCrLf & _
         "The size of some data found in " & ZZ & " requires that " & XX & " be defined as data type ""Long Text""." & vbCrLf & vbCrLf & _
         "However, " & XX & " is defined as """ & II & """.  This can be possibly corrected by changing the field to ""Long Text""." & vbCrLf & vbCrLf & _
         "Do you want to continue WITHOUT correcting the data type?" & vbCrLf & vbCrLf & _
         "Yes - Data will be truncated." & vbCrLf & _
         "No  - Process will be ABORTED. No data will be changed."
    Response = MsgBox(JJ, vbYesNo)
    If Response = 7 Then End   ' vbNo
  End If
Next I


'Delete  holdRowNumberColumn
strSql = "DELETE ETL123_table.F" & holdRowNumberColumn & " FROM ETL123_table " _
       & "WHERE (((ETL123_table.F" & holdRowNumberColumn & ")=""XXXXX""));"
DoCmd.RunSQL (strSql)   '  Delete the XXXXX row...
DebugPrintOn (" End of Step #1 ")


'Step #1a - Delete any extra columns (after Row_Number) from the import_table......
I = holdRowNumberColumn
Do
  I = I + 1
  fName = "F" & I
  If Not fieldexists(fName, "ETL123_table") Then Exit Do
  Call Remove_Table_Field(fName, "ETL123_table")
Loop

'Step #1aa - Delete all NULL rows
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And aData_Type(I) = "Date/Time" Then
    strSql = "UPDATE ETL123_table SET ETL123_table.F" & xlColNum(aExcel_Column_Number(I)) & " = Null " _
           & "WHERE (((ETL123_table.F" & xlColNum(aExcel_Column_Number(I)) & ")=""1/0/1900"")) OR (((ETL123_table.F" & xlColNum(aExcel_Column_Number(I)) & ")=""12/31/1899""));"
    DebugPrint ("Step #1aa - " & strSql)
    DoCmd.RunSQL (strSql)   '  Change these null date values to actual NULL value...
  End If
Next I

'Step #1b - Delete all NULL rows
If holdRowNumberColumn = 1 Then GoTo skipStep1b
strSql = "DELETE ETL123_table.* FROM ETL123_table WHERE ("
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then _
    strSql = strSql & "(((ETL123_table.F" & xlColNum(aExcel_Column_Number(I)) & ") Is Null) OR ((ETL123_table.F" & xlColNum(aExcel_Column_Number(I)) & ")="""")) AND "
Next I
If Right(strSql, 5) = " AND " Then strSql = Left(strSql, Len(strSql) - 5) & ");"
DebugPrint ("Step #1b - " & strSql)
DoCmd.RunSQL (strSql)   '  Delete all NULL rows...
skipStep1b:


'Step #1c - Delete the Excel heading line from the import_table......
strSql = "DELETE ETL123_table.F" & holdRowNumberColumn & " FROM ETL123_table " _
       & "WHERE (((ETL123_table.F" & holdRowNumberColumn & ")=""Row_Number""));"
DebugPrint ("Step #1c - " & strSql)
DoCmd.RunSQL (strSql)   '  Delete the heading line.


'Step #1d - Add an additional Specification Item to the numOfSpecs to capture the excel Row_Number column
Call Setup_Row_Numbers_Definitions

'Step #1e - Eliminate all dups from the ETL123_table (when keys are UNIQUE)
If RT_UniqueKey Then _
  Call Merge_Duplicates("ETL123_table", Output_Table)

'Step #1f - Look for any Boolean fields and translate Y/N to True/False
Call EditBoolFields

'Step #2 - Add SQL fields to ETL123_table as required, if Update Stamp is required, add update stamp fields to target table.
Call AddTableFields(Output_Table)

'Step #3 - Populate "ETL123_table" named data fields and check for excel data errors "null's", generate all error messages
'         Mark "reject_err_row" if entire row should be deleted.
Dim NumberReportedErrs As Variant:  NumberReportedErrs = 0
For I = 1 To numOfSpecs  ' Process all errors for defined fields.
   Call Process_Excel_Data_Error_Check(I, Worksheet_Name, _
                                       ReptFileName, _
                                       NumberReportedErrs)  ' Process all excel data errors for this field....
Next I


'Step #3a - Examine the "ETL123_table" and fill all null fields with 0 for the numeric fields and blanks for string fields.
'           so that JOINs will not have to handle NULL values later.
Dim hDefaultNullValue     As String
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then
    hDefaultNullValue = """ """
    If aData_Type(I) = "Number/Long Integer" Then hDefaultNullValue = "0"
    If aData_Type(I) = "Large Number/Big Integer" Then hDefaultNullValue = "0"
    If aData_Type(I) = "Currency" Then hDefaultNullValue = "0"
    If aData_Type(I) = "Date/Time" Then hDefaultNullValue = "1/1/1900"
    strSql = "UPDATE ETL123_table SET ETL123_table." & aField_Name_Output(I) & " = " & hDefaultNullValue & " WHERE (((ETL123_table." & aField_Name_Output(I) & ") Is Null));"
    DebugPrint ("Step #3a - " & strSql)
    DoCmd.RunSQL (strSql)   '  Now fill in all NULL values.
  End If
Next I


'Step #3b - Examine the "target" table and fill all null KEY fields with 0 for the numeric fields and blanks for string fields.
'           so that JOINs will not have to handle NULL values later.
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And aDup_Key_Field(I) Then
    hDefaultNullValue = """ """
    If aData_Type(I) = "Number/Long Integer" Then hDefaultNullValue = "0"
    If aData_Type(I) = "Currency" Then hDefaultNullValue = "0"
    If aData_Type(I) = "Large Number/Big Integer" Then hDefaultNullValue = "0"
    If aData_Type(I) = "Date/Time" Then hDefaultNullValue = "0"
    If aData_Type(I) <> "AutoNumber" Then
      strSql = "UPDATE " & Output_Table & " SET " & Output_Table & "." & aField_Name_Output(I) & " = " & _
        hDefaultNullValue & " WHERE (((" & Output_Table & "." & aField_Name_Output(I) & ") Is Null));"
      DebugPrint ("Step #3b - " & strSql)
      DoCmd.RunSQL (strSql)   '  Now fill in all NULL values.
    End If
  End If
Next I


'Step #4 - Delete any entire rows that have errors and that have been marked for error rejections.

strSql = "DELETE  ETL123_table.* FROM ETL123_table WHERE (((ETL123_table.reject_err_row)='Y'));"
DebugPrint ("Step #4 - " & strSql)
DoCmd.RunSQL (strSql)   '  Now delete all of the marked errors that qualify for delection from the input file.

' At this point, all appropriately marked rows have been deleted from the "ETL123_table" and
' we are ready now to add new rows and update existing rows in the target table.

'Step #5 - Mark all matching rows with matched_target_table = "Y" using an INNER JOIN to identify new rows that need
'          to be appended.
strSql = "UPDATE ETL123_table SET ETL123_table.matched_target_table = 'N';"
DoCmd.RunSQL (strSql)   '  Mark all rows...
strSql = "UPDATE ETL123_table INNER JOIN " & Output_Table & " ON "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then
    strSql = strSql & _
       "(ETL123_table." & Br(aField_Name_Output(I)) & " = " & Output_Table & "." & Br(aField_Name_Output(I)) & ") AND "
  End If
Next I
If Right(strSql, 5) = " AND " Then strSql = Left(strSql, Len(strSql) - 5)
strSql = strSql & " SET ETL123_table.matched_target_table = 'Y', "
If aMark_active_flag = "Imported" Or aMark_active_flag = "Changed" Then
  strSql = strSql & Output_Table & "." & Br(aActive_flag_name) & " = Yes, "
End If
strSql = strSql & Output_Table & ".etl123_row_number = [ETL123_table].[etl123_row_number];"
DebugPrint ("Step #5 - " & strSql)
DoCmd.RunSQL (strSql)   '  Mark all matching rows...

'Step #5a - When key values are NOT UNIQUE, all matching rows in the target table must first be deleted.
If Not RT_UniqueKey Then _
  Call Delete_Matching_From_Target("ETL123_table", Output_Table)

'Step #6 - Set all matched_target_table = "N" error'ed field values to null to avoid bad data from being added to the table.
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then
    strSql = "UPDATE ETL123_table SET ETL123_table." & Br(aField_Name_Output(I)) & " = Null " & _
             "WHERE (((ETL123_table.[" & aField_Name_Output(I) & "_err]) Is Not Null And (ETL123_table.[" & aField_Name_Output(I) & "_err])<>''));"
    DebugPrint ("Step #6 - '" & Br(aField_Name_Output(I)) & "' " & strSql)
    DoCmd.RunSQL (strSql)   '  Clear data on errored fields...
  End If
Next I

'Step #7 - Append new records with matched_target_table = "N"
strSql = "INSERT INTO " & Output_Table & " (etl123_row_number, "
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then strSql = strSql & Br(aField_Name_Output(I)) & ", "
Next I
If Add_Date_Changed_to_Rows Then _
  strSql = strSql & "update_date_time, update_user, update_program, update_file, create_date_time, create_user, create_program, create_file ) "
If Not Add_Date_Changed_to_Rows Then strSql = Left(strSql, Len(strSql) - 2) & " ) "

strSql = strSql & "SELECT ETL123_table.etl123_row_number, "
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then _
   strSql = strSql & "ETL123_table." & Br(aField_Name_Output(I)) & ", "
Next I
If Add_Date_Changed_to_Rows Then
  strSql = strSql & "ETL123_table.update_date_time, ETL123_table.update_user, ETL123_table.update_program, ETL123_table.update_file, "
  strSql = strSql & "ETL123_table.create_date_time, ETL123_table.create_user, ETL123_table.create_program, ETL123_table.create_file, "
End If
strSql = Left(strSql, Len(strSql) - 2) & " "
strSql = strSql & "FROM ETL123_table WHERE (((ETL123_table.matched_target_table)='N'));"
DebugPrint ("Step #7 - " & strSql)
DoCmd.RunSQL (strSql)   '  Append rows...

'Step #7a - Mark all active flags = Yes using an INNER JOIN to identify new rows added that need to be marked active.
strSql = "UPDATE ETL123_table INNER JOIN " & Output_Table & " ON "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then
    strSql = strSql & _
       "(ETL123_table." & Br(aField_Name_Output(I)) & " = " & Output_Table & "." & Br(aField_Name_Output(I)) & ") AND "
  End If
Next I
If Right(strSql, 5) = " AND " Then strSql = Left(strSql, Len(strSql) - 5)
strSql = strSql & " SET " & Output_Table & "." & Br(aActive_flag_name) & " = Yes "
strSql = strSql & "WHERE (((ETL123_table.matched_target_table)='N'));"
If aMark_active_flag = "New" Or aMark_active_flag = "Imported" Then
  DebugPrint ("Step #7a - " & strSql)
  DoCmd.RunSQL (strSql)   '  Mark Active Flags on all New rows...
End If

'Step 8 - evaluate the aAccept_Changes boolean to change the field to avoid an updated value being accepted in the data.
Dim Hold_Update_Clause     As String
Hold_Update_Clause = "UPDATE ETL123_table INNER JOIN " & Output_Table & " ON " & _
  "(ETL123_table.etl123_row_number = " & Output_Table & ".etl123_row_number) "

'  Now go field by field to fill in ETL123_table with original values from the target table....
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And Not aAccept_Changes(I) And Not aDup_Key_Field(I) Then
    strSql = Hold_Update_Clause & _
      "SET ETL123_table." & Br(aField_Name_Output(I)) & " = [" & Output_Table & "].[" & aField_Name_Output(I) & "] "
    If aData_Type(I) = "Short Text/Text" Then strSql = strSql & _
         "WHERE (((" & Output_Table & "." & Br(aField_Name_Output(I)) & ") Is Not Null And (" & _
           Output_Table & "." & Br(aField_Name_Output(I)) & ")<>''));"
    If aData_Type(I) <> "Short Text/Text" Then strSql = strSql & _
        "WHERE (" & Output_Table & "." & Br(aField_Name_Output(I)) & " Is Not Null);"
              
    DebugPrint ("Step #8 - " & strSql)
    DoCmd.RunSQL (strSql)   '  Pull the target fields and replace source fields...
  End If
Next I

'Step 8a - Find all "errored fields" and update the "ETL123_table" source field to match the target field....
Hold_Update_Clause = "UPDATE ETL123_table INNER JOIN " & Output_Table & " ON " & _
  "(ETL123_table.etl123_row_number = " & Output_Table & ".etl123_row_number) "

'  Now go field by field to fill in ETL123_table with original values from the target table....
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then
    strSql = Hold_Update_Clause & _
      "SET ETL123_table." & Br(aField_Name_Output(I)) & " = [" & Output_Table & "].[" & aField_Name_Output(I) & "] "
    strSql = strSql & "WHERE (((ETL123_table.[" & aField_Name_Output(I) & "_err]) Is Not Null And (ETL123_table.[" & aField_Name_Output(I) & "_err])<>''));"
    DebugPrint ("Step #8a - " & strSql)
    DoCmd.RunSQL (strSql)   '  Pull the target fields and replace source fields...
  End If
Next I

'Step 8b - evaluate the aAllowChange2Blank boolean to avoid existing values changing to blank or zero when the field is empty in the excel spreadsheet.
'          If aAllowChange2Blank is false, then replace field values in the ETL123_table when field is null or blank.
Hold_Update_Clause = "UPDATE ETL123_table INNER JOIN " & Output_Table & " ON " & _
  "(ETL123_table.etl123_row_number = " & Output_Table & ".etl123_row_number) "

'  Now go field by field to fill in ETL123_table with original values from the target table....
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And Not aAllowChange2Blank(I) Then
    strSql = Hold_Update_Clause & _
      "SET ETL123_table." & Br(aField_Name_Output(I)) & " = [" & Output_Table & "].[" & aField_Name_Output(I) & "] "
    
    strSql = strSql & _
        "WHERE (((ETL123_table.F" & xlColNum(aExcel_Column_Number(I)) & ") Is Null Or (ETL123_table.F" & xlColNum(aExcel_Column_Number(I)) & ")=''));"
              
    DebugPrint ("Step #8b - " & strSql)
    DoCmd.RunSQL (strSql)   '  Pull the target fields and replace source fields...
  End If
Next I



'Step 9 - Evaluate aSave_Date_Changed boolean and update the update stamp fields if the field value has changed. row_has_changed
Dim fieldValueCount   As Long
fieldValueCount = 0
strSql = "UPDATE ETL123_table INNER JOIN " & Output_Table & " ON " & _
  "(ETL123_table.etl123_row_number = " & Output_Table & ".etl123_row_number) "

strSql = strSql & "SET ETL123_table.row_has_changed = 'Y' WHERE "
For I = 1 To numOfSpecs
  If aSave_Date_Changed(I) And aField_Name_Output(I) <> "" Then
    strSql = strSql & "([" & Output_Table & "].[" & aField_Name_Output(I) & "]<>[ETL123_table].[" & _
                                                    aField_Name_Output(I) & "]) OR "
    fieldValueCount = fieldValueCount + 1
  End If
Next I
If Right(strSql, 4) = " OR " Then strSql = Left(strSql, Len(strSql) - 4)
strSql = strSql & ";"
If Add_Date_Changed_to_Rows And fieldValueCount > 0 Then
  DebugPrint ("Step #9 - " & strSql)
  DoCmd.RunSQL (strSql)   '  Update CHANGED rows with Change Stamp
End If

'Step 10 - Actually update all source fields from the ETL123_table to the target table with the new field values.
fieldValueCount = 0
strSql = "UPDATE " & Output_Table & " INNER JOIN ETL123_table ON " & _
  "(ETL123_table.etl123_row_number = " & Output_Table & ".etl123_row_number) "

strSql = strSql & "SET "
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And aData_Type(I) <> "AutoNumber" Then
    strSql = strSql & _
      Output_Table & "." & Br(aField_Name_Output(I)) & " = [ETL123_table].[" & aField_Name_Output(I) & "], "
    fieldValueCount = fieldValueCount + 1
  End If
Next I
If Right(strSql, 2) = ", " Then strSql = Left(strSql, Len(strSql) - 2)
strSql = strSql & ";"

If fieldValueCount > 0 Then
  DebugPrint ("Step #10 - " & strSql)
  'db.Execute (strSql)
  DoCmd.SetWarnings False
  DoCmd.RunSQL (strSql)   '  Update all rows...
  DoCmd.SetWarnings True
End If

'Step 11 - Verify that all fields were updated with new values.....  Use a LEFT JOIN to find ALL rows from ETL123_table
'          with all matching rows from the Target table.  Then all fields can be verified to ensure that the import worked.
strSql = "SELECT "

For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then _
    strSql = strSql & "[ETL123_table].[" & aField_Name_Output(I) & "] & ' / ' & "
Next I
If Right(strSql, 12) = "] & ' / ' & " Then strSql = Left(strSql, Len(strSql) - 11) & " AS key_fld, "

strSql = strSql & "[ETL123_table].[F" & xlColNum(aExcel_Column_Number(numOfSpecs)) & "] as excel_row, "

'  These key names will be used later to clear the row_has_changed field if there is a row level issue.
Dim holdKeyCtr   As Long
For I = LBound(holdKeyNames) To UBound(holdKeyNames)
  holdKeyNames(I) = ""  ' Initialize key names array
Next I
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And aDup_Key_Field(I) Then
    holdKeyCtr = holdKeyCtr + 1
    holdKeyNames(I) = "key_" & holdKeyCtr
    strSql = strSql & "ETL123_table." & Br(aField_Name_Output(I)) & " AS key_" & holdKeyCtr & ", "
  End If
Next I
 
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then
    strSql = strSql & Output_Table & "." & Br(aField_Name_Output(I)) & " AS tar_f" & xlColNum(aExcel_Column_Number(I)) & ", "
    strSql = strSql & "ETL123_table." & Br(aField_Name_Output(I)) & " AS src_f" & xlColNum(aExcel_Column_Number(I)) & ", "
    strSql = strSql & "ETL123_table.[" & aField_Name_Output(I) & "_err] AS err_f" & xlColNum(aExcel_Column_Number(I)) & ", "
  End If
Next I
strSql = Left(strSql, Len(strSql) - 2)  ' strip off the last comma.
strSql = strSql & " FROM ETL123_table LEFT JOIN " & Output_Table & " ON " & _
  "(ETL123_table.etl123_row_number = " & Output_Table & ".etl123_row_number) "

DebugPrint ("Step #11 - " & strSql)

Set Rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If Rst.RecordCount = 0 Then
   Rst.Close
   Set Rst = Nothing
   Exit Sub
End If
'  Now loop thru the query results and look for examples of mismatched data....

Dim MaxColNum  As Long:  MaxColNum = 0
For I = 1 To numOfSpecs
  J = xlColNum(aExcel_Column_Number(I))
  If J > MaxColNum Then MaxColNum = J
Next I

Do
  For I = 1 To numOfSpecs
    If aField_Name_Output(I) <> "" Then _
      Call Report_Cells_Not_Updated_Errors(MaxColNum, xlColNum(aExcel_Column_Number(I)), _
                                           Rst, Worksheet_Name, Output_Table, _
                                           ReptFileName, NumberReportedErrs)
  Next I

  Rst.MoveNext
  If Rst.EOF Then GoTo Finished_Do_Loop
Loop
Finished_Do_Loop:

If NumberReportedErrs = 0 Then
  JJ = "No Data Errors for this Worksheet were found..."
  Else
    JJ = "A total of " & NumberReportedErrs & " Data Errors were found in this Worksheet..."
End If
Call FileSave(vbCrLf & JJ, ReptFileName, , True)


'Step 12 - Evaluate imported_data row_has_changed = 'Y' and update the update stamp fields if the field value has changed.
strSql = "UPDATE ETL123_table INNER JOIN " & Output_Table & " ON " & _
  "(ETL123_table.etl123_row_number = " & Output_Table & ".etl123_row_number) "

strSql = strSql & "SET " & _
                   Output_Table & ".update_date_time = [ETL123_table].[update_date_time], " & _
                   Output_Table & ".update_user = [ETL123_table].[update_user], " & _
                   Output_Table & ".update_program = [ETL123_table].[update_program], " & _
                   Output_Table & ".update_file = [ETL123_table].[update_file] " & _
            "WHERE ETL123_table.row_has_changed = 'Y';"
If Add_Date_Changed_to_Rows Then
  DebugPrint ("Step #12 - " & strSql)
  DoCmd.RunSQL (strSql)   '  Update CHANGED rows with Change Stamp
End If

Exit Sub

End Sub

Private Sub Report_Cells_Not_Updated_Errors(ByVal MaxColNum As Long, _
                                            ByVal fldPtr As Long, _
                                            Rst As Recordset, _
                                            ByVal Worksheet_Name As String, _
                                            ByVal Output_Table As String, _
                                            ByVal ReptName As String, _
                                            ByRef NumErrs As Variant)
                                            
                       
Dim RR As String:  RR = ReptName
Dim I As Long
I = fldPtr    '
' This routine will double check each field that may be different and call an update routine.

If I = 1 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f1, ""), Nz(Rst!src_f1, ""), Nz(Rst!err_f1, ""), Worksheet_Name, RR, NumErrs)
If I = 2 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f2, ""), Nz(Rst!src_f2, ""), Nz(Rst!err_f2, ""), Worksheet_Name, RR, NumErrs)
If I = 3 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f3, ""), Nz(Rst!src_f3, ""), Nz(Rst!err_f3, ""), Worksheet_Name, RR, NumErrs)
If I = 4 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f4, ""), Nz(Rst!src_f4, ""), Nz(Rst!err_f4, ""), Worksheet_Name, RR, NumErrs)
If I = 5 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f5, ""), Nz(Rst!src_f5, ""), Nz(Rst!err_f5, ""), Worksheet_Name, RR, NumErrs)
If I = 6 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f6, ""), Nz(Rst!src_f6, ""), Nz(Rst!err_f6, ""), Worksheet_Name, RR, NumErrs)
If I = 7 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f7, ""), Nz(Rst!src_f7, ""), Nz(Rst!err_f7, ""), Worksheet_Name, RR, NumErrs)
If I = 8 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f8, ""), Nz(Rst!src_f8, ""), Nz(Rst!err_f8, ""), Worksheet_Name, RR, NumErrs)
If I = 9 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f9, ""), Nz(Rst!src_f9, ""), Nz(Rst!err_f9, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 9 Then Exit Sub

If I = 10 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f10, ""), Nz(Rst!src_f10, ""), Nz(Rst!err_f10, ""), Worksheet_Name, RR, NumErrs)
If I = 11 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f11, ""), Nz(Rst!src_f11, ""), Nz(Rst!err_f11, ""), Worksheet_Name, RR, NumErrs)
If I = 12 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f12, ""), Nz(Rst!src_f12, ""), Nz(Rst!err_f12, ""), Worksheet_Name, RR, NumErrs)
If I = 13 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f13, ""), Nz(Rst!src_f13, ""), Nz(Rst!err_f13, ""), Worksheet_Name, RR, NumErrs)
If I = 14 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f14, ""), Nz(Rst!src_f14, ""), Nz(Rst!err_f14, ""), Worksheet_Name, RR, NumErrs)
If I = 15 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f15, ""), Nz(Rst!src_f15, ""), Nz(Rst!err_f15, ""), Worksheet_Name, RR, NumErrs)
If I = 16 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f16, ""), Nz(Rst!src_f16, ""), Nz(Rst!err_f16, ""), Worksheet_Name, RR, NumErrs)
If I = 17 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f17, ""), Nz(Rst!src_f17, ""), Nz(Rst!err_f17, ""), Worksheet_Name, RR, NumErrs)
If I = 18 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f18, ""), Nz(Rst!src_f18, ""), Nz(Rst!err_f18, ""), Worksheet_Name, RR, NumErrs)
If I = 19 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f19, ""), Nz(Rst!src_f19, ""), Nz(Rst!err_f19, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 19 Then Exit Sub

If I = 20 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f20, ""), Nz(Rst!src_f20, ""), Nz(Rst!err_f20, ""), Worksheet_Name, RR, NumErrs)
If I = 21 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f21, ""), Nz(Rst!src_f21, ""), Nz(Rst!err_f21, ""), Worksheet_Name, RR, NumErrs)
If I = 22 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f22, ""), Nz(Rst!src_f22, ""), Nz(Rst!err_f22, ""), Worksheet_Name, RR, NumErrs)
If I = 23 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f23, ""), Nz(Rst!src_f23, ""), Nz(Rst!err_f23, ""), Worksheet_Name, RR, NumErrs)
If I = 24 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f24, ""), Nz(Rst!src_f24, ""), Nz(Rst!err_f24, ""), Worksheet_Name, RR, NumErrs)
If I = 25 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f25, ""), Nz(Rst!src_f25, ""), Nz(Rst!err_f25, ""), Worksheet_Name, RR, NumErrs)
If I = 26 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f26, ""), Nz(Rst!src_f26, ""), Nz(Rst!err_f26, ""), Worksheet_Name, RR, NumErrs)
If I = 27 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f27, ""), Nz(Rst!src_f27, ""), Nz(Rst!err_f27, ""), Worksheet_Name, RR, NumErrs)
If I = 28 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f28, ""), Nz(Rst!src_f28, ""), Nz(Rst!err_f28, ""), Worksheet_Name, RR, NumErrs)
If I = 29 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f29, ""), Nz(Rst!src_f29, ""), Nz(Rst!err_f29, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 29 Then Exit Sub

If I = 30 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f30, ""), Nz(Rst!src_f30, ""), Nz(Rst!err_f30, ""), Worksheet_Name, RR, NumErrs)
If I = 31 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f31, ""), Nz(Rst!src_f31, ""), Nz(Rst!err_f31, ""), Worksheet_Name, RR, NumErrs)
If I = 32 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f32, ""), Nz(Rst!src_f32, ""), Nz(Rst!err_f32, ""), Worksheet_Name, RR, NumErrs)
If I = 33 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f33, ""), Nz(Rst!src_f33, ""), Nz(Rst!err_f33, ""), Worksheet_Name, RR, NumErrs)
If I = 34 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f34, ""), Nz(Rst!src_f34, ""), Nz(Rst!err_f34, ""), Worksheet_Name, RR, NumErrs)
If I = 35 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f35, ""), Nz(Rst!src_f35, ""), Nz(Rst!err_f35, ""), Worksheet_Name, RR, NumErrs)
If I = 36 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f36, ""), Nz(Rst!src_f36, ""), Nz(Rst!err_f36, ""), Worksheet_Name, RR, NumErrs)
If I = 37 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f37, ""), Nz(Rst!src_f37, ""), Nz(Rst!err_f37, ""), Worksheet_Name, RR, NumErrs)
If I = 38 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f38, ""), Nz(Rst!src_f38, ""), Nz(Rst!err_f38, ""), Worksheet_Name, RR, NumErrs)
If I = 39 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f39, ""), Nz(Rst!src_f39, ""), Nz(Rst!err_f39, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 39 Then Exit Sub

If I = 40 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f40, ""), Nz(Rst!src_f40, ""), Nz(Rst!err_f40, ""), Worksheet_Name, RR, NumErrs)
If I = 41 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f41, ""), Nz(Rst!src_f41, ""), Nz(Rst!err_f41, ""), Worksheet_Name, RR, NumErrs)
If I = 42 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f42, ""), Nz(Rst!src_f42, ""), Nz(Rst!err_f42, ""), Worksheet_Name, RR, NumErrs)
If I = 43 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f43, ""), Nz(Rst!src_f43, ""), Nz(Rst!err_f43, ""), Worksheet_Name, RR, NumErrs)
If I = 44 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f44, ""), Nz(Rst!src_f44, ""), Nz(Rst!err_f44, ""), Worksheet_Name, RR, NumErrs)
If I = 45 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f45, ""), Nz(Rst!src_f45, ""), Nz(Rst!err_f45, ""), Worksheet_Name, RR, NumErrs)
If I = 46 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f46, ""), Nz(Rst!src_f46, ""), Nz(Rst!err_f46, ""), Worksheet_Name, RR, NumErrs)
If I = 47 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f47, ""), Nz(Rst!src_f47, ""), Nz(Rst!err_f47, ""), Worksheet_Name, RR, NumErrs)
If I = 48 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f48, ""), Nz(Rst!src_f48, ""), Nz(Rst!err_f48, ""), Worksheet_Name, RR, NumErrs)
If I = 49 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f49, ""), Nz(Rst!src_f49, ""), Nz(Rst!err_f49, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 49 Then Exit Sub

If I = 50 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f50, ""), Nz(Rst!src_f50, ""), Nz(Rst!err_f50, ""), Worksheet_Name, RR, NumErrs)
If I = 51 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f51, ""), Nz(Rst!src_f51, ""), Nz(Rst!err_f51, ""), Worksheet_Name, RR, NumErrs)
If I = 52 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f52, ""), Nz(Rst!src_f52, ""), Nz(Rst!err_f52, ""), Worksheet_Name, RR, NumErrs)
If I = 53 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f53, ""), Nz(Rst!src_f53, ""), Nz(Rst!err_f53, ""), Worksheet_Name, RR, NumErrs)
If I = 54 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f54, ""), Nz(Rst!src_f54, ""), Nz(Rst!err_f54, ""), Worksheet_Name, RR, NumErrs)
If I = 55 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f55, ""), Nz(Rst!src_f55, ""), Nz(Rst!err_f55, ""), Worksheet_Name, RR, NumErrs)
If I = 56 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f56, ""), Nz(Rst!src_f56, ""), Nz(Rst!err_f56, ""), Worksheet_Name, RR, NumErrs)
If I = 57 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f57, ""), Nz(Rst!src_f57, ""), Nz(Rst!err_f57, ""), Worksheet_Name, RR, NumErrs)
If I = 58 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f58, ""), Nz(Rst!src_f58, ""), Nz(Rst!err_f58, ""), Worksheet_Name, RR, NumErrs)
If I = 59 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f59, ""), Nz(Rst!src_f59, ""), Nz(Rst!err_f59, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 59 Then Exit Sub

If I = 60 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f60, ""), Nz(Rst!src_f60, ""), Nz(Rst!err_f60, ""), Worksheet_Name, RR, NumErrs)
If I = 61 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f61, ""), Nz(Rst!src_f61, ""), Nz(Rst!err_f61, ""), Worksheet_Name, RR, NumErrs)
If I = 62 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f62, ""), Nz(Rst!src_f62, ""), Nz(Rst!err_f62, ""), Worksheet_Name, RR, NumErrs)
If I = 63 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f63, ""), Nz(Rst!src_f63, ""), Nz(Rst!err_f63, ""), Worksheet_Name, RR, NumErrs)
If I = 64 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f64, ""), Nz(Rst!src_f64, ""), Nz(Rst!err_f64, ""), Worksheet_Name, RR, NumErrs)
If I = 65 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f65, ""), Nz(Rst!src_f65, ""), Nz(Rst!err_f65, ""), Worksheet_Name, RR, NumErrs)
If I = 66 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f66, ""), Nz(Rst!src_f66, ""), Nz(Rst!err_f66, ""), Worksheet_Name, RR, NumErrs)
If I = 67 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f67, ""), Nz(Rst!src_f67, ""), Nz(Rst!err_f67, ""), Worksheet_Name, RR, NumErrs)
If I = 68 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f68, ""), Nz(Rst!src_f68, ""), Nz(Rst!err_f68, ""), Worksheet_Name, RR, NumErrs)
If I = 69 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f69, ""), Nz(Rst!src_f69, ""), Nz(Rst!err_f69, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 69 Then Exit Sub

If I = 70 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f70, ""), Nz(Rst!src_f70, ""), Nz(Rst!err_f70, ""), Worksheet_Name, RR, NumErrs)
If I = 71 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f71, ""), Nz(Rst!src_f71, ""), Nz(Rst!err_f71, ""), Worksheet_Name, RR, NumErrs)
If I = 72 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f72, ""), Nz(Rst!src_f72, ""), Nz(Rst!err_f72, ""), Worksheet_Name, RR, NumErrs)
If I = 73 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f73, ""), Nz(Rst!src_f73, ""), Nz(Rst!err_f73, ""), Worksheet_Name, RR, NumErrs)
If I = 74 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f74, ""), Nz(Rst!src_f74, ""), Nz(Rst!err_f74, ""), Worksheet_Name, RR, NumErrs)
If I = 75 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f75, ""), Nz(Rst!src_f75, ""), Nz(Rst!err_f75, ""), Worksheet_Name, RR, NumErrs)
If I = 76 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f76, ""), Nz(Rst!src_f76, ""), Nz(Rst!err_f76, ""), Worksheet_Name, RR, NumErrs)
If I = 77 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f77, ""), Nz(Rst!src_f77, ""), Nz(Rst!err_f77, ""), Worksheet_Name, RR, NumErrs)
If I = 78 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f78, ""), Nz(Rst!src_f78, ""), Nz(Rst!err_f78, ""), Worksheet_Name, RR, NumErrs)
If I = 79 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f79, ""), Nz(Rst!src_f79, ""), Nz(Rst!err_f79, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 79 Then Exit Sub

If I = 80 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f80, ""), Nz(Rst!src_f80, ""), Nz(Rst!err_f80, ""), Worksheet_Name, RR, NumErrs)
If I = 81 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f81, ""), Nz(Rst!src_f81, ""), Nz(Rst!err_f81, ""), Worksheet_Name, RR, NumErrs)
If I = 82 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f82, ""), Nz(Rst!src_f82, ""), Nz(Rst!err_f82, ""), Worksheet_Name, RR, NumErrs)
If I = 83 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f83, ""), Nz(Rst!src_f83, ""), Nz(Rst!err_f83, ""), Worksheet_Name, RR, NumErrs)
If I = 84 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f84, ""), Nz(Rst!src_f84, ""), Nz(Rst!err_f84, ""), Worksheet_Name, RR, NumErrs)
If I = 85 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f85, ""), Nz(Rst!src_f85, ""), Nz(Rst!err_f85, ""), Worksheet_Name, RR, NumErrs)
If I = 86 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f86, ""), Nz(Rst!src_f86, ""), Nz(Rst!err_f86, ""), Worksheet_Name, RR, NumErrs)
If I = 87 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f87, ""), Nz(Rst!src_f87, ""), Nz(Rst!err_f87, ""), Worksheet_Name, RR, NumErrs)
If I = 88 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f88, ""), Nz(Rst!src_f88, ""), Nz(Rst!err_f88, ""), Worksheet_Name, RR, NumErrs)
If I = 89 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f89, ""), Nz(Rst!src_f89, ""), Nz(Rst!err_f89, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 89 Then Exit Sub

If I = 90 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f90, ""), Nz(Rst!src_f90, ""), Nz(Rst!err_f90, ""), Worksheet_Name, RR, NumErrs)
If I = 91 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f91, ""), Nz(Rst!src_f91, ""), Nz(Rst!err_f91, ""), Worksheet_Name, RR, NumErrs)
If I = 92 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f92, ""), Nz(Rst!src_f92, ""), Nz(Rst!err_f92, ""), Worksheet_Name, RR, NumErrs)
If I = 93 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f93, ""), Nz(Rst!src_f93, ""), Nz(Rst!err_f93, ""), Worksheet_Name, RR, NumErrs)
If I = 94 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f94, ""), Nz(Rst!src_f94, ""), Nz(Rst!err_f94, ""), Worksheet_Name, RR, NumErrs)
If I = 95 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f95, ""), Nz(Rst!src_f95, ""), Nz(Rst!err_f95, ""), Worksheet_Name, RR, NumErrs)
If I = 96 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f96, ""), Nz(Rst!src_f96, ""), Nz(Rst!err_f96, ""), Worksheet_Name, RR, NumErrs)
If I = 97 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f97, ""), Nz(Rst!src_f97, ""), Nz(Rst!err_f97, ""), Worksheet_Name, RR, NumErrs)
If I = 98 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f98, ""), Nz(Rst!src_f98, ""), Nz(Rst!err_f98, ""), Worksheet_Name, RR, NumErrs)
If I = 99 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f99, ""), Nz(Rst!src_f99, ""), Nz(Rst!err_f99, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 99 Then Exit Sub

If I = 100 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f100, ""), Nz(Rst!src_f100, ""), Nz(Rst!err_f100, ""), Worksheet_Name, RR, NumErrs)
If I = 101 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f101, ""), Nz(Rst!src_f101, ""), Nz(Rst!err_f101, ""), Worksheet_Name, RR, NumErrs)
If I = 102 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f102, ""), Nz(Rst!src_f102, ""), Nz(Rst!err_f102, ""), Worksheet_Name, RR, NumErrs)
If I = 103 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f103, ""), Nz(Rst!src_f103, ""), Nz(Rst!err_f103, ""), Worksheet_Name, RR, NumErrs)
If I = 104 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f104, ""), Nz(Rst!src_f104, ""), Nz(Rst!err_f104, ""), Worksheet_Name, RR, NumErrs)
If I = 105 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f105, ""), Nz(Rst!src_f105, ""), Nz(Rst!err_f105, ""), Worksheet_Name, RR, NumErrs)
If I = 106 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f106, ""), Nz(Rst!src_f106, ""), Nz(Rst!err_f106, ""), Worksheet_Name, RR, NumErrs)
If I = 107 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f107, ""), Nz(Rst!src_f107, ""), Nz(Rst!err_f107, ""), Worksheet_Name, RR, NumErrs)
If I = 108 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f108, ""), Nz(Rst!src_f108, ""), Nz(Rst!err_f108, ""), Worksheet_Name, RR, NumErrs)
If I = 109 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f109, ""), Nz(Rst!src_f109, ""), Nz(Rst!err_f109, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 109 Then Exit Sub

If I = 110 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f110, ""), Nz(Rst!src_f110, ""), Nz(Rst!err_f110, ""), Worksheet_Name, RR, NumErrs)
If I = 111 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f111, ""), Nz(Rst!src_f111, ""), Nz(Rst!err_f111, ""), Worksheet_Name, RR, NumErrs)
If I = 112 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f112, ""), Nz(Rst!src_f112, ""), Nz(Rst!err_f112, ""), Worksheet_Name, RR, NumErrs)
If I = 113 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f113, ""), Nz(Rst!src_f113, ""), Nz(Rst!err_f113, ""), Worksheet_Name, RR, NumErrs)
If I = 114 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f114, ""), Nz(Rst!src_f114, ""), Nz(Rst!err_f114, ""), Worksheet_Name, RR, NumErrs)
If I = 115 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f115, ""), Nz(Rst!src_f115, ""), Nz(Rst!err_f115, ""), Worksheet_Name, RR, NumErrs)
If I = 116 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f116, ""), Nz(Rst!src_f116, ""), Nz(Rst!err_f116, ""), Worksheet_Name, RR, NumErrs)
If I = 117 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f117, ""), Nz(Rst!src_f117, ""), Nz(Rst!err_f117, ""), Worksheet_Name, RR, NumErrs)
If I = 118 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f118, ""), Nz(Rst!src_f118, ""), Nz(Rst!err_f118, ""), Worksheet_Name, RR, NumErrs)
If I = 119 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f119, ""), Nz(Rst!src_f119, ""), Nz(Rst!err_f119, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 119 Then Exit Sub

If I = 120 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f120, ""), Nz(Rst!src_f120, ""), Nz(Rst!err_f120, ""), Worksheet_Name, RR, NumErrs)
If I = 121 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f121, ""), Nz(Rst!src_f121, ""), Nz(Rst!err_f121, ""), Worksheet_Name, RR, NumErrs)
If I = 122 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f122, ""), Nz(Rst!src_f122, ""), Nz(Rst!err_f122, ""), Worksheet_Name, RR, NumErrs)
If I = 123 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f123, ""), Nz(Rst!src_f123, ""), Nz(Rst!err_f123, ""), Worksheet_Name, RR, NumErrs)
If I = 124 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f124, ""), Nz(Rst!src_f124, ""), Nz(Rst!err_f124, ""), Worksheet_Name, RR, NumErrs)
If I = 125 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f125, ""), Nz(Rst!src_f125, ""), Nz(Rst!err_f125, ""), Worksheet_Name, RR, NumErrs)
If I = 126 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f126, ""), Nz(Rst!src_f126, ""), Nz(Rst!err_f126, ""), Worksheet_Name, RR, NumErrs)
If I = 127 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f127, ""), Nz(Rst!src_f127, ""), Nz(Rst!err_f127, ""), Worksheet_Name, RR, NumErrs)
If I = 128 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f128, ""), Nz(Rst!src_f128, ""), Nz(Rst!err_f128, ""), Worksheet_Name, RR, NumErrs)
If I = 129 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f129, ""), Nz(Rst!src_f129, ""), Nz(Rst!err_f129, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 129 Then Exit Sub

If I = 130 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f130, ""), Nz(Rst!src_f130, ""), Nz(Rst!err_f130, ""), Worksheet_Name, RR, NumErrs)
If I = 131 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f131, ""), Nz(Rst!src_f131, ""), Nz(Rst!err_f131, ""), Worksheet_Name, RR, NumErrs)
If I = 132 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f132, ""), Nz(Rst!src_f132, ""), Nz(Rst!err_f132, ""), Worksheet_Name, RR, NumErrs)
If I = 133 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f133, ""), Nz(Rst!src_f133, ""), Nz(Rst!err_f133, ""), Worksheet_Name, RR, NumErrs)
If I = 134 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f134, ""), Nz(Rst!src_f134, ""), Nz(Rst!err_f134, ""), Worksheet_Name, RR, NumErrs)
If I = 135 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f135, ""), Nz(Rst!src_f135, ""), Nz(Rst!err_f135, ""), Worksheet_Name, RR, NumErrs)
If I = 136 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f136, ""), Nz(Rst!src_f136, ""), Nz(Rst!err_f136, ""), Worksheet_Name, RR, NumErrs)
If I = 137 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f137, ""), Nz(Rst!src_f137, ""), Nz(Rst!err_f137, ""), Worksheet_Name, RR, NumErrs)
If I = 138 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f138, ""), Nz(Rst!src_f138, ""), Nz(Rst!err_f138, ""), Worksheet_Name, RR, NumErrs)
If I = 139 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f139, ""), Nz(Rst!src_f139, ""), Nz(Rst!err_f139, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 139 Then Exit Sub

If I = 140 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f140, ""), Nz(Rst!src_f140, ""), Nz(Rst!err_f140, ""), Worksheet_Name, RR, NumErrs)
If I = 141 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f141, ""), Nz(Rst!src_f141, ""), Nz(Rst!err_f141, ""), Worksheet_Name, RR, NumErrs)
If I = 142 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f142, ""), Nz(Rst!src_f142, ""), Nz(Rst!err_f142, ""), Worksheet_Name, RR, NumErrs)
If I = 143 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f143, ""), Nz(Rst!src_f143, ""), Nz(Rst!err_f143, ""), Worksheet_Name, RR, NumErrs)
If I = 144 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f144, ""), Nz(Rst!src_f144, ""), Nz(Rst!err_f144, ""), Worksheet_Name, RR, NumErrs)
If I = 145 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f145, ""), Nz(Rst!src_f145, ""), Nz(Rst!err_f145, ""), Worksheet_Name, RR, NumErrs)
If I = 146 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f146, ""), Nz(Rst!src_f146, ""), Nz(Rst!err_f146, ""), Worksheet_Name, RR, NumErrs)
If I = 147 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f147, ""), Nz(Rst!src_f147, ""), Nz(Rst!err_f147, ""), Worksheet_Name, RR, NumErrs)
If I = 148 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f148, ""), Nz(Rst!src_f148, ""), Nz(Rst!err_f148, ""), Worksheet_Name, RR, NumErrs)
If I = 149 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f149, ""), Nz(Rst!src_f149, ""), Nz(Rst!err_f149, ""), Worksheet_Name, RR, NumErrs)
If I = 150 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f150, ""), Nz(Rst!src_f150, ""), Nz(Rst!err_f150, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 150 Then Exit Sub

If I = 151 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f151, ""), Nz(Rst!src_f151, ""), Nz(Rst!err_f151, ""), Worksheet_Name, RR, NumErrs)
If I = 152 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f152, ""), Nz(Rst!src_f152, ""), Nz(Rst!err_f152, ""), Worksheet_Name, RR, NumErrs)
If I = 153 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f153, ""), Nz(Rst!src_f153, ""), Nz(Rst!err_f153, ""), Worksheet_Name, RR, NumErrs)
If I = 154 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f154, ""), Nz(Rst!src_f154, ""), Nz(Rst!err_f154, ""), Worksheet_Name, RR, NumErrs)
If I = 155 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f155, ""), Nz(Rst!src_f155, ""), Nz(Rst!err_f155, ""), Worksheet_Name, RR, NumErrs)
If I = 156 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f156, ""), Nz(Rst!src_f156, ""), Nz(Rst!err_f156, ""), Worksheet_Name, RR, NumErrs)
If I = 157 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f157, ""), Nz(Rst!src_f157, ""), Nz(Rst!err_f157, ""), Worksheet_Name, RR, NumErrs)
If I = 158 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f158, ""), Nz(Rst!src_f158, ""), Nz(Rst!err_f158, ""), Worksheet_Name, RR, NumErrs)
If I = 159 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f159, ""), Nz(Rst!src_f159, ""), Nz(Rst!err_f159, ""), Worksheet_Name, RR, NumErrs)
If I = 160 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f160, ""), Nz(Rst!src_f160, ""), Nz(Rst!err_f160, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 160 Then Exit Sub

If I = 161 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f161, ""), Nz(Rst!src_f161, ""), Nz(Rst!err_f161, ""), Worksheet_Name, RR, NumErrs)
If I = 162 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f162, ""), Nz(Rst!src_f162, ""), Nz(Rst!err_f162, ""), Worksheet_Name, RR, NumErrs)
If I = 163 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f163, ""), Nz(Rst!src_f163, ""), Nz(Rst!err_f163, ""), Worksheet_Name, RR, NumErrs)
If I = 164 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f164, ""), Nz(Rst!src_f164, ""), Nz(Rst!err_f164, ""), Worksheet_Name, RR, NumErrs)
If I = 165 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f165, ""), Nz(Rst!src_f165, ""), Nz(Rst!err_f165, ""), Worksheet_Name, RR, NumErrs)
If I = 166 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f166, ""), Nz(Rst!src_f166, ""), Nz(Rst!err_f166, ""), Worksheet_Name, RR, NumErrs)
If I = 167 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f167, ""), Nz(Rst!src_f167, ""), Nz(Rst!err_f167, ""), Worksheet_Name, RR, NumErrs)
If I = 168 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f168, ""), Nz(Rst!src_f168, ""), Nz(Rst!err_f168, ""), Worksheet_Name, RR, NumErrs)
If I = 169 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f169, ""), Nz(Rst!src_f169, ""), Nz(Rst!err_f169, ""), Worksheet_Name, RR, NumErrs)
If I = 170 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f170, ""), Nz(Rst!src_f170, ""), Nz(Rst!err_f170, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 170 Then Exit Sub

If I = 171 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f171, ""), Nz(Rst!src_f171, ""), Nz(Rst!err_f171, ""), Worksheet_Name, RR, NumErrs)
If I = 172 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f172, ""), Nz(Rst!src_f172, ""), Nz(Rst!err_f172, ""), Worksheet_Name, RR, NumErrs)
If I = 173 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f173, ""), Nz(Rst!src_f173, ""), Nz(Rst!err_f173, ""), Worksheet_Name, RR, NumErrs)
If I = 174 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f174, ""), Nz(Rst!src_f174, ""), Nz(Rst!err_f174, ""), Worksheet_Name, RR, NumErrs)
If I = 175 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f175, ""), Nz(Rst!src_f175, ""), Nz(Rst!err_f175, ""), Worksheet_Name, RR, NumErrs)
If I = 176 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f176, ""), Nz(Rst!src_f176, ""), Nz(Rst!err_f176, ""), Worksheet_Name, RR, NumErrs)
If I = 177 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f177, ""), Nz(Rst!src_f177, ""), Nz(Rst!err_f177, ""), Worksheet_Name, RR, NumErrs)
If I = 178 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f178, ""), Nz(Rst!src_f178, ""), Nz(Rst!err_f178, ""), Worksheet_Name, RR, NumErrs)
If I = 179 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f179, ""), Nz(Rst!src_f179, ""), Nz(Rst!err_f179, ""), Worksheet_Name, RR, NumErrs)
If I = 180 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f180, ""), Nz(Rst!src_f180, ""), Nz(Rst!err_f180, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 180 Then Exit Sub

If I = 181 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f181, ""), Nz(Rst!src_f181, ""), Nz(Rst!err_f181, ""), Worksheet_Name, RR, NumErrs)
If I = 182 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f182, ""), Nz(Rst!src_f182, ""), Nz(Rst!err_f182, ""), Worksheet_Name, RR, NumErrs)
If I = 183 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f183, ""), Nz(Rst!src_f183, ""), Nz(Rst!err_f183, ""), Worksheet_Name, RR, NumErrs)
If I = 184 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f184, ""), Nz(Rst!src_f184, ""), Nz(Rst!err_f184, ""), Worksheet_Name, RR, NumErrs)
If I = 185 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f185, ""), Nz(Rst!src_f185, ""), Nz(Rst!err_f185, ""), Worksheet_Name, RR, NumErrs)
If I = 186 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f186, ""), Nz(Rst!src_f186, ""), Nz(Rst!err_f186, ""), Worksheet_Name, RR, NumErrs)
If I = 187 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f187, ""), Nz(Rst!src_f187, ""), Nz(Rst!err_f187, ""), Worksheet_Name, RR, NumErrs)
If I = 188 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f188, ""), Nz(Rst!src_f188, ""), Nz(Rst!err_f188, ""), Worksheet_Name, RR, NumErrs)
If I = 189 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f189, ""), Nz(Rst!src_f189, ""), Nz(Rst!err_f189, ""), Worksheet_Name, RR, NumErrs)
If I = 190 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f190, ""), Nz(Rst!src_f190, ""), Nz(Rst!err_f190, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 190 Then Exit Sub

If I = 191 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f191, ""), Nz(Rst!src_f191, ""), Nz(Rst!err_f191, ""), Worksheet_Name, RR, NumErrs)
If I = 192 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f192, ""), Nz(Rst!src_f192, ""), Nz(Rst!err_f192, ""), Worksheet_Name, RR, NumErrs)
If I = 193 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f193, ""), Nz(Rst!src_f193, ""), Nz(Rst!err_f193, ""), Worksheet_Name, RR, NumErrs)
If I = 194 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f194, ""), Nz(Rst!src_f194, ""), Nz(Rst!err_f194, ""), Worksheet_Name, RR, NumErrs)
If I = 195 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f195, ""), Nz(Rst!src_f195, ""), Nz(Rst!err_f195, ""), Worksheet_Name, RR, NumErrs)
If I = 196 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f196, ""), Nz(Rst!src_f196, ""), Nz(Rst!err_f196, ""), Worksheet_Name, RR, NumErrs)
If I = 197 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f197, ""), Nz(Rst!src_f197, ""), Nz(Rst!err_f197, ""), Worksheet_Name, RR, NumErrs)
If I = 198 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f198, ""), Nz(Rst!src_f198, ""), Nz(Rst!err_f198, ""), Worksheet_Name, RR, NumErrs)
If I = 199 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f199, ""), Nz(Rst!src_f199, ""), Nz(Rst!err_f199, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 199 Then Exit Sub
  
Call Process_Field_Update_Data_Error_Check300(MaxColNum, fldPtr, Rst, Worksheet_Name, Output_Table, RR, NumErrs)
If MaxColNum <= 300 Then Exit Sub

Call Process_Field_Update_Data_Error_Check400(MaxColNum, fldPtr, Rst, Worksheet_Name, Output_Table, RR, NumErrs)
If MaxColNum <= 400 Then Exit Sub

Call Process_Field_Update_Data_Error_Check500(MaxColNum, fldPtr, Rst, Worksheet_Name, Output_Table, RR, NumErrs)
If MaxColNum <= 500 Then Exit Sub

Call Process_Field_Update_Data_Error_Check600(MaxColNum, fldPtr, Rst, Worksheet_Name, Output_Table, RR, NumErrs)
If MaxColNum <= 600 Then Exit Sub

Call Process_Field_Update_Data_Error_Check700(MaxColNum, fldPtr, Rst, Worksheet_Name, Output_Table, RR, NumErrs)
If MaxColNum <= 700 Then Exit Sub

Call MsgBox("Report_Cells_Not_Updated_Errors - has been overrun.  Excel file has too Many Columns to process." & vbCrLf & vbCrLf & _
            "Import Process will ABORT.")
End

End Sub

Private Sub Process_Field_Update_Data_Error_Check300(MaxColNum As Long, fldPtr As Long, _
                Rst As Recordset, _
                Worksheet_Name As String, _
                ByVal Output_Table As String, _
                ByVal ReptName As String, _
                ByRef NumErrs As Variant)
                       
Dim RR As String:  RR = ReptName
Dim I As Long
I = fldPtr

If MaxColNum > 301 Then Exit Sub

If I = 200 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f200, ""), Nz(Rst!src_f200, ""), Nz(Rst!err_f200, ""), Worksheet_Name, RR, NumErrs)
If I = 201 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f201, ""), Nz(Rst!src_f201, ""), Nz(Rst!err_f201, ""), Worksheet_Name, RR, NumErrs)
If I = 202 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f202, ""), Nz(Rst!src_f202, ""), Nz(Rst!err_f202, ""), Worksheet_Name, RR, NumErrs)
If I = 203 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f203, ""), Nz(Rst!src_f203, ""), Nz(Rst!err_f203, ""), Worksheet_Name, RR, NumErrs)
If I = 204 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f204, ""), Nz(Rst!src_f204, ""), Nz(Rst!err_f204, ""), Worksheet_Name, RR, NumErrs)
If I = 205 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f205, ""), Nz(Rst!src_f205, ""), Nz(Rst!err_f205, ""), Worksheet_Name, RR, NumErrs)
If I = 206 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f206, ""), Nz(Rst!src_f206, ""), Nz(Rst!err_f206, ""), Worksheet_Name, RR, NumErrs)
If I = 207 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f207, ""), Nz(Rst!src_f207, ""), Nz(Rst!err_f207, ""), Worksheet_Name, RR, NumErrs)
If I = 208 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f208, ""), Nz(Rst!src_f208, ""), Nz(Rst!err_f208, ""), Worksheet_Name, RR, NumErrs)
If I = 209 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f209, ""), Nz(Rst!src_f209, ""), Nz(Rst!err_f209, ""), Worksheet_Name, RR, NumErrs)
If I = 210 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f210, ""), Nz(Rst!src_f210, ""), Nz(Rst!err_f210, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 210 Then Exit Sub

If I = 211 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f211, ""), Nz(Rst!src_f211, ""), Nz(Rst!err_f211, ""), Worksheet_Name, RR, NumErrs)
If I = 212 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f212, ""), Nz(Rst!src_f212, ""), Nz(Rst!err_f212, ""), Worksheet_Name, RR, NumErrs)
If I = 213 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f213, ""), Nz(Rst!src_f213, ""), Nz(Rst!err_f213, ""), Worksheet_Name, RR, NumErrs)
If I = 214 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f214, ""), Nz(Rst!src_f214, ""), Nz(Rst!err_f214, ""), Worksheet_Name, RR, NumErrs)
If I = 215 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f215, ""), Nz(Rst!src_f215, ""), Nz(Rst!err_f215, ""), Worksheet_Name, RR, NumErrs)
If I = 216 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f216, ""), Nz(Rst!src_f216, ""), Nz(Rst!err_f216, ""), Worksheet_Name, RR, NumErrs)
If I = 217 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f217, ""), Nz(Rst!src_f217, ""), Nz(Rst!err_f217, ""), Worksheet_Name, RR, NumErrs)
If I = 218 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f218, ""), Nz(Rst!src_f218, ""), Nz(Rst!err_f218, ""), Worksheet_Name, RR, NumErrs)
If I = 219 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f219, ""), Nz(Rst!src_f219, ""), Nz(Rst!err_f219, ""), Worksheet_Name, RR, NumErrs)
If I = 220 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f220, ""), Nz(Rst!src_f220, ""), Nz(Rst!err_f220, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 220 Then Exit Sub

If I = 221 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f221, ""), Nz(Rst!src_f221, ""), Nz(Rst!err_f221, ""), Worksheet_Name, RR, NumErrs)
If I = 222 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f222, ""), Nz(Rst!src_f222, ""), Nz(Rst!err_f222, ""), Worksheet_Name, RR, NumErrs)
If I = 223 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f223, ""), Nz(Rst!src_f223, ""), Nz(Rst!err_f223, ""), Worksheet_Name, RR, NumErrs)
If I = 224 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f224, ""), Nz(Rst!src_f224, ""), Nz(Rst!err_f224, ""), Worksheet_Name, RR, NumErrs)
If I = 225 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f225, ""), Nz(Rst!src_f225, ""), Nz(Rst!err_f225, ""), Worksheet_Name, RR, NumErrs)
If I = 226 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f226, ""), Nz(Rst!src_f226, ""), Nz(Rst!err_f226, ""), Worksheet_Name, RR, NumErrs)
If I = 227 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f227, ""), Nz(Rst!src_f227, ""), Nz(Rst!err_f227, ""), Worksheet_Name, RR, NumErrs)
If I = 228 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f228, ""), Nz(Rst!src_f228, ""), Nz(Rst!err_f228, ""), Worksheet_Name, RR, NumErrs)
If I = 229 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f229, ""), Nz(Rst!src_f229, ""), Nz(Rst!err_f229, ""), Worksheet_Name, RR, NumErrs)
If I = 230 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f230, ""), Nz(Rst!src_f230, ""), Nz(Rst!err_f230, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 230 Then Exit Sub

If I = 231 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f231, ""), Nz(Rst!src_f231, ""), Nz(Rst!err_f231, ""), Worksheet_Name, RR, NumErrs)
If I = 232 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f232, ""), Nz(Rst!src_f232, ""), Nz(Rst!err_f232, ""), Worksheet_Name, RR, NumErrs)
If I = 233 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f233, ""), Nz(Rst!src_f233, ""), Nz(Rst!err_f233, ""), Worksheet_Name, RR, NumErrs)
If I = 234 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f234, ""), Nz(Rst!src_f234, ""), Nz(Rst!err_f234, ""), Worksheet_Name, RR, NumErrs)
If I = 235 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f235, ""), Nz(Rst!src_f235, ""), Nz(Rst!err_f235, ""), Worksheet_Name, RR, NumErrs)
If I = 236 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f236, ""), Nz(Rst!src_f236, ""), Nz(Rst!err_f236, ""), Worksheet_Name, RR, NumErrs)
If I = 237 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f237, ""), Nz(Rst!src_f237, ""), Nz(Rst!err_f237, ""), Worksheet_Name, RR, NumErrs)
If I = 238 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f238, ""), Nz(Rst!src_f238, ""), Nz(Rst!err_f238, ""), Worksheet_Name, RR, NumErrs)
If I = 239 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f239, ""), Nz(Rst!src_f239, ""), Nz(Rst!err_f239, ""), Worksheet_Name, RR, NumErrs)
If I = 240 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f240, ""), Nz(Rst!src_f240, ""), Nz(Rst!err_f240, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 240 Then Exit Sub

If I = 241 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f241, ""), Nz(Rst!src_f241, ""), Nz(Rst!err_f241, ""), Worksheet_Name, RR, NumErrs)
If I = 242 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f242, ""), Nz(Rst!src_f242, ""), Nz(Rst!err_f242, ""), Worksheet_Name, RR, NumErrs)
If I = 243 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f243, ""), Nz(Rst!src_f243, ""), Nz(Rst!err_f243, ""), Worksheet_Name, RR, NumErrs)
If I = 244 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f244, ""), Nz(Rst!src_f244, ""), Nz(Rst!err_f244, ""), Worksheet_Name, RR, NumErrs)
If I = 245 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f245, ""), Nz(Rst!src_f245, ""), Nz(Rst!err_f245, ""), Worksheet_Name, RR, NumErrs)
If I = 246 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f246, ""), Nz(Rst!src_f246, ""), Nz(Rst!err_f246, ""), Worksheet_Name, RR, NumErrs)
If I = 247 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f247, ""), Nz(Rst!src_f247, ""), Nz(Rst!err_f247, ""), Worksheet_Name, RR, NumErrs)
If I = 248 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f248, ""), Nz(Rst!src_f248, ""), Nz(Rst!err_f248, ""), Worksheet_Name, RR, NumErrs)
If I = 249 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f249, ""), Nz(Rst!src_f249, ""), Nz(Rst!err_f249, ""), Worksheet_Name, RR, NumErrs)
If I = 250 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f250, ""), Nz(Rst!src_f250, ""), Nz(Rst!err_f250, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 250 Then Exit Sub

If I = 251 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f251, ""), Nz(Rst!src_f251, ""), Nz(Rst!err_f251, ""), Worksheet_Name, RR, NumErrs)
If I = 252 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f252, ""), Nz(Rst!src_f252, ""), Nz(Rst!err_f252, ""), Worksheet_Name, RR, NumErrs)
If I = 253 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f253, ""), Nz(Rst!src_f253, ""), Nz(Rst!err_f253, ""), Worksheet_Name, RR, NumErrs)
If I = 254 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f254, ""), Nz(Rst!src_f254, ""), Nz(Rst!err_f254, ""), Worksheet_Name, RR, NumErrs)
If I = 255 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f255, ""), Nz(Rst!src_f255, ""), Nz(Rst!err_f255, ""), Worksheet_Name, RR, NumErrs)
If I = 256 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f256, ""), Nz(Rst!src_f256, ""), Nz(Rst!err_f256, ""), Worksheet_Name, RR, NumErrs)
If I = 257 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f257, ""), Nz(Rst!src_f257, ""), Nz(Rst!err_f257, ""), Worksheet_Name, RR, NumErrs)
If I = 258 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f258, ""), Nz(Rst!src_f258, ""), Nz(Rst!err_f258, ""), Worksheet_Name, RR, NumErrs)
If I = 259 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f259, ""), Nz(Rst!src_f259, ""), Nz(Rst!err_f259, ""), Worksheet_Name, RR, NumErrs)
If I = 260 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f260, ""), Nz(Rst!src_f260, ""), Nz(Rst!err_f260, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 260 Then Exit Sub

If I = 261 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f261, ""), Nz(Rst!src_f261, ""), Nz(Rst!err_f261, ""), Worksheet_Name, RR, NumErrs)
If I = 262 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f262, ""), Nz(Rst!src_f262, ""), Nz(Rst!err_f262, ""), Worksheet_Name, RR, NumErrs)
If I = 263 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f263, ""), Nz(Rst!src_f263, ""), Nz(Rst!err_f263, ""), Worksheet_Name, RR, NumErrs)
If I = 264 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f264, ""), Nz(Rst!src_f264, ""), Nz(Rst!err_f264, ""), Worksheet_Name, RR, NumErrs)
If I = 265 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f265, ""), Nz(Rst!src_f265, ""), Nz(Rst!err_f265, ""), Worksheet_Name, RR, NumErrs)
If I = 266 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f266, ""), Nz(Rst!src_f266, ""), Nz(Rst!err_f266, ""), Worksheet_Name, RR, NumErrs)
If I = 267 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f267, ""), Nz(Rst!src_f267, ""), Nz(Rst!err_f267, ""), Worksheet_Name, RR, NumErrs)
If I = 268 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f268, ""), Nz(Rst!src_f268, ""), Nz(Rst!err_f268, ""), Worksheet_Name, RR, NumErrs)
If I = 269 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f269, ""), Nz(Rst!src_f269, ""), Nz(Rst!err_f269, ""), Worksheet_Name, RR, NumErrs)
If I = 270 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f270, ""), Nz(Rst!src_f270, ""), Nz(Rst!err_f270, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 270 Then Exit Sub

If I = 271 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f271, ""), Nz(Rst!src_f271, ""), Nz(Rst!err_f271, ""), Worksheet_Name, RR, NumErrs)
If I = 272 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f272, ""), Nz(Rst!src_f272, ""), Nz(Rst!err_f272, ""), Worksheet_Name, RR, NumErrs)
If I = 273 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f273, ""), Nz(Rst!src_f273, ""), Nz(Rst!err_f273, ""), Worksheet_Name, RR, NumErrs)
If I = 274 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f274, ""), Nz(Rst!src_f274, ""), Nz(Rst!err_f274, ""), Worksheet_Name, RR, NumErrs)
If I = 275 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f275, ""), Nz(Rst!src_f275, ""), Nz(Rst!err_f275, ""), Worksheet_Name, RR, NumErrs)
If I = 276 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f276, ""), Nz(Rst!src_f276, ""), Nz(Rst!err_f276, ""), Worksheet_Name, RR, NumErrs)
If I = 277 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f277, ""), Nz(Rst!src_f277, ""), Nz(Rst!err_f277, ""), Worksheet_Name, RR, NumErrs)
If I = 278 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f278, ""), Nz(Rst!src_f278, ""), Nz(Rst!err_f278, ""), Worksheet_Name, RR, NumErrs)
If I = 279 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f279, ""), Nz(Rst!src_f279, ""), Nz(Rst!err_f279, ""), Worksheet_Name, RR, NumErrs)
If I = 280 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f280, ""), Nz(Rst!src_f280, ""), Nz(Rst!err_f280, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 280 Then Exit Sub

If I = 281 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f281, ""), Nz(Rst!src_f281, ""), Nz(Rst!err_f281, ""), Worksheet_Name, RR, NumErrs)
If I = 282 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f282, ""), Nz(Rst!src_f282, ""), Nz(Rst!err_f282, ""), Worksheet_Name, RR, NumErrs)
If I = 283 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f283, ""), Nz(Rst!src_f283, ""), Nz(Rst!err_f283, ""), Worksheet_Name, RR, NumErrs)
If I = 284 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f284, ""), Nz(Rst!src_f284, ""), Nz(Rst!err_f284, ""), Worksheet_Name, RR, NumErrs)
If I = 285 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f285, ""), Nz(Rst!src_f285, ""), Nz(Rst!err_f285, ""), Worksheet_Name, RR, NumErrs)
If I = 286 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f286, ""), Nz(Rst!src_f286, ""), Nz(Rst!err_f286, ""), Worksheet_Name, RR, NumErrs)
If I = 287 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f287, ""), Nz(Rst!src_f287, ""), Nz(Rst!err_f287, ""), Worksheet_Name, RR, NumErrs)
If I = 288 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f288, ""), Nz(Rst!src_f288, ""), Nz(Rst!err_f288, ""), Worksheet_Name, RR, NumErrs)
If I = 289 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f289, ""), Nz(Rst!src_f289, ""), Nz(Rst!err_f289, ""), Worksheet_Name, RR, NumErrs)
If I = 290 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f290, ""), Nz(Rst!src_f290, ""), Nz(Rst!err_f290, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 290 Then Exit Sub

If I = 291 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f291, ""), Nz(Rst!src_f291, ""), Nz(Rst!err_f291, ""), Worksheet_Name, RR, NumErrs)
If I = 292 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f292, ""), Nz(Rst!src_f292, ""), Nz(Rst!err_f292, ""), Worksheet_Name, RR, NumErrs)
If I = 293 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f293, ""), Nz(Rst!src_f293, ""), Nz(Rst!err_f293, ""), Worksheet_Name, RR, NumErrs)
If I = 294 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f294, ""), Nz(Rst!src_f294, ""), Nz(Rst!err_f294, ""), Worksheet_Name, RR, NumErrs)
If I = 295 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f295, ""), Nz(Rst!src_f295, ""), Nz(Rst!err_f295, ""), Worksheet_Name, RR, NumErrs)
If I = 296 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f296, ""), Nz(Rst!src_f296, ""), Nz(Rst!err_f296, ""), Worksheet_Name, RR, NumErrs)
If I = 297 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f297, ""), Nz(Rst!src_f297, ""), Nz(Rst!err_f297, ""), Worksheet_Name, RR, NumErrs)
If I = 298 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f298, ""), Nz(Rst!src_f298, ""), Nz(Rst!err_f298, ""), Worksheet_Name, RR, NumErrs)
If I = 299 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f299, ""), Nz(Rst!src_f299, ""), Nz(Rst!err_f299, ""), Worksheet_Name, RR, NumErrs)
If I = 300 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f300, ""), Nz(Rst!src_f300, ""), Nz(Rst!err_f300, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 300 Then Exit Sub

End Sub

Private Sub Process_Field_Update_Data_Error_Check400(MaxColNum As Long, fldPtr As Long, _
                Rst As Recordset, _
                Worksheet_Name As String, _
                ByVal Output_Table As String, _
                ByVal ReptName As String, _
                ByRef NumErrs As Variant)
                       
Dim RR As String:  RR = ReptName
Dim I As Long
I = fldPtr

If MaxColNum > 401 Then Exit Sub

If I = 301 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f301, ""), Nz(Rst!src_f301, ""), Nz(Rst!err_f301, ""), Worksheet_Name, RR, NumErrs)
If I = 302 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f302, ""), Nz(Rst!src_f302, ""), Nz(Rst!err_f302, ""), Worksheet_Name, RR, NumErrs)
If I = 303 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f303, ""), Nz(Rst!src_f303, ""), Nz(Rst!err_f303, ""), Worksheet_Name, RR, NumErrs)
If I = 304 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f304, ""), Nz(Rst!src_f304, ""), Nz(Rst!err_f304, ""), Worksheet_Name, RR, NumErrs)
If I = 305 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f305, ""), Nz(Rst!src_f305, ""), Nz(Rst!err_f305, ""), Worksheet_Name, RR, NumErrs)
If I = 306 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f306, ""), Nz(Rst!src_f306, ""), Nz(Rst!err_f306, ""), Worksheet_Name, RR, NumErrs)
If I = 307 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f307, ""), Nz(Rst!src_f307, ""), Nz(Rst!err_f307, ""), Worksheet_Name, RR, NumErrs)
If I = 308 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f308, ""), Nz(Rst!src_f308, ""), Nz(Rst!err_f308, ""), Worksheet_Name, RR, NumErrs)
If I = 309 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f309, ""), Nz(Rst!src_f309, ""), Nz(Rst!err_f309, ""), Worksheet_Name, RR, NumErrs)
If I = 310 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f310, ""), Nz(Rst!src_f310, ""), Nz(Rst!err_f310, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 310 Then Exit Sub

If I = 311 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f311, ""), Nz(Rst!src_f311, ""), Nz(Rst!err_f311, ""), Worksheet_Name, RR, NumErrs)
If I = 312 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f312, ""), Nz(Rst!src_f312, ""), Nz(Rst!err_f312, ""), Worksheet_Name, RR, NumErrs)
If I = 313 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f313, ""), Nz(Rst!src_f313, ""), Nz(Rst!err_f313, ""), Worksheet_Name, RR, NumErrs)
If I = 314 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f314, ""), Nz(Rst!src_f314, ""), Nz(Rst!err_f314, ""), Worksheet_Name, RR, NumErrs)
If I = 315 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f315, ""), Nz(Rst!src_f315, ""), Nz(Rst!err_f315, ""), Worksheet_Name, RR, NumErrs)
If I = 316 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f316, ""), Nz(Rst!src_f316, ""), Nz(Rst!err_f316, ""), Worksheet_Name, RR, NumErrs)
If I = 317 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f317, ""), Nz(Rst!src_f317, ""), Nz(Rst!err_f317, ""), Worksheet_Name, RR, NumErrs)
If I = 318 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f318, ""), Nz(Rst!src_f318, ""), Nz(Rst!err_f318, ""), Worksheet_Name, RR, NumErrs)
If I = 319 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f319, ""), Nz(Rst!src_f319, ""), Nz(Rst!err_f319, ""), Worksheet_Name, RR, NumErrs)
If I = 320 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f320, ""), Nz(Rst!src_f320, ""), Nz(Rst!err_f320, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 320 Then Exit Sub

If I = 321 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f321, ""), Nz(Rst!src_f321, ""), Nz(Rst!err_f321, ""), Worksheet_Name, RR, NumErrs)
If I = 322 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f322, ""), Nz(Rst!src_f322, ""), Nz(Rst!err_f322, ""), Worksheet_Name, RR, NumErrs)
If I = 323 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f323, ""), Nz(Rst!src_f323, ""), Nz(Rst!err_f323, ""), Worksheet_Name, RR, NumErrs)
If I = 324 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f324, ""), Nz(Rst!src_f324, ""), Nz(Rst!err_f324, ""), Worksheet_Name, RR, NumErrs)
If I = 325 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f325, ""), Nz(Rst!src_f325, ""), Nz(Rst!err_f325, ""), Worksheet_Name, RR, NumErrs)
If I = 326 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f326, ""), Nz(Rst!src_f326, ""), Nz(Rst!err_f326, ""), Worksheet_Name, RR, NumErrs)
If I = 327 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f327, ""), Nz(Rst!src_f327, ""), Nz(Rst!err_f327, ""), Worksheet_Name, RR, NumErrs)
If I = 328 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f328, ""), Nz(Rst!src_f328, ""), Nz(Rst!err_f328, ""), Worksheet_Name, RR, NumErrs)
If I = 329 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f329, ""), Nz(Rst!src_f329, ""), Nz(Rst!err_f329, ""), Worksheet_Name, RR, NumErrs)
If I = 330 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f330, ""), Nz(Rst!src_f330, ""), Nz(Rst!err_f330, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 330 Then Exit Sub

If I = 331 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f331, ""), Nz(Rst!src_f331, ""), Nz(Rst!err_f331, ""), Worksheet_Name, RR, NumErrs)
If I = 332 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f332, ""), Nz(Rst!src_f332, ""), Nz(Rst!err_f332, ""), Worksheet_Name, RR, NumErrs)
If I = 333 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f333, ""), Nz(Rst!src_f333, ""), Nz(Rst!err_f333, ""), Worksheet_Name, RR, NumErrs)
If I = 334 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f334, ""), Nz(Rst!src_f334, ""), Nz(Rst!err_f334, ""), Worksheet_Name, RR, NumErrs)
If I = 335 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f335, ""), Nz(Rst!src_f335, ""), Nz(Rst!err_f335, ""), Worksheet_Name, RR, NumErrs)
If I = 336 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f336, ""), Nz(Rst!src_f336, ""), Nz(Rst!err_f336, ""), Worksheet_Name, RR, NumErrs)
If I = 337 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f337, ""), Nz(Rst!src_f337, ""), Nz(Rst!err_f337, ""), Worksheet_Name, RR, NumErrs)
If I = 338 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f338, ""), Nz(Rst!src_f338, ""), Nz(Rst!err_f338, ""), Worksheet_Name, RR, NumErrs)
If I = 339 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f339, ""), Nz(Rst!src_f339, ""), Nz(Rst!err_f339, ""), Worksheet_Name, RR, NumErrs)
If I = 340 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f340, ""), Nz(Rst!src_f340, ""), Nz(Rst!err_f340, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 340 Then Exit Sub

If I = 341 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f341, ""), Nz(Rst!src_f341, ""), Nz(Rst!err_f341, ""), Worksheet_Name, RR, NumErrs)
If I = 342 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f342, ""), Nz(Rst!src_f342, ""), Nz(Rst!err_f342, ""), Worksheet_Name, RR, NumErrs)
If I = 343 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f343, ""), Nz(Rst!src_f343, ""), Nz(Rst!err_f343, ""), Worksheet_Name, RR, NumErrs)
If I = 344 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f344, ""), Nz(Rst!src_f344, ""), Nz(Rst!err_f344, ""), Worksheet_Name, RR, NumErrs)
If I = 345 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f345, ""), Nz(Rst!src_f345, ""), Nz(Rst!err_f345, ""), Worksheet_Name, RR, NumErrs)
If I = 346 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f346, ""), Nz(Rst!src_f346, ""), Nz(Rst!err_f346, ""), Worksheet_Name, RR, NumErrs)
If I = 347 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f347, ""), Nz(Rst!src_f347, ""), Nz(Rst!err_f347, ""), Worksheet_Name, RR, NumErrs)
If I = 348 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f348, ""), Nz(Rst!src_f348, ""), Nz(Rst!err_f348, ""), Worksheet_Name, RR, NumErrs)
If I = 349 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f349, ""), Nz(Rst!src_f349, ""), Nz(Rst!err_f349, ""), Worksheet_Name, RR, NumErrs)
If I = 350 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f350, ""), Nz(Rst!src_f350, ""), Nz(Rst!err_f350, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 350 Then Exit Sub

If I = 351 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f351, ""), Nz(Rst!src_f351, ""), Nz(Rst!err_f351, ""), Worksheet_Name, RR, NumErrs)
If I = 352 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f352, ""), Nz(Rst!src_f352, ""), Nz(Rst!err_f352, ""), Worksheet_Name, RR, NumErrs)
If I = 353 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f353, ""), Nz(Rst!src_f353, ""), Nz(Rst!err_f353, ""), Worksheet_Name, RR, NumErrs)
If I = 354 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f354, ""), Nz(Rst!src_f354, ""), Nz(Rst!err_f354, ""), Worksheet_Name, RR, NumErrs)
If I = 355 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f355, ""), Nz(Rst!src_f355, ""), Nz(Rst!err_f355, ""), Worksheet_Name, RR, NumErrs)
If I = 356 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f356, ""), Nz(Rst!src_f356, ""), Nz(Rst!err_f356, ""), Worksheet_Name, RR, NumErrs)
If I = 357 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f357, ""), Nz(Rst!src_f357, ""), Nz(Rst!err_f357, ""), Worksheet_Name, RR, NumErrs)
If I = 358 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f358, ""), Nz(Rst!src_f358, ""), Nz(Rst!err_f358, ""), Worksheet_Name, RR, NumErrs)
If I = 359 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f359, ""), Nz(Rst!src_f359, ""), Nz(Rst!err_f359, ""), Worksheet_Name, RR, NumErrs)
If I = 360 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f360, ""), Nz(Rst!src_f360, ""), Nz(Rst!err_f360, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 360 Then Exit Sub

If I = 361 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f361, ""), Nz(Rst!src_f361, ""), Nz(Rst!err_f361, ""), Worksheet_Name, RR, NumErrs)
If I = 362 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f362, ""), Nz(Rst!src_f362, ""), Nz(Rst!err_f362, ""), Worksheet_Name, RR, NumErrs)
If I = 363 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f363, ""), Nz(Rst!src_f363, ""), Nz(Rst!err_f363, ""), Worksheet_Name, RR, NumErrs)
If I = 364 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f364, ""), Nz(Rst!src_f364, ""), Nz(Rst!err_f364, ""), Worksheet_Name, RR, NumErrs)
If I = 365 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f365, ""), Nz(Rst!src_f365, ""), Nz(Rst!err_f365, ""), Worksheet_Name, RR, NumErrs)
If I = 366 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f366, ""), Nz(Rst!src_f366, ""), Nz(Rst!err_f366, ""), Worksheet_Name, RR, NumErrs)
If I = 367 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f367, ""), Nz(Rst!src_f367, ""), Nz(Rst!err_f367, ""), Worksheet_Name, RR, NumErrs)
If I = 368 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f368, ""), Nz(Rst!src_f368, ""), Nz(Rst!err_f368, ""), Worksheet_Name, RR, NumErrs)
If I = 369 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f369, ""), Nz(Rst!src_f369, ""), Nz(Rst!err_f369, ""), Worksheet_Name, RR, NumErrs)
If I = 370 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f370, ""), Nz(Rst!src_f370, ""), Nz(Rst!err_f370, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 370 Then Exit Sub

If I = 371 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f371, ""), Nz(Rst!src_f371, ""), Nz(Rst!err_f371, ""), Worksheet_Name, RR, NumErrs)
If I = 372 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f372, ""), Nz(Rst!src_f372, ""), Nz(Rst!err_f372, ""), Worksheet_Name, RR, NumErrs)
If I = 373 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f373, ""), Nz(Rst!src_f373, ""), Nz(Rst!err_f373, ""), Worksheet_Name, RR, NumErrs)
If I = 374 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f374, ""), Nz(Rst!src_f374, ""), Nz(Rst!err_f374, ""), Worksheet_Name, RR, NumErrs)
If I = 375 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f375, ""), Nz(Rst!src_f375, ""), Nz(Rst!err_f375, ""), Worksheet_Name, RR, NumErrs)
If I = 376 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f376, ""), Nz(Rst!src_f376, ""), Nz(Rst!err_f376, ""), Worksheet_Name, RR, NumErrs)
If I = 377 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f377, ""), Nz(Rst!src_f377, ""), Nz(Rst!err_f377, ""), Worksheet_Name, RR, NumErrs)
If I = 378 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f378, ""), Nz(Rst!src_f378, ""), Nz(Rst!err_f378, ""), Worksheet_Name, RR, NumErrs)
If I = 379 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f379, ""), Nz(Rst!src_f379, ""), Nz(Rst!err_f379, ""), Worksheet_Name, RR, NumErrs)
If I = 380 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f380, ""), Nz(Rst!src_f380, ""), Nz(Rst!err_f380, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 380 Then Exit Sub

If I = 381 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f381, ""), Nz(Rst!src_f381, ""), Nz(Rst!err_f381, ""), Worksheet_Name, RR, NumErrs)
If I = 382 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f382, ""), Nz(Rst!src_f382, ""), Nz(Rst!err_f382, ""), Worksheet_Name, RR, NumErrs)
If I = 383 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f383, ""), Nz(Rst!src_f383, ""), Nz(Rst!err_f383, ""), Worksheet_Name, RR, NumErrs)
If I = 384 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f384, ""), Nz(Rst!src_f384, ""), Nz(Rst!err_f384, ""), Worksheet_Name, RR, NumErrs)
If I = 385 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f385, ""), Nz(Rst!src_f385, ""), Nz(Rst!err_f385, ""), Worksheet_Name, RR, NumErrs)
If I = 386 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f386, ""), Nz(Rst!src_f386, ""), Nz(Rst!err_f386, ""), Worksheet_Name, RR, NumErrs)
If I = 387 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f387, ""), Nz(Rst!src_f387, ""), Nz(Rst!err_f387, ""), Worksheet_Name, RR, NumErrs)
If I = 388 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f388, ""), Nz(Rst!src_f388, ""), Nz(Rst!err_f388, ""), Worksheet_Name, RR, NumErrs)
If I = 389 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f389, ""), Nz(Rst!src_f389, ""), Nz(Rst!err_f389, ""), Worksheet_Name, RR, NumErrs)
If I = 390 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f390, ""), Nz(Rst!src_f390, ""), Nz(Rst!err_f390, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 390 Then Exit Sub

If I = 391 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f391, ""), Nz(Rst!src_f391, ""), Nz(Rst!err_f391, ""), Worksheet_Name, RR, NumErrs)
If I = 392 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f392, ""), Nz(Rst!src_f392, ""), Nz(Rst!err_f392, ""), Worksheet_Name, RR, NumErrs)
If I = 393 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f393, ""), Nz(Rst!src_f393, ""), Nz(Rst!err_f393, ""), Worksheet_Name, RR, NumErrs)
If I = 394 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f394, ""), Nz(Rst!src_f394, ""), Nz(Rst!err_f394, ""), Worksheet_Name, RR, NumErrs)
If I = 395 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f395, ""), Nz(Rst!src_f395, ""), Nz(Rst!err_f395, ""), Worksheet_Name, RR, NumErrs)
If I = 396 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f396, ""), Nz(Rst!src_f396, ""), Nz(Rst!err_f396, ""), Worksheet_Name, RR, NumErrs)
If I = 397 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f397, ""), Nz(Rst!src_f397, ""), Nz(Rst!err_f397, ""), Worksheet_Name, RR, NumErrs)
If I = 398 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f398, ""), Nz(Rst!src_f398, ""), Nz(Rst!err_f398, ""), Worksheet_Name, RR, NumErrs)
If I = 399 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f399, ""), Nz(Rst!src_f399, ""), Nz(Rst!err_f399, ""), Worksheet_Name, RR, NumErrs)
If I = 400 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f400, ""), Nz(Rst!src_f400, ""), Nz(Rst!err_f400, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 400 Then Exit Sub

End Sub

Private Sub Process_Field_Update_Data_Error_Check500(MaxColNum As Long, fldPtr As Long, _
                Rst As Recordset, _
                Worksheet_Name As String, _
                ByVal Output_Table As String, _
                ByVal ReptName As String, _
                ByRef NumErrs As Variant)
                       
Dim RR As String:  RR = ReptName
Dim I As Long
I = fldPtr

If MaxColNum > 501 Then Exit Sub

If I = 401 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f401, ""), Nz(Rst!src_f401, ""), Nz(Rst!err_f401, ""), Worksheet_Name, RR, NumErrs)
If I = 402 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f402, ""), Nz(Rst!src_f402, ""), Nz(Rst!err_f402, ""), Worksheet_Name, RR, NumErrs)
If I = 403 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f403, ""), Nz(Rst!src_f403, ""), Nz(Rst!err_f403, ""), Worksheet_Name, RR, NumErrs)
If I = 404 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f404, ""), Nz(Rst!src_f404, ""), Nz(Rst!err_f404, ""), Worksheet_Name, RR, NumErrs)
If I = 405 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f405, ""), Nz(Rst!src_f405, ""), Nz(Rst!err_f405, ""), Worksheet_Name, RR, NumErrs)
If I = 406 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f406, ""), Nz(Rst!src_f406, ""), Nz(Rst!err_f406, ""), Worksheet_Name, RR, NumErrs)
If I = 407 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f407, ""), Nz(Rst!src_f407, ""), Nz(Rst!err_f407, ""), Worksheet_Name, RR, NumErrs)
If I = 408 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f408, ""), Nz(Rst!src_f408, ""), Nz(Rst!err_f408, ""), Worksheet_Name, RR, NumErrs)
If I = 409 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f409, ""), Nz(Rst!src_f409, ""), Nz(Rst!err_f409, ""), Worksheet_Name, RR, NumErrs)
If I = 410 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f410, ""), Nz(Rst!src_f410, ""), Nz(Rst!err_f410, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 410 Then Exit Sub

If I = 411 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f411, ""), Nz(Rst!src_f411, ""), Nz(Rst!err_f411, ""), Worksheet_Name, RR, NumErrs)
If I = 412 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f412, ""), Nz(Rst!src_f412, ""), Nz(Rst!err_f412, ""), Worksheet_Name, RR, NumErrs)
If I = 413 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f413, ""), Nz(Rst!src_f413, ""), Nz(Rst!err_f413, ""), Worksheet_Name, RR, NumErrs)
If I = 414 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f414, ""), Nz(Rst!src_f414, ""), Nz(Rst!err_f414, ""), Worksheet_Name, RR, NumErrs)
If I = 415 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f415, ""), Nz(Rst!src_f415, ""), Nz(Rst!err_f415, ""), Worksheet_Name, RR, NumErrs)
If I = 416 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f416, ""), Nz(Rst!src_f416, ""), Nz(Rst!err_f416, ""), Worksheet_Name, RR, NumErrs)
If I = 417 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f417, ""), Nz(Rst!src_f417, ""), Nz(Rst!err_f417, ""), Worksheet_Name, RR, NumErrs)
If I = 418 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f418, ""), Nz(Rst!src_f418, ""), Nz(Rst!err_f418, ""), Worksheet_Name, RR, NumErrs)
If I = 419 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f419, ""), Nz(Rst!src_f419, ""), Nz(Rst!err_f419, ""), Worksheet_Name, RR, NumErrs)
If I = 420 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f420, ""), Nz(Rst!src_f420, ""), Nz(Rst!err_f420, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 420 Then Exit Sub

If I = 421 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f421, ""), Nz(Rst!src_f421, ""), Nz(Rst!err_f421, ""), Worksheet_Name, RR, NumErrs)
If I = 422 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f422, ""), Nz(Rst!src_f422, ""), Nz(Rst!err_f422, ""), Worksheet_Name, RR, NumErrs)
If I = 423 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f423, ""), Nz(Rst!src_f423, ""), Nz(Rst!err_f423, ""), Worksheet_Name, RR, NumErrs)
If I = 424 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f424, ""), Nz(Rst!src_f424, ""), Nz(Rst!err_f424, ""), Worksheet_Name, RR, NumErrs)
If I = 425 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f425, ""), Nz(Rst!src_f425, ""), Nz(Rst!err_f425, ""), Worksheet_Name, RR, NumErrs)
If I = 426 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f426, ""), Nz(Rst!src_f426, ""), Nz(Rst!err_f426, ""), Worksheet_Name, RR, NumErrs)
If I = 427 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f427, ""), Nz(Rst!src_f427, ""), Nz(Rst!err_f427, ""), Worksheet_Name, RR, NumErrs)
If I = 428 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f428, ""), Nz(Rst!src_f428, ""), Nz(Rst!err_f428, ""), Worksheet_Name, RR, NumErrs)
If I = 429 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f429, ""), Nz(Rst!src_f429, ""), Nz(Rst!err_f429, ""), Worksheet_Name, RR, NumErrs)
If I = 430 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f430, ""), Nz(Rst!src_f430, ""), Nz(Rst!err_f430, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 430 Then Exit Sub

If I = 431 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f431, ""), Nz(Rst!src_f431, ""), Nz(Rst!err_f431, ""), Worksheet_Name, RR, NumErrs)
If I = 432 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f432, ""), Nz(Rst!src_f432, ""), Nz(Rst!err_f432, ""), Worksheet_Name, RR, NumErrs)
If I = 433 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f433, ""), Nz(Rst!src_f433, ""), Nz(Rst!err_f433, ""), Worksheet_Name, RR, NumErrs)
If I = 434 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f434, ""), Nz(Rst!src_f434, ""), Nz(Rst!err_f434, ""), Worksheet_Name, RR, NumErrs)
If I = 435 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f435, ""), Nz(Rst!src_f435, ""), Nz(Rst!err_f435, ""), Worksheet_Name, RR, NumErrs)
If I = 436 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f436, ""), Nz(Rst!src_f436, ""), Nz(Rst!err_f436, ""), Worksheet_Name, RR, NumErrs)
If I = 437 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f437, ""), Nz(Rst!src_f437, ""), Nz(Rst!err_f437, ""), Worksheet_Name, RR, NumErrs)
If I = 438 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f438, ""), Nz(Rst!src_f438, ""), Nz(Rst!err_f438, ""), Worksheet_Name, RR, NumErrs)
If I = 439 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f439, ""), Nz(Rst!src_f439, ""), Nz(Rst!err_f439, ""), Worksheet_Name, RR, NumErrs)
If I = 440 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f440, ""), Nz(Rst!src_f440, ""), Nz(Rst!err_f440, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 440 Then Exit Sub

If I = 441 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f441, ""), Nz(Rst!src_f441, ""), Nz(Rst!err_f441, ""), Worksheet_Name, RR, NumErrs)
If I = 442 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f442, ""), Nz(Rst!src_f442, ""), Nz(Rst!err_f442, ""), Worksheet_Name, RR, NumErrs)
If I = 443 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f443, ""), Nz(Rst!src_f443, ""), Nz(Rst!err_f443, ""), Worksheet_Name, RR, NumErrs)
If I = 444 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f444, ""), Nz(Rst!src_f444, ""), Nz(Rst!err_f444, ""), Worksheet_Name, RR, NumErrs)
If I = 445 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f445, ""), Nz(Rst!src_f445, ""), Nz(Rst!err_f445, ""), Worksheet_Name, RR, NumErrs)
If I = 446 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f446, ""), Nz(Rst!src_f446, ""), Nz(Rst!err_f446, ""), Worksheet_Name, RR, NumErrs)
If I = 447 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f447, ""), Nz(Rst!src_f447, ""), Nz(Rst!err_f447, ""), Worksheet_Name, RR, NumErrs)
If I = 448 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f448, ""), Nz(Rst!src_f448, ""), Nz(Rst!err_f448, ""), Worksheet_Name, RR, NumErrs)
If I = 449 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f449, ""), Nz(Rst!src_f449, ""), Nz(Rst!err_f449, ""), Worksheet_Name, RR, NumErrs)
If I = 450 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f450, ""), Nz(Rst!src_f450, ""), Nz(Rst!err_f450, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 450 Then Exit Sub

If I = 451 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f451, ""), Nz(Rst!src_f451, ""), Nz(Rst!err_f451, ""), Worksheet_Name, RR, NumErrs)
If I = 452 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f452, ""), Nz(Rst!src_f452, ""), Nz(Rst!err_f452, ""), Worksheet_Name, RR, NumErrs)
If I = 453 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f453, ""), Nz(Rst!src_f453, ""), Nz(Rst!err_f453, ""), Worksheet_Name, RR, NumErrs)
If I = 454 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f454, ""), Nz(Rst!src_f454, ""), Nz(Rst!err_f454, ""), Worksheet_Name, RR, NumErrs)
If I = 455 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f455, ""), Nz(Rst!src_f455, ""), Nz(Rst!err_f455, ""), Worksheet_Name, RR, NumErrs)
If I = 456 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f456, ""), Nz(Rst!src_f456, ""), Nz(Rst!err_f456, ""), Worksheet_Name, RR, NumErrs)
If I = 457 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f457, ""), Nz(Rst!src_f457, ""), Nz(Rst!err_f457, ""), Worksheet_Name, RR, NumErrs)
If I = 458 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f458, ""), Nz(Rst!src_f458, ""), Nz(Rst!err_f458, ""), Worksheet_Name, RR, NumErrs)
If I = 459 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f459, ""), Nz(Rst!src_f459, ""), Nz(Rst!err_f459, ""), Worksheet_Name, RR, NumErrs)
If I = 460 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f460, ""), Nz(Rst!src_f460, ""), Nz(Rst!err_f460, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 460 Then Exit Sub

If I = 461 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f461, ""), Nz(Rst!src_f461, ""), Nz(Rst!err_f461, ""), Worksheet_Name, RR, NumErrs)
If I = 462 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f462, ""), Nz(Rst!src_f462, ""), Nz(Rst!err_f462, ""), Worksheet_Name, RR, NumErrs)
If I = 463 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f463, ""), Nz(Rst!src_f463, ""), Nz(Rst!err_f463, ""), Worksheet_Name, RR, NumErrs)
If I = 464 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f464, ""), Nz(Rst!src_f464, ""), Nz(Rst!err_f464, ""), Worksheet_Name, RR, NumErrs)
If I = 465 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f465, ""), Nz(Rst!src_f465, ""), Nz(Rst!err_f465, ""), Worksheet_Name, RR, NumErrs)
If I = 466 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f466, ""), Nz(Rst!src_f466, ""), Nz(Rst!err_f466, ""), Worksheet_Name, RR, NumErrs)
If I = 467 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f467, ""), Nz(Rst!src_f467, ""), Nz(Rst!err_f467, ""), Worksheet_Name, RR, NumErrs)
If I = 468 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f468, ""), Nz(Rst!src_f468, ""), Nz(Rst!err_f468, ""), Worksheet_Name, RR, NumErrs)
If I = 469 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f469, ""), Nz(Rst!src_f469, ""), Nz(Rst!err_f469, ""), Worksheet_Name, RR, NumErrs)
If I = 470 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f470, ""), Nz(Rst!src_f470, ""), Nz(Rst!err_f470, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 470 Then Exit Sub

If I = 471 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f471, ""), Nz(Rst!src_f471, ""), Nz(Rst!err_f471, ""), Worksheet_Name, RR, NumErrs)
If I = 472 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f472, ""), Nz(Rst!src_f472, ""), Nz(Rst!err_f472, ""), Worksheet_Name, RR, NumErrs)
If I = 473 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f473, ""), Nz(Rst!src_f473, ""), Nz(Rst!err_f473, ""), Worksheet_Name, RR, NumErrs)
If I = 474 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f474, ""), Nz(Rst!src_f474, ""), Nz(Rst!err_f474, ""), Worksheet_Name, RR, NumErrs)
If I = 475 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f475, ""), Nz(Rst!src_f475, ""), Nz(Rst!err_f475, ""), Worksheet_Name, RR, NumErrs)
If I = 476 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f476, ""), Nz(Rst!src_f476, ""), Nz(Rst!err_f476, ""), Worksheet_Name, RR, NumErrs)
If I = 477 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f477, ""), Nz(Rst!src_f477, ""), Nz(Rst!err_f477, ""), Worksheet_Name, RR, NumErrs)
If I = 478 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f478, ""), Nz(Rst!src_f478, ""), Nz(Rst!err_f478, ""), Worksheet_Name, RR, NumErrs)
If I = 479 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f479, ""), Nz(Rst!src_f479, ""), Nz(Rst!err_f479, ""), Worksheet_Name, RR, NumErrs)
If I = 480 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f480, ""), Nz(Rst!src_f480, ""), Nz(Rst!err_f480, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 480 Then Exit Sub

If I = 481 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f481, ""), Nz(Rst!src_f481, ""), Nz(Rst!err_f481, ""), Worksheet_Name, RR, NumErrs)
If I = 482 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f482, ""), Nz(Rst!src_f482, ""), Nz(Rst!err_f482, ""), Worksheet_Name, RR, NumErrs)
If I = 483 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f483, ""), Nz(Rst!src_f483, ""), Nz(Rst!err_f483, ""), Worksheet_Name, RR, NumErrs)
If I = 484 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f484, ""), Nz(Rst!src_f484, ""), Nz(Rst!err_f484, ""), Worksheet_Name, RR, NumErrs)
If I = 485 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f485, ""), Nz(Rst!src_f485, ""), Nz(Rst!err_f485, ""), Worksheet_Name, RR, NumErrs)
If I = 486 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f486, ""), Nz(Rst!src_f486, ""), Nz(Rst!err_f486, ""), Worksheet_Name, RR, NumErrs)
If I = 487 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f487, ""), Nz(Rst!src_f487, ""), Nz(Rst!err_f487, ""), Worksheet_Name, RR, NumErrs)
If I = 488 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f488, ""), Nz(Rst!src_f488, ""), Nz(Rst!err_f488, ""), Worksheet_Name, RR, NumErrs)
If I = 489 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f489, ""), Nz(Rst!src_f489, ""), Nz(Rst!err_f489, ""), Worksheet_Name, RR, NumErrs)
If I = 490 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f490, ""), Nz(Rst!src_f490, ""), Nz(Rst!err_f490, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 490 Then Exit Sub

If I = 491 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f491, ""), Nz(Rst!src_f491, ""), Nz(Rst!err_f491, ""), Worksheet_Name, RR, NumErrs)
If I = 492 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f492, ""), Nz(Rst!src_f492, ""), Nz(Rst!err_f492, ""), Worksheet_Name, RR, NumErrs)
If I = 493 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f493, ""), Nz(Rst!src_f493, ""), Nz(Rst!err_f493, ""), Worksheet_Name, RR, NumErrs)
If I = 494 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f494, ""), Nz(Rst!src_f494, ""), Nz(Rst!err_f494, ""), Worksheet_Name, RR, NumErrs)
If I = 495 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f495, ""), Nz(Rst!src_f495, ""), Nz(Rst!err_f495, ""), Worksheet_Name, RR, NumErrs)
If I = 496 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f496, ""), Nz(Rst!src_f496, ""), Nz(Rst!err_f496, ""), Worksheet_Name, RR, NumErrs)
If I = 497 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f497, ""), Nz(Rst!src_f497, ""), Nz(Rst!err_f497, ""), Worksheet_Name, RR, NumErrs)
If I = 498 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f498, ""), Nz(Rst!src_f498, ""), Nz(Rst!err_f498, ""), Worksheet_Name, RR, NumErrs)
If I = 499 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f499, ""), Nz(Rst!src_f499, ""), Nz(Rst!err_f499, ""), Worksheet_Name, RR, NumErrs)
If I = 500 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f500, ""), Nz(Rst!src_f500, ""), Nz(Rst!err_f500, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 500 Then Exit Sub

End Sub

Private Sub Process_Field_Update_Data_Error_Check600(MaxColNum As Long, fldPtr As Long, _
                Rst As Recordset, _
                Worksheet_Name As String, _
                ByVal Output_Table As String, _
                ByVal ReptName As String, _
                ByRef NumErrs As Variant)
                       
Dim RR As String:  RR = ReptName
Dim I As Long
I = fldPtr

If MaxColNum > 601 Then Exit Sub

If I = 501 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f501, ""), Nz(Rst!src_f501, ""), Nz(Rst!err_f501, ""), Worksheet_Name, RR, NumErrs)
If I = 502 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f502, ""), Nz(Rst!src_f502, ""), Nz(Rst!err_f502, ""), Worksheet_Name, RR, NumErrs)
If I = 503 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f503, ""), Nz(Rst!src_f503, ""), Nz(Rst!err_f503, ""), Worksheet_Name, RR, NumErrs)
If I = 504 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f504, ""), Nz(Rst!src_f504, ""), Nz(Rst!err_f504, ""), Worksheet_Name, RR, NumErrs)
If I = 505 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f505, ""), Nz(Rst!src_f505, ""), Nz(Rst!err_f505, ""), Worksheet_Name, RR, NumErrs)
If I = 506 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f506, ""), Nz(Rst!src_f506, ""), Nz(Rst!err_f506, ""), Worksheet_Name, RR, NumErrs)
If I = 507 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f507, ""), Nz(Rst!src_f507, ""), Nz(Rst!err_f507, ""), Worksheet_Name, RR, NumErrs)
If I = 508 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f508, ""), Nz(Rst!src_f508, ""), Nz(Rst!err_f508, ""), Worksheet_Name, RR, NumErrs)
If I = 509 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f509, ""), Nz(Rst!src_f509, ""), Nz(Rst!err_f509, ""), Worksheet_Name, RR, NumErrs)
If I = 510 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f510, ""), Nz(Rst!src_f510, ""), Nz(Rst!err_f510, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 510 Then Exit Sub

If I = 511 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f511, ""), Nz(Rst!src_f511, ""), Nz(Rst!err_f511, ""), Worksheet_Name, RR, NumErrs)
If I = 512 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f512, ""), Nz(Rst!src_f512, ""), Nz(Rst!err_f512, ""), Worksheet_Name, RR, NumErrs)
If I = 513 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f513, ""), Nz(Rst!src_f513, ""), Nz(Rst!err_f513, ""), Worksheet_Name, RR, NumErrs)
If I = 514 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f514, ""), Nz(Rst!src_f514, ""), Nz(Rst!err_f514, ""), Worksheet_Name, RR, NumErrs)
If I = 515 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f515, ""), Nz(Rst!src_f515, ""), Nz(Rst!err_f515, ""), Worksheet_Name, RR, NumErrs)
If I = 516 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f516, ""), Nz(Rst!src_f516, ""), Nz(Rst!err_f516, ""), Worksheet_Name, RR, NumErrs)
If I = 517 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f517, ""), Nz(Rst!src_f517, ""), Nz(Rst!err_f517, ""), Worksheet_Name, RR, NumErrs)
If I = 518 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f518, ""), Nz(Rst!src_f518, ""), Nz(Rst!err_f518, ""), Worksheet_Name, RR, NumErrs)
If I = 519 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f519, ""), Nz(Rst!src_f519, ""), Nz(Rst!err_f519, ""), Worksheet_Name, RR, NumErrs)
If I = 520 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f520, ""), Nz(Rst!src_f520, ""), Nz(Rst!err_f520, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 520 Then Exit Sub

If I = 521 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f521, ""), Nz(Rst!src_f521, ""), Nz(Rst!err_f521, ""), Worksheet_Name, RR, NumErrs)
If I = 522 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f522, ""), Nz(Rst!src_f522, ""), Nz(Rst!err_f522, ""), Worksheet_Name, RR, NumErrs)
If I = 523 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f523, ""), Nz(Rst!src_f523, ""), Nz(Rst!err_f523, ""), Worksheet_Name, RR, NumErrs)
If I = 524 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f524, ""), Nz(Rst!src_f524, ""), Nz(Rst!err_f524, ""), Worksheet_Name, RR, NumErrs)
If I = 525 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f525, ""), Nz(Rst!src_f525, ""), Nz(Rst!err_f525, ""), Worksheet_Name, RR, NumErrs)
If I = 526 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f526, ""), Nz(Rst!src_f526, ""), Nz(Rst!err_f526, ""), Worksheet_Name, RR, NumErrs)
If I = 527 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f527, ""), Nz(Rst!src_f527, ""), Nz(Rst!err_f527, ""), Worksheet_Name, RR, NumErrs)
If I = 528 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f528, ""), Nz(Rst!src_f528, ""), Nz(Rst!err_f528, ""), Worksheet_Name, RR, NumErrs)
If I = 529 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f529, ""), Nz(Rst!src_f529, ""), Nz(Rst!err_f529, ""), Worksheet_Name, RR, NumErrs)
If I = 530 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f530, ""), Nz(Rst!src_f530, ""), Nz(Rst!err_f530, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 530 Then Exit Sub

If I = 531 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f531, ""), Nz(Rst!src_f531, ""), Nz(Rst!err_f531, ""), Worksheet_Name, RR, NumErrs)
If I = 532 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f532, ""), Nz(Rst!src_f532, ""), Nz(Rst!err_f532, ""), Worksheet_Name, RR, NumErrs)
If I = 533 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f533, ""), Nz(Rst!src_f533, ""), Nz(Rst!err_f533, ""), Worksheet_Name, RR, NumErrs)
If I = 534 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f534, ""), Nz(Rst!src_f534, ""), Nz(Rst!err_f534, ""), Worksheet_Name, RR, NumErrs)
If I = 535 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f535, ""), Nz(Rst!src_f535, ""), Nz(Rst!err_f535, ""), Worksheet_Name, RR, NumErrs)
If I = 536 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f536, ""), Nz(Rst!src_f536, ""), Nz(Rst!err_f536, ""), Worksheet_Name, RR, NumErrs)
If I = 537 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f537, ""), Nz(Rst!src_f537, ""), Nz(Rst!err_f537, ""), Worksheet_Name, RR, NumErrs)
If I = 538 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f538, ""), Nz(Rst!src_f538, ""), Nz(Rst!err_f538, ""), Worksheet_Name, RR, NumErrs)
If I = 539 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f539, ""), Nz(Rst!src_f539, ""), Nz(Rst!err_f539, ""), Worksheet_Name, RR, NumErrs)
If I = 540 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f540, ""), Nz(Rst!src_f540, ""), Nz(Rst!err_f540, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 540 Then Exit Sub

If I = 541 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f541, ""), Nz(Rst!src_f541, ""), Nz(Rst!err_f541, ""), Worksheet_Name, RR, NumErrs)
If I = 542 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f542, ""), Nz(Rst!src_f542, ""), Nz(Rst!err_f542, ""), Worksheet_Name, RR, NumErrs)
If I = 543 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f543, ""), Nz(Rst!src_f543, ""), Nz(Rst!err_f543, ""), Worksheet_Name, RR, NumErrs)
If I = 544 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f544, ""), Nz(Rst!src_f544, ""), Nz(Rst!err_f544, ""), Worksheet_Name, RR, NumErrs)
If I = 545 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f545, ""), Nz(Rst!src_f545, ""), Nz(Rst!err_f545, ""), Worksheet_Name, RR, NumErrs)
If I = 546 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f546, ""), Nz(Rst!src_f546, ""), Nz(Rst!err_f546, ""), Worksheet_Name, RR, NumErrs)
If I = 547 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f547, ""), Nz(Rst!src_f547, ""), Nz(Rst!err_f547, ""), Worksheet_Name, RR, NumErrs)
If I = 548 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f548, ""), Nz(Rst!src_f548, ""), Nz(Rst!err_f548, ""), Worksheet_Name, RR, NumErrs)
If I = 549 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f549, ""), Nz(Rst!src_f549, ""), Nz(Rst!err_f549, ""), Worksheet_Name, RR, NumErrs)
If I = 550 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f550, ""), Nz(Rst!src_f550, ""), Nz(Rst!err_f550, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 550 Then Exit Sub

If I = 551 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f551, ""), Nz(Rst!src_f551, ""), Nz(Rst!err_f551, ""), Worksheet_Name, RR, NumErrs)
If I = 552 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f552, ""), Nz(Rst!src_f552, ""), Nz(Rst!err_f552, ""), Worksheet_Name, RR, NumErrs)
If I = 553 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f553, ""), Nz(Rst!src_f553, ""), Nz(Rst!err_f553, ""), Worksheet_Name, RR, NumErrs)
If I = 554 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f554, ""), Nz(Rst!src_f554, ""), Nz(Rst!err_f554, ""), Worksheet_Name, RR, NumErrs)
If I = 555 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f555, ""), Nz(Rst!src_f555, ""), Nz(Rst!err_f555, ""), Worksheet_Name, RR, NumErrs)
If I = 556 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f556, ""), Nz(Rst!src_f556, ""), Nz(Rst!err_f556, ""), Worksheet_Name, RR, NumErrs)
If I = 557 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f557, ""), Nz(Rst!src_f557, ""), Nz(Rst!err_f557, ""), Worksheet_Name, RR, NumErrs)
If I = 558 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f558, ""), Nz(Rst!src_f558, ""), Nz(Rst!err_f558, ""), Worksheet_Name, RR, NumErrs)
If I = 559 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f559, ""), Nz(Rst!src_f559, ""), Nz(Rst!err_f559, ""), Worksheet_Name, RR, NumErrs)
If I = 560 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f560, ""), Nz(Rst!src_f560, ""), Nz(Rst!err_f560, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 560 Then Exit Sub

If I = 561 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f561, ""), Nz(Rst!src_f561, ""), Nz(Rst!err_f561, ""), Worksheet_Name, RR, NumErrs)
If I = 562 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f562, ""), Nz(Rst!src_f562, ""), Nz(Rst!err_f562, ""), Worksheet_Name, RR, NumErrs)
If I = 563 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f563, ""), Nz(Rst!src_f563, ""), Nz(Rst!err_f563, ""), Worksheet_Name, RR, NumErrs)
If I = 564 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f564, ""), Nz(Rst!src_f564, ""), Nz(Rst!err_f564, ""), Worksheet_Name, RR, NumErrs)
If I = 565 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f565, ""), Nz(Rst!src_f565, ""), Nz(Rst!err_f565, ""), Worksheet_Name, RR, NumErrs)
If I = 566 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f566, ""), Nz(Rst!src_f566, ""), Nz(Rst!err_f566, ""), Worksheet_Name, RR, NumErrs)
If I = 567 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f567, ""), Nz(Rst!src_f567, ""), Nz(Rst!err_f567, ""), Worksheet_Name, RR, NumErrs)
If I = 568 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f568, ""), Nz(Rst!src_f568, ""), Nz(Rst!err_f568, ""), Worksheet_Name, RR, NumErrs)
If I = 569 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f569, ""), Nz(Rst!src_f569, ""), Nz(Rst!err_f569, ""), Worksheet_Name, RR, NumErrs)
If I = 570 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f570, ""), Nz(Rst!src_f570, ""), Nz(Rst!err_f570, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 570 Then Exit Sub

If I = 571 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f571, ""), Nz(Rst!src_f571, ""), Nz(Rst!err_f571, ""), Worksheet_Name, RR, NumErrs)
If I = 572 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f572, ""), Nz(Rst!src_f572, ""), Nz(Rst!err_f572, ""), Worksheet_Name, RR, NumErrs)
If I = 573 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f573, ""), Nz(Rst!src_f573, ""), Nz(Rst!err_f573, ""), Worksheet_Name, RR, NumErrs)
If I = 574 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f574, ""), Nz(Rst!src_f574, ""), Nz(Rst!err_f574, ""), Worksheet_Name, RR, NumErrs)
If I = 575 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f575, ""), Nz(Rst!src_f575, ""), Nz(Rst!err_f575, ""), Worksheet_Name, RR, NumErrs)
If I = 576 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f576, ""), Nz(Rst!src_f576, ""), Nz(Rst!err_f576, ""), Worksheet_Name, RR, NumErrs)
If I = 577 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f577, ""), Nz(Rst!src_f577, ""), Nz(Rst!err_f577, ""), Worksheet_Name, RR, NumErrs)
If I = 578 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f578, ""), Nz(Rst!src_f578, ""), Nz(Rst!err_f578, ""), Worksheet_Name, RR, NumErrs)
If I = 579 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f579, ""), Nz(Rst!src_f579, ""), Nz(Rst!err_f579, ""), Worksheet_Name, RR, NumErrs)
If I = 580 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f580, ""), Nz(Rst!src_f580, ""), Nz(Rst!err_f580, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 580 Then Exit Sub

If I = 581 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f581, ""), Nz(Rst!src_f581, ""), Nz(Rst!err_f581, ""), Worksheet_Name, RR, NumErrs)
If I = 582 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f582, ""), Nz(Rst!src_f582, ""), Nz(Rst!err_f582, ""), Worksheet_Name, RR, NumErrs)
If I = 583 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f583, ""), Nz(Rst!src_f583, ""), Nz(Rst!err_f583, ""), Worksheet_Name, RR, NumErrs)
If I = 584 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f584, ""), Nz(Rst!src_f584, ""), Nz(Rst!err_f584, ""), Worksheet_Name, RR, NumErrs)
If I = 585 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f585, ""), Nz(Rst!src_f585, ""), Nz(Rst!err_f585, ""), Worksheet_Name, RR, NumErrs)
If I = 586 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f586, ""), Nz(Rst!src_f586, ""), Nz(Rst!err_f586, ""), Worksheet_Name, RR, NumErrs)
If I = 587 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f587, ""), Nz(Rst!src_f587, ""), Nz(Rst!err_f587, ""), Worksheet_Name, RR, NumErrs)
If I = 588 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f588, ""), Nz(Rst!src_f588, ""), Nz(Rst!err_f588, ""), Worksheet_Name, RR, NumErrs)
If I = 589 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f589, ""), Nz(Rst!src_f589, ""), Nz(Rst!err_f589, ""), Worksheet_Name, RR, NumErrs)
If I = 590 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f590, ""), Nz(Rst!src_f590, ""), Nz(Rst!err_f590, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 590 Then Exit Sub

If I = 591 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f591, ""), Nz(Rst!src_f591, ""), Nz(Rst!err_f591, ""), Worksheet_Name, RR, NumErrs)
If I = 592 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f592, ""), Nz(Rst!src_f592, ""), Nz(Rst!err_f592, ""), Worksheet_Name, RR, NumErrs)
If I = 593 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f593, ""), Nz(Rst!src_f593, ""), Nz(Rst!err_f593, ""), Worksheet_Name, RR, NumErrs)
If I = 594 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f594, ""), Nz(Rst!src_f594, ""), Nz(Rst!err_f594, ""), Worksheet_Name, RR, NumErrs)
If I = 595 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f595, ""), Nz(Rst!src_f595, ""), Nz(Rst!err_f595, ""), Worksheet_Name, RR, NumErrs)
If I = 596 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f596, ""), Nz(Rst!src_f596, ""), Nz(Rst!err_f596, ""), Worksheet_Name, RR, NumErrs)
If I = 597 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f597, ""), Nz(Rst!src_f597, ""), Nz(Rst!err_f597, ""), Worksheet_Name, RR, NumErrs)
If I = 598 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f598, ""), Nz(Rst!src_f598, ""), Nz(Rst!err_f598, ""), Worksheet_Name, RR, NumErrs)
If I = 599 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f599, ""), Nz(Rst!src_f599, ""), Nz(Rst!err_f599, ""), Worksheet_Name, RR, NumErrs)
If I = 600 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f600, ""), Nz(Rst!src_f600, ""), Nz(Rst!err_f600, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 600 Then Exit Sub

End Sub

Private Sub Process_Field_Update_Data_Error_Check700(MaxColNum As Long, fldPtr As Long, _
                Rst As Recordset, _
                Worksheet_Name As String, _
                ByVal Output_Table As String, _
                ByVal ReptName As String, _
                ByRef NumErrs As Variant)
                       
Dim RR As String:  RR = ReptName
Dim I As Long
I = fldPtr

If MaxColNum > 701 Then Exit Sub

If I = 601 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f601, ""), Nz(Rst!src_f601, ""), Nz(Rst!err_f601, ""), Worksheet_Name, RR, NumErrs)
If I = 602 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f602, ""), Nz(Rst!src_f602, ""), Nz(Rst!err_f602, ""), Worksheet_Name, RR, NumErrs)
If I = 603 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f603, ""), Nz(Rst!src_f603, ""), Nz(Rst!err_f603, ""), Worksheet_Name, RR, NumErrs)
If I = 604 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f604, ""), Nz(Rst!src_f604, ""), Nz(Rst!err_f604, ""), Worksheet_Name, RR, NumErrs)
If I = 605 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f605, ""), Nz(Rst!src_f605, ""), Nz(Rst!err_f605, ""), Worksheet_Name, RR, NumErrs)
If I = 606 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f606, ""), Nz(Rst!src_f606, ""), Nz(Rst!err_f606, ""), Worksheet_Name, RR, NumErrs)
If I = 607 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f607, ""), Nz(Rst!src_f607, ""), Nz(Rst!err_f607, ""), Worksheet_Name, RR, NumErrs)
If I = 608 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f608, ""), Nz(Rst!src_f608, ""), Nz(Rst!err_f608, ""), Worksheet_Name, RR, NumErrs)
If I = 609 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f609, ""), Nz(Rst!src_f609, ""), Nz(Rst!err_f609, ""), Worksheet_Name, RR, NumErrs)
If I = 610 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f610, ""), Nz(Rst!src_f610, ""), Nz(Rst!err_f610, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 610 Then Exit Sub

If I = 611 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f611, ""), Nz(Rst!src_f611, ""), Nz(Rst!err_f611, ""), Worksheet_Name, RR, NumErrs)
If I = 612 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f612, ""), Nz(Rst!src_f612, ""), Nz(Rst!err_f612, ""), Worksheet_Name, RR, NumErrs)
If I = 613 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f613, ""), Nz(Rst!src_f613, ""), Nz(Rst!err_f613, ""), Worksheet_Name, RR, NumErrs)
If I = 614 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f614, ""), Nz(Rst!src_f614, ""), Nz(Rst!err_f614, ""), Worksheet_Name, RR, NumErrs)
If I = 615 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f615, ""), Nz(Rst!src_f615, ""), Nz(Rst!err_f615, ""), Worksheet_Name, RR, NumErrs)
If I = 616 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f616, ""), Nz(Rst!src_f616, ""), Nz(Rst!err_f616, ""), Worksheet_Name, RR, NumErrs)
If I = 617 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f617, ""), Nz(Rst!src_f617, ""), Nz(Rst!err_f617, ""), Worksheet_Name, RR, NumErrs)
If I = 618 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f618, ""), Nz(Rst!src_f618, ""), Nz(Rst!err_f618, ""), Worksheet_Name, RR, NumErrs)
If I = 619 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f619, ""), Nz(Rst!src_f619, ""), Nz(Rst!err_f619, ""), Worksheet_Name, RR, NumErrs)
If I = 620 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f620, ""), Nz(Rst!src_f620, ""), Nz(Rst!err_f620, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 620 Then Exit Sub

If I = 621 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f621, ""), Nz(Rst!src_f621, ""), Nz(Rst!err_f621, ""), Worksheet_Name, RR, NumErrs)
If I = 622 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f622, ""), Nz(Rst!src_f622, ""), Nz(Rst!err_f622, ""), Worksheet_Name, RR, NumErrs)
If I = 623 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f623, ""), Nz(Rst!src_f623, ""), Nz(Rst!err_f623, ""), Worksheet_Name, RR, NumErrs)
If I = 624 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f624, ""), Nz(Rst!src_f624, ""), Nz(Rst!err_f624, ""), Worksheet_Name, RR, NumErrs)
If I = 625 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f625, ""), Nz(Rst!src_f625, ""), Nz(Rst!err_f625, ""), Worksheet_Name, RR, NumErrs)
If I = 626 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f626, ""), Nz(Rst!src_f626, ""), Nz(Rst!err_f626, ""), Worksheet_Name, RR, NumErrs)
If I = 627 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f627, ""), Nz(Rst!src_f627, ""), Nz(Rst!err_f627, ""), Worksheet_Name, RR, NumErrs)
If I = 628 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f628, ""), Nz(Rst!src_f628, ""), Nz(Rst!err_f628, ""), Worksheet_Name, RR, NumErrs)
If I = 629 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f629, ""), Nz(Rst!src_f629, ""), Nz(Rst!err_f629, ""), Worksheet_Name, RR, NumErrs)
If I = 630 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f630, ""), Nz(Rst!src_f630, ""), Nz(Rst!err_f630, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 630 Then Exit Sub

If I = 631 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f631, ""), Nz(Rst!src_f631, ""), Nz(Rst!err_f631, ""), Worksheet_Name, RR, NumErrs)
If I = 632 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f632, ""), Nz(Rst!src_f632, ""), Nz(Rst!err_f632, ""), Worksheet_Name, RR, NumErrs)
If I = 633 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f633, ""), Nz(Rst!src_f633, ""), Nz(Rst!err_f633, ""), Worksheet_Name, RR, NumErrs)
If I = 634 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f634, ""), Nz(Rst!src_f634, ""), Nz(Rst!err_f634, ""), Worksheet_Name, RR, NumErrs)
If I = 635 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f635, ""), Nz(Rst!src_f635, ""), Nz(Rst!err_f635, ""), Worksheet_Name, RR, NumErrs)
If I = 636 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f636, ""), Nz(Rst!src_f636, ""), Nz(Rst!err_f636, ""), Worksheet_Name, RR, NumErrs)
If I = 637 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f637, ""), Nz(Rst!src_f637, ""), Nz(Rst!err_f637, ""), Worksheet_Name, RR, NumErrs)
If I = 638 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f638, ""), Nz(Rst!src_f638, ""), Nz(Rst!err_f638, ""), Worksheet_Name, RR, NumErrs)
If I = 639 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f639, ""), Nz(Rst!src_f639, ""), Nz(Rst!err_f639, ""), Worksheet_Name, RR, NumErrs)
If I = 640 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f640, ""), Nz(Rst!src_f640, ""), Nz(Rst!err_f640, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 640 Then Exit Sub

If I = 641 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f641, ""), Nz(Rst!src_f641, ""), Nz(Rst!err_f641, ""), Worksheet_Name, RR, NumErrs)
If I = 642 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f642, ""), Nz(Rst!src_f642, ""), Nz(Rst!err_f642, ""), Worksheet_Name, RR, NumErrs)
If I = 643 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f643, ""), Nz(Rst!src_f643, ""), Nz(Rst!err_f643, ""), Worksheet_Name, RR, NumErrs)
If I = 644 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f644, ""), Nz(Rst!src_f644, ""), Nz(Rst!err_f644, ""), Worksheet_Name, RR, NumErrs)
If I = 645 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f645, ""), Nz(Rst!src_f645, ""), Nz(Rst!err_f645, ""), Worksheet_Name, RR, NumErrs)
If I = 646 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f646, ""), Nz(Rst!src_f646, ""), Nz(Rst!err_f646, ""), Worksheet_Name, RR, NumErrs)
If I = 647 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f647, ""), Nz(Rst!src_f647, ""), Nz(Rst!err_f647, ""), Worksheet_Name, RR, NumErrs)
If I = 648 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f648, ""), Nz(Rst!src_f648, ""), Nz(Rst!err_f648, ""), Worksheet_Name, RR, NumErrs)
If I = 649 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f649, ""), Nz(Rst!src_f649, ""), Nz(Rst!err_f649, ""), Worksheet_Name, RR, NumErrs)
If I = 650 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f650, ""), Nz(Rst!src_f650, ""), Nz(Rst!err_f650, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 650 Then Exit Sub

If I = 651 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f651, ""), Nz(Rst!src_f651, ""), Nz(Rst!err_f651, ""), Worksheet_Name, RR, NumErrs)
If I = 652 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f652, ""), Nz(Rst!src_f652, ""), Nz(Rst!err_f652, ""), Worksheet_Name, RR, NumErrs)
If I = 653 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f653, ""), Nz(Rst!src_f653, ""), Nz(Rst!err_f653, ""), Worksheet_Name, RR, NumErrs)
If I = 654 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f654, ""), Nz(Rst!src_f654, ""), Nz(Rst!err_f654, ""), Worksheet_Name, RR, NumErrs)
If I = 655 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f655, ""), Nz(Rst!src_f655, ""), Nz(Rst!err_f655, ""), Worksheet_Name, RR, NumErrs)
If I = 656 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f656, ""), Nz(Rst!src_f656, ""), Nz(Rst!err_f656, ""), Worksheet_Name, RR, NumErrs)
If I = 657 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f657, ""), Nz(Rst!src_f657, ""), Nz(Rst!err_f657, ""), Worksheet_Name, RR, NumErrs)
If I = 658 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f658, ""), Nz(Rst!src_f658, ""), Nz(Rst!err_f658, ""), Worksheet_Name, RR, NumErrs)
If I = 659 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f659, ""), Nz(Rst!src_f659, ""), Nz(Rst!err_f659, ""), Worksheet_Name, RR, NumErrs)
If I = 660 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f660, ""), Nz(Rst!src_f660, ""), Nz(Rst!err_f660, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 660 Then Exit Sub

If I = 661 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f661, ""), Nz(Rst!src_f661, ""), Nz(Rst!err_f661, ""), Worksheet_Name, RR, NumErrs)
If I = 662 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f662, ""), Nz(Rst!src_f662, ""), Nz(Rst!err_f662, ""), Worksheet_Name, RR, NumErrs)
If I = 663 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f663, ""), Nz(Rst!src_f663, ""), Nz(Rst!err_f663, ""), Worksheet_Name, RR, NumErrs)
If I = 664 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f664, ""), Nz(Rst!src_f664, ""), Nz(Rst!err_f664, ""), Worksheet_Name, RR, NumErrs)
If I = 665 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f665, ""), Nz(Rst!src_f665, ""), Nz(Rst!err_f665, ""), Worksheet_Name, RR, NumErrs)
If I = 666 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f666, ""), Nz(Rst!src_f666, ""), Nz(Rst!err_f666, ""), Worksheet_Name, RR, NumErrs)
If I = 667 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f667, ""), Nz(Rst!src_f667, ""), Nz(Rst!err_f667, ""), Worksheet_Name, RR, NumErrs)
If I = 668 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f668, ""), Nz(Rst!src_f668, ""), Nz(Rst!err_f668, ""), Worksheet_Name, RR, NumErrs)
If I = 669 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f669, ""), Nz(Rst!src_f669, ""), Nz(Rst!err_f669, ""), Worksheet_Name, RR, NumErrs)
If I = 670 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f670, ""), Nz(Rst!src_f670, ""), Nz(Rst!err_f670, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 670 Then Exit Sub

If I = 671 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f671, ""), Nz(Rst!src_f671, ""), Nz(Rst!err_f671, ""), Worksheet_Name, RR, NumErrs)
If I = 672 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f672, ""), Nz(Rst!src_f672, ""), Nz(Rst!err_f672, ""), Worksheet_Name, RR, NumErrs)
If I = 673 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f673, ""), Nz(Rst!src_f673, ""), Nz(Rst!err_f673, ""), Worksheet_Name, RR, NumErrs)
If I = 674 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f674, ""), Nz(Rst!src_f674, ""), Nz(Rst!err_f674, ""), Worksheet_Name, RR, NumErrs)
If I = 675 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f675, ""), Nz(Rst!src_f675, ""), Nz(Rst!err_f675, ""), Worksheet_Name, RR, NumErrs)
If I = 676 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f676, ""), Nz(Rst!src_f676, ""), Nz(Rst!err_f676, ""), Worksheet_Name, RR, NumErrs)
If I = 677 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f677, ""), Nz(Rst!src_f677, ""), Nz(Rst!err_f677, ""), Worksheet_Name, RR, NumErrs)
If I = 678 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f678, ""), Nz(Rst!src_f678, ""), Nz(Rst!err_f678, ""), Worksheet_Name, RR, NumErrs)
If I = 679 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f679, ""), Nz(Rst!src_f679, ""), Nz(Rst!err_f679, ""), Worksheet_Name, RR, NumErrs)
If I = 680 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f680, ""), Nz(Rst!src_f680, ""), Nz(Rst!err_f680, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 680 Then Exit Sub

If I = 681 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f681, ""), Nz(Rst!src_f681, ""), Nz(Rst!err_f681, ""), Worksheet_Name, RR, NumErrs)
If I = 682 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f682, ""), Nz(Rst!src_f682, ""), Nz(Rst!err_f682, ""), Worksheet_Name, RR, NumErrs)
If I = 683 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f683, ""), Nz(Rst!src_f683, ""), Nz(Rst!err_f683, ""), Worksheet_Name, RR, NumErrs)
If I = 684 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f684, ""), Nz(Rst!src_f684, ""), Nz(Rst!err_f684, ""), Worksheet_Name, RR, NumErrs)
If I = 685 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f685, ""), Nz(Rst!src_f685, ""), Nz(Rst!err_f685, ""), Worksheet_Name, RR, NumErrs)
If I = 686 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f686, ""), Nz(Rst!src_f686, ""), Nz(Rst!err_f686, ""), Worksheet_Name, RR, NumErrs)
If I = 687 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f687, ""), Nz(Rst!src_f687, ""), Nz(Rst!err_f687, ""), Worksheet_Name, RR, NumErrs)
If I = 688 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f688, ""), Nz(Rst!src_f688, ""), Nz(Rst!err_f688, ""), Worksheet_Name, RR, NumErrs)
If I = 689 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f689, ""), Nz(Rst!src_f689, ""), Nz(Rst!err_f689, ""), Worksheet_Name, RR, NumErrs)
If I = 690 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f690, ""), Nz(Rst!src_f690, ""), Nz(Rst!err_f690, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 690 Then Exit Sub

If I = 691 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f691, ""), Nz(Rst!src_f691, ""), Nz(Rst!err_f691, ""), Worksheet_Name, RR, NumErrs)
If I = 692 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f692, ""), Nz(Rst!src_f692, ""), Nz(Rst!err_f692, ""), Worksheet_Name, RR, NumErrs)
If I = 693 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f693, ""), Nz(Rst!src_f693, ""), Nz(Rst!err_f693, ""), Worksheet_Name, RR, NumErrs)
If I = 694 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f694, ""), Nz(Rst!src_f694, ""), Nz(Rst!err_f694, ""), Worksheet_Name, RR, NumErrs)
If I = 695 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f695, ""), Nz(Rst!src_f695, ""), Nz(Rst!err_f695, ""), Worksheet_Name, RR, NumErrs)
If I = 696 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f696, ""), Nz(Rst!src_f696, ""), Nz(Rst!err_f696, ""), Worksheet_Name, RR, NumErrs)
If I = 697 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f697, ""), Nz(Rst!src_f697, ""), Nz(Rst!err_f697, ""), Worksheet_Name, RR, NumErrs)
If I = 698 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f698, ""), Nz(Rst!src_f698, ""), Nz(Rst!err_f698, ""), Worksheet_Name, RR, NumErrs)
If I = 699 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f699, ""), Nz(Rst!src_f699, ""), Nz(Rst!err_f699, ""), Worksheet_Name, RR, NumErrs)
If I = 700 Then Call Process_Field_Update_Data_Error_Check(Rst, I, Nz(Rst!tar_f700, ""), Nz(Rst!src_f700, ""), Nz(Rst!err_f700, ""), Worksheet_Name, RR, NumErrs)
If MaxColNum <= 700 Then Exit Sub

End Sub


Private Sub Merge_Duplicates(ETL123_table As String, Output_Table As String)
                       
'  This subroutine will merge duplicates in the "ETL123_table" based on the aDup_Key_Field from the Spec-ifications

Dim Rst As DAO.Recordset
Dim I As Long


'  Step 1 - Create a table to count the number occurances of the key values and then merge it into the ETL123_table
'           to populate the "row_count" field.
DelTbl ("ETL123_table_count")
strSql = "SELECT "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then
    strSql = strSql & "ETL123_table.F" & xlColNum(aExcel_Column_Number(I)) & ", "
  End If
Next I
If strSql = "SELECT " Then
  ErrorMsg = Output_Table & " import specification has no Key Fields marked.  Correct this and restart the import."
  MsgBox (ErrorMsg)
  DebugPrintOn (vbCrLf & "****" & ErrorMsg & vbCrLf)
  End
End If

strSql = Left(strSql, Len(strSql) - 2)
strSql = strSql & ", Count(ETL123_table.[F1]) AS countx INTO ETL123_table_count  FROM ETL123_table GROUP BY "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then
    strSql = strSql & "ETL123_table.F" & xlColNum(aExcel_Column_Number(I)) & ", "
  End If
Next I
strSql = Left(strSql, Len(strSql) - 2) & ";"
DebugPrint ("Step #1 - 'Merge_Duplicates' " & strSql)
DoCmd.RunSQL (strSql)

' Step 2 - Update the record counts in the ETL123_table
strSql = "UPDATE ETL123_table INNER JOIN ETL123_table_count ON "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then
    strSql = strSql & "(ETL123_table.F" & xlColNum(aExcel_Column_Number(I)) & _
        " = ETL123_table_count.F" & xlColNum(aExcel_Column_Number(I)) & ") AND "
  End If
Next I
strSql = Left(strSql, Len(strSql) - 4) & "SET ETL123_table.row_count = [ETL123_table_count].[countx];"
DebugPrint ("Step #2 - 'Merge_Duplicates' " & strSql)
DoCmd.RunSQL (strSql)

'Step 3 - Collect all of the records that have multiple keys that need to be merged.
strSql = "SELECT ETL123_table.* FROM ETL123_table WHERE (((ETL123_table.row_count) > 1)) " _
       & "ORDER BY ETL123_table.etl123_row_number;"
       
Dim holdKeySql As String
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then holdKeySql = holdKeySql & "[F" & xlColNum(aExcel_Column_Number(I)) & "] & ',' & "
Next I
If Right(holdKeySql, 9) = " & ',' & " Then holdKeySql = Left(holdKeySql, Len(holdKeySql) - 9)

strSql = "SELECT ETL123_table.*, [F1] & ',' & [F2] AS tab_key FROM ETL123_table WHERE (((ETL123_table.row_count) > 1)) ORDER BY [F1] & ',' & [F2], ETL123_table.etl123_row_number;"
DebugPrint ("Step #3 - 'Merge_Duplicates' " & strSql)
       
strSql = "SELECT ETL123_table.*, " & holdKeySql _
       & " AS tab_key FROM ETL123_table WHERE (((ETL123_table.row_count) > 1)) ORDER BY " _
       & holdKeySql & ", ETL123_table.etl123_row_number;"
DebugPrint ("Step #4 - 'Merge_Duplicates' " & strSql)
       
'''Dim holdData(1 To numOfSpecs) As String  '  This array will hold the contents of one row...
Dim lastKey As String, holdCount As Long, excelRowNumber As String
Dim LastRowNumber   As Double

Set Rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If Rst.RecordCount = 0 Then
   Rst.Close
   Set Rst = Nothing
   Exit Sub
End If
'  Now loop thru the query results and populate the array table
lastKey = "?"  ' Force Initial Break
Do
  If lastKey <> Rst!tab_key Then Call Merge_Break(Rst!tab_key, lastKey, holdCount, excelRowNumber, LastRowNumber)
    
  'Process an individual record.
  LastRowNumber = Nz(Rst!etl123_row_number, 0)
  Call merge_values(Rst, holdCount, excelRowNumber)
  
  Rst.MoveNext
  If Rst.EOF Then GoTo Finished_Do_Loop
Loop
Finished_Do_Loop:
Call Merge_Break("?", lastKey, holdCount, excelRowNumber, LastRowNumber)  ' Take the final break....

' Delete all of the original rows that have been merged.......
strSql = "DELETE ETL123_table.row_count FROM ETL123_table WHERE (((ETL123_table.row_count)>1));"
DebugPrint ("'Merge_Break' " & strSql)
DoCmd.RunSQL (strSql)


Rst.Close
Set Rst = Nothing

Exit Sub
DelTbl ("ETL123_table_count")


End Sub
Private Sub Merge_Break(thisKey As String, _
                lastKey As String, _
                holdCount As Long, _
                excelRowNumber As String, _
                LastRowNumber As Double)
                       
Dim I As Long, x As String

If lastKey = "?" Then GoTo First_Time
' Here we write out new merged row...

strSql = "INSERT INTO ETL123_table ( "
For I = 1 To (numOfSpecs)
  If holdMergedData(I) <> "#NULL#" Then strSql = strSql & "F" & I & ", "
Next I
strSql = strSql & "row_count, etl123_row_number) VALUES ( "
For I = 1 To (numOfSpecs - 1)
  If holdMergedData(I) <> "#NULL#" Then strSql = strSql & Scrub(holdMergedData(I)) & ", "
Next I
strSql = strSql & "'" & excelRowNumber & "', " & holdCount & ", " & LastRowNumber & ");"
DebugPrint ("'Merge_Break' " & strSql)
DoCmd.RunSQL (strSql)

First_Time:
If thisKey = "?" Then Exit Sub

' Now prepare for the next set of records.
holdCount = 0
excelRowNumber = ""  ' Clear Previous Values
lastKey = thisKey
For I = 1 To numOfSpecs
  holdMergedData(I) = ""  ' Clear old values....
Next I

End Sub

Private Sub merge_values(Rst As Recordset, holdCount As Long, excelRowNumber As String)
                       
Dim holdData() As String
ReDim holdData(1 To numOfSpecs) As String
Dim I As Long

If excelRowNumber = "" Then
  excelRowNumber = Nz(Rst!etl123_row_number, "")
 Else
  excelRowNumber = excelRowNumber & ", " & Nz(Rst!etl123_row_number, "")
End If

holdCount = Rst!row_count * -1

On Error GoTo Continue_Merge
holdData(1) = Nz(Rst!f1, "#NULL#")
holdData(2) = Nz(Rst!f2, "#NULL#")
holdData(3) = Nz(Rst!f3, "#NULL#")
holdData(4) = Nz(Rst!f4, "#NULL#")
holdData(5) = Nz(Rst!f5, "#NULL#")
holdData(6) = Nz(Rst!f6, "#NULL#")
holdData(7) = Nz(Rst!f7, "#NULL#")
holdData(8) = Nz(Rst!f8, "#NULL#")
holdData(9) = Nz(Rst!f9, "#NULL#")
holdData(10) = Nz(Rst!f10, "#NULL#")
holdData(11) = Nz(Rst!F11, "#NULL#")
holdData(12) = Nz(Rst!f12, "#NULL#")
holdData(13) = Nz(Rst!F13, "#NULL#")
holdData(14) = Nz(Rst!f14, "#NULL#")
holdData(15) = Nz(Rst!f15, "#NULL#")
holdData(16) = Nz(Rst!f16, "#NULL#")
holdData(17) = Nz(Rst!f17, "#NULL#")
holdData(18) = Nz(Rst!f18, "#NULL#")
holdData(19) = Nz(Rst!f19, "#NULL#")
holdData(20) = Nz(Rst!f20, "#NULL#")
holdData(21) = Nz(Rst!f21, "#NULL#")
holdData(22) = Nz(Rst!f22, "#NULL#")
holdData(23) = Nz(Rst!f23, "#NULL#")
holdData(24) = Nz(Rst!f24, "#NULL#")
holdData(25) = Nz(Rst!f25, "#NULL#")
holdData(26) = Nz(Rst!f26, "#NULL#")
holdData(27) = Nz(Rst!f27, "#NULL#")
holdData(28) = Nz(Rst!f28, "#NULL#")
holdData(29) = Nz(Rst!f29, "#NULL#")
holdData(30) = Nz(Rst!f30, "#NULL#")
holdData(31) = Nz(Rst!f31, "#NULL#")
holdData(32) = Nz(Rst!f32, "#NULL#")
holdData(33) = Nz(Rst!f33, "#NULL#")
holdData(34) = Nz(Rst!f34, "#NULL#")
holdData(35) = Nz(Rst!f35, "#NULL#")
holdData(36) = Nz(Rst!f36, "#NULL#")
holdData(37) = Nz(Rst!f37, "#NULL#")
holdData(38) = Nz(Rst!f38, "#NULL#")
holdData(39) = Nz(Rst!f39, "#NULL#")
holdData(40) = Nz(Rst!f40, "#NULL#")
holdData(41) = Nz(Rst!f41, "#NULL#")
holdData(42) = Nz(Rst!f42, "#NULL#")
holdData(43) = Nz(Rst!f43, "#NULL#")
holdData(44) = Nz(Rst!f44, "#NULL#")
holdData(45) = Nz(Rst!f45, "#NULL#")
holdData(46) = Nz(Rst!f46, "#NULL#")
holdData(47) = Nz(Rst!f47, "#NULL#")
holdData(48) = Nz(Rst!f48, "#NULL#")
holdData(49) = Nz(Rst!f49, "#NULL#")
holdData(50) = Nz(Rst!f50, "#NULL#")
holdData(51) = Nz(Rst!f51, "#NULL#")
holdData(52) = Nz(Rst!f52, "#NULL#")
holdData(53) = Nz(Rst!f53, "#NULL#")
holdData(54) = Nz(Rst!f54, "#NULL#")
holdData(55) = Nz(Rst!f55, "#NULL#")
holdData(56) = Nz(Rst!f56, "#NULL#")
holdData(57) = Nz(Rst!f57, "#NULL#")
holdData(58) = Nz(Rst!f58, "#NULL#")
holdData(59) = Nz(Rst!f59, "#NULL#")
holdData(60) = Nz(Rst!f60, "#NULL#")
holdData(61) = Nz(Rst!f61, "#NULL#")
holdData(62) = Nz(Rst!f62, "#NULL#")
holdData(63) = Nz(Rst!f63, "#NULL#")
holdData(64) = Nz(Rst!f64, "#NULL#")
holdData(65) = Nz(Rst!f65, "#NULL#")
holdData(66) = Nz(Rst!f66, "#NULL#")
holdData(67) = Nz(Rst!f67, "#NULL#")
holdData(68) = Nz(Rst!f68, "#NULL#")
holdData(69) = Nz(Rst!f69, "#NULL#")
holdData(70) = Nz(Rst!f70, "#NULL#")
holdData(71) = Nz(Rst!f71, "#NULL#")
holdData(72) = Nz(Rst!f72, "#NULL#")
holdData(73) = Nz(Rst!f73, "#NULL#")
holdData(74) = Nz(Rst!f74, "#NULL#")
holdData(75) = Nz(Rst!f75, "#NULL#")
holdData(76) = Nz(Rst!f76, "#NULL#")
holdData(77) = Nz(Rst!f77, "#NULL#")
holdData(78) = Nz(Rst!f78, "#NULL#")
holdData(79) = Nz(Rst!f79, "#NULL#")
holdData(80) = Nz(Rst!f80, "#NULL#")
holdData(81) = Nz(Rst!f81, "#NULL#")
holdData(82) = Nz(Rst!f82, "#NULL#")
holdData(83) = Nz(Rst!f83, "#NULL#")
holdData(84) = Nz(Rst!f84, "#NULL#")
holdData(85) = Nz(Rst!f85, "#NULL#")
holdData(86) = Nz(Rst!f86, "#NULL#")
holdData(87) = Nz(Rst!f87, "#NULL#")
holdData(88) = Nz(Rst!f88, "#NULL#")
holdData(89) = Nz(Rst!f89, "#NULL#")
holdData(90) = Nz(Rst!f90, "#NULL#")
holdData(91) = Nz(Rst!f91, "#NULL#")
holdData(92) = Nz(Rst!f92, "#NULL#")
holdData(93) = Nz(Rst!f93, "#NULL#")
holdData(94) = Nz(Rst!f94, "#NULL#")
holdData(95) = Nz(Rst!f95, "#NULL#")
holdData(96) = Nz(Rst!f96, "#NULL#")
holdData(97) = Nz(Rst!f97, "#NULL#")
holdData(98) = Nz(Rst!f98, "#NULL#")
holdData(99) = Nz(Rst!f99, "#NULL#")
holdData(100) = Nz(Rst!f100, "#NULL#")
holdData(101) = Nz(Rst!f101, "#NULL#")
holdData(102) = Nz(Rst!f102, "#NULL#")
holdData(103) = Nz(Rst!f103, "#NULL#")
holdData(104) = Nz(Rst!f104, "#NULL#")
holdData(105) = Nz(Rst!f105, "#NULL#")
holdData(106) = Nz(Rst!f106, "#NULL#")
holdData(107) = Nz(Rst!f107, "#NULL#")
holdData(108) = Nz(Rst!f108, "#NULL#")
holdData(109) = Nz(Rst!f109, "#NULL#")
holdData(110) = Nz(Rst!f110, "#NULL#")
holdData(111) = Nz(Rst!f111, "#NULL#")
holdData(112) = Nz(Rst!f112, "#NULL#")
holdData(113) = Nz(Rst!f113, "#NULL#")
holdData(114) = Nz(Rst!f114, "#NULL#")
holdData(115) = Nz(Rst!f115, "#NULL#")
holdData(116) = Nz(Rst!f116, "#NULL#")
holdData(117) = Nz(Rst!f117, "#NULL#")
holdData(118) = Nz(Rst!f118, "#NULL#")
holdData(119) = Nz(Rst!f119, "#NULL#")
holdData(120) = Nz(Rst!f120, "#NULL#")
holdData(121) = Nz(Rst!f121, "#NULL#")
holdData(122) = Nz(Rst!f122, "#NULL#")
holdData(123) = Nz(Rst!f123, "#NULL#")
holdData(124) = Nz(Rst!f124, "#NULL#")
holdData(125) = Nz(Rst!f125, "#NULL#")
holdData(126) = Nz(Rst!f126, "#NULL#")
holdData(127) = Nz(Rst!f127, "#NULL#")
holdData(128) = Nz(Rst!f128, "#NULL#")
holdData(129) = Nz(Rst!f129, "#NULL#")
holdData(130) = Nz(Rst!f130, "#NULL#")
holdData(131) = Nz(Rst!f131, "#NULL#")
holdData(132) = Nz(Rst!f132, "#NULL#")
holdData(133) = Nz(Rst!f133, "#NULL#")
holdData(134) = Nz(Rst!f134, "#NULL#")
holdData(135) = Nz(Rst!f135, "#NULL#")
holdData(136) = Nz(Rst!f136, "#NULL#")
holdData(137) = Nz(Rst!f137, "#NULL#")
holdData(138) = Nz(Rst!f138, "#NULL#")
holdData(139) = Nz(Rst!f139, "#NULL#")
holdData(140) = Nz(Rst!f140, "#NULL#")
holdData(141) = Nz(Rst!f141, "#NULL#")
holdData(142) = Nz(Rst!f142, "#NULL#")
holdData(143) = Nz(Rst!f143, "#NULL#")
holdData(144) = Nz(Rst!f144, "#NULL#")
holdData(145) = Nz(Rst!f145, "#NULL#")
holdData(146) = Nz(Rst!f146, "#NULL#")
holdData(147) = Nz(Rst!f147, "#NULL#")
holdData(148) = Nz(Rst!f148, "#NULL#")
holdData(149) = Nz(Rst!f149, "#NULL#")
holdData(150) = Nz(Rst!f150, "#NULL#")
holdData(151) = Nz(Rst!f151, "#NULL#")
holdData(152) = Nz(Rst!f152, "#NULL#")
holdData(153) = Nz(Rst!f153, "#NULL#")
holdData(154) = Nz(Rst!f154, "#NULL#")
holdData(155) = Nz(Rst!f155, "#NULL#")
holdData(156) = Nz(Rst!f156, "#NULL#")
holdData(157) = Nz(Rst!f157, "#NULL#")
holdData(158) = Nz(Rst!f158, "#NULL#")
holdData(159) = Nz(Rst!f159, "#NULL#")
holdData(160) = Nz(Rst!f160, "#NULL#")
holdData(161) = Nz(Rst!f161, "#NULL#")
holdData(162) = Nz(Rst!f162, "#NULL#")
holdData(163) = Nz(Rst!f163, "#NULL#")
holdData(164) = Nz(Rst!f164, "#NULL#")
holdData(165) = Nz(Rst!f165, "#NULL#")
holdData(166) = Nz(Rst!f166, "#NULL#")
holdData(167) = Nz(Rst!f167, "#NULL#")
holdData(168) = Nz(Rst!f168, "#NULL#")
holdData(169) = Nz(Rst!f169, "#NULL#")
holdData(170) = Nz(Rst!f170, "#NULL#")
holdData(171) = Nz(Rst!f171, "#NULL#")
holdData(172) = Nz(Rst!f172, "#NULL#")
holdData(173) = Nz(Rst!f173, "#NULL#")
holdData(174) = Nz(Rst!f174, "#NULL#")
holdData(175) = Nz(Rst!f175, "#NULL#")
holdData(176) = Nz(Rst!f176, "#NULL#")
holdData(177) = Nz(Rst!f177, "#NULL#")
holdData(178) = Nz(Rst!f178, "#NULL#")
holdData(179) = Nz(Rst!f179, "#NULL#")
holdData(180) = Nz(Rst!f180, "#NULL#")
holdData(181) = Nz(Rst!f181, "#NULL#")
holdData(182) = Nz(Rst!f182, "#NULL#")
holdData(183) = Nz(Rst!f183, "#NULL#")
holdData(184) = Nz(Rst!f184, "#NULL#")
holdData(185) = Nz(Rst!f185, "#NULL#")
holdData(186) = Nz(Rst!f186, "#NULL#")
holdData(187) = Nz(Rst!f187, "#NULL#")
holdData(188) = Nz(Rst!f188, "#NULL#")
holdData(189) = Nz(Rst!f189, "#NULL#")
holdData(190) = Nz(Rst!f190, "#NULL#")
holdData(191) = Nz(Rst!f191, "#NULL#")
holdData(192) = Nz(Rst!f192, "#NULL#")
holdData(193) = Nz(Rst!f193, "#NULL#")
holdData(194) = Nz(Rst!f194, "#NULL#")
holdData(195) = Nz(Rst!f195, "#NULL#")
holdData(196) = Nz(Rst!f196, "#NULL#")
holdData(197) = Nz(Rst!f197, "#NULL#")
holdData(198) = Nz(Rst!f198, "#NULL#")
holdData(199) = Nz(Rst!f199, "#NULL#")
holdData(200) = Nz(Rst!f200, "#NULL#")
holdData(201) = Nz(Rst!f201, "#NULL#")
holdData(202) = Nz(Rst!f202, "#NULL#")
holdData(203) = Nz(Rst!f203, "#NULL#")
holdData(204) = Nz(Rst!f204, "#NULL#")
holdData(205) = Nz(Rst!f205, "#NULL#")
holdData(206) = Nz(Rst!f206, "#NULL#")
holdData(207) = Nz(Rst!f207, "#NULL#")
holdData(208) = Nz(Rst!f208, "#NULL#")
holdData(209) = Nz(Rst!f209, "#NULL#")
holdData(210) = Nz(Rst!f210, "#NULL#")
holdData(211) = Nz(Rst!f211, "#NULL#")
holdData(212) = Nz(Rst!f212, "#NULL#")
holdData(213) = Nz(Rst!f213, "#NULL#")
holdData(214) = Nz(Rst!f214, "#NULL#")
holdData(215) = Nz(Rst!f215, "#NULL#")
holdData(216) = Nz(Rst!f216, "#NULL#")
holdData(217) = Nz(Rst!f217, "#NULL#")
holdData(218) = Nz(Rst!f218, "#NULL#")
holdData(219) = Nz(Rst!f219, "#NULL#")
holdData(220) = Nz(Rst!f220, "#NULL#")
holdData(221) = Nz(Rst!f221, "#NULL#")
holdData(222) = Nz(Rst!f222, "#NULL#")
holdData(223) = Nz(Rst!f223, "#NULL#")
holdData(224) = Nz(Rst!f224, "#NULL#")
holdData(225) = Nz(Rst!f225, "#NULL#")
holdData(226) = Nz(Rst!f226, "#NULL#")
holdData(227) = Nz(Rst!f227, "#NULL#")
holdData(228) = Nz(Rst!f228, "#NULL#")
holdData(229) = Nz(Rst!f229, "#NULL#")
holdData(230) = Nz(Rst!f230, "#NULL#")
holdData(231) = Nz(Rst!f231, "#NULL#")
holdData(232) = Nz(Rst!f232, "#NULL#")
holdData(233) = Nz(Rst!f233, "#NULL#")
holdData(234) = Nz(Rst!f234, "#NULL#")
holdData(235) = Nz(Rst!f235, "#NULL#")
holdData(236) = Nz(Rst!f236, "#NULL#")
holdData(237) = Nz(Rst!f237, "#NULL#")
holdData(238) = Nz(Rst!f238, "#NULL#")
holdData(239) = Nz(Rst!f239, "#NULL#")
holdData(240) = Nz(Rst!f240, "#NULL#")
holdData(241) = Nz(Rst!f241, "#NULL#")
holdData(242) = Nz(Rst!f242, "#NULL#")
holdData(243) = Nz(Rst!f243, "#NULL#")
holdData(244) = Nz(Rst!f244, "#NULL#")
holdData(245) = Nz(Rst!f245, "#NULL#")
holdData(246) = Nz(Rst!f246, "#NULL#")
holdData(247) = Nz(Rst!f247, "#NULL#")
holdData(248) = Nz(Rst!f248, "#NULL#")
holdData(249) = Nz(Rst!f249, "#NULL#")
holdData(250) = Nz(Rst!f250, "#NULL#")
holdData(251) = Nz(Rst!f251, "#NULL#")
holdData(252) = Nz(Rst!f252, "#NULL#")
holdData(253) = Nz(Rst!f253, "#NULL#")
holdData(254) = Nz(Rst!f254, "#NULL#")
holdData(255) = Nz(Rst!f255, "#NULL#")
holdData(256) = Nz(Rst!f256, "#NULL#")
holdData(257) = Nz(Rst!f257, "#NULL#")
holdData(258) = Nz(Rst!f258, "#NULL#")
holdData(259) = Nz(Rst!f259, "#NULL#")
holdData(260) = Nz(Rst!f260, "#NULL#")
holdData(261) = Nz(Rst!f261, "#NULL#")
holdData(262) = Nz(Rst!f262, "#NULL#")
holdData(263) = Nz(Rst!f263, "#NULL#")
holdData(264) = Nz(Rst!f264, "#NULL#")
holdData(265) = Nz(Rst!f265, "#NULL#")
holdData(266) = Nz(Rst!f266, "#NULL#")
holdData(267) = Nz(Rst!f267, "#NULL#")
holdData(268) = Nz(Rst!f268, "#NULL#")
holdData(269) = Nz(Rst!f269, "#NULL#")
holdData(270) = Nz(Rst!f270, "#NULL#")
holdData(271) = Nz(Rst!f271, "#NULL#")
holdData(272) = Nz(Rst!f272, "#NULL#")
holdData(273) = Nz(Rst!f273, "#NULL#")
holdData(274) = Nz(Rst!f274, "#NULL#")
holdData(275) = Nz(Rst!f275, "#NULL#")
holdData(276) = Nz(Rst!f276, "#NULL#")
holdData(277) = Nz(Rst!f277, "#NULL#")
holdData(278) = Nz(Rst!f278, "#NULL#")
holdData(279) = Nz(Rst!f279, "#NULL#")
holdData(280) = Nz(Rst!f280, "#NULL#")
holdData(281) = Nz(Rst!f281, "#NULL#")
holdData(282) = Nz(Rst!f282, "#NULL#")
holdData(283) = Nz(Rst!f283, "#NULL#")
holdData(284) = Nz(Rst!f284, "#NULL#")
holdData(285) = Nz(Rst!f285, "#NULL#")
holdData(286) = Nz(Rst!f286, "#NULL#")
holdData(287) = Nz(Rst!f287, "#NULL#")
holdData(288) = Nz(Rst!f288, "#NULL#")
holdData(289) = Nz(Rst!f289, "#NULL#")
holdData(290) = Nz(Rst!f290, "#NULL#")
holdData(291) = Nz(Rst!f291, "#NULL#")
holdData(292) = Nz(Rst!f292, "#NULL#")
holdData(293) = Nz(Rst!f293, "#NULL#")
holdData(294) = Nz(Rst!f294, "#NULL#")
holdData(295) = Nz(Rst!f295, "#NULL#")
holdData(296) = Nz(Rst!f296, "#NULL#")
holdData(297) = Nz(Rst!f297, "#NULL#")
holdData(298) = Nz(Rst!f298, "#NULL#")
holdData(299) = Nz(Rst!f299, "#NULL#")
holdData(300) = Nz(Rst!f300, "#NULL#")
holdData(301) = Nz(Rst!f301, "#NULL#")
holdData(302) = Nz(Rst!f302, "#NULL#")
holdData(303) = Nz(Rst!f303, "#NULL#")
holdData(304) = Nz(Rst!f304, "#NULL#")
holdData(305) = Nz(Rst!f305, "#NULL#")
holdData(306) = Nz(Rst!f306, "#NULL#")
holdData(307) = Nz(Rst!f307, "#NULL#")
holdData(308) = Nz(Rst!f308, "#NULL#")
holdData(309) = Nz(Rst!f309, "#NULL#")
holdData(310) = Nz(Rst!f310, "#NULL#")
holdData(311) = Nz(Rst!f311, "#NULL#")
holdData(312) = Nz(Rst!f312, "#NULL#")
holdData(313) = Nz(Rst!f313, "#NULL#")
holdData(314) = Nz(Rst!f314, "#NULL#")
holdData(315) = Nz(Rst!f315, "#NULL#")
holdData(316) = Nz(Rst!f316, "#NULL#")
holdData(317) = Nz(Rst!f317, "#NULL#")
holdData(318) = Nz(Rst!f318, "#NULL#")
holdData(319) = Nz(Rst!f319, "#NULL#")
holdData(320) = Nz(Rst!f320, "#NULL#")
holdData(321) = Nz(Rst!f321, "#NULL#")
holdData(322) = Nz(Rst!f322, "#NULL#")
holdData(323) = Nz(Rst!f323, "#NULL#")
holdData(324) = Nz(Rst!f324, "#NULL#")
holdData(325) = Nz(Rst!f325, "#NULL#")
holdData(326) = Nz(Rst!f326, "#NULL#")
holdData(327) = Nz(Rst!f327, "#NULL#")
holdData(328) = Nz(Rst!f328, "#NULL#")
holdData(329) = Nz(Rst!f329, "#NULL#")
holdData(330) = Nz(Rst!f330, "#NULL#")
holdData(331) = Nz(Rst!f331, "#NULL#")
holdData(332) = Nz(Rst!f332, "#NULL#")
holdData(333) = Nz(Rst!f333, "#NULL#")
holdData(334) = Nz(Rst!f334, "#NULL#")
holdData(335) = Nz(Rst!f335, "#NULL#")
holdData(336) = Nz(Rst!f336, "#NULL#")
holdData(337) = Nz(Rst!f337, "#NULL#")
holdData(338) = Nz(Rst!f338, "#NULL#")
holdData(339) = Nz(Rst!f339, "#NULL#")
holdData(340) = Nz(Rst!f340, "#NULL#")
holdData(341) = Nz(Rst!f341, "#NULL#")
holdData(342) = Nz(Rst!f342, "#NULL#")
holdData(343) = Nz(Rst!f343, "#NULL#")
holdData(344) = Nz(Rst!f344, "#NULL#")
holdData(345) = Nz(Rst!f345, "#NULL#")
holdData(346) = Nz(Rst!f346, "#NULL#")
holdData(347) = Nz(Rst!f347, "#NULL#")
holdData(348) = Nz(Rst!f348, "#NULL#")
holdData(349) = Nz(Rst!f349, "#NULL#")
holdData(350) = Nz(Rst!f350, "#NULL#")
holdData(351) = Nz(Rst!f351, "#NULL#")
holdData(352) = Nz(Rst!f352, "#NULL#")
holdData(353) = Nz(Rst!f353, "#NULL#")
holdData(354) = Nz(Rst!f354, "#NULL#")
holdData(355) = Nz(Rst!f355, "#NULL#")
holdData(356) = Nz(Rst!f356, "#NULL#")
holdData(357) = Nz(Rst!f357, "#NULL#")
holdData(358) = Nz(Rst!f358, "#NULL#")
holdData(359) = Nz(Rst!f359, "#NULL#")
holdData(360) = Nz(Rst!f360, "#NULL#")
holdData(361) = Nz(Rst!f361, "#NULL#")
holdData(362) = Nz(Rst!f362, "#NULL#")
holdData(363) = Nz(Rst!f363, "#NULL#")
holdData(364) = Nz(Rst!f364, "#NULL#")
holdData(365) = Nz(Rst!f365, "#NULL#")
holdData(366) = Nz(Rst!f366, "#NULL#")
holdData(367) = Nz(Rst!f367, "#NULL#")
holdData(368) = Nz(Rst!f368, "#NULL#")
holdData(369) = Nz(Rst!f369, "#NULL#")
holdData(370) = Nz(Rst!f370, "#NULL#")
holdData(371) = Nz(Rst!f371, "#NULL#")
holdData(372) = Nz(Rst!f372, "#NULL#")
holdData(373) = Nz(Rst!f373, "#NULL#")
holdData(374) = Nz(Rst!f374, "#NULL#")
holdData(375) = Nz(Rst!f375, "#NULL#")
holdData(376) = Nz(Rst!f376, "#NULL#")
holdData(377) = Nz(Rst!f377, "#NULL#")
holdData(378) = Nz(Rst!f378, "#NULL#")
holdData(379) = Nz(Rst!f379, "#NULL#")
holdData(380) = Nz(Rst!f380, "#NULL#")
holdData(381) = Nz(Rst!f381, "#NULL#")
holdData(382) = Nz(Rst!f382, "#NULL#")
holdData(383) = Nz(Rst!f383, "#NULL#")
holdData(384) = Nz(Rst!f384, "#NULL#")
holdData(385) = Nz(Rst!f385, "#NULL#")
holdData(386) = Nz(Rst!f386, "#NULL#")
holdData(387) = Nz(Rst!f387, "#NULL#")
holdData(388) = Nz(Rst!f388, "#NULL#")
holdData(389) = Nz(Rst!f389, "#NULL#")
holdData(390) = Nz(Rst!f390, "#NULL#")
holdData(391) = Nz(Rst!f391, "#NULL#")
holdData(392) = Nz(Rst!f392, "#NULL#")
holdData(393) = Nz(Rst!f393, "#NULL#")
holdData(394) = Nz(Rst!f394, "#NULL#")
holdData(395) = Nz(Rst!f395, "#NULL#")
holdData(396) = Nz(Rst!f396, "#NULL#")
holdData(397) = Nz(Rst!f397, "#NULL#")
holdData(398) = Nz(Rst!f398, "#NULL#")
holdData(399) = Nz(Rst!f399, "#NULL#")
holdData(400) = Nz(Rst!f400, "#NULL#")
holdData(401) = Nz(Rst!f401, "#NULL#")
holdData(402) = Nz(Rst!f402, "#NULL#")
holdData(403) = Nz(Rst!f403, "#NULL#")
holdData(404) = Nz(Rst!f404, "#NULL#")
holdData(405) = Nz(Rst!f405, "#NULL#")
holdData(406) = Nz(Rst!f406, "#NULL#")
holdData(407) = Nz(Rst!f407, "#NULL#")
holdData(408) = Nz(Rst!f408, "#NULL#")
holdData(409) = Nz(Rst!f409, "#NULL#")
holdData(410) = Nz(Rst!f410, "#NULL#")
holdData(411) = Nz(Rst!f411, "#NULL#")
holdData(412) = Nz(Rst!f412, "#NULL#")
holdData(413) = Nz(Rst!f413, "#NULL#")
holdData(414) = Nz(Rst!f414, "#NULL#")
holdData(415) = Nz(Rst!f415, "#NULL#")
holdData(416) = Nz(Rst!f416, "#NULL#")
holdData(417) = Nz(Rst!f417, "#NULL#")
holdData(418) = Nz(Rst!f418, "#NULL#")
holdData(419) = Nz(Rst!f419, "#NULL#")
holdData(420) = Nz(Rst!f420, "#NULL#")
holdData(421) = Nz(Rst!f421, "#NULL#")
holdData(422) = Nz(Rst!f422, "#NULL#")
holdData(423) = Nz(Rst!f423, "#NULL#")
holdData(424) = Nz(Rst!f424, "#NULL#")
holdData(425) = Nz(Rst!f425, "#NULL#")
holdData(426) = Nz(Rst!f426, "#NULL#")
holdData(427) = Nz(Rst!f427, "#NULL#")
holdData(428) = Nz(Rst!f428, "#NULL#")
holdData(429) = Nz(Rst!f429, "#NULL#")
holdData(430) = Nz(Rst!f430, "#NULL#")
holdData(431) = Nz(Rst!f431, "#NULL#")
holdData(432) = Nz(Rst!f432, "#NULL#")
holdData(433) = Nz(Rst!f433, "#NULL#")
holdData(434) = Nz(Rst!f434, "#NULL#")
holdData(435) = Nz(Rst!f435, "#NULL#")
holdData(436) = Nz(Rst!f436, "#NULL#")
holdData(437) = Nz(Rst!f437, "#NULL#")
holdData(438) = Nz(Rst!f438, "#NULL#")
holdData(439) = Nz(Rst!f439, "#NULL#")
holdData(440) = Nz(Rst!f440, "#NULL#")
holdData(441) = Nz(Rst!f441, "#NULL#")
holdData(442) = Nz(Rst!f442, "#NULL#")
holdData(443) = Nz(Rst!f443, "#NULL#")
holdData(444) = Nz(Rst!f444, "#NULL#")
holdData(445) = Nz(Rst!f445, "#NULL#")
holdData(446) = Nz(Rst!f446, "#NULL#")
holdData(447) = Nz(Rst!f447, "#NULL#")
holdData(448) = Nz(Rst!f448, "#NULL#")
holdData(449) = Nz(Rst!f449, "#NULL#")
holdData(450) = Nz(Rst!f450, "#NULL#")
holdData(451) = Nz(Rst!f451, "#NULL#")
holdData(452) = Nz(Rst!f452, "#NULL#")
holdData(453) = Nz(Rst!f453, "#NULL#")
holdData(454) = Nz(Rst!f454, "#NULL#")
holdData(455) = Nz(Rst!f455, "#NULL#")
holdData(456) = Nz(Rst!f456, "#NULL#")
holdData(457) = Nz(Rst!f457, "#NULL#")
holdData(458) = Nz(Rst!f458, "#NULL#")
holdData(459) = Nz(Rst!f459, "#NULL#")
holdData(460) = Nz(Rst!f460, "#NULL#")
holdData(461) = Nz(Rst!f461, "#NULL#")
holdData(462) = Nz(Rst!f462, "#NULL#")
holdData(463) = Nz(Rst!f463, "#NULL#")
holdData(464) = Nz(Rst!f464, "#NULL#")
holdData(465) = Nz(Rst!f465, "#NULL#")
holdData(466) = Nz(Rst!f466, "#NULL#")
holdData(467) = Nz(Rst!f467, "#NULL#")
holdData(468) = Nz(Rst!f468, "#NULL#")
holdData(469) = Nz(Rst!f469, "#NULL#")
holdData(470) = Nz(Rst!f470, "#NULL#")
holdData(471) = Nz(Rst!f471, "#NULL#")
holdData(472) = Nz(Rst!f472, "#NULL#")
holdData(473) = Nz(Rst!f473, "#NULL#")
holdData(474) = Nz(Rst!f474, "#NULL#")
holdData(475) = Nz(Rst!f475, "#NULL#")
holdData(476) = Nz(Rst!f476, "#NULL#")
holdData(477) = Nz(Rst!f477, "#NULL#")
holdData(478) = Nz(Rst!f478, "#NULL#")
holdData(479) = Nz(Rst!f479, "#NULL#")
holdData(480) = Nz(Rst!f480, "#NULL#")
holdData(481) = Nz(Rst!f481, "#NULL#")
holdData(482) = Nz(Rst!f482, "#NULL#")
holdData(483) = Nz(Rst!f483, "#NULL#")
holdData(484) = Nz(Rst!f484, "#NULL#")
holdData(485) = Nz(Rst!f485, "#NULL#")
holdData(486) = Nz(Rst!f486, "#NULL#")
holdData(487) = Nz(Rst!f487, "#NULL#")
holdData(488) = Nz(Rst!f488, "#NULL#")
holdData(489) = Nz(Rst!f489, "#NULL#")
holdData(490) = Nz(Rst!f490, "#NULL#")
holdData(491) = Nz(Rst!f491, "#NULL#")
holdData(492) = Nz(Rst!f492, "#NULL#")
holdData(493) = Nz(Rst!f493, "#NULL#")
holdData(494) = Nz(Rst!f494, "#NULL#")
holdData(495) = Nz(Rst!f495, "#NULL#")
holdData(496) = Nz(Rst!f496, "#NULL#")
holdData(497) = Nz(Rst!f497, "#NULL#")
holdData(498) = Nz(Rst!f498, "#NULL#")
holdData(499) = Nz(Rst!f499, "#NULL#")
holdData(500) = Nz(Rst!f500, "#NULL#")
holdData(501) = Nz(Rst!f501, "#NULL#")
holdData(502) = Nz(Rst!f502, "#NULL#")
holdData(503) = Nz(Rst!f503, "#NULL#")
holdData(504) = Nz(Rst!f504, "#NULL#")
holdData(505) = Nz(Rst!f505, "#NULL#")
holdData(506) = Nz(Rst!f506, "#NULL#")
holdData(507) = Nz(Rst!f507, "#NULL#")
holdData(508) = Nz(Rst!f508, "#NULL#")
holdData(509) = Nz(Rst!f509, "#NULL#")
holdData(510) = Nz(Rst!f510, "#NULL#")
holdData(511) = Nz(Rst!f511, "#NULL#")
holdData(512) = Nz(Rst!f512, "#NULL#")
holdData(513) = Nz(Rst!f513, "#NULL#")
holdData(514) = Nz(Rst!f514, "#NULL#")
holdData(515) = Nz(Rst!f515, "#NULL#")
holdData(516) = Nz(Rst!f516, "#NULL#")
holdData(517) = Nz(Rst!f517, "#NULL#")
holdData(518) = Nz(Rst!f518, "#NULL#")
holdData(519) = Nz(Rst!f519, "#NULL#")
holdData(520) = Nz(Rst!f520, "#NULL#")
holdData(521) = Nz(Rst!f521, "#NULL#")
holdData(522) = Nz(Rst!f522, "#NULL#")
holdData(523) = Nz(Rst!f523, "#NULL#")
holdData(524) = Nz(Rst!f524, "#NULL#")
holdData(525) = Nz(Rst!f525, "#NULL#")
holdData(526) = Nz(Rst!f526, "#NULL#")
holdData(527) = Nz(Rst!f527, "#NULL#")
holdData(528) = Nz(Rst!f528, "#NULL#")
holdData(529) = Nz(Rst!f529, "#NULL#")
holdData(530) = Nz(Rst!f530, "#NULL#")
holdData(531) = Nz(Rst!f531, "#NULL#")
holdData(532) = Nz(Rst!f532, "#NULL#")
holdData(533) = Nz(Rst!f533, "#NULL#")
holdData(534) = Nz(Rst!f534, "#NULL#")
holdData(535) = Nz(Rst!f535, "#NULL#")
holdData(536) = Nz(Rst!f536, "#NULL#")
holdData(537) = Nz(Rst!f537, "#NULL#")
holdData(538) = Nz(Rst!f538, "#NULL#")
holdData(539) = Nz(Rst!f539, "#NULL#")
holdData(540) = Nz(Rst!f540, "#NULL#")
holdData(541) = Nz(Rst!f541, "#NULL#")
holdData(542) = Nz(Rst!f542, "#NULL#")
holdData(543) = Nz(Rst!f543, "#NULL#")
holdData(544) = Nz(Rst!f544, "#NULL#")
holdData(545) = Nz(Rst!f545, "#NULL#")
holdData(546) = Nz(Rst!f546, "#NULL#")
holdData(547) = Nz(Rst!f547, "#NULL#")
holdData(548) = Nz(Rst!f548, "#NULL#")
holdData(549) = Nz(Rst!f549, "#NULL#")
holdData(550) = Nz(Rst!f550, "#NULL#")
holdData(551) = Nz(Rst!f551, "#NULL#")
holdData(552) = Nz(Rst!f552, "#NULL#")
holdData(553) = Nz(Rst!f553, "#NULL#")
holdData(554) = Nz(Rst!f554, "#NULL#")
holdData(555) = Nz(Rst!f555, "#NULL#")
holdData(556) = Nz(Rst!f556, "#NULL#")
holdData(557) = Nz(Rst!f557, "#NULL#")
holdData(558) = Nz(Rst!f558, "#NULL#")
holdData(559) = Nz(Rst!f559, "#NULL#")
holdData(560) = Nz(Rst!f560, "#NULL#")
holdData(561) = Nz(Rst!f561, "#NULL#")
holdData(562) = Nz(Rst!f562, "#NULL#")
holdData(563) = Nz(Rst!f563, "#NULL#")
holdData(564) = Nz(Rst!f564, "#NULL#")
holdData(565) = Nz(Rst!f565, "#NULL#")
holdData(566) = Nz(Rst!f566, "#NULL#")
holdData(567) = Nz(Rst!f567, "#NULL#")
holdData(568) = Nz(Rst!f568, "#NULL#")
holdData(569) = Nz(Rst!f569, "#NULL#")
holdData(570) = Nz(Rst!f570, "#NULL#")
holdData(571) = Nz(Rst!f571, "#NULL#")
holdData(572) = Nz(Rst!f572, "#NULL#")
holdData(573) = Nz(Rst!f573, "#NULL#")
holdData(574) = Nz(Rst!f574, "#NULL#")
holdData(575) = Nz(Rst!f575, "#NULL#")
holdData(576) = Nz(Rst!f576, "#NULL#")
holdData(577) = Nz(Rst!f577, "#NULL#")
holdData(578) = Nz(Rst!f578, "#NULL#")
holdData(579) = Nz(Rst!f579, "#NULL#")
holdData(580) = Nz(Rst!f580, "#NULL#")
holdData(581) = Nz(Rst!f581, "#NULL#")
holdData(582) = Nz(Rst!f582, "#NULL#")
holdData(583) = Nz(Rst!f583, "#NULL#")
holdData(584) = Nz(Rst!f584, "#NULL#")
holdData(585) = Nz(Rst!f585, "#NULL#")
holdData(586) = Nz(Rst!f586, "#NULL#")
holdData(587) = Nz(Rst!f587, "#NULL#")
holdData(588) = Nz(Rst!f588, "#NULL#")
holdData(589) = Nz(Rst!f589, "#NULL#")
holdData(590) = Nz(Rst!f590, "#NULL#")
holdData(591) = Nz(Rst!f591, "#NULL#")
holdData(592) = Nz(Rst!f592, "#NULL#")
holdData(593) = Nz(Rst!f593, "#NULL#")
holdData(594) = Nz(Rst!f594, "#NULL#")
holdData(595) = Nz(Rst!f595, "#NULL#")
holdData(596) = Nz(Rst!f596, "#NULL#")
holdData(597) = Nz(Rst!f597, "#NULL#")
holdData(598) = Nz(Rst!f598, "#NULL#")
holdData(599) = Nz(Rst!f599, "#NULL#")
holdData(600) = Nz(Rst!f600, "#NULL#")
holdData(601) = Nz(Rst!f601, "#NULL#")
holdData(602) = Nz(Rst!f602, "#NULL#")
holdData(603) = Nz(Rst!f603, "#NULL#")
holdData(604) = Nz(Rst!f604, "#NULL#")
holdData(605) = Nz(Rst!f605, "#NULL#")
holdData(606) = Nz(Rst!f606, "#NULL#")
holdData(607) = Nz(Rst!f607, "#NULL#")
holdData(608) = Nz(Rst!f608, "#NULL#")
holdData(609) = Nz(Rst!f609, "#NULL#")
holdData(610) = Nz(Rst!f610, "#NULL#")
holdData(611) = Nz(Rst!f611, "#NULL#")
holdData(612) = Nz(Rst!f612, "#NULL#")
holdData(613) = Nz(Rst!f613, "#NULL#")
holdData(614) = Nz(Rst!f614, "#NULL#")
holdData(615) = Nz(Rst!f615, "#NULL#")
holdData(616) = Nz(Rst!f616, "#NULL#")
holdData(617) = Nz(Rst!f617, "#NULL#")
holdData(618) = Nz(Rst!f618, "#NULL#")
holdData(619) = Nz(Rst!f619, "#NULL#")
holdData(620) = Nz(Rst!f620, "#NULL#")
holdData(621) = Nz(Rst!f621, "#NULL#")
holdData(622) = Nz(Rst!f622, "#NULL#")
holdData(623) = Nz(Rst!f623, "#NULL#")
holdData(624) = Nz(Rst!f624, "#NULL#")
holdData(625) = Nz(Rst!f625, "#NULL#")
holdData(626) = Nz(Rst!f626, "#NULL#")
holdData(627) = Nz(Rst!f627, "#NULL#")
holdData(628) = Nz(Rst!f628, "#NULL#")
holdData(629) = Nz(Rst!f629, "#NULL#")
holdData(630) = Nz(Rst!f630, "#NULL#")
holdData(631) = Nz(Rst!f631, "#NULL#")
holdData(632) = Nz(Rst!f632, "#NULL#")
holdData(633) = Nz(Rst!f633, "#NULL#")
holdData(634) = Nz(Rst!f634, "#NULL#")
holdData(635) = Nz(Rst!f635, "#NULL#")
holdData(636) = Nz(Rst!f636, "#NULL#")
holdData(637) = Nz(Rst!f637, "#NULL#")
holdData(638) = Nz(Rst!f638, "#NULL#")
holdData(639) = Nz(Rst!f639, "#NULL#")
holdData(640) = Nz(Rst!f640, "#NULL#")
holdData(641) = Nz(Rst!f641, "#NULL#")
holdData(642) = Nz(Rst!f642, "#NULL#")
holdData(643) = Nz(Rst!f643, "#NULL#")
holdData(644) = Nz(Rst!f644, "#NULL#")
holdData(645) = Nz(Rst!f645, "#NULL#")
holdData(646) = Nz(Rst!f646, "#NULL#")
holdData(647) = Nz(Rst!f647, "#NULL#")
holdData(648) = Nz(Rst!f648, "#NULL#")
holdData(649) = Nz(Rst!f649, "#NULL#")
holdData(650) = Nz(Rst!f650, "#NULL#")
holdData(651) = Nz(Rst!f651, "#NULL#")
holdData(652) = Nz(Rst!f652, "#NULL#")
holdData(653) = Nz(Rst!f653, "#NULL#")
holdData(654) = Nz(Rst!f654, "#NULL#")
holdData(655) = Nz(Rst!f655, "#NULL#")
holdData(656) = Nz(Rst!f656, "#NULL#")
holdData(657) = Nz(Rst!f657, "#NULL#")
holdData(658) = Nz(Rst!f658, "#NULL#")
holdData(659) = Nz(Rst!f659, "#NULL#")
holdData(660) = Nz(Rst!f660, "#NULL#")
holdData(661) = Nz(Rst!f661, "#NULL#")
holdData(662) = Nz(Rst!f662, "#NULL#")
holdData(663) = Nz(Rst!f663, "#NULL#")
holdData(664) = Nz(Rst!f664, "#NULL#")
holdData(665) = Nz(Rst!f665, "#NULL#")
holdData(666) = Nz(Rst!f666, "#NULL#")
holdData(667) = Nz(Rst!f667, "#NULL#")
holdData(668) = Nz(Rst!f668, "#NULL#")
holdData(669) = Nz(Rst!f669, "#NULL#")
holdData(670) = Nz(Rst!f670, "#NULL#")
holdData(671) = Nz(Rst!f671, "#NULL#")
holdData(672) = Nz(Rst!f672, "#NULL#")
holdData(673) = Nz(Rst!f673, "#NULL#")
holdData(674) = Nz(Rst!f674, "#NULL#")
holdData(675) = Nz(Rst!f675, "#NULL#")
holdData(676) = Nz(Rst!f676, "#NULL#")
holdData(677) = Nz(Rst!f677, "#NULL#")
holdData(678) = Nz(Rst!f678, "#NULL#")
holdData(679) = Nz(Rst!f679, "#NULL#")
holdData(680) = Nz(Rst!f680, "#NULL#")
holdData(681) = Nz(Rst!f681, "#NULL#")
holdData(682) = Nz(Rst!f682, "#NULL#")
holdData(683) = Nz(Rst!f683, "#NULL#")
holdData(684) = Nz(Rst!f684, "#NULL#")
holdData(685) = Nz(Rst!f685, "#NULL#")
holdData(686) = Nz(Rst!f686, "#NULL#")
holdData(687) = Nz(Rst!f687, "#NULL#")
holdData(688) = Nz(Rst!f688, "#NULL#")
holdData(689) = Nz(Rst!f689, "#NULL#")
holdData(690) = Nz(Rst!f690, "#NULL#")
holdData(691) = Nz(Rst!f691, "#NULL#")
holdData(692) = Nz(Rst!f692, "#NULL#")
holdData(693) = Nz(Rst!f693, "#NULL#")
holdData(694) = Nz(Rst!f694, "#NULL#")
holdData(695) = Nz(Rst!f695, "#NULL#")
holdData(696) = Nz(Rst!f696, "#NULL#")
holdData(697) = Nz(Rst!f697, "#NULL#")
holdData(698) = Nz(Rst!f698, "#NULL#")
holdData(699) = Nz(Rst!f699, "#NULL#")
holdData(700) = Nz(Rst!f700, "#NULL#")
Continue_Merge:
  On Error GoTo 0
  
For I = 1 To numOfSpecs  '  This loop will merge the values....
  If excelRowNumber = Rst!etl123_row_number Then _
     holdMergedData(I) = holdData(I)  ' Load the initial values.
          
  If Not aAllowChange2Blank(I) And (holdData(I) = "#NULL#" Or holdData(I) = "") Then _
    holdData(I) = holdMergedData(I) ' Don't allow blanks to survive the merge (except in the case of the first item.
     
  If aAccept_Changes(I) Then _
    holdMergedData(I) = holdData(I)  ' Load the changed values.
Next I
  
End Sub
Private Function FieldTypeEnum(ByVal aTable As String, _
                              ByVal aFieldName As String, _
                              Optional ByRef RetAlfaType As String = "") As Long

Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim fld As DAO.Field

Set db = CurrentDb()
Set tdf = db.TableDefs(aTable)

For Each fld In tdf.Fields
  If fld.Name = aFieldName Then
    FieldTypeEnum = fld.Type
    RetAlfaType = FieldTypeName(fld)
    Exit Function
  End If
Next
FieldTypeEnum = -1  ' Field not found

End Function


Private Function Field_Type(aTable As String, _
                           aFieldName As String, _
                           Optional ByRef FldType As Variant) As String
                       
Dim AlfaType  As String
FldType = FieldTypeEnum(aTable, aFieldName, AlfaType)

If FldType < 0 Then
  Field_Type = aFieldName & " was not found in Table-" & aTable
  Exit Function
End If

Select Case AlfaType
  Case "Big Integer": Field_Type = "Large Number/Big Integer"
  Case "Long Integer": Field_Type = "Number/Long Integer"
  Case "Text": Field_Type = "Short Text/Text"
  Case "Memo": Field_Type = "Long Text/Memo"
  Case Else: Field_Type = AlfaType
End Select

End Function




Private Function PrepXL(xlfile As String, wkSheetNum As Long)
                       
   Dim ETL123Prep As String
   'Dim xlApp As Excel.Application
   'Dim xlBk As Excel.Workbook
   'Dim xlSht As Excel.Worksheet
   Dim xlApp As Object
   Dim xlBk As Object
   Dim xlSht As Object
   Dim ChkXL As String
   Dim I   As Long, J As Long
   Dim LastCol As Long
      
   Dim HoldFilePath        As String
   I = InStrRev(xlfile, "\")
   HoldFilePath = Left(xlfile, I)
   
'   ETL123Prep = CurrentProject.Path & "\ETL123Prep." & holdXtension
   ETL123Prep = HoldFilePath & "ETL123Prep.xlsm"  ' A Macro Enabled Spreadsheet...
                                                  ' Should allow for processing all types of Excel.
   If Dir(ETL123Prep) <> "" Then
      KillFile (ETL123Prep)
   End If
  ' FileCopy xlfile, ETL123Prep
   DebugPrint ("xlfile=" & xlfile)
   DebugPrint ("ETL123Prep=" & ETL123Prep)
   Set xlApp = CreateObject("Excel.Application")
   Set xlBk = xlApp.Workbooks.Open(xlfile)
   Set xlSht = xlBk.Sheets(1)
   xlSht.Activate
   
   Dim lastCell    As String
   
   With xlApp
   .Sheets(wkSheetNum).Select
   .ActiveSheet.UnProtect
   .Cells.Select
   .Range("A1").Activate
   .Selection.EntireColumn.Hidden = False
   .Selection.EntireRow.Hidden = False
   .Selection.AutoFilter     '  Turn off filters.
   
   .ActiveCell.SpecialCells(xlLastCell).Select
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .ActiveCell.FormulaR1C1 = " "
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .ActiveCell.FormulaR1C1 = " "
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .ActiveCell.FormulaR1C1 = " "
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .ActiveCell.FormulaR1C1 = " "
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .Range("A1").Select
   .Selection.End(xlToRight).Select
   .ActiveCell.Offset(0, 1).Range("A1").Select
   .ActiveCell.FormulaR1C1 = "##Row_Number##"
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .ActiveCell.FormulaR1C1 = "2"
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .ActiveCell.FormulaR1C1 = "=R[-1]C+1"
   .ActiveCell.Select
   .Selection.Copy
   .Range(.Selection, .ActiveCell.SpecialCells(xlLastCell)).Select
   .ActiveSheet.Paste
   .Application.CutCopyMode = False
   .Range("A1").Select
   
   
  ' Now find column location of "Row_Number"
    LastCol = -1
    For I = 1 To 700
      .Range(xlColAlfa(I) & "1").Select
      If .ActiveCell.FormulaR1C1 = "##Row_Number##" Then
        .ActiveCell.FormulaR1C1 = "Row_Number"
        LastCol = I
        Exit For
      End If
    Next I

    If LastCol < 1 Then
      xlBk.SaveAs FileName:=ETL123Prep _
        , FileFormat:=52, CreateBackup:=False
      xlBk.Close
      Set xlBk = Nothing
      Set xlApp = Nothing
      Set xlSht = Nothing
      Call MsgBox("Error in PrepXL - Could not locate ""Row_Number"".  Process will ABORT.")
      End
    End If
    
   
   ' Insert a row and fill each cell with "X" to force definition of database field to be string.
    .Rows("1:1").Select
    .Selection.Insert Shift:=xlDown
  ' Write a row of all "X" cells.
    For I = 1 To LastCol
      .Range(xlColAlfa(I) & "1").Select
      .ActiveCell.FormulaR1C1 = "XXXXX"
    Next I
   
  End With
   
'    xlBk.Save
'    xlBk.SaveAs FileName:=ETL123Prep _
'        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    xlBk.SaveAs FileName:=ETL123Prep _
        , FileFormat:=52, CreateBackup:=False

    xlBk.Close
    Set xlBk = Nothing
    Set xlApp = Nothing
    Set xlSht = Nothing
    PrepXL = ETL123Prep  ' Return back with the Prep'd file
End Function



Private Sub Delete_Matching_From_Target(ETL123_table As String, Output_Table As String)
'  This subroutine will delete all of the Matching rows from the Target table to prepare of adding back to the target.

Dim tdf       As TableDef
Dim db As DAO.Database, Rst As Recordset
Dim I As Long

Set db = CurrentDb
Set tdf = db.TableDefs(Output_Table)
On Error Resume Next
tdf.Fields.Append tdf.CreateField("delete_matching", 10)     ' dbText
On Error GoTo 0

'Step 1 - Mark all of the matching rows in the Target Table that are matched.
'UPDATE bank_account INNER JOIN ETL123_table ON (bank_account.bank_account_id = ETL123_table.bank_account_id) SET bank_account.delete_matching = 'Y';
strSql = "UPDATE " & Output_Table & " INNER JOIN " & ETL123_table & " ON "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then _
    strSql = strSql & _
       "(" & ETL123_table & "." & Br(aField_Name_Output(I)) & " = " & Output_Table & "." & Br(aField_Name_Output(I)) & ") AND "
Next I
If Right(strSql, 5) = " AND " Then strSql = Left(strSql, Len(strSql) - 5)
strSql = strSql & " SET " & Output_Table & ".delete_matching = 'Y';"
DebugPrint ("(Delete_Matching_From_Target) Step #1 - " & strSql)
DoCmd.RunSQL (strSql)   '  Mark all matching rows...


'Step 2 - Save the Create Stamp values from the Target Table.
'UPDATE ETL123_table INNER JOIN bank_account ON ETL123_table.bank_account_id = bank_account.bank_account_id
'SET ETL123_table.create_date_time = [bank_account].[create_date_time], ETL123_table.create_user = [bank_account].[create_user], ETL123_table.create_program = [bank_account].[create_program], ETL123_table.create_file = [bank_account].[create_file];

If fieldexists("create_date_time", Output_Table) And _
   fieldexists("create_user", Output_Table) And _
   fieldexists("create_program", Output_Table) And _
   fieldexists("create_file", Output_Table) Then
  strSql = "UPDATE " & ETL123_table & " INNER JOIN " & Output_Table & " ON "
  For I = 1 To numOfSpecs
    If aDup_Key_Field(I) Then _
      strSql = strSql & _
         "(" & ETL123_table & "." & Br(aField_Name_Output(I)) & " = " & Output_Table & "." & Br(aField_Name_Output(I)) & ") AND "
  Next I
  If Right(strSql, 5) = " AND " Then strSql = Left(strSql, Len(strSql) - 5)
  strSql = strSql & " SET "
  strSql = strSql & ETL123_table & ".create_date_time = [" & Output_Table & "].[create_date_time], "
  strSql = strSql & ETL123_table & ".create_user = [" & Output_Table & "].[create_user], "
  strSql = strSql & ETL123_table & ".create_program = [" & Output_Table & "].[create_program], "
  strSql = strSql & ETL123_table & ".create_file = [" & Output_Table & "].[create_file];"
  DebugPrint ("(Delete_Matching_From_Target) Step #2 - " & strSql)
  DoCmd.RunSQL (strSql)   '  Save Create Date....
End If

'Step 3 - Actually delete the records from the target table. (Will be added back from ETL123_table)
strSql = "DELETE " & Output_Table & ".delete_matching FROM " & Output_Table & " WHERE ((" & Output_Table & ".delete_matching)='Y');"
DebugPrint ("(Delete_Matching_From_Target) Step #3 - " & strSql)
DoCmd.RunSQL (strSql)

'Step 4 - Actually delete the records from the target table. (Will be added back from ETL123_table)
strSql = "UPDATE " & ETL123_table & " SET " & ETL123_table & ".matched_target_table = 'N';"
DebugPrint ("(Delete_Matching_From_Target) Step #4 - " & strSql)
DoCmd.RunSQL (strSql)   '  Mark all rows...

tdf.Fields.Delete ("delete_matching") ' Remove the field column from the table.
Set tdf = Nothing
Set db = Nothing

End Sub



' Verify data found in the specification.
Private Function Specification_Is_Valid(Output_Table As String)
                       
Dim I As Long, J As Long
Dim ErrorMsg  As String

Specification_Is_Valid = True

' Check to see if all table names are valid for this table.
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And Not fieldexists(aField_Name_Output(I), Output_Table) Then
    ErrorMsg = "Field name-'" & aField_Name_Output(I) & "' for Excel-'" & aExcel_Heading_Text(I) & "' DOES NOT EXIST in Table-'" & _
            Output_Table & "' specification." & vbCrLf & vbCrLf & _
            "Import process will be terminated."
    MsgBox (ErrorMsg)
    DebugPrintOn (vbCrLf & "****" & ErrorMsg & vbCrLf)
    Specification_Is_Valid = False
  End If
Next I


' aActive_flag_name and aMark_active_flag work together.  If one is filled in then both are required.
If aActive_flag_name = "" And aMark_active_flag = "" Then GoTo Skip_Active_Validation
' Check to see if the active_flag_name is valid for this table.
If Not fieldexists(aActive_flag_name, Output_Table) Then
    ErrorMsg = "Field name-'" & aActive_flag_name & "' defining an Active Flag, DOES NOT EXIST in Table-'" & _
            Output_Table & "' specification." & vbCrLf & vbCrLf & _
            "A valid field name is required when mark_active_flag='" & aMark_active_flag & "' is specified. " & vbCrLf & vbCrLf & _
            "Import process will be terminated."
    MsgBox (ErrorMsg)
    Debug.Print (vbCrLf & "****" & ErrorMsg & vbCrLf)
    Specification_Is_Valid = False
End If

If (aMark_active_flag <> "Changed") And (aMark_active_flag <> "New") And (aMark_active_flag <> "Imported") Then
    ErrorMsg = "mark_active_flag='" & aMark_active_flag & "' is not valid value for Table-'" & _
            Output_Table & "' in import specification." & vbCrLf & vbCrLf & _
            "A valid mark_active_flag is required when active_flag_name-'" & aActive_flag_name & "' is filled in." & vbCrLf & vbCrLf & _
            "Import process will be terminated."
    MsgBox (ErrorMsg)
    DebugPrintOn (vbCrLf & "****" & ErrorMsg & vbCrLf)
    Specification_Is_Valid = False
End If

Skip_Active_Validation:

' Now check to make sure that key fields are valid.
J = 0
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then J = J + 1
Next I
If J = 0 Then
  ErrorMsg = Output_Table & " import specification has no Key Fields marked.  Correct this and restart the import." & vbCrLf & vbCrLf & _
            "Import process will be terminated."
  MsgBox (ErrorMsg)
  DebugPrintOn (vbCrLf & "****" & ErrorMsg & vbCrLf)
  Specification_Is_Valid = False
End If

For I = 1 To numOfSpecs
  If aDup_Key_Field(I) And aField_Name_Output(I) = "" Then
    ErrorMsg = "Key Field is marked for " & aExcel_Heading_Text(I) & " and Table-" & Output_Table & ", but no field name is given." & vbCrLf & vbCrLf & _
            "Correct this and restart the import." & vbCrLf & vbCrLf & _
            "Import process will be terminated."
    MsgBox (ErrorMsg)
    DebugPrintOn (vbCrLf & "****" & ErrorMsg & vbCrLf)
  Specification_Is_Valid = False
  End If
Next I

End Function


Private Function Find_Row_Number_Column(prepedFile As String, _
                                        wrkSheetNum As Long)
                       
Dim ExcelApp As Object
Dim ExcelBook As Object
Dim ExcelSheet As Object
Dim I As Long, J As Long, ColumnNumber
Dim holdColumnHeading    As String

Dim excelColumns() As String  '  Excel column letters array......
Call Init_Excel_Names(excelColumns)
   
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelBook = ExcelApp.Workbooks.Open(prepedFile)

Set ExcelSheet = ExcelBook.Sheets(wrkSheetNum)
ExcelSheet.Activate
For I = 1 To UBound(excelColumns)
   ColumnNumber = excelColumns(I) & "2"  ' Construct the cell id for proper row 2 heading....
   holdColumnHeading = ExcelSheet.Range(ColumnNumber).Value
   If holdColumnHeading = "Row_Number" Then
     Find_Row_Number_Column = I
    ' GoTo leaveFunction
   End If
Next I

leaveFunction:
    ExcelBook.Saved = True  ' Avoid the user message when closing Excel Workbook
    ExcelBook.Close
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing
    Set ExcelSheet = Nothing
    Exit Function

End Function


Private Function Br(sqlName As String)  ' Function used to insert brackets around SQL field names.
                       
Br = "[" & sqlName & "]"
End Function

Private Sub ClearLongTextFormatProp(TableName As String)
Dim db As DAO.Database
Dim tbl As TableDef
Dim fld As Field
Dim strName As String
Set db = CurrentDb
Set tbl = db.TableDefs(TableName)
    
Dim MsgBoxTitle:  MsgBoxTitle = "ClearLongTextFormatProp Error Handler"
Dim ErrorHasOccured:  ErrorHasOccured = False
On Error GoTo Error_Handler
    
Dim FldName As String, x
For Each fld In tbl.Fields
  FldName = fld.Name
  If CurrentDb.TableDefs(TableName).Fields(FldName).Type = 12 Then
    ' If Long Text then make sure that Format is cleared.
    x = CurrentDb.TableDefs(TableName).Fields(FldName).Properties!Format
    If x <> "None" Then
      CurrentDb.TableDefs(TableName).Fields(FldName).Properties.Delete ("Format")
    End If
  End If
Next
    
GoTo EndOfRoutine   '  Error Handling should always be the last thing in a routine.
Dim Errmsg, errResponse ' Call Err.Raise(????) The range 513 - 65535 is available for user errors.
Error_Handler:
  ErrorHasOccured = True
  Errmsg = "Error number: " & Str(err.Number) & vbNewLine & _
           "Source: " & err.source & vbNewLine & _
           "Description: " & err.Description & vbCrLf & vbCrLf
  Select Case err.Number
  Case 3270
    x = "None"
    Resume Next
  Case Else
    Errmsg = Errmsg & "No specific Handling.. " & vbCrLf & vbCrLf & _
                      "Abort will launch standard error handling. (Use to Debug)" & vbCrLf & _
                      "Retry will Try Again." & vbCrLf & _
                      "Ignore will END the process/program."
    errResponse = MsgBox(Errmsg, vbAbortRetryIgnore, MsgBoxTitle)
    ' 3 Abort, 4 Retry, 5 Ignore
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
    If errResponse = 3 Then
      On Error GoTo 0 ' Turn off error trap.
      Resume  ' Ignore To continue...
    End If
    If errResponse = 4 Then Resume  ' Retry....
    If errResponse = 5 Then
      DebugPrintOn ("Process Aborted by User")
      End     ' Ignore will end the process.
    End If
    DebugPrintOn ("Process Aborted by User")
    End
  End Select
Resume  ' Extra Resume for debug.  Locate source of the error..
EndOfRoutine:
    
    
End Sub

Private Function StrNormalize(ByVal Inp) As Variant

Dim II
If Not IsArray(Inp) Then
  StrNormalize = Inp
  Exit Function
End If

StrNormalize = ""
For II = LBound(Inp) To UBound(Inp) Step 2
  StrNormalize = StrNormalize & Chr(Inp(II))
Next II

End Function

Private Sub EditBoolFields()

' This routine will allow boolean values to come in as Y, N, Yes, No, T, True, F, or False values
Dim Fxx  As String, II: Fxx = "F13"

'xlColNum
For II = LBound(aData_Type) To UBound(aData_Type)
  If aData_Type(II) = "Yes/No" Then
    Fxx = "F" & xlColNum(aExcel_Column_Number(II))
    strSql = "UPDATE ETL123_table SET ETL123_table." & Fxx & " = ""True"" " _
           & "WHERE (((ETL123_table." & Fxx & ")=""Y"")) OR (((ETL123_table." & Fxx & ")=""Yes"")) OR (((ETL123_table." & Fxx & ")=""T""));"
    DebugPrint ("EditBoolFields - " & strSql)
    DoCmd.RunSQL (strSql)

    strSql = "UPDATE ETL123_table SET ETL123_table." & Fxx & " = ""False"" " _
           & "WHERE (((ETL123_table." & Fxx & ")=""N"")) OR (((ETL123_table." & Fxx & ")=""No"")) OR (((ETL123_table." & Fxx & ")=""F""));"
    DebugPrint ("EditBoolFields - " & strSql)
    DoCmd.RunSQL (strSql)
  End If
Next II

End Sub

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
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
    If Response = 4 Then Resume
    End
  Case Else
    Resume Next
    Resume  ' Extra Resume
  End Select
  
End Sub


Private Function DelTblS(tblName As String)

' This routine is similar to "DelTbl", but has support for WildCard "*" and can delete multiple tables.

Dim db As DAO.Database
Dim tdf As DAO.TableDef
Set db = CurrentDb
For Each tdf In db.TableDefs
    ' ignore system and temporary tables
    If Not (tdf.Name Like "MSys*" Or tdf.Name Like "~*") And (tdf.Name Like tblName) Then
        DebugPrint ("Deleting DB Table - " & tdf.Name)
        Call DelTbl(tdf.Name)
    End If
Next
Set tdf = Nothing
Set db = Nothing

End Function


Private Function Scrub(x As String)
'  Ensure that all apostrophe characters are translated out before passing to SQL
Dim I   As Long, ret As String, aChar  As String, aChr As String
For I = 1 To Len(x)
  aChar = Mid(x, I, 1)
  aChr = ""
  If aChar = "'" Then aChr = "chr(39)"
 ' If aChar = ":" Then aChr = "chr(158)"
  If aChr <> "" Then
    If Len(ret) = 0 Then
      ret = aChr & " & '"
    Else
    If Mid(ret, Len(ret), 1) = "'" Then
      ret = Left(ret, Len(ret) - 1) & aChr & " & '"
    Else
    ret = ret & "' & " & aChr & " & '"
    End If
    End If
  End If
  If aChr = "" Then
    If ret = "" Then
      ret = "'"
    End If
    ret = ret & aChar
  End If
Next I
If Right(ret, 4) = " & '" Then
  ret = Left(ret, Len(ret) - 4)
Else
  ret = ret & "'"
End If

Scrub = ret
End Function

Private Sub Remove_Table_Field(fieldname As String, TableName As String)

Dim tdf       As TableDef
Dim db As DAO.Database
Dim fld As DAO.Field
Dim prop As DAO.Property

If Not fieldexists(fieldname, TableName) Then Exit Sub

Set db = CurrentDb
Set tdf = db.TableDefs(TableName)

tdf.Fields.Delete (fieldname)

Set tdf = Nothing
Set tdf = Nothing

End Sub

' test if fieldName field exists in tableName table
Private Function fieldexists(fieldname As String, TableName As String) As Boolean
    Dim db As DAO.Database
    Dim tbl As TableDef
    Dim fld As Field
    Dim strName As String
    Set db = CurrentDb
    Set tbl = db.TableDefs(TableName)
    fieldexists = False
    For Each fld In tbl.Fields
        If fld.Name = fieldname Then
            fieldexists = True
            Exit For
        End If
    Next
End Function
Private Function TableExists(tblName As String)

' This routine is similar to "DelTbl", but has support for WildCard "*" and can delete multiple tables.

Dim db As DAO.Database
Dim tdf As DAO.TableDef
Set db = CurrentDb
TableExists = False
For Each tdf In db.TableDefs
  If tdf.Name = tblName Then
    TableExists = True
    Exit Function
  End If
Next
Set tdf = Nothing
Set db = Nothing

End Function


Private Sub Create_A_Table(theTableName As String)
 
' Create a table with an AutoNumber Id primary key so that new fields can be added.  (Delete existing table first)
Dim strSql As String
DelTbl (theTableName)
strSql = "CREATE TABLE " & theTableName & " (id COUNTER PRIMARY KEY);"
'DebugPrint (strSql)
DoCmd.RunSQL (strSql)
 
End Sub

Private Function RunFile(strFile As String, strWndStyle As String)
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


Private Function getFileNames2(ByVal aPath As String, ByVal aExt As String) As Variant
'?getFileNames("G:\My Drive\Joel's Files\qbo files")
'Create an array of all files in folder with a given extension and return list of values.

If Left(aExt, 1) <> "." Then aExt = "." & aExt
If Right(aPath, 1) <> "\" Then aPath = aPath & "\"

Dim strFileName As String
'TODO: Specify path and file spec

'Dim strFolder As String: strFolder = "G:\My Drive\Joel's Files\qbo files\"
Dim strFolder As String: strFolder = aPath
If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"  ' Make sure that path ends in a \

Dim strFileSpec As String: strFileSpec = strFolder & "*" & aExt
Dim FileList() As Variant
Dim intFoundFiles As Long
strFileName = Dir(strFileSpec)
Do While Len(strFileName) > 0
    ReDim Preserve FileList(intFoundFiles)
    FileList(intFoundFiles) = strFileName
   'DebugPrint (FileList(intFoundFiles))
    intFoundFiles = intFoundFiles + 1
    strFileName = Dir
Loop
getFileNames2 = FileList()
End Function
Private Function getFolderNames(ByVal aPath As String) As Variant
'?getFileNames("G:\My Drive\Joel's Files\qbo files")
'Create an array of all folders in a folder and return array list of path values.

If Right(aPath, 1) <> "\" Then aPath = aPath & "\"

Dim FS As New FileSystemObject
Dim FSfolder As Folder
Dim subfolder As Folder
Dim FileList() As Variant
Dim intFoundFiles As Long

Set FSfolder = FS.GetFolder(aPath)
    
For Each subfolder In FSfolder.SubFolders
  DoEvents
  'DebugPrint (subfolder)
  ReDim Preserve FileList(intFoundFiles)
  FileList(intFoundFiles) = subfolder
  intFoundFiles = intFoundFiles + 1
Next
    
getFolderNames = FileList()
Set FSfolder = Nothing
End Function


Private Function replacex()
Dim str1  As String
str1 = "One fish, two fish, red fish, blue fish"
DebugPrint (str1)
str1 = Replace(str1, "fish", "cat")

DebugPrint (str1)
End Function

Private Function strCount(ByVal XX As String, ByVal findStr As String) As Long

Dim I         As Long
Dim xStart    As Long: xStart = 1

I = InStr(xStart, XX, findStr)
Do While (I <> 0)
  strCount = strCount + 1
  xStart = I + 1
  I = InStr(xStart, XX, findStr)
Loop
End Function

Private Function GetFileNames(ByVal pFileName As String) As Variant
 
'Create an array of all files in folder within a given pattern............
 
Dim I As Long
I = InStrRev(pFileName, "\")
If I = 0 Then
  ErrorMsg = "Invalid Path name """ & pFileName & """  passed to getFileNames." & _
              vbCrLf & vbCrLf & "Process will abort..........."
  Call MsgBox(ErrorMsg, vbOKOnly, "getFileNames *******************")
  DebugPrint (vbCrLf & "****" & ErrorMsg & vbCrLf)
  End
End If

Dim aFolder As String
aFolder = Left(pFileName, I)

Dim FileList() As Variant
Dim intFoundFiles As Long
pFileName = Dir(pFileName)

Do While Len(pFileName) > 0
    ReDim Preserve FileList(intFoundFiles)
    FileList(intFoundFiles) = aFolder & pFileName
    intFoundFiles = intFoundFiles + 1
    pFileName = Dir
Loop
GetFileNames = FileList()
End Function

Private Function RenameFileOrDir(ByVal strSource As String, ByVal strTarget As String, _
  Optional fOverwriteTarget As Boolean = False) As Boolean
 
  On Error GoTo PROC_ERR
 
  Dim fRenameOK As Boolean
  Dim fRemoveTarget As Boolean
  Dim strFirstDrive As String
  Dim strSecondDrive As String
  Dim fOK As Boolean
 
  If Not ((Len(strSource) = 0) Or (Len(strTarget) = 0) Or (Not (FileOrDirExists(strSource)))) Then
 
    ' Check if the target exists
    If FileOrDirExists(strTarget) Then
 
      If fOverwriteTarget Then
        fRemoveTarget = True
      Else
        If MsgBox("Do you wish to overwrite the target file?", vbExclamation + vbYesNo, "Overwrite confirmation") = vbYes Then
          fRemoveTarget = True
        End If
      End If
 
      If fRemoveTarget Then
        ' Check that it's not a directory
        If ((GetAttr(strTarget) And vbDirectory)) <> vbDirectory Then
          Kill strTarget
          fRenameOK = True
        Else
          MsgBox "Cannot overwrite a directory", vbOKOnly, "Cannot perform operation"
        End If
      End If
    Else
      ' The target does not exist
      ' Check if source is a directory
      If ((GetAttr(strSource) And vbDirectory) = vbDirectory) Then
        ' Source is a directory, see if drives are the same
        strFirstDrive = Left(strSource, InStr(strSource, ":\"))
        strSecondDrive = Left(strTarget, InStr(strTarget, ":\"))
        If strFirstDrive = strSecondDrive Then
          fRenameOK = True
        Else
          MsgBox "Cannot rename directories across drives", vbOKOnly, "Cannot perform operation"
        End If
      Else
        ' It's a file, ok to proceed
        fRenameOK = True
      End If
    End If
 
    If fRenameOK Then
      Name strSource As strTarget
      fOK = True
    End If
  End If
 
  RenameFileOrDir = fOK
 
PROC_EXIT:
  Exit Function
 
PROC_ERR:
  MsgBox "Error: " & err.Number & ". " & err.Description, , "RenameFileOrDir"
  Resume PROC_EXIT
  Resume ' Extra Resume
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

Private Sub MakeMyFolder(FolderAndPath As String)
'Updateby Extendoffice 20161109
    Dim fdObj As Object

   ' Application.ScreenUpdating = False
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    If Not fdObj.FolderExists(FolderAndPath) Then
        fdObj.CreateFolder (FolderAndPath)
        MsgBox "Git folder-" & FolderAndPath & " has been created.", vbInformation
    End If
   ' Application.ScreenUpdating = True
End Sub

Private Function Verify(ByVal aString, ByVal ValidChars)
' Similar to PLI Verify Function
' Evaluate left to right that all char's found in aString are found in table of ValidChars
' Returns the position of the first INVALID char
Dim I, x
Verify = 0
For I = 1 To Len(aString)
  If InStr(1, ValidChars, Mid(aString, I, 1)) = 0 Then
    Verify = I
    Exit Function
  End If
Next I
End Function
Private Function Translate(ByVal aString, Optional ByVal ReplChar As String = "", Optional ByVal OrigChar As String = "")
' Similar to PLI Translate Function
' Evaluate left to right replacing characters from the OrigChar list with cooresponding characters from ReplChar
' Returns translated string.
Dim I, J
Translate = aString
If OrigChar = "" Then Exit Function
' Ensure ReplChar has at least as many characters as OrigChar
Do
  If Len(ReplChar) >= Len(OrigChar) Then Exit Do
  ReplChar = ReplChar & " "  ' Pad with blanks....
Loop

For I = 1 To Len(aString)
  J = InStr(1, OrigChar, Mid(aString, I, 1))
  If J <> 0 Then _
    aString = Mid(aString, 1, I - 1) & Mid(ReplChar, J, 1) & Mid(aString, I + 1)
Next I
Translate = aString

End Function

Private Function LTrimChars(ByVal XX, Optional ByVal chars = "0")
LTrimChars = TrimChars(XX, chars, True, False)
End Function

Private Function RTrimChars(ByVal XX, Optional ByVal chars = " ")
RTrimChars = TrimChars(XX, chars, False, True)
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



Private Function ShortFilePath(ByVal FilePath) As String

ShortFilePath = Left(FilePath, 15) & "...." & Right(FilePath, 15)
If Len(FilePath) <= 30 Then ShortFilePath = FilePath

End Function

Private Function AutoNumberFieldName(TableName As String) As String

     Dim dbs As DAO.Database
     Dim tdf As DAO.TableDef
     Dim fld As DAO.Field

     Set dbs = CurrentDb
     Set tdf = dbs.TableDefs(TableName)
     For Each fld In tdf.Fields
         If fld.Attributes And dbAutoIncrField Then
             AutoNumberFieldName = fld.Name
             Exit Function
         End If
     Next fld
     
End Function





Private Function RemoveWhiteSpace(ByVal XX As String) As String

Dim I, J, L, aChr

RemoveWhiteSpace = XX

XX = RemoveComments(XX)
' First translate all special characters.  Escape chars (") and (') are not recognized.
XX = Replace(XX, Chr(9), " ")
XX = Replace(XX, Chr(10), " ")
XX = Replace(XX, Chr(11), " ")
XX = Replace(XX, Chr(12), " ")
XX = Replace(XX, Chr(13), " ")

'  Now Always insert at least one space after a comma....
'  Escape chars (") and (') are not recognized.
aChr = " "   ' Init with one space.
For I = 1 To Len(XX)
  aChr = aChr & Mid(XX, I, 1)
  If Mid(XX, I, 1) = "," Then aChr = aChr & " " ' Always add 1 space after a comma.
Next
XX = Trim(aChr)

'  Turn double spaces into a single space..
Do
  RemoveWhiteSpace = xReplace(XX, "  ", " ")
  If RemoveWhiteSpace = XX Then Exit Do
  XX = RemoveWhiteSpace
Loop
XX = RemoveWhiteSpace

'  Remove spaces from in front of comma's
Do
  RemoveWhiteSpace = xReplace(XX, " ,", ",")
  If RemoveWhiteSpace = XX Then Exit Do
  XX = RemoveWhiteSpace
Loop
XX = RemoveWhiteSpace

End Function

Private Function RemoveComments(ByVal XX As String) As String

' This routine will return the input string with comments removed.
' Comments are specified as starting with  /* and ending with */
' Can also be single line comments beginning with // and ending with a CR - Carriage Return...
Dim II, JJ

RemoveComments = XX

' Remove multi-line comments....
Do
  II = InStr(1, XX, "/*")
  If II = 0 Then Exit Do
  JJ = InStr(II, XX, "*/")
  If JJ = 0 Then Exit Do
  XX = Mid(XX, 1, II - 1) & Mid(XX, JJ + 2)
  'DebugPrintOn ("xx=""" & xx & """ ")
  RemoveComments = XX
Loop

' Remove single-line comments....
Do
  II = InStr(1, XX, "//")
  If II = 0 Then Exit Do
  JJ = InStr(II, XX, vbCr)
  If JJ = 0 Then Exit Do
  XX = Mid(XX, 1, II - 1) & Mid(XX, JJ)  ' Leave the CR at the end of the line.
  RemoveComments = XX
Loop

End Function

Private Function xReplace(ByRef XX, ByVal xSearch, ByVal xRepl) As String

' This version of the replace will recognize double quote (") and single quote (') as escape chars
' and not replace the found string that is bracketed by either of these values.

Dim I, J, esc, xchr
esc = ""
xReplace = XX
If Len(xSearch) = Len(xRepl) And xSearch = xRepl Then Exit Function

J = 0
For I = 1 To Len(XX)
  xchr = Mid(XX, I, 1)
  If esc = "" Then
    If xchr = "'" Or xchr = """" Then
      esc = xchr
      GoTo NextForLoop
    End If
  End If
  If esc = xchr Then
    esc = ""
    GoTo NextForLoop
  End If
    
  If esc = "" And (Mid(XX, I, Len(xSearch)) = xSearch) Then
    J = I
    Exit For
  End If
NextForLoop:
Next I

If J <> 0 Then _
  xReplace = Mid(XX, 1, J - 1) & xRepl & Mid(XX, J + Len(xSearch))

End Function

Private Function ExtractKeyWord(ByRef XX, ByVal KeyW, Optional ByVal AbreviatedKeyW As String = "") As String

' This routine has an input string that will be searched for a Keyword=
' The result of the function is the actual data extracted from the keyword.  The data can be framed in ONE set
' of parentheses.  Nesting of parentheses is NOT supported.

' The input string is altered by removing the found keyword, so that multiple same keywords can be extracted
' White Space is removed (but not in strings framed with (), Double Quotes or Single Quotes.

' Ensure that the keyword ends in an "="
If Right(KeyW, 1) <> "=" Then KeyW = KeyW & "="


' Remove all White Space
XX = RemoveComments(XX)
XX = RemoveWhiteSpace(XX)

Dim II, JJ, LL, Data1st, DataLast, KW1st
ExtractKeyWord = ""

Dim DelmPreceeds  As Boolean

Do
    JJ = InStr(JJ + 1, XX, KeyW)
    If JJ = 0 Then Exit Do
    'DebugPrintOn (JJ & "  """ & Mid(xx, JJ) & """")
    DelmPreceeds = False
    If JJ = 1 Then DelmPreceeds = True
    If JJ > 1 Then
      If Verify(Mid(XX, JJ - 1, 1), "(),""' ") = 0 Then DelmPreceeds = True
    End If
    If JJ <> 0 And DelmPreceeds Then Exit Do
Loop
'DebugPrintOn ("len(xx)=" & Len(xx))
If JJ = 0 Then
  If AbreviatedKeyW = "" Then Exit Function
  KeyW = AbreviatedKeyW
  If Right(KeyW, 1) <> "=" Then KeyW = KeyW & "="
  Do
    JJ = InStr(JJ + 1, XX, KeyW)
    If JJ = 0 Then Exit Do
    'DebugPrintOn (JJ & "  """ & Mid(xx, JJ) & """")
    DelmPreceeds = False
    If JJ = 1 Then DelmPreceeds = True
    If JJ > 1 Then
      If Verify(Mid(XX, JJ - 1, 1), "(),""' ") = 0 Then DelmPreceeds = True
    End If
    If JJ <> 0 And DelmPreceeds Then Exit Do
  Loop
  If JJ = 0 Then Exit Function
End If

KW1st = JJ

Data1st = InStr(JJ, XX, "=") + 1
If Mid(XX, Data1st, 1) = "(" Then
  DataLast = InStr(Data1st, XX, ")")
  GoTo FoundTheEnd
End If
If Mid(XX, Data1st, 1) = "'" Then
  DataLast = InStr(Data1st + 1, XX, "'")
  GoTo FoundTheEnd
End If
If Mid(XX, Data1st, 1) = """" Then
  DataLast = InStr(Data1st + 1, XX, """")
  GoTo FoundTheEnd
End If

JJ = InStr(Data1st, XX, ",")
If JJ <> 0 Then
  DataLast = JJ
  GoTo FoundTheEnd
End If

JJ = InStr(Data1st, XX, " ")
If JJ <> 0 Then
  DataLast = JJ
  GoTo FoundTheEnd
End If

DataLast = Len(XX) ' Data is the last in the string, with no delimiter.

FoundTheEnd:
If DataLast = 0 Then Exit Function ' Proper ending was not found....

LL = (DataLast - Data1st) + 1
ExtractKeyWord = Mid(XX, Data1st, LL)
If Left(ExtractKeyWord, 1) = "(" Or _
   Left(ExtractKeyWord, 1) = """" Or _
   Left(ExtractKeyWord, 1) = "'" Then _
  ExtractKeyWord = Mid(ExtractKeyWord, 2, Len(ExtractKeyWord) - 2)
If Right(ExtractKeyWord, 1) = "," Then _
  ExtractKeyWord = Left(ExtractKeyWord, Len(ExtractKeyWord) - 1)
ExtractKeyWord = Trim(ExtractKeyWord)

'Remove keyword and data from the search string....
LL = (DataLast - KW1st) + 1
XX = Trim(Mid(XX, 1, KW1st - 1) & Mid(XX, DataLast + 1))

End Function

Private Function ExtractDelimited(ByRef XX) As String

' This routine has an input string that will be searched for a delimited value
' The result of the function is the actual data extracted from the Delimited String.

' The input string is altered by removing the found value.
' White Space is removed (but not in strings framed with (), Double Quotes or Single Quotes.

' Remove all White Space
XX = RemoveComments(XX)
XX = RemoveWhiteSpace(XX)

Dim II, JJ, LL, Data1st, DataLast, KW1st
ExtractDelimited = ""

Data1st = 1
If Mid(XX, Data1st, 1) = "'" Then
  DataLast = InStr(Data1st + 1, XX, "'")
  GoTo FoundTheEnd
End If
If Mid(XX, Data1st, 1) = """" Then
  DataLast = InStr(Data1st + 1, XX, """")
  GoTo FoundTheEnd
End If

JJ = InStr(Data1st, XX, ",")
If JJ <> 0 Then
  DataLast = JJ
  GoTo FoundTheEnd
End If

JJ = InStr(Data1st, XX, " ")
If JJ <> 0 Then
  DataLast = JJ
  GoTo FoundTheEnd
End If

DataLast = Len(XX) ' Data is the last in the string, with no delimiter.

FoundTheEnd:
If DataLast = 0 Then Exit Function ' Proper ending was not found....

LL = (DataLast - Data1st) + 1
ExtractDelimited = Mid(XX, Data1st, LL)
If Left(ExtractDelimited, 1) = """" Or _
   Left(ExtractDelimited, 1) = "'" Then _
  ExtractDelimited = Mid(ExtractDelimited, 2, Len(ExtractDelimited) - 2)
If Right(ExtractDelimited, 1) = "," Then _
  ExtractDelimited = Left(ExtractDelimited, Len(ExtractDelimited) - 1)
ExtractDelimited = Trim(ExtractDelimited)

'Remove data from the input string....
XX = Mid(XX, DataLast + 1)

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

Private Function TemplateReadLineByLine()

' This example found in https://analystcave.com/vba-read-file-vba/
' Also good reference - https://bettersolutions.com/index.htm#
Dim Errmsg, Response, MsgBoxTitle:  MsgBoxTitle = "Standard Error_Handler:"
On Error GoTo Error_Handler

Dim FileName As String, textData As String, textRow As String, fileNo As Integer
FileName = "G:\My Drive\Joel's Files\VA Services Design\VA Services Data Repository\Normalized Data Tables\WorkingOnFiduciaryAccess\ImportSpecTest.txt"
fileNo = FreeFile 'Get first free file number
    
Open FileName For Input As #fileNo
Do While Not EOF(fileNo)
   Line Input #fileNo, textRow
   textData = textData & textRow
Loop
Close #fileNo

Error_Handler:
  Errmsg = "Error number: " & Str(err.Number) & vbNewLine & _
           "Source: " & err.source & vbNewLine & _
           "Description: " & err.Description & vbCrLf & vbCrLf
  Select Case err.Number
  Case 53
    Errmsg = Errmsg & FileName
    Response = MsgBox(Errmsg, vbRetryCancel, MsgBoxTitle)
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
    If Response = 4 Then Resume
    End
  Case Else
    Resume Next
    Resume ' Extra Resume
  End Select


End Function

Private Function MsgBox2(Prompt As String, _
    Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional Title As String = "Microsoft Excel", _
    Optional HelpFile As String, _
    Optional Context As Long) As VbMsgBoxResult
     '
     '****************************************************************************************
     '       Title       MsgBox2
     '       Target Application: any
     '       Function:   substitute for standard MsgBox; displays more text than the ~1024 character
     '                   limit of MsgBox.  Displays blocks of approx 900 characters (properly split
     '                   at blanks or line feeds or "returns" and adds some "special text" to suggest
     '                   that more text is coming for each block except the last.  Special text is
     '                   easily changed.
     '
     '                   An EndOfBlack separator is also supported.  If found, MsgBox2 will only
     '                   display the characters through the EndOfBlock separator.  This provides
     '                   complete control over how text is displayed.  The current separator is
     '                   "||".
     '       Limitations:  the optional values for MsgBox display, i.e., Buttons, Title, HelpFile,
     '                     and Context  are the same for each block of text displayed.
     '       Passed Values:  same arguement list and type as standard MsgBox
     '
     '****************************************************************************************
     '
     '
    Dim CurLocn         As Long
    Dim EndOfBlock      As String
    Dim EOBIndex        As Integer
    Dim EOBLen          As Integer
    Dim Index           As Integer
    Dim MaxLen          As Integer
    Dim OldIndex        As Integer
    Dim strMoreToCome   As String
    Dim strTemp         As String
    Dim ThisChar        As String
    Dim TotLen          As Integer
     
     '
     '           set procedure variable that control how/what procedure does:
     '
     '       EndOfBlock is the string variable containing the character or characters
     '           that denote the end of a block of text.  These characters are not displayed.
     '           Do not use a character or characters that might be used in normal text.
     '       MaxLen is the maximum number of characters to be displayed at one time.  The
     '           limit for MsgBox is approx 1024, but that depends on the particular chars
     '           in the prompt string.  900 is a safe number as long as the len(strMoreToCome)
     '           is reasonable.
     '       strMoreToCome is text displayed at the bottom of each block indicating that more
     '           text/data is coming.
     '
    EndOfBlock = "||"
    MaxLen = 900
    strMoreToCome = " ... press any button except CANCEL to see next block of text ... "
     
    EOBLen = Len(EndOfBlock)
    CurLocn = 0
    OldIndex = 1
    TotLen = 0
     
NextBlock:
     '
     '           test for special break and, if found, that it is not the last chars in Prompt
     '
    EOBIndex = InStr(1, Mid(Prompt, OldIndex, MaxLen), EndOfBlock)
    If EOBIndex > 0 And CurLocn < Len(Prompt) - 1 Then
        CurLocn = EOBIndex + OldIndex - 1
        strTemp = Mid(Prompt, OldIndex, CurLocn - OldIndex)
        TotLen = TotLen + Len(strTemp) + EOBLen
        OldIndex = CurLocn + EOBLen
        GoTo MidDisplay
    End If
     '
     '           no special break, handle as normal block
     '
    Index = OldIndex + MaxLen
     '
     '           test for last block
     '
    If Index > Len(Prompt) Then
        strTemp = Mid(Prompt, OldIndex, Len(Prompt) - OldIndex + 1)
LastDisplay:
        MsgBox2 = MsgBox(strTemp, Buttons, Title, HelpFile, Context)
        Exit Function
    End If
     '
     '           not last display; process block
     '
    CurLocn = Index
NextIndex:
    ThisChar = Mid(Prompt, CurLocn, 1)
    If ThisChar = " " Or _
    ThisChar = Chr(10) Or _
    ThisChar = Chr(13) Then
         '
         '           block break found
         '
        strTemp = Mid(Prompt, OldIndex, CurLocn - OldIndex + 1)
        TotLen = TotLen + Len(strTemp)
        OldIndex = CurLocn + 1
MidDisplay:
         '
         '           display current block of text appending string indicating that
         '           more text is to come.  Then test if user hit Cancel button or
         '           equivalent; if so, exit MsgBox2 without further processing
         '
        MsgBox2 = MsgBox(strTemp & vbCrLf & strMoreToCome, _
        Buttons, Title, HelpFile, Context)
        If MsgBox2 = vbCancel Then Exit Function
        GoTo NextBlock
    End If
    CurLocn = CurLocn - 1
    If CurLocn > OldIndex Then GoTo NextIndex
     '
     '           no blanks, CR's, LF's or special breaks found in previous block
     '           display these characters and move on
     '
    strTemp = Mid(Prompt, OldIndex, MaxLen)
    CurLocn = OldIndex + MaxLen
    TotLen = TotLen + Len(strTemp)
    OldIndex = CurLocn + 1
    GoTo MidDisplay
     
End Function
 
Private Sub MsgBox2_Test(TestNum)
     '
     '****************************************************************************************
     '       Title       MsgBox2_Test
     '       Target Application: any
     '       Function;   demos use of MsgBox2
     '       Limitations:    none
     '       Passed Values:  none
     '****************************************************************************************
     '
     '
    Dim I           As Long
    Dim Answer      As VbMsgBoxResult
    Dim strPrompt   As String
     
    Select Case TestNum
    Case Is = 1
        strPrompt = "Initial stuff ..." & vbCrLf & vbCrLf
        For I = 48 To 122
            strPrompt = strPrompt & String(25, Chr(I)) & vbCrLf
        Next I
        strPrompt = strPrompt & vbCrLf & "... final stuff"
        Answer = MsgBox2(strPrompt, vbYesNoCancel, "1st Demo of MsgBox2")
    Case Is = 2
        strPrompt = "Initial stuff ..." & vbCrLf & vbCrLf
        For I = 48 To 122
            strPrompt = strPrompt & String(25, Chr(I))
        Next I
        strPrompt = strPrompt & vbCrLf & "... final stuff"
        Answer = MsgBox2(strPrompt, vbYesNoCancel, "2nd Demo of MsgBox2")
    Case Is = 3, 4
        strPrompt = "MsgBox is one of the most useful VB/VBA functions and it would be unlikely " & _
        "to find a VB/VBA application that did not use MsgBox at least once.  Unfortunately " & _
        "MsgBox has several not-easily-solved limitations, e.g., text size, text font, " & _
        "colors, and amount of text.  The former are irritating, but probably not fatal.  " & _
        "The latter, i.e., the amount of text that can be easily displayed via the Prompt " & _
        "string, is non-trivial.  MsgBox limits the number of characters to ~ 1024 (the " & _
        "exact number depends on the actual characters displayed).  If the length of Prompt " & _
        "is greater, the remaining characters are not displayed.  This can be particularly " & _
        "annoying (and possibly disastrous) if the last few words clarify an important " & _
        "result or what options are available or what is expected of the user." & vbCrLf & vbCrLf & "||"
        strPrompt = strPrompt & _
        "An alternative to MsgBox is a custom UserForm.  This is a good solution if one " & _
        "wants to improve several of MsgBox's limitations, but may be overkill if just " & _
        "displaying more text is desired." & vbCrLf & vbCrLf & "||" & _
        "MsgBox2 eliminates this limit by breaking the Prompt string into displayed blocks " & _
        "of approx 900 characters each.  For each block except the last, MsgBox2 displays " & _
        "the block and adds a line feed and special text suggesting that 'more data' is " & _
        "coming.  The special text is defined by the appl developer.  The current text is " & _
        vbCrLf & "       ... press any button except CANCEL to see next block of text ..." & vbCrLf & _
        "Text blocks are broken at " & _
        "logical separators: blanks; line feeds; or 'returns'.  Thus a Prompt string of, " & _
        "say, 2000 characters would be displayed in 3 blocks, the first two of approximately " & _
        "900 characters (ending with CrLf and '? more ?') and a final block with " & _
        "approximately 200 characters.  Each display is tested for 'Cancel' and, if " & _
        "encountered, MsgBox2 exits with a functional value equal to vbCancel or 2 (the " & _
        "numerical value for vbCancel)" & vbCrLf & vbCrLf & "||" & _
        "MsgBox2 also supports an 'end-of-block' option.  If the end-of-block character " & _
        "sequence is encountered (see code for current setting), MsgBox2 will automatically " & _
        "display the current buffer regardless of length." & vbCrLf & vbCrLf & "||" & _
        "Although simple is concept and execution, MsgBox2 is a very handy and" & vbCrLf & _
        "useful function.   MsgBox2 can be used in any VBA application." & vbCrLf & _
        "The demo is Excel based."
        If TestNum = 3 Then Answer = MsgBox2(strPrompt, vbYesNoCancel, "3rd Demo of MsgBox2")
        If TestNum = 4 Then MsgBox2 strPrompt, vbYesNoCancel, "4th Demo of MsgBox2"
         
    Case Else
        MsgBox "Invalid case fo MsgBox2_Test", vbCritical
    End Select
    If TestNum < 4 Then MsgBox "MsgBox2 return = " & MsgBoxResult(Answer)
     
End Sub
 
Private Function MsgBoxResult(result As VbMsgBoxResult) As String
     '
     '****************************************************************************************
     '       Title       MsgBoxResult
     '       Target Application: any
     '       Function:   returns (as a string) the "vb constant" associated with a MsgBox result
     '       Limitations:    none
     '       Passed Values:
     '           Result  [input, type=vbMsgBoxResult] result or from call to MsgBox
     '****************************************************************************************
     '
     '
    Select Case result
    Case Is = 1
        MsgBoxResult = "vbOK"
    Case Is = 2
        MsgBoxResult = "vbCancel"
    Case Is = 3
        MsgBoxResult = "vbAbort"
    Case Is = 4
        MsgBoxResult = "vbRetry"
    Case Is = 5
        MsgBoxResult = "vbAbort"
    Case Is = 6
        MsgBoxResult = "vbYes"
    Case Is = 7
        MsgBoxResult = "vbNo"
    Case Else
        MsgBoxResult = "UNKNOWN"
    End Select
     
End Function


Private Function FileSaveTest()
Dim FileName As String

'FileName = FileSaveAs("My test data for file...", _
'                      "G:\My Drive\Joel's Files\Work\JohnsTestFile.txt")
'                      True, _
'                      "Johns*.txt", _     Pattern for file...
'                      "Box Title Desc", _
'                      "Johns*.txt", _
'  "G:\My Drive\Joel's Files\Work\Johns*.txt"
'  "G:\My Drive\Joel's Files\Workxx\JohnsTestData*.txt,JohnsTestData2*.txt,JohnsTestFile*.txt,",
Dim fsatest As String
Dim FsaTestName As String: FsaTestName = "G:\My Drive\Joel's Files\Work\MyTest\MySubFolder\Johns.txt"
fsatest = FileSave("More Test data.....   xx", FsaTestName, _
            "Box Title", True)
fsatest = FileSave("More Test data.....   xx", FsaTestName, _
            "Box Title", True)
fsatest = FileSave("More Test data.....   xx", FsaTestName, _
            "Box Title", True)
FsaTestName = RenameAddDateTime(FsaTestName)
End Function

Private Function RenameAddDateTime(ByVal FilePath As String, _
                           Optional ByVal UseThisDateTime As Date = 0) As String
'  Rename an existing file and add the date and time to the name of the file.
Dim NewFilePath As String
If UseThisDateTime = 0 Then UseThisDateTime = Now()
NewFilePath = AddDateTimeToFileName(FilePath, UseThisDateTime)

If Not FileOrDirExists(FilePath) Then
  Call MsgBox("RenameAddDateTime is attempting to rename:" & vbCrLf & _
              FilePath & vbCrLf & vbCrLf & _
              "The file cannot be found.  Program will ABORT.")
  End
End If

If Not RenameFileOrDir(FilePath, NewFilePath) Then
  Call MsgBox("RenameAddDateTime is attempting to rename:" & vbCrLf & _
              FilePath & vbCrLf & vbCrLf & _
              "The file cannot be RENAMED.  Program will ABORT.")
  End
End If

RenameAddDateTime = NewFilePath

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
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
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
DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
If errResponse = 3 Then
  On Error GoTo 0 ' Turn off error trap.
  Resume  ' Ignore To continue...
End If
If errResponse = 4 Then Resume  ' Retry....
If errResponse = 5 Then
  DebugPrintOn ("Process Aborted by User")
  End     ' Ignore will end the process.
End If
DebugPrintOn ("Process Aborted by User")
End
  
Resume  ' Extra Resume for debug.  Locate source of the error..
EndOfRoutine:


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
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
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
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
    If errResponse = 3 Then
      On Error GoTo 0 ' Turn off error trap.
      Resume  ' Ignore To continue...
    End If
    If errResponse = 4 Then Resume  ' Retry....
    If errResponse = 5 Then
      DebugPrintOn ("Process Aborted by User")
      End     ' Ignore will end the process.
    End If
    DebugPrintOn ("Process Aborted by User")
    End
  End Select
Resume  ' Extra Resume for debug.  Locate source of the error..
EndOfRoutine:
  

End Function

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
  DebugPrintOn ("Process Aborted by User")
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
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
    On Error GoTo 0  ' Turn off error handling.
    Resume  ' Retry....
  Case 53
    Errmsg = Errmsg & "My custom message...   Cancel will continue with next statement.."
    errResponse = MsgBox(Errmsg, vbRetryCancel, MsgBoxTitle)
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
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
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
    If errResponse = 3 Then
      On Error GoTo 0 ' Turn off error trap.
      Resume  ' Ignore To continue...
    End If
    If errResponse = 4 Then Resume  ' Retry....
    If errResponse = 5 Then
      DebugPrintOn ("Process Aborted by User")
      End     ' Ignore will end the process.
    End If
    DebugPrintOn ("Process Aborted by User")
    End
  End Select
Resume  ' Extra Resume for debug.  Locate source of the error..
EndOfRoutine:

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

Private Function Rpad(ByVal strInput As String, ByVal NewLen As Long, Optional ByVal PadChar As String = " ") As String
If Len(PadChar) > 1 Then PadChar = Left(PadChar, 1)
If PadChar = "" Then PadChar = " "
Rpad = strInput
Do
  If Len(Rpad) >= NewLen Then Exit Function
  Rpad = Rpad & PadChar
Loop
End Function

Private Function ExtractDelimitedStr(ByRef DelimitedStr As String, ByVal SearchVal) As String

' The input string is searched from the left to right, up to the search value.

Dim II

DelimitedStr = TrimChars(RemoveComments(RemoveWhiteSpace(DelimitedStr)), ",")

II = InStr(1, DelimitedStr, SearchVal)
If II = 0 Then
  ExtractDelimitedStr = TrimChars(DelimitedStr, ",")
  DelimitedStr = ""
  Exit Function
End If

If II = 1 Then II = InStr(2, DelimitedStr, SearchVal)
If II = 0 Then
  ExtractDelimitedStr = TrimChars(DelimitedStr, ",")
  DelimitedStr = ""
  Exit Function
End If

ExtractDelimitedStr = TrimChars(Mid(DelimitedStr, 1, II - 1), ",")
DelimitedStr = Mid(DelimitedStr, II)

End Function

Private Function TstWs()
Dim x  As String, z As String

x = GetEntireFile("G:\My Drive\Joel's Files\Work\WS.txt")
DebugPrintOn (x)
z = RemoveWhiteSpace(x)
DebugPrintOn ("-----------------------------------------------------------------------------------------")
DebugPrintOn (z)
End Function

Private Function GetExcelWkShtNames(ByVal ExcelPathAndFile As String, _
                                 Optional ByVal StripParms As Boolean = False) As Variant
'  aWorkSheetNames is the main output array of WkSht names.

Call GetPopulateExcelGbl(ExcelPathAndFile) ' Ensure that Global Keeper of Excel WkShts and Headings Populated.

Dim II, JJ, KK, ReturnArray() As String, MyWkShtElmt As String
For II = LBound(GblKeepExcelWkShtHeadings) To UBound(GblKeepExcelWkShtHeadings)
  JJ = InStr(1, GblKeepExcelWkShtHeadings(II), "~") + 1
  KK = InStr(JJ, GblKeepExcelWkShtHeadings(II), "~")
  If KK <> 0 Then
    KK = KK - JJ  ' Calculate the length.
    MyWkShtElmt = Mid(GblKeepExcelWkShtHeadings(II), JJ, KK)
  Else
    MyWkShtElmt = Mid(GblKeepExcelWkShtHeadings(II), JJ)
  End If
  Call InsertNewElementIntoArray(ReturnArray, MyWkShtElmt)
Next II

If StripParms Then
  For II = LBound(ReturnArray) To UBound(ReturnArray)
    JJ = InStr(1, ReturnArray(II), ":")
    If JJ = 0 Then JJ = InStr(1, ReturnArray(II), ";")
    If JJ > 1 Then
      ReturnArray(II) = Trim(Left(ReturnArray(II), JJ - 1))
    End If
  Next II
End If

GetExcelWkShtNames = ReturnArray

End Function


Private Function GetExcelHeadingText(ByVal ExcelPathAndFile As String, _
                                    ByVal WkShtName As String, _
                                    Optional ByVal StripParms As Boolean = False) As Variant

Call GetPopulateExcelGbl(ExcelPathAndFile) ' Ensure that Global Keeper of Excel WkShts and Headings Populated.

Dim II, JJ, KK, ReturnArray() As String, MyWkShtElmt As String, FoundThisWkSht: FoundThisWkSht = -1
Dim I

If WkShtName <> "" And Verify(WkShtName, "0123456789") = 0 Then FoundThisWkSht = WkShtName - 1
  
If FoundThisWkSht < 0 Then
  For II = LBound(GblKeepExcelWkShtHeadings) To UBound(GblKeepExcelWkShtHeadings)
    JJ = InStr(1, GblKeepExcelWkShtHeadings(II), "~") + 1
    KK = InStr(JJ, GblKeepExcelWkShtHeadings(II), "~")
    If KK <> 0 Then
      KK = KK - JJ  ' Calculate the length.
      MyWkShtElmt = Mid(GblKeepExcelWkShtHeadings(II), JJ, KK)
    Else
      MyWkShtElmt = Mid(GblKeepExcelWkShtHeadings(II), JJ)
    End If
    If MyWkShtElmt = WkShtName Then
      FoundThisWkSht = II
      Exit For
    End If
  Next II
End If

' Did not find an exact match on WkSht name.  Now check for a match after extra data is stripped.
If FoundThisWkSht < 0 Then
  For II = LBound(GblKeepExcelWkShtHeadings) To UBound(GblKeepExcelWkShtHeadings)
    JJ = InStr(1, GblKeepExcelWkShtHeadings(II), "~") + 1
    KK = InStr(JJ, GblKeepExcelWkShtHeadings(II), "~")
    If KK <> 0 Then
      KK = KK - JJ  ' Calculate the length.
      MyWkShtElmt = Mid(GblKeepExcelWkShtHeadings(II), JJ, KK)
    Else
      MyWkShtElmt = Mid(GblKeepExcelWkShtHeadings(II), JJ)
    End If
    
    I = InStr(1, MyWkShtElmt, ":")
    If I = 0 Then I = InStr(1, MyWkShtElmt, ";")
    If I <> 0 Then MyWkShtElmt = Trim(Left(MyWkShtElmt, I - 1))
    
    If MyWkShtElmt = WkShtName Then
      FoundThisWkSht = II
      Exit For
    End If
  Next II
End If

If FoundThisWkSht < LBound(GblKeepExcelWkShtHeadings) Or FoundThisWkSht > UBound(GblKeepExcelWkShtHeadings) Then _
  FoundThisWkSht = -1

If FoundThisWkSht < 0 Then
  Call MsgBox("GetExcelHeadingText could not locate the requested Excel WkSht-""" & WkShtName & """ for:" _
             & vbCrLf & ExcelPathAndFile & "" _
             & vbCrLf & vbCrLf & "Process will abort.")
  End
  GoTo ReturnResults
End If

Dim strHeadings As String
II = FoundThisWkSht
JJ = InStr(1, GblKeepExcelWkShtHeadings(II), "~") + 1
KK = InStr(JJ, GblKeepExcelWkShtHeadings(II), "~")
If KK <> 0 Then
  strHeadings = Mid(GblKeepExcelWkShtHeadings(II), KK + 1)
  ReturnArray = Split(strHeadings, "~")
 Else
  Call InsertNewElementIntoArray(ReturnArray, "")
End If

If StripParms Then
  For II = LBound(ReturnArray) To UBound(ReturnArray)
    JJ = InStr(1, ReturnArray(II), ":")
    If JJ = 0 Then JJ = InStr(1, ReturnArray(II), ";")
    If JJ > 1 Then
      ReturnArray(II) = Trim(Left(ReturnArray(II), JJ - 1))
    End If
  Next II
End If

ReturnResults:
GetExcelHeadingText = ReturnArray

End Function



Private Sub GetPopulateExcelGbl(ByVal ExcelPathAndFile As String)
'  aWorkSheetNames is the main output array of WkSht names.
Dim II

' If GblKeepExcelWkShtHeadings is already allocated for this Excel file then continue.
Dim ExcelFileNameAndTime: ExcelFileNameAndTime = ExcelPathAndFile & "*" & FileDateTime(ExcelPathAndFile)
If IsArrayAllocated(GblKeepExcelWkShtHeadings) Then
  II = InStr(1, GblKeepExcelWkShtHeadings(0), "~")
  If II <> 0 And Left(GblKeepExcelWkShtHeadings(0), II - 1) = ExcelFileNameAndTime Then Exit Sub
  Erase GblKeepExcelWkShtHeadings
End If

Dim MsgBoxTitle:  MsgBoxTitle = "GetExcelWkShtNames Template Error Handler"
Dim ErrorHasOccured:  ErrorHasOccured = False
Dim Errmsg, errResponse ' Call Err.Raise(????) The range 513 - 65535 is available for user errors.

On Error GoTo Error_Handler
 
Dim AppExcel As New Excel.Application
Dim Wkb As Workbook
Dim Wksh As Worksheet
Dim I As Long
Dim aElement As String

Dim Headings:  Headings = ""
Dim FoundLast
Dim CellContents: CellContents = ""

I = InStrRev(ExcelPathAndFile, ".")
If Mid(ExcelPathAndFile, I) Like ".xls*" Or Mid(ExcelPathAndFile, I) Like ".csv" Then
  Set Wkb = AppExcel.Workbooks.Open(ExcelPathAndFile, False, True)
  For Each Wksh In Wkb.Worksheets
    FoundLast = False
    Headings = ""
    ' First work backward to find first heading that contains data.
    For I = 700 To 1 Step -1
      CellContents = Trim(Wksh.Range(GetColumnHeadingName(I)).Value)
      If CellContents <> "" Then FoundLast = True
      If FoundLast Then
        If InStr(1, CellContents, "~") <> 0 Then
          MsgBox ("Cell(" & GetColumnHeadingName(I) & ")  Contents-""" & CellContents & """ contains an illegal ""~"" character." & vbCrLf & vbCrLf _
                & "Process will ABORT.")
          GoTo AbortQuit
        End If
        Headings = "~" & CellContents & Headings
      End If
    Next I
    If InStr(1, ExcelFileNameAndTime, "~") <> 0 Then
      MsgBox ("File-""" & ExcelFileNameAndTime & """ contains an illegal ""~"" character." & vbCrLf & vbCrLf _
            & "Process will ABORT.")
      GoTo AbortQuit
    End If
    If InStr(1, Wksh.Name, "~") <> 0 Then
      MsgBox ("WorkSheet-""" & Wksh.Name & """ contains an illegal ""~"" character." & vbCrLf & vbCrLf _
            & "Process will ABORT.")
      GoTo AbortQuit
    End If
    aElement = ExcelFileNameAndTime & "~" & Wksh.Name & Headings
    ''DebugPrintOn (aElement)
    Call InsertNewElementIntoArray(GblKeepExcelWkShtHeadings, aElement)
  Next Wksh
  
  Wkb.Saved = True
  Wkb.Close
  AppExcel.Quit
  Set Wkb = Nothing
  Set AppExcel = Nothing
  Exit Sub
End If

GoTo EndOfRoutine   '  Error Handling should always be the last thing in a routine.

AbortQuit:
  Wkb.Saved = True
  Wkb.Close
  AppExcel.Quit
  Set Wkb = Nothing
  Set AppExcel = Nothing
  End


Error_Handler:
  ErrorHasOccured = True
  Errmsg = "Error number: " & Str(err.Number) & vbNewLine & _
           "Source: " & err.source & vbNewLine & _
           "Description: " & err.Description & vbCrLf & vbCrLf
  Select Case err.Number
  Case 53
    Errmsg = Errmsg & "My custom message...   Cancel will continue with next statement.."
    errResponse = MsgBox(Errmsg, vbRetryCancel, MsgBoxTitle)
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
    If errResponse = 4 Then
      On Error GoTo 0  ' Turn off error handling.
      Resume  ' Retry....
    End If
    Resume Next  ' To continue...
  Case Else
    If Not IsNull(Wkb) = False Then
      Wkb.Saved = True
      Wkb.Close
      AppExcel.Quit
      Set Wkb = Nothing
      Set AppExcel = Nothing
    End If
    Errmsg = Errmsg & "No specific Handling.. " & vbCrLf & vbCrLf & _
                      "Abort will launch standard error handling. (Use to Debug)" & vbCrLf & _
                      "Retry will Try Again." & vbCrLf & _
                      "Ignore will Close Excel and END the process/program."
    errResponse = MsgBox(Errmsg, vbAbortRetryIgnore, MsgBoxTitle)
    ' 3 Abort, 4 Retry, 5 Ignore
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
    If errResponse = 3 Then
      On Error GoTo 0 ' Turn off error trap.
      Resume  ' Ignore To continue...
    End If
    If errResponse = 4 Then Resume  ' Retry....
    If errResponse = 5 Then
      DebugPrintOn ("Process Aborted by User")
      End     ' Ignore will end the process.
    End If
    DebugPrintOn ("Process Aborted by User")
    End
  End Select
Resume  ' Extra Resume for debug.  Locate source of the error..
EndOfRoutine:


End Sub

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



Private Function ElapsedTimeSince(ByVal TimerName As String, _
                             Optional ByVal BeginTimeValue As Date, _
                             Optional ByVal ResetBeginTime As Boolean = False, _
                             Optional ByRef ReturnElapsedSeconds As Long = 0, _
                             Optional ByRef ReturnBeginTime As Date) As String

Dim II, JJ, xNow As Date:  xNow = Now()

JJ = -1
If IsArrayAllocated(ElapsedTimerName) Then
  For II = LBound(ElapsedTimerName) To UBound(ElapsedTimerName)
    If TimerName = ElapsedTimerName(II) Then
      JJ = II
      Exit For
    End If
  Next II
End If
If JJ = -1 Then  ' Start a new timer.
  Call InsertNewElementIntoArray(ElapsedTimerName, TimerName)
  Call InsertNewElementIntoArray(ElapsedTimerBegin, xNow)
  JJ = UBound(ElapsedTimerBegin)
End If

ElapsedTimeSince = TimeDiff(ElapsedTimerBegin(JJ), xNow, ReturnElapsedSeconds)
If ResetBeginTime Then ElapsedTimerBegin(JJ) = xNow  ' Reset the timer....

End Function


Private Function TimeDiff(ByVal BeginTime As Date, _
                         ByVal EndTime As Date, _
                         Optional ByRef Seconds As Long, _
                         Optional ByRef Minutes As Long, _
                         Optional ByRef Hours As Long, _
                         Optional ByRef Days As Long) As String
                         
Dim HoldTime, HoldStart, HoldEnd, aSign As Long: aSign = 1

HoldStart = BeginTime: HoldEnd = EndTime
If BeginTime > EndTime Then  ' Check to see if result should be negative.
  aSign = -1
  HoldTime = HoldStart  ' Swap the values..
  HoldStart = HoldEnd
  HoldEnd = HoldTime
End If
                         
Dim diff As Date
diff = HoldEnd - HoldStart
TimeDiff = Format(diff, "hh:ss")

Days = CLng(Int(diff))
Hours = (Days * 24) + Hour(diff)
Minutes = (Hours * 60) + Minute(diff)
Seconds = (Minutes * 60) + Second(diff)
Days = Days * aSign
Hours = Hours * aSign
Minutes = Minutes * aSign
Seconds = Seconds * aSign

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
    DebugPrint (vbCrLf & "****" & Errmsg & vbCrLf)
    If errResponse = 3 Then
      On Error GoTo 0 ' Turn off error trap.
      Resume  ' Ignore To continue...
    End If
    If errResponse = 4 Then Resume  ' Retry....
    If errResponse = 5 Then
      DebugPrintOn ("Process Aborted by User")
      End     ' Ignore will end the process.
    End If
    DebugPrintOn ("Process Aborted by User")
    End
  End Select
Resume  ' Extra Resume for debug.  Locate source of the error..
EndOfRoutine:

End Function

Private Function xlColNum(ColumnName As String)
                       
' Input is the Alpha column number, output is the numeric equivilant
Dim alpha  As String
Dim I As Long, J As Long, JJ As Long
alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

I = InStr(alpha, Mid(ColumnName, 1, 1))
xlColNum = I
If Len(ColumnName) = 2 Then
  J = InStr(alpha, Mid(ColumnName, 2, 1))
  xlColNum = (I * 26) + J
End If
      
End Function

Private Function xlColAlfa(ColumnNum As Long)
                       
' Input to function is numeric excel column, and output is the Alpha equivilant.
Dim I As Long, J As Long, JJ As Long, alpha As String
alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

J = ColumnNum \ 26
I = ColumnNum Mod 26
If I = 0 Then
  I = I + 26
  J = J - 1
End If
xlColAlfa = Mid(alpha, I, 1)
If J > 0 Then xlColAlfa = Mid(alpha, J, 1) & xlColAlfa

End Function

Private Function FilePrePicker(Optional ByVal PrePickedFile As String = "", _
                              Optional ByVal FilePattern As String = "", _
                              Optional BoxTitle As String = "FilePrePicker") As String

' 2) Check prePickedFile to verify that it exists...
' 1) Sort out the file pattern to be used to verify picked file.
' ?) If PrePicked is not valid then go"*.xlsx, *.xls, *.csv" and manually pick file with FilePickex

Dim Response, II, HoldFilePath, HoldFileName, HoldExt

FilePrePicker = ""   ' Initial value.....
If PrePickedFile = "" Then GoTo PrePickedFileNotValid

Dim Ext() As String
If FilePattern <> "" Then
  If InStr(1, FilePattern, ",") Then
      Ext = Split(FilePattern, ",")
    Else
     Ext = Split(FilePattern)
  End If
End If

II = InStrRev(PrePickedFile, "\")
If II = 0 Then II = InStrRev(PrePickedFile, ":")
If II = 0 Then
  Response = MsgBox("Pre-Picked file does not have a valid PATH syntax:" & vbCrLf & vbCrLf & _
                    """" & PrePickedFile & """" & vbCrLf & vbCrLf & _
                    "Do want to pick the file manually? Program with END if you pick ""No"". ", vbYesNo, BoxTitle)
  PrePickedFile = ""
  If Response = 6 Then GoTo PrePickedFileNotValid
  End
End If
HoldFilePath = Left(PrePickedFile, II)
HoldFileName = Mid(PrePickedFile, II + 1)
Dim NameMatchesPattern  As Boolean:  NameMatchesPattern = False
If FilePattern = "" Then NameMatchesPattern = True
For II = LBound(Ext) To UBound(Ext)
  If HoldFileName Like Trim(Ext(II)) Then NameMatchesPattern = True
Next II

If Not NameMatchesPattern Then
  Response = MsgBox("Pre-Picked file does not match the file name pattern(s):" & vbCrLf & vbCrLf & _
                    """" & HoldFileName & """ does not match with """ & FilePattern & """" & vbCrLf & vbCrLf & _
                    "Do want to pick the file manually? Program with END if you pick ""No"". ", vbYesNo, BoxTitle)
  PrePickedFile = ""
  If Response = 6 Then GoTo PrePickedFileNotValid
  End
End If

If FileOrDirExists(PrePickedFile) Then
  FilePrePicker = PrePickedFile
  Exit Function
End If

If Not FileOrDirExists(PrePickedFile) Then
  Response = MsgBox("Pre-Picked file does not exist:" & vbCrLf & vbCrLf & _
                    """" & PrePickedFile & """" & vbCrLf & vbCrLf & _
                    "Do want to pick the file manually? Program with END if you pick ""No"". ", vbYesNo, BoxTitle)
  PrePickedFile = ""
  If Response = 6 Then GoTo PrePickedFileNotValid
  End
End If


PrePickedFileNotValid:
Dim PickedFiles() As Variant
If PrePickedFile = "" Then
  Call FilePicker(PickedFiles, BoxTitle, FilePattern, , False)
  If IsEmpty(PickedFiles) Or Not IsArrayAllocated(PickedFiles) Then
    FilePrePicker = ""
    Exit Function
  End If
  FilePrePicker = PickedFiles(0)
  Exit Function
End If

End Function

Private Sub FilePicker(ByRef PickedFilesArray() As Variant, _
                 Optional ByVal BoxTitle As String = "Files Picked by FilePicker", _
                 Optional ByVal Filters As String = "*.*", _
                 Optional ByVal FilterDesc As String = "Files", _
                 Optional ByVal AllowMultiples As Boolean = True, _
                 Optional ByVal StartingFolderPath As String = "")

' FilePattern can be multiples comma separated.  Example:  "*.xlsx, *.csv, *.txt"
                      
Dim myObj As FileDialog
Dim HoldShow, II
If GblLastFilePickerPath = "" Then GblLastFilePickerPath = CurrentProject.Path & "\"
If StartingFolderPath = "" Then StartingFolderPath = GblLastFilePickerPath

Erase PickedFilesArray
    
''''Set myObj = Application.FileDialog(msoFileDialogFilePicker)
    Set myObj = Application.FileDialog(3)
    myObj.InitialFileName = StartingFolderPath
    With myObj
        .AllowMultiSelect = AllowMultiples
        .Title = BoxTitle
       
        .Filters.Clear
        .Filters.Add FilterDesc, Filters
        .FilterIndex = 1
        HoldShow = .Show
        ' -1 when file selected.
        ' 0 when cancel is pressed.
        If HoldShow = 0 Then GoTo ExitSubroutine
        
        If HoldShow = -1 Then
          Dim vrtSelectedItem, SelectedFile: SelectedFile = ""
          For Each vrtSelectedItem In myObj.SelectedItems
             SelectedFile = vrtSelectedItem
             Call InsertNewElementIntoArray(PickedFilesArray, SelectedFile)
          Next vrtSelectedItem
        End If
        
    End With

If SelectedFile <> "" Then
  II = InStrRev(SelectedFile, "\")
  If II <> 0 Then GblLastFilePickerPath = Left(SelectedFile, II)
End If

ExitSubroutine:
  Set myObj = Nothing

End Sub

Private Sub FolderPicker(ByRef PickedFoldersArray() As Variant, _
                   Optional ByVal BoxTitle As String = "Files Picked by FolderPicker", _
                   Optional ByVal AllowMultiples As Boolean = True, _
                   Optional ByVal StartingFolderPath As String = "")

Dim myObj As FileDialog
Dim HoldShow, II
If GblLastFolderPickerPath = "" Then GblLastFolderPickerPath = CurrentProject.Path & "\"
If StartingFolderPath = "" Then StartingFolderPath = GblLastFolderPickerPath

Erase PickedFoldersArray
    
''''Set myObj = Application.FileDialog(msoFileDialogFolderPicker)
    Set myObj = Application.FileDialog(4)
    myObj.InitialFileName = StartingFolderPath
    With myObj
        .AllowMultiSelect = AllowMultiples
        .Title = BoxTitle
        .ButtonName = "Pick Folder"
       
        HoldShow = .Show
        ' -1 when file selected.
        ' 0 when cancel is pressed.
        If HoldShow = 0 Then GoTo ExitSubroutine
        
        If HoldShow = -1 Then
          Dim vrtSelectedItem, SelectedFolder: SelectedFolder = ""
          For Each vrtSelectedItem In myObj.SelectedItems
             SelectedFolder = vrtSelectedItem
             Call InsertNewElementIntoArray(PickedFoldersArray, SelectedFolder)
          Next vrtSelectedItem
        End If
        
    End With

If SelectedFolder <> "" Then
  II = InStrRev(SelectedFolder, "\")
  If II <> 0 Then GblLastFolderPickerPath = Left(SelectedFolder, II)
End If

ExitSubroutine:
  Set myObj = Nothing

End Sub

Private Function TableInfo(strTableName As String)
On Error GoTo TableInfoErr
   ' Purpose:   Display the field names, types, sizes and descriptions for a table.
   ' Argument:  Name of a table in the current daxabase.
   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   Dim fld As DAO.Field
   
   Set db = CurrentDb()
   Set tdf = db.TableDefs(strTableName)
   Debug.Print "FIELD NAME", "FIELD TYPE", "ENUM", "SIZE", "DESCRIPTION"
   Debug.Print "==========", "==========", "====", "====", "==========="

   For Each fld In tdf.Fields
      Debug.Print fld.Name,
      Debug.Print FieldTypeName(fld),
      Debug.Print fld.Type,
      Debug.Print fld.Size,
      Debug.Print GetDescrip(fld)
   Next
   Debug.Print "==========", "==========", "====", "====", "==========="

TableInfoExit:
   Set db = Nothing
   Exit Function

TableInfoErr:
   Select Case err
   Case 3265&  'Table name invalid
      MsgBox strTableName & " table doesn't exist"
   Case Else
      Debug.Print "TableInfo() Error " & err & ": " & Error
   End Select
   Resume TableInfoExit
End Function


Private Function GetDescrip(obj As Object) As String
    On Error Resume Next
    GetDescrip = obj.Properties("Description")
End Function

Private Function FieldTypeName(fld As DAO.Field) As String
    'Purpose: Converts the numeric results of DAO Field.Type to text.
    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn
End Function



Private Function Contains(ByVal aString, ByVal InValidChars) As Long
' Evaluate left to right that all char's found in aString are NOT found in table of InValidChars
' Returns the position of the first INVALID char
Dim I, x
Contains = 0
For I = 1 To Len(aString)
  If InStr(1, InValidChars, Mid(aString, I, 1)) <> 0 Then
    Contains = I   ' Pos of first invalid character.
    Exit Function
  End If
Next I
End Function

Private Function FrameQuotes(ByVal XX As String) As String

FrameQuotes = XX
If Contains(XX, " ,-:/\~!@#$%^&*'""") Then FrameQuotes = """" & XX & """"

End Function

'''''''''''''''  Beginning of Array Handling Routines.......

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

Private Function FindFirstDupValue(ByVal vArray As Variant, _
                                  Optional ByRef DupValueReturned As Variant, _
                                  Optional ByRef DupValuePosReturned As Long) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FindDupValue - JGray 10/31/2018
' This subroutine will find the first duplicate value in a single Dimension array.
' DupValueReturned will return back with the value of the first duplicate item.  Otherwise, its empty.
' DupValuePosReturned will return back position of the first duplicate item.  Otherwise, it's -1
' Returns True or False indicating that a duplicate value was found.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim II, JJ
FindFirstDupValue = False ' Default value
DupValueReturned = Empty    ' Initial value
DupValuePosReturned = -1  ' Initial values

If Not IsVariantArrayConsistent(vArray) Then Exit Function
If LBound(vArray) = UBound(vArray) Then Exit Function

For II = LBound(vArray) To UBound(vArray) - 1
  For JJ = II + 1 To UBound(vArray)
    If Trim(vArray(II)) = Trim(vArray(JJ)) Then
      FindFirstDupValue = True ' Found Duplicate.
      DupValueReturned = vArray(II) ' Found Value.
      DupValuePosReturned = II  ' Found Value.
      Exit Function
    End If
  Next JJ
Next II

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


Private Function FindPosition(ByVal vArray As Variant, ByVal aValue) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FindPosition - JGray 8/12/2018
' This subroutine will search for a duplicate value in a single Dimension array and return
' the POSITION of the FIRST value.
' If the value is NOT found then a (-1) will be returned.
' A (-2) will be returned if there is an error.
' Returns True or False indicating that a duplicate value was found.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim II
If Not IsVariantArrayConsistent(vArray) Then
  FindPosition = -2
  Exit Function
End If

For II = LBound(vArray) To UBound(vArray)
  If vArray(II) = aValue Then
    FindPosition = II
    Exit Function
  End If
Next II
FindPosition = -1
End Function

Private Function FindNonUniquePos(ByVal vArray As Variant, ByVal aValue) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FindNonUniquePos - JGray 8/12/2018
' This subroutine will search for a duplicate value in a single Dimension array and return
' the POSITION of the SECOND occurance of a duplicate value.
' If the value is unique and NOT duplicate then a (-1) will be returned.
' A (-2) will be returned if there is an error.
' Returns True or False indicating that a duplicate value was found.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim II, count As Long
If Not IsVariantArrayConsistent(vArray) Then  ' Not an array?
  FindNonUniquePos = -2
  Exit Function
End If

For II = LBound(vArray) To UBound(vArray)
  If vArray(II) = aValue Then count = count + 1
  If count = 2 Then
    FindNonUniquePos = II
    Exit Function
  End If
Next II
FindNonUniquePos = -1
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


Private Function IsVariantArrayConsistent(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsVariantArrayConsistent
'
' This returns TRUE or FALSE indicating whether an array of variants
' contains all the same data types. Returns FALSE under the following
' circumstances:
'       Arr is not an array
'       Arr is an array but is unallocated
'       Arr is a multidimensional array
'       Arr is allocated but does not contain consistant data types.
'
' If Arr is an array of objects, objects that are Nothing are ignored.
' As long as all non-Nothing objects are the same object type, the
' function returns True.
'
' It returns TRUE if all the elements of the array have the same
' data type. If Arr is an array of a specific data types, not variants,
' (E.g., Dim V(1 To 3) As Long), the function will return True. If
' an array of variants contains an uninitialized element (VarType =
' vbEmpty) that element is skipped and not used in the comparison. The
' reasoning behind this is that an empty variable will return the
' data type of the variable to which it is assigned (e.g., it will
' return vbNullString to a String and 0 to a Double).
'
' The function does not support arrays of User Defined Types.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim FirstDataType As VbVarType
Dim Ndx As Long
'''''''''''''''''''''''''''''''''''''''''
' Exit with False if Arr is not an array.
'''''''''''''''''''''''''''''''''''''''''
If IsArray(Arr) = False Then
    IsVariantArrayConsistent = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Exit with False if Arr is not allocated.
''''''''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(Arr) = False Then
    IsVariantArrayConsistent = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''
' Exit with false on multi-dimensional
' arrays.
''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr) <> 1 Then
    IsVariantArrayConsistent = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Test if we have an array of a specific
' type rather than Variants. If so,
' return TRUE and get out.
''''''''''''''''''''''''''''''''''''''''''
If (VarType(Arr) <= vbArray) And _
    (VarType(Arr) <> vbVariant) Then
    IsVariantArrayConsistent = True
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Get the data type of the first element.
''''''''''''''''''''''''''''''''''''''''''
FirstDataType = VarType(Arr(LBound(Arr)))
''''''''''''''''''''''''''''''''''''''''''
' Loop through the array and exit if
' a differing data type if found.
''''''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(Arr) + 1 To UBound(Arr)
    If VarType(Arr(Ndx)) <> vbEmpty Then
        If IsObject(Arr(Ndx)) = True Then
            If Not Arr(Ndx) Is Nothing Then
                If VarType(Arr(Ndx)) <> FirstDataType Then
                    IsVariantArrayConsistent = False
                    Exit Function
                End If
            End If
        Else
            If VarType(Arr(Ndx)) <> FirstDataType Then
                IsVariantArrayConsistent = False
                Exit Function
            End If
        End If
    End If
Next Ndx

''''''''''''''''''''''''''''''''''''''''''
' If we make it out of the loop,
' then the array is consistent.
''''''''''''''''''''''''''''''''''''''''''
IsVariantArrayConsistent = True

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


