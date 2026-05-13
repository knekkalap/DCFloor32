Attribute VB_Name = "Startup"
Option Explicit

Public Const debugging = 1

Public Const bristol = 1
Public Const monterrey = 2
Public Const syracuse = 3
Public Const corry = 7
Public Const saline = 10
Public Const milw = 12
Public Const burl = 17
Public Const remesa = 36
Public Const hyson_products = 76
Public Const stroms = 77
Public Const singapore = 30
Public Const pv_singapore = 60
Public Const Synventive = 99

Public gSQLConnection As ADODB.Connection

Public gError As CError

Public gErrLocation As String       'Path and File Name of Error File
Public gAttdLogFile As String
Public gLaborLogFile As String
Public gAttdLogLocal As String

 Public gDateformat As String
 Public gDecimal    As String       ' Decimal char
 Public gThousand   As String       ' Thousands char
 Public gDateMask   As String       ' Mask for input boxes
 Public gDateSep    As String

 Public gTimeFormat As String
 Public gOneDecPos As String
 Public gTwoDecPos As String
  
 Public gLanguageOffset As Integer
 
 Public Const CONTINUE = 1
 Public Const ABORT = 2
 
 Public div As Integer   'set in LoadApp()

Rem Flags for handling Division variations in Direct Labor
Public gbEOShift As Boolean        'inpEndOfShift Call made
Public gOKToComplete As Boolean
Public gbNumbMach As Boolean       '>=1 machines, Ask # machines
Public gbInpSetOrMach As Boolean   'Ask for Setup or Machine Hours
Public gbInpSetAndMach As Boolean  'Ask for both Setup & Machine Hours
Public gbAddonHrs As Boolean       'Ask for Addon Hours
Public gbZeroQty As Boolean        'Zero ActQty for some operations
Public gbCallCompl As Boolean      'Ask if complete
Public gbNoLoads As Boolean        'Ask # loads, Cycle time & Batch Incentive
Public gbAnotherRec As Boolean     '>=1 machines, Ask if more
Public gbMachRecY As Boolean       'Set first MachRec = "Y"
Public gbInpStds As Boolean        'Ask for Standards in inpJobAndSeq
Public gbRecTypeR As Boolean       'Set Rec Type to R for Rework
Public gbScrapPounds As Boolean
Public gbInpIndirHrs As Boolean
'Ask for Lbs or Pieces for Scrap Entry

Public strconnect As String

'Setup for Entry Modules
Public gDone As Boolean             'Escape up to Main Menu flag
Public Const gEntryAll = 0          'Entry of only digits allowed
Public Const gEntryNumeric = 1      'Entry of printable characters allowed

'global variables for by ScrapWind
'   (was originally for passing between ScrapWind and frmLookup)
Public gSelCode As String           'The selected Code from frmLookup
Public gSelDesc As String           'The selected Description from frmLookup
Public gSelIndx As Integer          'The index of the selected item in the
Public ListOK As Boolean                                    '   collection used by frmLookup

Public gBatchProcess As Boolean
Public okToPrompt As Boolean

'***** GLOBAL SQL Server DECLARATIONS *****
Public gSQLServer       As String
Public gSQLDatabase     As String
Public gbConnected      As Boolean
Public gbStartupMode    As Boolean     'Used to shut down app if problem starting up

'***** GLOBAL App Variables *****

Public gstrHelpDeskInfo As String

'Win32 or Win 31 API Calls to write to .INI files
#If Win32 Then
    Public Declare Function wu_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function wu_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
    Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
#End If

''Win32 or Win 31 API Calls to Set the Window Position. Originally
'   used to try to resolve problem with Border Style 0 not getting
'   control back to app after alt-tab.  Must have control box set
'   to true to solve this.
'
'#If Win32 Then
'    Declare Function SetWindowPos Lib "user32" _
'        (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
'        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
'        ByVal cy As Long, ByVal wFlags As Long)
'#Else
'Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, _
'    ByVal hWndInsertAfter As Integer, ByVal X As Integer, _
'    ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, _
'    ByVal wFlags As Integer)
'
'#End If
'
'Public Const SWP_HIDEWINDOW = &H80
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOSIZE = &H1
'Public Const SWP_SHOWWINDOW = &H40
'Public Const HWND_NOTOPMOST = -2
'Public Const HWND_TOP = 0
'Public Const HWND_TOPMOST = -1
 
Public Type UserType
    UserNum As Integer
    UserID As String
    UserName As String
    FName As String
    MI As String
    LName As String
    DivNum As Integer
    EstEdit As Integer
    EstApprove As Integer
    MngPPGD As Integer
    MngUsers As Integer
    Disabled As Integer
 End Type

'***** GLOBAL CONSTANTS *****
'SQL Login ID
Public Const SQL_LOGIN_ID = "datacollection"
Public Const SQL_LOGIN_PSWD = "cherrygarcia"
Public Const MB_ICONEXCLAMATION = 48
Public Const MB_ICONSTOP = 16
Public Const MB_OK = 0
Public Const MB_YESNO = 4
Public Const vbYes = 6
Public Const vbNo = 7
Public Const CARRIAGE_RETURN = 13
Public Const TAB_KEY = 9
Public Const MOUSE_HOUR_GLASS = 11
Public Const MOUSE_NORMAL = 0
Public Const WINDOW_MAXIMIZE = 2

Public Const CLOCK_IN = 1
Public Const CLOCK_OUT = 2
Public Const OUT_FOR_LUNCH = 3
Public Const BACK_FROM_LUNCH = 4
Public Const SUSPEND_ATTD = 5
Public Const RESUME_ATTD = 6
Public Const START_CALL_IN = 7
Public Const END_CALL_IN = 8
Public Const INIT_RND_TM = -9999

'Misc
Public dBlankDate As Date
Public Type LaborType
   companyCd As Integer
   EmplNo As Long
   Actdate As Date
   OrderNo As String
   SeqNumb As Integer
   IndirCd As String
   WCID As String
   Oper As String
   SetStd As Currency
   MchStd As Currency
   EOrSStd As String
   ActQty As Long
   ActHrs As Currency
   ScrapQty1 As Long
   ScrapCd1 As String
   ScrapQty2 As Long
   ScrapCd2 As String
   ScrapQty3 As Long
   ScrapCd3 As String
   ScrapQty4 As Long
   ScrapCd4 As String
   ScrapQty5 As Long
   ScrapCd5 As String
   ScrapQty6 As Long
   ScrapCd6 As String
   ScrapQty7 As Long
   ScrapCd7 As String
   ScrapQty8 As Long
   ScrapCd8 As String
   CCID As String
   MachRec As String
   PayRate As Currency
   Earnings As Currency
   EmplEff As Currency
   EarnHrs As Currency
   EarnType As String
   ShiftNo As Byte
   OperStat As String
   IncenCd As Byte
   RecType As String
   MMachCd As String
   ExtractCd As String
   ActHrsUnrnd As Currency
   MachHrs As Currency
   ErrorLst As String
   ErrorFlag As String
   UnplnReas As String
   TranDate As Date
   LineNo As Long
   AddonHrs As Currency
   AddonReas As String
   NumbLoads As Integer
   CycleTime As Integer
   batchMeth As Integer
   box_nbr_1 As Long
   box_nbr_2 As Long
   box_nbr_3 As Long
   box_nbr_4 As Long
   box_nbr_5 As Long
   box_nbr_6 As Long
   box_nbr_7 As Long
   box_nbr_8 As Long
End Type


Public Type AttendanceType
   companyCd As Integer
   EmplNo As Long
   Actdate As Date
   ShiftCd As Byte
   CCID As String
   PdLunch As String
   HrsUnrnd As Currency
   ActHrs As Currency
   ErrorFlag As String
   ErrorLst As String
   ClockIn As Date
   ClockOut As Date
   LunchStart As Date
   LunchEnd As Date
   SuspLbr1 As Date
   ResmLbr1 As Date
   SuspLbr2 As Date
   ResmLbr2 As Date
   CIStart As Date
   CIEnd As Date
   ClockInRnd As Currency
   ClockOutRnd As Currency
   LunchOutRnd As Currency
   LunchInRnd As Currency
   SuspLbr1Rnd As Currency
   ResmLbr1Rnd As Currency
   SuspLbr2Rnd As Currency
   ResmLbr2Rnd As Currency
   CIStartRnd As Currency
   CIEndRnd As Currency
   CallInHrs As Currency
   Earnings As Currency
   PayRate As Currency
End Type

Public Type DivParamType
   ParamKey As String
   ParamData As String
End Type

Public Type BatchIncenType
   companyCd As Integer
   WCID As String
   BatchMethod As Integer
   MethodDesc As String
   IncenMin As Integer
End Type

Public Type CarryoverType
   companyCd As Integer
   EmplNo As Long
   Actdate As Date
   ShiftCd As Byte
   OrderNo As String
   SeqNumb As Integer
   WCID As String
   ActHrs As Currency
   OperStatus As String
   ActQty As Long
   transdate As Date
   CCID As String
   IncenType As Integer
   MachNumb As Integer
   SetStd As Currency
   IndirCd As String
   OperCode As String
   CloseDate As Date
End Type
   
Public Type CostCenterType
   companyCd As Integer
   CCID As String
   Descr As String
   SetupRate As Currency
   OvrhdRate As Currency
   MachRate As Currency
   PaidLunch As String
End Type


Public Type EmployeeType
   companyCd As Integer
   EmplNo As Long
   shiftGrp As Byte
   EmplName As String
   HomeCCID As String
   ClassWrkr As Byte
   TermCd As String
   PdLunch As String
   HourlyRate As Currency
   IncenRate As Currency
   GMRate As Currency
   AERRate As Currency
   SpeclRate As Currency
   Actdate As Date
   AsgnShift As Byte
   GenShift As Byte
   IndirCCID As String
   IndirCd As String
   AttdStartDate As Date
   AttdEndDate As Date
   AttdStatus As Integer
   LaborStartDate As Date
   ShiftStartDate As Date
   NewHrlyRate As Currency
   NewRateEff As Date
   LeadHand As String
End Type

Public Type indirIncenType
   companyCd As Integer
   CCID As String
   IncenType As Integer
   MachNumb As Integer
   WorkStd As Currency
   Descr As String
End Type


Public Type JobMasterType
    companyCd As Integer
    OrderNo As String
    OrderType As String
    OrderQty As Long
    PartNo As String
    PartRev As String
    CustCode As String
    CustName As String
    Actdate As Date
    OrderStat As String
    MatlWgt As Currency
End Type

Public Type JobOpersType
    companyCd As Integer
    OrderNo As String
    OperSeqNo As Integer
    WCID As String
    Oper As String
    PlnQty As Long
    PlnHrs As Currency
    SetStd As Currency
    MchStd As Currency
    EOrSStd As String
    TimeBasisCd As String
    ActQty As Long
    ActHrs As Currency
    ScrapQty As Long
    OperStatus As String
End Type

Public Type WorkCenterType
    companyCd As Integer
    WCID As String
    WcGroup As String
    Location As String
    Descr As String
    CCID As String
    PaidLunch As String
End Type

Public Type OpenLaborType
     companyCd As Integer
      EmplNo As Long
      Actdate As Date
      OrderNo As String
      SeqNumb As Integer
      OperCd As String
      IndirCd As String
      WCID As String
      CCID As String
      ShiftCd As Integer
      StartDateTime As Date
      MachRec As String
      IncenCd As Integer
      MMachCd As Integer
      UnplnReas As String
      ElapsedHrs As Currency
 End Type
          
 
Public Type MaintDataType
    companyCd As Integer
    EmplNo As Long
    Actdate As Date
    MaintReas As String
    WCID As String
    ActHrs As Currency
End Type

'Public Type CodeLookupType
'    Code As String
'    Desc As String
'    CodeDesc As String
'End Type

Public Type ShiftRecType
    startTime As String
    endTime As String
    LunchOut As String
    LunchIn As String
    StParam As String
    EndParam As String
    lunchGrace As String
    ACT As String
    EOSStart As String
    EOSEnd As String
End Type

Public Type StdChgsType
    companyCd As Integer
    EmplNo As Long
    OrderNo As String
    SeqNumb As Integer
    OrigFixStd As Currency
    RevFixedStd As Currency
    origMchStd As Currency
    RevMchStd As Currency
    ReasonText As String
    ShiftCd As Byte
    Actdate As Date
    transdate As Date
    ProcessFlg As String
    ProcessDate As Date
End Type

Public Type IIDataType
    IncnType As Integer
    MachNumb As Integer
End Type

Public Type IILookupType
    WorkStd As String
    Descr As String
End Type

Public Type BatIncnLkupType
    Method As Integer
    Descr As String
End Type

Public Type ScrapDataType
    Qty As Long
    ScrapType As String
    Code As String
End Type

Public Type OperCodeType
    OperNumb As Integer
    OperDesc As String
End Type

Public Type WorkDataType
    EmplNo As Long
    IndirCd As String
    WCID As String
    CCID As String
    OrderNo As String
    MatlWgt As Currency
    inHours As Currency
    machHours As Currency
    setupHours As Currency
    addonHours As Currency
    totScrapQty As Long
    ActQty As Long
    Scrap(1 To 8) As ScrapDataType
    SeqNo As Integer
    NumbMachs As Integer
    MachRec As String
    OperStatus As String
    Oper As String
    EOrS As String
    UnplnReas As String
    RecType As String
    SetStd As Currency
    MchStd As Currency
    incenRec As Integer
    NoOfLoads As Integer
    CycleTime As Integer
    biMethod As Integer
    ErrorLst As String
    BoxNumber(1 To 8) As Long
 End Type
 
Public Type PaEmplyDataType
    companyCd As Byte
    EmplNo As Long
    RecCode As String
    Actdate As String
    ShiftCd As Byte
    OrderNo As String
    SeqNo As Integer
    IndirCd As String
    CCID As String
    WCID As String
    ActHrs As Currency
    HrsUnrnd As Currency
    MachHrs As Currency
    ActQty As Long
    RecType As String
    EmplEff As Currency
    MchStd As Currency
    ScrapQty As Long
    IncenCd As Byte
    transdate As Date
    Oper_Cd As String
End Type
 
Public Type OTOfferingType
   companyCd As Byte
   EmplNo As Long
   Actdate As String
   OTCCID As String
   OfferedHrs As Currency
   AcceptedHrs As Currency
End Type

Public Type NotepadType
    companyCd As Byte
    EmplNo As Long
    Actdate As Date
    MsgText As String
    PrintFlag As String
End Type

Public Type messagetype
   msg_type As Integer
   message As String
   message_id As Long
   msg_answer As String
End Type
   
Public Type FRMSIZE
   Height As Long
   Width As Long
End Type

Public Const G_RETURNCOUNT = 360    '3 minutes - this count extends the timer interval as
                                                             ' the timer code is entered every half second, and the
                                                             ' count is reduced by 1 each time.
                                                             ' 600 / 2 per second = 300 seconds = 5 minutes
Public Const G_SETDATETIME = 600  'reset pc clock every 10 minutes

'Public Const C_SNAMESIZE = 48
'Public Const SUCCESSFUL = &H0
'Public Const CONNECTION_LOGGED_IN = &H4
'Declare Function NWGetPrimaryConnectionID Lib "NWCALLS.DLL" (conn As Integer) As Integer
Declare Function sendmessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal vmsg As Long, ByVal vParam As Long, lparam As String) As Long
Public gNotepad As NotepadType
Public gMessage As messagetype
Public gPaEmplyData(250) As PaEmplyDataType

Public gOTOffering() As OTOfferingType

Public gLaborIn(1 To 15) As WorkDataType
 
Public gCarryOver           As CarryoverType
Public gOperCode            As OperCodeType
Public gStdChg              As StdChgsType
Public gWorkCenter          As WorkCenterType
Public gCostCenter          As CostCenterType
Public gAttendance          As AttendanceType
Public gOpenLabor          As OpenLaborType
Public nAttendance          As AttendanceType
Public gLabor               As LaborType
Public gJobMaster           As JobMasterType
Public gJobOpers            As JobOpersType
Public gEmployee            As EmployeeType
Public uEmployee            As EmployeeType
Public gDivParam            As DivParamType
Public gUserInfo            As UserType
Public gBatchIncen          As BatchIncenType

Public gcolIndirCodes       As New clsCodeTable         'Lookup Table Codes
Public gcolScrapCodes       As New clsCodeTable
Public gcolUnplnCodes       As New clsCodeTable
Public gcolTlmntCodes       As New clsCodeTable
Public gcolOperCodes        As New clsCodeTable

#If Not gbJobPrint Then
    Public gcolAllIICodes   As New clsAllIITable        'Contains all the II(Indirect Incentives) Codes accross Cost Centers
    Public gcolCcIICodes    As New clsCCIITable         'Contains the II Codes for a specific Cost Center
    
    Public gcolAllBICodes   As New clsAllBITable        'Contains all the BI(Batch Incentives) Codes accross Work Centers
    Public gcolWcBICodes    As New clsWcBITable         'Contains the BI Codes for a specific Work Center
#End If
Public gcolIICodes          As New clsCodeTable         'Indirect Incentive :These are the LookUp tables for the Entry System Combo box
Public gcolBICodes          As New clsCodeTable         'Batch Incentive :These are the LookUp tables for the Entry System Combo box

Public gShiftData(20, 3) As ShiftRecType
Public gShiftRecData(1 To 3) As ShiftRecType
Public gAttdAdj(10) As Currency
'Public gIIData() As IIDataType

'Public gIndIncCd() As IILookupType
'Public gBatIncnCd() As String
'Public gIncMin() As Integer

Public AttdShift As Integer
Public AttdActDate As Date
Public AttdStatus As Integer
Public AttdDayCd As Integer
Public autolunch As Boolean
Public AttdIndirLbr As Boolean
Public mComment As String
Public gCompanyCd As Integer
Public gCompanyName As String * 20
Public gWCId As String
Public gEmplNo As Long
Public gOrderNo As String
Public gSeqNo As Integer
Public gCIGraceBef As Integer
Public gCIGraceAft As Integer
Public gCOGraceBef As Integer
Public gCOGraceAft As Integer
Public gGracePer As Integer
Public gAssyOperCode(1 To 7) As String
Public gDCSpecialOpers(1 To 20) As String

Public HalfHourRounding As Boolean
Public QtrHourRounding As Boolean

Public gScrapColl As Boolean
Public empClkOutOK As Boolean
Public gEosFlag As Boolean

Public gQtyOverrun As Integer
Public gClockOutMins As Integer
Public gAutoLunch As Boolean

'System date and time variables

Public gSysDate As Date   'Holds only date, no time
Public gSysTime As String 'holds the system time as HH:MM:SS
Public gSysTimeStrp As String 'Current system time without colons
Public gSysTimeSecs As Long  'Current system time expressed as number of seconds since midnight
Public gSysDateTimeStr As String 'Current date and time contained in a string
Public gSysDateTime As Date 'Current date and time as a date variable
Public gSysDateStr As String 'contains the current system date as a string mm/dd/yyyy

Public gTransType As Integer
Public Attd_Disp As String
Public gClockOutFlag As Boolean
Public inpClockOut As Boolean
Public x, Y As Integer
Public mintReturnCount As Integer

Public Int_Flag As Boolean
Public LoopCnt As Integer
Public Attd_sel As Integer

Public gInString As String
Public gRetnKey As Integer

Public gSrvrTimeStr As String

Public gErrBoxDuration As Integer
Public Const module_name = "Startup"
Public MyForm As FRMSIZE
Public DesignX As Integer
Public DesignY As Integer
Public ScaleFactorX As Single
Public ScaleFactorY As Single  ' Scaling factors
Public scalemode As Integer
Public Xtwips As Integer
Public Ytwips As Integer
Public Xpixels As Integer
Public Ypixels As Integer

Public RePosForm As Boolean
Public DoResize As Boolean
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
' Constants for Windows 32-bit Registry API
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Private Const REG_SZ = 1                        ' Unicode null terminated string
Private Const ERROR_SUCCESS = 0&



Public Function ReadIniTimeFormat() As Long
Const strFunction = "ReadIniTimeFormat"
Dim strErrorMsg         As String   'Var to Contain the Error Message
Dim sIniFile            As String   'Path and File name for the Apps ini file
Dim lrc                 As Long     'Var for checking return parms
Dim lpReturnedString    As String   'Place holder for string returned from API Call
Dim lSize               As Long     'Place holder for Size of string Passed to API

    ReadIniTimeFormat = vbError          'Assume the worst
    
    Rem .INI file in directory of program executed, with the
    Rem     same name as the executing program with appended .INI
    sIniFile = App.Path & "\" & Trim$(App.EXEName) & ".ini"
    lpReturnedString = Space(100)
    lSize = Len(lpReturnedString)
 
    '===============================================
    '   DSN Infor for DataColl
    '===============================================
    'Name
    lrc = GetPrivateProfileString("General", "TimeFormat", "", lpReturnedString, lSize, sIniFile)
    If lrc > 0 Then
        gTimeFormat = Left$(lpReturnedString, lrc)
        ReadIniTimeFormat = vbOK
    Else
        gTimeFormat = "HH:MM:SS AM/PM"
    End If

End Function
Public Function ReadIniOneDecFormat() As Long
Const strFunction = "ReadIniOneDecFormat"
Dim strErrorMsg         As String   'Var to Contain the Error Message
Dim sIniFile            As String   'Path and File name for the Apps ini file
Dim lrc                 As Long     'Var for checking return parms
Dim lpReturnedString    As String   'Place holder for string returned from API Call
Dim lSize               As Long     'Place holder for Size of string Passed to API

    ReadIniOneDecFormat = vbError          'Assume the worst
    
    Rem .INI file in directory of program executed, with the
    Rem     same name as the executing program with appended .INI
    sIniFile = App.Path & "\" & Trim$(App.EXEName) & ".ini"
    lpReturnedString = Space(100)
    lSize = Len(lpReturnedString)
 
    '===============================================
    '   DSN Infor for DataColl
    '===============================================
    'Name
    lrc = GetPrivateProfileString("General", "OneDecPos", "", lpReturnedString, lSize, sIniFile)
    If lrc > 0 Then
        gOneDecPos = Left$(lpReturnedString, lrc)
        ReadIniOneDecFormat = vbOK
    Else
        gOneDecPos = "#####.0"
        Err.Raise -994, , "Missing decimal format string in INI file"
    End If

End Function
Public Function ReadIniTwoDecFormat() As Long
Const strFunction = "ReadIniTwoDecFormat"
Dim strErrorMsg         As String   'Var to Contain the Error Message
Dim sIniFile            As String   'Path and File name for the Apps ini file
Dim lrc                 As Long     'Var for checking return parms
Dim lpReturnedString    As String   'Place holder for string returned from API Call
Dim lSize               As Long     'Place holder for Size of string Passed to API

    ReadIniTwoDecFormat = vbError          'Assume the worst
    
    Rem .INI file in directory of program executed, with the
    Rem     same name as the executing program with appended .INI
    sIniFile = App.Path & "\" & Trim$(App.EXEName) & ".ini"
    lpReturnedString = Space(100)
    lSize = Len(lpReturnedString)
 
    '===============================================
    '   DSN Infor for DataColl
    '===============================================
    'Name
    lrc = GetPrivateProfileString("General", "TwoDecPos", "", lpReturnedString, lSize, sIniFile)
    If lrc > 0 Then
        gTwoDecPos = Left$(lpReturnedString, lrc)
        ReadIniTwoDecFormat = vbOK
    Else
        gTwoDecPos = "#####.00"
        Err.Raise -994, , "Missing decimal format string in INI file"
    End If
End Function

Public Function ReadIniDateFormat() As Long
Const strFunction = "ReadIniDateFormat"
Dim strErrorMsg         As String   'Var to Contain the Error Message
Dim sIniFile            As String   'Path and File name for the Apps ini file
Dim lrc                 As Long     'Var for checking return parms
Dim lpReturnedString    As String   'Place holder for string returned from API Call
Dim lSize               As Long     'Place holder for Size of string Passed to API

    ReadIniDateFormat = vbError          'Assume the worst
    
    Rem .INI file in directory of program executed, with the
    Rem     same name as the executing program with appended .INI
    sIniFile = App.Path & "\" & Trim$(App.EXEName) & ".ini"
    lpReturnedString = Space(100)
    lSize = Len(lpReturnedString)
 
    '===============================================
    '   DSN Infor for DataColl
    '===============================================
    'Name
    lrc = GetPrivateProfileString("General", "DateFormat", "", lpReturnedString, lSize, sIniFile)
    If lrc > 0 Then
        gDateformat = Left$(lpReturnedString, lrc)
        ReadIniDateFormat = vbOK
    Else
        gDateformat = "MM/DD/YYYY"
    End If

End Function


Public Function ReadIniServer() As Long
'==============================================================================
'Name : John Canny & David Taylor / PRT Group
'Date : 06/15/98
'Comments : Reads the Ini file for this app and sets the Public Vars accordingly
'Parameters : None
'Returns : vbError for Error, vbOk for good to go
'==============================================================================
Const strFunction = "ReadIniServer"
Dim strErrorMsg         As String   'Var to Contain the Error Message
Dim sIniFile            As String   'Path and File name for the Apps ini file
Dim lrc                 As Long     'Var for checking return parms
Dim lpReturnedString    As String   'Place holder for string returned from API Call
Dim lSize               As Long     'Place holder for Size of string Passed to API

    ReadIniServer = vbError           'Assume the worst
    
    Rem .INI file in directory of program executed, with the
    Rem     same name as the executing program with appended .INI
    sIniFile = App.Path & "\" & Trim$(App.EXEName) & ".ini"
    lpReturnedString = Space(100)
    lSize = Len(lpReturnedString)
 
    '===============================================
    '   Server Info
    '===============================================
    'Name
    lrc = GetPrivateProfileString("SQL Server", "Server", "", _
        lpReturnedString, lSize, sIniFile)
    If lrc > 0 Then
        gSQLServer = Left$(lpReturnedString, lrc)
        ReadIniServer = vbOK
    Else
        gSQLServer = ""
    End If
    
    
End Function

Public Function ReadIniLangOffset() As Long
'==============================================================================
'Name : John Canny & David Taylor / PRT Group
'Date : 06/15/98
'Comments : Reads the Ini file for this app and sets the Public Vars accordingly
'Parameters : None
'Returns : vbError for Error, vbOk for good to go
'==============================================================================
Const strFunction = "ReadIniLangOffset"
Dim strErrorMsg         As String   'Var to Contain the Error Message
Dim sIniFile            As String   'Path and File name for the Apps ini file
Dim lrc                 As Long     'Var for checking return parms
Dim lpReturnedString    As String   'Place holder for string returned from API Call
Dim lSize               As Long     'Place holder for Size of string Passed to API

    ReadIniLangOffset = vbError           'Assume the worst
    
    Rem .INI file in directory of program executed, with the
    Rem     same name as the executing program with appended .INI
    sIniFile = App.Path & "\" & Trim$(App.EXEName) & ".ini"
    lpReturnedString = Space(100)
    lSize = Len(lpReturnedString)
 
    '===============================================
    '   Language Offset
    '===============================================
    'Name
    lrc = GetPrivateProfileString("General", "LanguageOffset", "", _
        lpReturnedString, lSize, sIniFile)
    If lrc > 0 Then
        gLanguageOffset = CInt(Left$(lpReturnedString, lrc))
        ReadIniLangOffset = vbOK
    Else
        gLanguageOffset = 0
    End If
    
    
End Function


Public Function ReadIniDatabase() As Long
'==============================================================================
'Name : John Canny & David Taylor / PRT Group
'Date : 06/15/98
'Comments : Reads the Ini file for this app and sets the Public Vars accordingly
'Parameters : None
'Returns : vbError for Error, vbOk for good to go
'==============================================================================
Const strfunctionname = "ReadIniDatabase"
Dim strErrorMsg         As String   'Var to Contain the Error Message
Dim sIniFile            As String   'Path and File name for the Apps ini file
Dim lrc                 As Long     'Var for checking return parms
Dim lpReturnedString    As String   'Place holder for string returned from API Call
Dim lSize               As Long     'Place holder for Size of string Passed to API
    ReadIniDatabase = vbError           'Assume the worst
    
    Rem .INI file in directory of program executed, with the
    Rem     same name as the executing program with appended .INI
    sIniFile = App.Path & "\" & Trim$(App.EXEName) & ".ini"
    lpReturnedString = Space(100)
    lSize = Len(lpReturnedString)
 
    '===============================================
    '   Server Info
    '===============================================
    'Name
    lrc = GetPrivateProfileString("SQL Server", "Database", "", _
        lpReturnedString, lSize, sIniFile)
    If lrc > 0 Then
        gSQLDatabase = Left$(lpReturnedString, lrc)
        ReadIniDatabase = vbOK
    Else
        gSQLDatabase = ""
        strErrorMsg = strfunctionname & _
            LoadResString(gLanguageOffset + 22) & sIniFile _
            & Chr(13) & gstrHelpDeskInfo
    End If
   
End Function


Public Function ReadIniErrorFile() As String
'==============================================================================
'Name : John Canny & David Taylor / PRT Group
'Date : 06/15/98
'Comments : Reads the Ini file for this app and sets the Public Vars accordingly
'Parameters : None
'Returns : vbError for Error, vbOk for good to go
'==============================================================================
Const strFunction = "ReadIniErrorFile"
Dim strErrorMsg         As String   'Var to Contain the Error Message
Dim sIniFile            As String   'Path and File name for the Apps ini file
Dim lrc                 As Long     'Var for checking return parms
Dim lpReturnedString    As String   'Place holder for string returned from API Call
Dim lSize               As Long     'Place holder for Size of string Passed to API

    ReadIniErrorFile = vbError           'Assume the worst
    
    Rem .INI file in directory of program executed, with the
    Rem     same name as the executing program with appended .INI
    sIniFile = App.Path & "\" & Trim$(App.EXEName) & ".ini"
    lpReturnedString = Space(100)
    lSize = Len(lpReturnedString)
 
    '===============================================
    '   Server Info
    '===============================================
    'Name
    lrc = GetPrivateProfileString("General", "ErrorFile", "", _
        lpReturnedString, lSize, sIniFile)
    If lrc > 0 Then
        gErrLocation = Left$(lpReturnedString, lrc)
        ReadIniErrorFile = vbOK
    Else
        gErrLocation = ""
    End If
    
    
End Function


Public Sub Main()
    Dim x, asmbcounter  As Integer
    Dim strServerTime As String
    Dim strconnect As String
    On Error GoTo ErrorHandler
    Set gError = New CError
          ' Size of Form in Pixels at design resolution
    DesignX = 1024
    DesignY = 768
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
      ' Set up the screen values
    Xtwips = Screen.TwipsPerPixelX
    Ytwips = Screen.TwipsPerPixelY
    Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
    Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

      ' Determine scaling factors
    ScaleFactorX = (Xpixels / DesignX)
    ScaleFactorY = (Ypixels / DesignY)
    scalemode = 1  ' twips
   
    'preset common error string trailer
    gstrHelpDeskInfo = LoadResString(gLanguageOffset + 7)
            
    'Check to see if program is already running
    If App.PrevInstance Then
        Err.Raise -992, , "Application is already running"
        End
     End If
         
    'Get INI File information
    
    'Set Error Log Location
    If ReadIniErrorFile() = vbError Then
        gErrLocation = App.Path & "\" & Trim$(App.EXEName) & ".Err"
        Err.Raise -998, , "Missing error log file name in INI file"
    End If
   'Get Ini file setting for language
    If ReadIniLangOffset() = vbError Then
          Err.Raise -997, , "Missing language offset in INI file"
    End If
    If ReadIniDateFormat() = vbError Then
           Err.Raise -995, , "Missing date format string in INI file"
    End If
    If ReadIniTimeFormat() = vbError Then
         Err.Raise -993, , "Missing time format string in INI file"
    End If
    gOneDecPos = "#####0.0"
    gTwoDecPos = "######0.00"
    
    If ReadIniServer = vbError Then
           Err.Raise -996, , "Missing Sql Server name in INI file"
    End If
    If ReadIniDatabase() = vbError Then
           Err.Raise -999, , "Missing Database name in INI file"
    End If
       
    gAttdLogFile = App.Path & "\" & Trim$(App.EXEName) & ".Log"
    gLaborLogFile = App.Path & "\DCLabor.Log"
    gAttdLogLocal = "C:\AttdLog\DCFloor.log"
    
    'Start loading up tables - shut down if SQL Error
    gbStartupMode = True
    gDateformat = GetSetting(HKEY_CURRENT_USER, "Control Panel\International", "sshortdate")
    gDecimal = GetSetting(HKEY_CURRENT_USER, "Control Panel\International", "sDecimal")
    gThousand = GetSetting(HKEY_CURRENT_USER, "Control Panel\International", "sMonThousandSep")
    gDateSep = GetSetting(HKEY_CURRENT_USER, "Control Panel\International", "sDate")
    gTimeFormat = GetSetting(HKEY_CURRENT_USER, "Control Panel\International", "sTimeFormat")
    'Open SQL Database
    If ConnectToDB() = vbError Then
           Err.Raise -990, , "Could not connect to database User ID=" & SQL_LOGIN_ID & " Password=" & SQL_LOGIN_PSWD & _
                              " Initial Catalog=" & gSQLDatabase & " Data Source=" & gSQLServer
    End If
    
    gSrvrTimeStr = "NET TIME \\" & gSQLServer & " /set /y"
    Call SetupDiv
    
    Shell gSrvrTimeStr
   
    frmDCMain.lblDateDisp = Format(Date, gDateformat)
    frmDCMain.lblTimeDisp = Format(Time, gTimeFormat)

    
    If GetDivParam("CLOCKOUT") = True Then
        gClockOutMins = Left(gDivParam.ParamData, 2)
    Else
        gClockOutMins = 0
    End If

    If GetDivParam("ASMBOPER") = True Then
          asmbcounter = 1
          x = 1
          Do While asmbcounter <= 7
               If IsNumeric(Mid(gDivParam.ParamData, x, 3)) Then
                      gAssyOperCode(asmbcounter) = Mid(gDivParam.ParamData, x, 3)
               Else
                      gAssyOperCode(asmbcounter) = 999
                End If
                asmbcounter = asmbcounter + 1
                x = x + 3
          Loop
    Else
         gAssyOperCode(1) = 999
    End If
    
    If GetDivParam("DC-SPECL") = True Then
          asmbcounter = 1
          x = 1
          Do While asmbcounter <= 20
               If IsNumeric(Mid(gDivParam.ParamData, x, 3)) Then
                      gDCSpecialOpers(asmbcounter) = Mid(gDivParam.ParamData, x, 3)
               Else
                      gDCSpecialOpers(asmbcounter) = 999
                End If
                asmbcounter = asmbcounter + 1
                x = x + 3
          Loop
    Else
         gDCSpecialOpers(1) = 999
    End If
    
    If GetDivParam("QTYOVRRN") = True Then
        gQtyOverrun = Left(gDivParam.ParamData, 2)
    Else
        gQtyOverrun = 5
    End If
    
    If GetDivParam("GRACEPER") = True Then
        gGracePer = Left(gDivParam.ParamData, 2)
        gCIGraceBef = Mid(gDivParam.ParamData, 3, 2)
        gCIGraceAft = Mid(gDivParam.ParamData, 5, 2)
        gCOGraceBef = Mid(gDivParam.ParamData, 7, 2)
        gCOGraceAft = Mid(gDivParam.ParamData, 9, 2)
    Else
        gGracePer = 1
        gCIGraceBef = 3
        gCIGraceAft = 3
        gCOGraceBef = 3
        gCOGraceAft = 3
    End If
          
    If GetDivParam("ATTD_ADJ") = False Then
        gDivParam.ParamData = "0.00.00.00.00.00.00.00.00.00.0"
    End If
    For x = 1 To 10
       If IsNumeric(Mid(gDivParam.ParamData, ((x * 3) - 2), 3)) Then
           gAttdAdj(x) = CCur(Mid(gDivParam.ParamData, ((x * 3) - 2), 3))
       Else
           gAttdAdj(x) = 0#
       End If
    Next x
    
    gScrapColl = GetDivParam("SCRPCOLL")
    
    If GetDivParam("AUTOLNCH") = True Then
          If Left(gDivParam.ParamData, 1) = "Y" Then
                gAutoLunch = True
          Else
               gAutoLunch = False
          End If
     Else
          gAutoLunch = False
    End If
    
        'No Error checking needed here, handled by Error and Message Events
    Call LoadLookupCols("INDIR", 2, gcolIndirCodes)
    Call LoadLookupCols("SCRAP", 2, gcolScrapCodes)
    Call LoadLookupCols("TLMNT", 2, gcolTlmntCodes)
    Call LoadLookupCols("UNPLAN", 2, gcolUnplnCodes)
    Call LoadAllBatchIncnCols(gcolAllBICodes)
    Call LoadAllIndirIncnCols(gcolAllIICodes)
    Call LoadOperCol(gcolOperCodes)
 '  Call LoadShiftCodes(gShiftData())
        
    'Done loading up tables for startup - don't shut down if SQL Error
    gbStartupMode = False
    
    DisconnectDB
'    If gSQLConnection = adStateOpen Then
 '       gSQLConnection.Close
  '  End If
    
    
'    frmDCMain.Show 1
     frmDCMain.Show
 '    UnloadApp
    Exit Sub
ErrorHandler:
    ' get error info
    gError.ErrNum = Err.Number
    gError.ErrDesc = Err.Description
    gError.ModuleName = module_name
    gError.FunctionName = "Main"
    gError.ErrorRoutine
'    gError.Clear
    gError.ShowMsg True
    End
End Sub
Public Sub UnloadApp()

 DisconnectDB
 
 
' If gSQLConnection.State = adStateOpen Then
'      gSQLConnection.Close
' End If
'If Not (gSQLConnection Is Nothing) Then
'        Set gSQLConnection = Nothing
'End If

End

End Sub
Public Sub GetSysDateTime()
   gSysDateTime = Now

   gSysDate = DateValue(gSysDateTime)   'Date
   gSysDateStr = CStr(DateValue(gSysDateTime))
   gSysTime = Format(gSysDateTime, "HH:MM:SS")
   
'   gSysDateTimeStr = gSysDateStr & " " & gSysTime
   gSysDateTimeStr = CStr(gSysDateTime)
   
   If Len(gSysTime) = 7 Then
      gSysTimeStrp = "0" & Mid(gSysTime, 1, 1) & _
      Mid(gSysTime, 3, 2) & Mid(gSysTime, 6, 2)
   Else
      gSysTimeStrp = Mid(gSysTime, 1, 2) & _
      Mid(gSysTime, 4, 2) & Mid(gSysTime, 7, 2)
   End If
   gSysTimeSecs = ConvTimeToSecs(gSysTimeStrp)
End Sub
Public Sub RecordError(ByVal sNewErrCode As String)
Dim lenDiff As Integer
  
  sNewErrCode = RTrim(sNewErrCode)
'if string to be added is empty or null exit sub without further
'action
  If IsEmpty(sNewErrCode) Or IsNull(sNewErrCode) Or Len(sNewErrCode) = 0 Then
      Exit Sub
  End If
'if receiving string is empty or null initialize to a single
'space. The Rtrim function will remove it and its length will be
'evaluated as zero.

  If IsEmpty(gLaborIn(LoopCnt).ErrorLst) Or IsNull(gLaborIn(LoopCnt).ErrorLst) Then
        gLaborIn(LoopCnt).ErrorLst = " "
  End If
  gLaborIn(LoopCnt).ErrorLst = RTrim(gLaborIn(LoopCnt).ErrorLst)

  If (Len(gLaborIn(LoopCnt).ErrorLst) + 1) <= 10 Then
     gLaborIn(LoopCnt).ErrorLst = gLaborIn(LoopCnt).ErrorLst & sNewErrCode
     gLaborIn(LoopCnt).ErrorLst = RTrim(gLaborIn(LoopCnt).ErrorLst)
  End If
End Sub
Public Sub Init_varbs()
Dim x, Y As Integer

x = 1
For x = 1 To 15
   gLaborIn(x).EmplNo = 0
   gLaborIn(x).IndirCd = "  "
   gLaborIn(x).WCID = "     "
   gLaborIn(x).CCID = "     "
   gLaborIn(x).OrderNo = "       "
   gLaborIn(x).MatlWgt = 0
   gLaborIn(x).inHours = 0
   gLaborIn(x).machHours = 0
   gLaborIn(x).setupHours = 0
   gLaborIn(x).addonHours = 0
   gLaborIn(x).totScrapQty = 0
   gLaborIn(x).ActQty = 0
   For Y = 1 To 8
      gLaborIn(x).Scrap(Y).Qty = 0
      gLaborIn(x).Scrap(Y).ScrapType = " "
      gLaborIn(x).Scrap(Y).Code = "  "
      gLaborIn(x).BoxNumber(Y) = 0
   Next Y
   gLaborIn(x).SeqNo = 0
   gLaborIn(x).NumbMachs = 0
   gLaborIn(x).MachRec = " "
   gLaborIn(x).OperStatus = " "
   gLaborIn(x).Oper = " "
   gLaborIn(x).EOrS = " "
   gLaborIn(x).UnplnReas = "  "
   gLaborIn(x).RecType = " "
   gLaborIn(x).SetStd = 0
   gLaborIn(x).MchStd = 0
   gLaborIn(x).incenRec = 0
   gLaborIn(x).NoOfLoads = 0
   gLaborIn(x).CycleTime = 0
   gLaborIn(x).biMethod = 0
   gLaborIn(x).ErrorLst = "          "
 Next x
 LoopCnt = 1
 empClkOutOK = False
 gClockOutFlag = False
 gEosFlag = False
 AttdIndirLbr = False
 gOKToComplete = False
End Sub

Public Sub Init_labor_rec()
 gLabor.companyCd = 0
 gLabor.EmplNo = 0
 gLabor.Actdate = Date
 gLabor.OrderNo = "       "
 gLabor.SeqNumb = 0
 gLabor.IndirCd = "  "
 gLabor.WCID = "     "
 gLabor.Oper = "                         "
 gLabor.SetStd = 0
 gLabor.MchStd = 0
 gLabor.EOrSStd = " "
 gLabor.ActQty = 0
 gLabor.ActHrs = 0
 gLabor.ScrapQty1 = 0
 gLabor.ScrapCd1 = "  "
 gLabor.ScrapQty2 = 0
 gLabor.ScrapCd2 = "  "
 gLabor.ScrapQty3 = 0
 gLabor.ScrapCd3 = "  "
 gLabor.ScrapQty4 = 0
 gLabor.ScrapCd4 = "  "
 gLabor.ScrapQty5 = 0
 gLabor.ScrapCd5 = "  "
 gLabor.ScrapQty6 = 0
 gLabor.ScrapCd6 = "  "
 gLabor.ScrapQty7 = 0
 gLabor.ScrapCd7 = "  "
 gLabor.ScrapQty8 = 0
 gLabor.ScrapCd8 = "  "
 gLabor.CCID = "     "
 gLabor.MachRec = " "
 gLabor.PayRate = 0
 gLabor.Earnings = 0
 gLabor.EmplEff = 0
 gLabor.EarnHrs = 0
 gLabor.EarnType = " "
 gLabor.ShiftNo = 0
 gLabor.OperStat = " "
 gLabor.IncenCd = 0
 gLabor.RecType = " "
 gLabor.MMachCd = 0
 gLabor.ExtractCd = " "
 gLabor.ActHrsUnrnd = 0
 gLabor.MachHrs = 0
 gLabor.ErrorLst = "          "
 gLabor.ErrorFlag = " "
 gLabor.UnplnReas = "  "
 gLabor.TranDate = Date
 gLabor.LineNo = 0
 gLabor.AddonHrs = 0
 gLabor.AddonReas = "  "
 gLabor.NumbLoads = 0
 gLabor.CycleTime = 0
 gLabor.batchMeth = 0
 gLabor.box_nbr_1 = 0
 gLabor.box_nbr_2 = 0
 gLabor.box_nbr_3 = 0
 gLabor.box_nbr_4 = 0
 gLabor.box_nbr_5 = 0
 gLabor.box_nbr_6 = 0
 gLabor.box_nbr_7 = 0
 gLabor.box_nbr_8 = 0

End Sub



Public Sub Build_OSSendOut()
Dim aryNumb As Integer

For aryNumb = 1 To LoopCnt
   Init_labor_rec
   'need to send only four fields to the sp... the other fields in the labor record
   'are assigned default values for this transaction.
   gLabor.EmplNo = gEmployee.EmplNo
   gLabor.OrderNo = gLaborIn(aryNumb).OrderNo
   gLabor.SeqNumb = gLaborIn(aryNumb).SeqNo
   gLabor.ActQty = gLaborIn(aryNumb).ActQty
   Ins_OSSendOut
Next aryNumb

End Sub


Public Sub Build_Assembly_Recs()
Dim aryNumb As Integer

For aryNumb = 1 To LoopCnt
   Init_labor_rec
   'need to send only four fields to the sp... the other fields in the labor record
   'are assigned default values for this transaction.
   gLabor.EmplNo = gEmployee.EmplNo
   gLabor.OrderNo = gLaborIn(aryNumb).OrderNo
   gLabor.SeqNumb = gLaborIn(aryNumb).SeqNo
   gLabor.ActQty = gLaborIn(aryNumb).ActQty
   Ins_Assembly_Labor
Next aryNumb

End Sub
Public Sub Build_Std_Lbr_Recs()
Dim aryNumb As Integer

For aryNumb = 1 To LoopCnt
   Init_labor_rec
   gLabor.EmplNo = gEmployee.EmplNo
   gLabor.OrderNo = gLaborIn(aryNumb).OrderNo
   gLabor.SeqNumb = gLaborIn(aryNumb).SeqNo
   gLabor.ActQty = gLaborIn(aryNumb).ActQty
   gLabor.companyCd = gCompanyCd
   If gLaborIn(aryNumb).Scrap(1).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty1 = ConvScrapType(gLaborIn(aryNumb).Scrap(1).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
  Else
          gLabor.ScrapQty1 = gLaborIn(aryNumb).Scrap(1).Qty
  End If
   gLabor.ScrapCd1 = gLaborIn(aryNumb).Scrap(1).Code
   If gLaborIn(aryNumb).Scrap(2).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty2 = ConvScrapType(gLaborIn(aryNumb).Scrap(2).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
   Else
          gLabor.ScrapQty2 = gLaborIn(aryNumb).Scrap(2).Qty
   End If
   gLabor.ScrapCd2 = gLaborIn(aryNumb).Scrap(2).Code
   If gLaborIn(aryNumb).Scrap(3).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty3 = ConvScrapType(gLaborIn(aryNumb).Scrap(3).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
   Else
          gLabor.ScrapQty3 = gLaborIn(aryNumb).Scrap(3).Qty
   End If
   gLabor.ScrapCd3 = gLaborIn(aryNumb).Scrap(3).Code
       If gLaborIn(aryNumb).Scrap(4).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty4 = ConvScrapType(gLaborIn(aryNumb).Scrap(4).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
   Else
          gLabor.ScrapQty4 = gLaborIn(aryNumb).Scrap(4).Qty
   End If
   gLabor.ScrapCd4 = gLaborIn(aryNumb).Scrap(4).Code
   If gLaborIn(aryNumb).Scrap(5).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty5 = ConvScrapType(gLaborIn(aryNumb).Scrap(5).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
   Else
          gLabor.ScrapQty5 = gLaborIn(aryNumb).Scrap(5).Qty
   End If
   gLabor.ScrapCd5 = gLaborIn(aryNumb).Scrap(5).Code
   If gLaborIn(aryNumb).Scrap(6).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty6 = ConvScrapType(gLaborIn(aryNumb).Scrap(6).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
   Else
          gLabor.ScrapQty6 = gLaborIn(aryNumb).Scrap(6).Qty
   End If
   gLabor.ScrapCd6 = gLaborIn(aryNumb).Scrap(6).Code
   If gLaborIn(aryNumb).Scrap(7).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty7 = ConvScrapType(gLaborIn(aryNumb).Scrap(7).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
   Else
          gLabor.ScrapQty7 = gLaborIn(aryNumb).Scrap(7).Qty
   End If
   gLabor.ScrapCd7 = gLaborIn(aryNumb).Scrap(7).Code
   If gLaborIn(aryNumb).Scrap(8).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty8 = ConvScrapType(gLaborIn(aryNumb).Scrap(8).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
   Else
          gLabor.ScrapQty8 = gLaborIn(aryNumb).Scrap(8).Qty
   End If
   gLabor.ScrapCd8 = gLaborIn(aryNumb).Scrap(8).Code
   gLabor.OperStat = gLaborIn(aryNumb).OperStatus
   gLabor.ActHrsUnrnd = 0
   Ins_Std_Labor
Next aryNumb

End Sub

Public Sub Build_Labor_Recs()
Dim AttdUpdFlag As Boolean
Dim x, shiftGrp, shiftNum As Integer
Dim aryNumb As Integer
If First_Labor_Rec(gEmployee.EmplNo, gSysDate) = True Then
   AttdUpdFlag = True
End If
For aryNumb = 1 To LoopCnt
   Init_labor_rec
   If gCompanyCd = 1 And gEmployee.ClassWrkr = 2 Then
       checkIncen (aryNumb)
   End If
   gLabor.companyCd = gCompanyCd
   gLabor.EmplNo = gEmployee.EmplNo
   gLabor.Actdate = gEmployee.Actdate
   gLabor.CCID = gLaborIn(aryNumb).CCID
   gLabor.TranDate = gSysDateTime
   gLabor.ShiftNo = gEmployee.GenShift
   gLabor.ErrorLst = gLaborIn(aryNumb).ErrorLst
   gLabor.RecType = gLaborIn(aryNumb).RecType
   If AttdUpdFlag Then
       gAttendance.CCID = gLabor.CCID
   End If
   If gLabor.RecType = "I" Or gLabor.RecType = "C" Then
       gLabor.IndirCd = gLaborIn(aryNumb).IndirCd
       gLabor.ActHrsUnrnd = gLaborIn(aryNumb).inHours
       gLabor.ActHrs = gLaborIn(aryNumb).inHours
       If gCompanyCd = hyson_products And gLabor.IndirCd = "20" Then
           gLabor.OrderNo = gLaborIn(aryNumb).OrderNo
       End If
       If gCompanyCd = corry And gLabor.IndirCd = "50" Then
           gLabor.OrderNo = gLaborIn(aryNumb).OrderNo
       End If
       If gCompanyCd = bristol And (Left(gLabor.IndirCd, 1) = "A" Or _
       Left(gLabor.IndirCd, 1) = "B") Then
           gLabor.OrderNo = gLaborIn(aryNumb).OrderNo
       End If
       If gCompanyCd = corry And gLabor.IndirCd = "49" Then
          gLabor.SetStd = gLaborIn(aryNumb).SetStd
          gLabor.OperStat = gLaborIn(aryNumb).OperStatus
          gLabor.IncenCd = gLaborIn(aryNumb).incenRec
          gLabor.MMachCd = gLaborIn(aryNumb).NumbMachs
       End If
       If gCompanyCd = milw And gLabor.IndirCd = "12" Then
          gLabor.UnplnReas = gLaborIn(aryNumb).UnplnReas
          gLabor.WCID = gLaborIn(aryNumb).WCID
       End If
    Else
       gLabor.WCID = gLaborIn(aryNumb).WCID
       gLabor.OrderNo = gLaborIn(aryNumb).OrderNo
       gLabor.SeqNumb = gLaborIn(aryNumb).SeqNo
       gLabor.SetStd = gLaborIn(aryNumb).SetStd
       gLabor.MchStd = gLaborIn(aryNumb).MchStd
       gLabor.EOrSStd = gLaborIn(aryNumb).EOrS
       gLabor.Oper = gLaborIn(aryNumb).Oper
       If gLaborIn(aryNumb).Scrap(1).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty1 = ConvScrapType(gLaborIn(aryNumb).Scrap(1).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty1 = gLaborIn(aryNumb).Scrap(1).Qty
       End If
       gLabor.ScrapCd1 = gLaborIn(aryNumb).Scrap(1).Code
       If gLaborIn(aryNumb).Scrap(2).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty2 = ConvScrapType(gLaborIn(aryNumb).Scrap(2).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty2 = gLaborIn(aryNumb).Scrap(2).Qty
       End If
       gLabor.ScrapCd2 = gLaborIn(aryNumb).Scrap(2).Code
       If gLaborIn(aryNumb).Scrap(3).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty3 = ConvScrapType(gLaborIn(aryNumb).Scrap(3).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty3 = gLaborIn(aryNumb).Scrap(3).Qty
       End If
       gLabor.ScrapCd3 = gLaborIn(aryNumb).Scrap(3).Code
       If gLaborIn(aryNumb).Scrap(4).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty4 = ConvScrapType(gLaborIn(aryNumb).Scrap(4).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty4 = gLaborIn(aryNumb).Scrap(4).Qty
       End If
       gLabor.ScrapCd4 = gLaborIn(aryNumb).Scrap(4).Code
       If gLaborIn(aryNumb).Scrap(5).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty5 = ConvScrapType(gLaborIn(aryNumb).Scrap(5).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty5 = gLaborIn(aryNumb).Scrap(5).Qty
       End If
       gLabor.ScrapCd5 = gLaborIn(aryNumb).Scrap(5).Code
       If gLaborIn(aryNumb).Scrap(6).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty6 = ConvScrapType(gLaborIn(aryNumb).Scrap(6).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty6 = gLaborIn(aryNumb).Scrap(6).Qty
       End If
       gLabor.ScrapCd6 = gLaborIn(aryNumb).Scrap(6).Code
       If gLaborIn(aryNumb).Scrap(7).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty7 = ConvScrapType(gLaborIn(aryNumb).Scrap(7).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty7 = gLaborIn(aryNumb).Scrap(7).Qty
       End If
       gLabor.ScrapCd7 = gLaborIn(aryNumb).Scrap(7).Code
       If gLaborIn(aryNumb).Scrap(8).ScrapType = "L" And (gCompanyCd = milw Or gCompanyCd = saline) Then
          gLabor.ScrapQty8 = ConvScrapType(gLaborIn(aryNumb).Scrap(8).Qty, _
                               gLaborIn(aryNumb).MatlWgt)
       Else
          gLabor.ScrapQty8 = gLaborIn(aryNumb).Scrap(8).Qty
       End If
       gLabor.ScrapCd8 = gLaborIn(aryNumb).Scrap(8).Code
       gLabor.ActQty = gLaborIn(aryNumb).ActQty
       gLabor.OperStat = gLaborIn(aryNumb).OperStatus
       gLabor.ActHrsUnrnd = 0
       gLabor.ActHrs = gLaborIn(aryNumb).setupHours
       gLabor.MachHrs = gLaborIn(aryNumb).machHours
       gLabor.UnplnReas = gLaborIn(aryNumb).UnplnReas
       gLabor.MachRec = gLaborIn(aryNumb).MachRec
       gLabor.MMachCd = gLaborIn(aryNumb).NumbMachs
       gLabor.IncenCd = gLaborIn(aryNumb).incenRec
       gLabor.AddonHrs = gLaborIn(aryNumb).addonHours
       gLabor.NumbLoads = gLaborIn(aryNumb).NoOfLoads
       gLabor.CycleTime = gLaborIn(aryNumb).CycleTime
       gLabor.batchMeth = gLaborIn(aryNumb).biMethod
       If gCompanyCd = 7 And gLabor.MMachCd = 1 And gLabor.MachRec = "Y" Then
         gLabor.EarnHrs = gLaborIn(aryNumb).inHours
       End If
       gLabor.box_nbr_1 = gLaborIn(aryNumb).BoxNumber(1)
       gLabor.box_nbr_2 = gLaborIn(aryNumb).BoxNumber(2)
       gLabor.box_nbr_3 = gLaborIn(aryNumb).BoxNumber(3)
       gLabor.box_nbr_4 = gLaborIn(aryNumb).BoxNumber(4)
       gLabor.box_nbr_5 = gLaborIn(aryNumb).BoxNumber(5)
       gLabor.box_nbr_6 = gLaborIn(aryNumb).BoxNumber(6)
       gLabor.box_nbr_7 = gLaborIn(aryNumb).BoxNumber(7)
       gLabor.box_nbr_8 = gLaborIn(aryNumb).BoxNumber(8)
    End If
    Call Ins_Labor_Rec
Next aryNumb
Call GetEmployee(gEmployee.EmplNo)
End Sub
Public Sub Build_OpenLabor()
Dim AttdUpdFlag As Boolean
Dim x, shiftGrp, shiftNum As Integer
Dim aryNumb As Integer
If First_Labor_Rec(gEmployee.EmplNo, gSysDate) = True Then
   AttdUpdFlag = True
End If
For aryNumb = 1 To LoopCnt
   Init_labor_rec
   gOpenLabor.companyCd = gCompanyCd
   gOpenLabor.EmplNo = gEmployee.EmplNo
   gOpenLabor.Actdate = gEmployee.Actdate
   gOpenLabor.CCID = gLaborIn(aryNumb).CCID
   gOpenLabor.ShiftCd = gEmployee.GenShift
   gOpenLabor.IndirCd = gLaborIn(aryNumb).IndirCd
   gOpenLabor.OrderNo = gLaborIn(aryNumb).OrderNo
   gOpenLabor.SeqNumb = gLaborIn(aryNumb).SeqNo
   gOpenLabor.IncenCd = gLaborIn(aryNumb).incenRec
   gOpenLabor.MMachCd = gLaborIn(aryNumb).NumbMachs
   gOpenLabor.UnplnReas = gLaborIn(aryNumb).UnplnReas
   gOpenLabor.WCID = gLaborIn(aryNumb).WCID
   gOpenLabor.OperCd = gLaborIn(aryNumb).Oper
   gOpenLabor.MachRec = "N"

    Call Ins_OpenLabor_Rec
Next aryNumb
Call GetEmployee(gEmployee.EmplNo)
End Sub


Public Sub SetupDiv()
        'Set Current Division to user's default division
 '      gDivNum = gUserInfo.DivNum
 
        If GetDivParam("DIVISION") = True Then
            gCompanyCd = Left(gDivParam.ParamData, 2)
            gCompanyName = Mid(gDivParam.ParamData, 3, 20)
            div = CInt(gCompanyCd)
        Else
            MsgBox (LoadResString(gLanguageOffset + 43))
        End If
        
        Rem Preset Labor Options to False
        gbEOShift = False
        gbNumbMach = False
        gbInpSetOrMach = False
        gbInpSetAndMach = False
        gbAddonHrs = False
        gbZeroQty = False
        gbCallCompl = False
        gbNoLoads = False
        gbAnotherRec = False
        gbMachRecY = False
        gbInpStds = False
        gbRecTypeR = False
        gbScrapPounds = False
        gbInpIndirHrs = False
        gbInpIndirHrs = False
        
        Rem Set Unchanging Labor Options by division
        Select Case div
        Case bristol
            gbZeroQty = True
            gbCallCompl = False
            gbInpIndirHrs = True
            
        Case syracuse
            gbInpSetOrMach = True
            gbZeroQty = True
            gbInpIndirHrs = True
        Case corry
            gbEOShift = True
            gbAddonHrs = True
            gbCallCompl = True
            gbInpStds = True
        Case saline
            gbInpSetOrMach = True
            gbZeroQty = True
            gbScrapPounds = True
            gbInpIndirHrs = True
            Rem No Direct Labor
        Case milw
            gbInpSetOrMach = True
            gbZeroQty = True
            gbScrapPounds = True
            gbCallCompl = False
            gbInpIndirHrs = True
       Case monterrey
           gbInpSetOrMach = True
            gbInpIndirHrs = True
        Case remesa
           gbInpSetOrMach = True
            gbInpIndirHrs = True
        Case hyson_products
            gbCallCompl = True
'            gbInpSetOrMach = True
            gbInpIndirHrs = False
        Case stroms
            gbCallCompl = True
        Case singapore
            gbInpSetOrMach = True
            gbZeroQty = True
            gbInpIndirHrs = True
        Case pv_singapore
            gbInpSetOrMach = True
            gbZeroQty = True
            gbInpIndirHrs = True
        End Select

End Sub

Public Sub InitPaEmplyData()
'==============================================================================
'Author           : David A. Taylor
'Date             : June, 1998
'Procedure Name   : Startup.InitPaEmplyData
'Description      : Initialize the Array of type Data used to hold the
'                   Check Labor Entries Menu Option.
'Parameters       : None
'Procedures Used  : None
'Objects Used     : gPaEmplyData Array
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    Dim i As Integer
    For i = 0 To 249
        gPaEmplyData(i).companyCd = 0
        gPaEmplyData(i).EmplNo = 0
        gPaEmplyData(i).RecCode = ""
        gPaEmplyData(i).Actdate = ""
        gPaEmplyData(i).ShiftCd = 0
        gPaEmplyData(i).OrderNo = ""
        gPaEmplyData(i).SeqNo = 0
        gPaEmplyData(i).IndirCd = ""
        gPaEmplyData(i).CCID = ""
        gPaEmplyData(i).WCID = ""
        gPaEmplyData(i).ActHrs = 0
        gPaEmplyData(i).HrsUnrnd = 0
        gPaEmplyData(i).MachHrs = 0
        gPaEmplyData(i).ActQty = 0
        gPaEmplyData(i).RecType = ""
        gPaEmplyData(i).EmplEff = 0
        gPaEmplyData(i).MchStd = 0
        gPaEmplyData(i).Oper_Cd = ""
    Next i
End Sub

Public Function checkIncen(ByVal i As Integer)
  If gLaborIn(i).MachRec = "Y" Then
     Exit Function
  End If
  
  gLaborIn(i).incenRec = 2
  If gLaborIn(i).MchStd > 0 And _
     gLaborIn(i).EOrS = "S" And _
     gLaborIn(i).SeqNo > 0 And _
     gLaborIn(i).machHours > 1# Then
       gLaborIn(i).incenRec = 1
  End If
End Function

Public Sub MainError(strError)
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Startup.MainError
'Description      : Display Error onto Main Form (Big Bold and Red).
'                   Error message is cleared when successful dbOpen called
'                   using the ClearMainError routine.
'Parameters       : None
'Procedures Used  : None
'Objects Used     : frmDCMain Error Labels
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    'Display Error onto Main Form (Big Bold and Red)
    '   Error message is cleared when successful dbOpen called
    frmDCMain.lblSysError.Visible = True
    frmDCMain.lblSysErrorDet.Visible = True
    frmDCMain.lblSysErrorDet = strError & Chr(13) _
        & gstrHelpDeskInfo
    gDone = True
End Sub

Public Sub ClearMainError()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : Startup.ClearMainError
'Description      : Clear the Error displays used by MainError routine.
'Parameters       : None
'Procedures Used  : None
'Objects Used     : frmDCMain Error Labels
'Forms Loaded     : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    frmDCMain.lblSysError.Visible = False
    frmDCMain.lblSysErrorDet.Visible = False
    frmDCMain.lblSysErrorDet = ""
End Sub



Public Function ConvScrapType(lScrapQty As Long, lMatlWgt As Currency) As Long
 Dim lInterimQty As Currency
 Dim lRndStr  As String
  
  ConvScrapType = 0
  If lMatlWgt > 0 Then
      lInterimQty = lScrapQty / lMatlWgt
      lRndStr = CStr(lInterimQty)
      ConvScrapType = CLng(lRndStr)
  End If
End Function

Public Function GetSetting(ByVal Hive As Long, ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
    ' Hive      Required. API Constant indicating which HKEY hive is needed
    ' Section   Required. String expression containing the name of the section where the key setting is found.
    '           If omitted, key setting is assumed to be in default subkey.
    ' Key       Required. String expression containing the name of the key setting to return.
    ' Default   Optional. Expression containing the value to return if no value is set in the key setting.
    '           If omitted, default is assumed to be a zero-length string ("").
    Dim nRet As Long
    Dim hKey As Long
    Dim nType As Long
    Dim nBytes As Long
    Dim Buffer As String
    
    ' Assume failure and set return to Default
    GetSetting = Default
    
    ' Open key
    nRet = RegOpenKeyEx(Hive, Section, 0&, KEY_QUERY_VALUE, hKey)
    If nRet = ERROR_SUCCESS Then
       ' Set appropriate value for default query
       If Key = "*" Then Key = vbNullString
       
       ' Determine how large the buffer needs to be
       nRet = RegQueryValueEx(hKey, Key, 0&, nType, ByVal Buffer, nBytes)
       If nRet = ERROR_SUCCESS Then
          ' Build buffer and get data
          If nBytes > 0 Then
             Buffer = Space(nBytes)
             nRet = RegQueryValueEx(hKey, Key, 0&, nType, ByVal Buffer, Len(Buffer))
             If nRet = ERROR_SUCCESS Then
                ' Trim NULL and return successful query!
                GetSetting = Left(Buffer, nBytes - 1)
             End If
          End If
       Call RegCloseKey(hKey)
       End If
    End If
End Function

Public Function SaveSetting(ByVal Hive As Long, ByVal Section As String, ByVal Key As String, ByVal Setting As String) As Boolean
    ' Hive      Required. API Constant indicating which HKEY hive is needed
    ' Section   Required. String expression containing the name of the section where the key setting is being saved.
    ' Key       Required. String expression containing the name of the key setting being saved.
    ' Setting   Required. Expression containing the value that key is being set to.
   Dim nRet As Long
   Dim hKey As Long
   Dim nResult As Long
   
   ' Open (or create and open) key
   nRet = RegCreateKeyEx(Hive, Section, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
   If nRet = ERROR_SUCCESS Then
      ' Set appropriate value for default query
      If Key = "*" Then Key = vbNullString
      ' Null-terminate setting, in case it's empty.
      ' Strange mirroring can occur otherwise.
      Setting = Setting & vbNullChar
      ' Write new value to registry
      nRet = RegSetValueEx(hKey, Key, 0&, REG_SZ, ByVal Setting, Len(Setting))
      Call RegCloseKey(hKey)
   End If
   SaveSetting = (nRet = ERROR_SUCCESS)
End Function

