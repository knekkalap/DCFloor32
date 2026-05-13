VERSION 5.00
Begin VB.Form frmLabor 
   Caption         =   "Data Collection Labor Menu"
   ClientHeight    =   6885
   ClientLeft      =   3300
   ClientTop       =   3840
   ClientWidth     =   8445
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6885
   ScaleWidth      =   8445
   WindowState     =   2  'Maximized
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   990
      TabIndex        =   10
      Top             =   4875
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   990
      Picture         =   "Frmlabor.frx":0000
      Stretch         =   -1  'True
      Top             =   1305
      Width           =   870
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   990
      TabIndex        =   9
      Top             =   4475
      Width           =   60
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   990
      TabIndex        =   8
      Top             =   4075
      Width           =   60
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   990
      TabIndex        =   7
      Top             =   3675
      Width           =   60
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   990
      TabIndex        =   6
      Top             =   3275
      Width           =   60
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   990
      TabIndex        =   5
      Top             =   2875
      Width           =   60
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   450
      TabIndex        =   4
      Top             =   5265
      Width           =   75
   End
   Begin VB.Label lblPressCaption 
      AutoSize        =   -1  'True
      Caption         =   "Press:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   2115
      Width           =   750
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   1080
      TabIndex        =   2
      Top             =   5310
      Width           =   60
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   990
      TabIndex        =   1
      Top             =   2475
      Width           =   60
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      Caption         =   "Labor Selections Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   510
      Left            =   1980
      TabIndex        =   0
      Top             =   1440
      Width           =   4800
   End
End
Attribute VB_Name = "frmLabor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Rem Menu selection constants
Const lmDirLabor = 1
Const lmReworkLabor = 2
Const lmMachLabor = 3
Const lmDirBatch = 4
Const lmMachBatch = 5
Const lmIndirCalc = 6
Const lmScrap = 7
Const lmDirCC = 8
Const lmStartIndir = 9
Const lmStartDir = 10
Const lmEndIndir = 11
Const lmEndDir = 12
Const lmOSSendOut = 13
Const lmOSReturn = 14
Const lmStdLbr = 15
Const lmIndirEnt = 16
Const lmEsc = 99
Dim mintLoaded As Integer



Private Sub Form_Activate()
    Dim strEntrySub As String
    Dim strEntryEmpNo As String

    If gDone Then
' Unload Me
    gDone = False
    End If
    
    'Code to make sure only one version of this is running
    If mintLoaded <> 1 Then
        mintLoaded = 1
        'Bring up the code to show the menu
        While gDone = False
            Call basEntry
        Wend
        mintLoaded = 0
        Unload Me
    End If
End Sub


'
'
Public Sub basEntry()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmLabor.basEntry
'Description      : Routine that processes direct labor for all divisions.
'                   Boolean values are globally set by division and type of
'                   Labor entry, to determine which options are used for the
'                   labor process.
'Parameters       : None
'Called From      : frmDirLabor.Form_Activate
'                   This Form is Loaded directly from frmDCMain during scanning
'                   of labor, and also from frmInpEmplNo after entering
'                   employee number in the normal labor input process.
'Procedures Used  : Init_varbs, GetSysDateTime, gEntry.Add, basGetEntryValue,
'                   checkClockOut
'Forms Loaded     : frmDirLabor, frmIndLabor, frmScrap, frmDCCLabor
'                   getBatchIncn, DirLaborReEnter, CalcActHrs, AnotherRec,
'                   Build_Labor_Recs, ClockOut, Upd_Empl_Attd, ErrorBeep, ErrBox
'Return Parameters: None
'Tables Updates   : Updates done in Build_Labor_Recs, Upd_Empl_Attd
'
'Revisions :
'   10/01/98    JC - Commented out the gDone and the Unload lines, Bristol wants _
                The app to just return to the Labor Menu
                
'==============================================================================
    'Procedure called upon form activation to call up the labor
    '    menu selections based upon company
    Dim strEntrySub As String
    Dim vEntrySub As Variant
    Dim strEntryEmpNo As String
    
    Call Init_varbs         'Initialize gLaborIn Array
    Call GetSysDateTime     'save current time
    
    Set gEntry = New clsEntry
    gEntry.Title = LoadResString(gLanguageOffset + 138)
    
    If Me.Tag = 1 Then 'engineering labor - used by employees who are not required to clock in/out
            Call GetSysDateTime     'save current time
             gEmployee.Actdate = gSysDate 'Needed since employee does not clock in or out
             gEmployee.GenShift = 1    'Also needed since no clock in/out
   End If
    
    If div = bristol Then
        gEntry.Add "S", "Single", lmMachLabor
        lblSelect(0).Caption = "S  for Single"
        gEntry.Add "D", "Direct Batch", lmDirBatch
        lblSelect(1).Caption = "D  for Direct Batch"
        gEntry.Add "I", "Indirect Batch", lmIndirEnt
        lblSelect(2).Caption = "I  for Indirect Batch"
        gEntry.Add "C", "scrap Counts", lmScrap
        lblSelect(3).Caption = "C  for scrap Counts"
'        gEntry.Add "1", LoadResString(gLanguageOffset + 202), lmStartIndir
'        lblSelect(0).Caption = LoadResString(gLanguageOffset + 140)
'        gEntry.Add "2", LoadResString(gLanguageOffset + 203), lmStartDir
'        lblSelect(1).Caption = LoadResString(gLanguageOffset + 141)
'        gEntry.Add "3", LoadResString(gLanguageOffset + 204), lmEndIndir
'        lblSelect(2).Caption = LoadResString(gLanguageOffset + 142)
'        gEntry.Add "4", LoadResString(gLanguageOffset + 205), lmEndDir
'        lblSelect(3).Caption = LoadResString(gLanguageOffset + 143)
        gEntry.Add "E", "E/Esc - Escape to Main Menu", lmEsc
        lblSelect(4).Caption = "E/Esc  to  Escape to Main Menu"
    ElseIf div = monterrey Then
        gEntry.Add "I", "Mano de Obra Indirecta", lmIndirEnt
        lblSelect(0).Caption = LoadResString(gLanguageOffset + 300)
        gEntry.Add "D", "Mano de Obra Directa", lmDirLabor
        lblSelect(1).Caption = LoadResString(gLanguageOffset + 301)
        gEntry.Add "C", "Cargo Directo a Centro de Costos", lmDirCC
        lblSelect(2).Caption = "C para Cargo Directo a Centro de Costos"
        gEntry.Add "S", "Peso de Desperdicio", lmScrap
        lblSelect(3).Caption = LoadResString(gLanguageOffset + 302)
         gEntry.Add "E", "Regresal al Menu Principal", lmEsc
         lblSelect(4).Caption = LoadResString(gLanguageOffset + 303)
        lblSelect(5).Caption = ""
        lblSelect(6).Caption = ""
    ElseIf div = syracuse Then
        gEntry.Add "I", "Indirect Labor", lmIndirEnt
        lblSelect(0).Caption = "I for Indirect Labor"
        gEntry.Add "D", "Direct Labor", lmDirLabor
        lblSelect(1).Caption = "D for Direct Labor"
        gEntry.Add "S", "Scrap counts", lmScrap
        lblSelect(2).Caption = "S for additional Scrap counts"
        gEntry.Add "E", "E/Esc - Escape to Main Menu", lmEsc
        lblSelect(3).Caption = "E/Esc to Escape to Main Menu"
        lblSelect(4).Caption = ""
        lblSelect(5).Caption = ""
    ElseIf div = corry Then
        gEntry.Add "L", "Direct Labor; One or More Mach.; hours are Calculated", lmDirLabor
        lblSelect(0).Caption = "L for Direct Labor; One or More Mach.; hours are Calculated"
        gEntry.Add "D", "Direct Cost Center Charge", lmDirCC
        lblSelect(1).Caption = "D for Direct Cost Center Charge"
        gEntry.Add "I", "Indirect Labor", lmIndirCalc
        lblSelect(2).Caption = "I for Indirect Labor"
        gEntry.Add "B", "Batch Direct Labor; One or More Mach.; hours are Entered", lmDirBatch
        lblSelect(3).Caption = "B for Batch Direct Labor; One or More Mach.; hours are Entered"
       'gEntry.Add "S", "Additional Scrap Counts", lmScrap
       'lblSelect(4).Caption = "Additional Scrap Counts"
        gEntry.Add "E", "E/Esc - Escape to Main Menu", lmEsc
        lblSelect(5).Caption = "E/Esc to Escape to Main Menu"
    ElseIf div = saline Then
        gEntry.Add "I", "Indirect Labor", lmIndirEnt
        lblSelect(0).Caption = "I for Indirect Labor"
        gEntry.Add "D", "Direct Labor", lmDirLabor
        lblSelect(1).Caption = "D for Direct Labor"
        gEntry.Add "S", "Scrap counts", lmScrap
        lblSelect(2).Caption = "S for additional Scrap counts"
        gEntry.Add "E", "E/Esc - Escape to Main Menu", lmEsc
        lblSelect(3).Caption = "E/Esc to Escape to Main Menu"
        lblSelect(4).Caption = ""
        lblSelect(5).Caption = ""
    ElseIf div = milw Then
        gEntry.Add "I", "Indirect Labor", lmIndirEnt
        lblSelect(0).Caption = "I for Indirect Labor"
        gEntry.Add "D", "Direct Labor", lmDirLabor
        lblSelect(1).Caption = "D for Direct Labor"
        gEntry.Add "E", "E/Esc - Escape to Main Menu", lmEsc
        lblSelect(2).Caption = "E/Esc to Escape to Main Menu"
        lblSelect(3).Caption = ""
        lblSelect(4).Caption = ""
        lblSelect(5).Caption = ""
    ElseIf div = remesa Then
        gEntry.Add "I", "Mano de Obra Indirecta", lmIndirEnt
        lblSelect(0).Caption = LoadResString(gLanguageOffset + 300)
        gEntry.Add "D", "Mano de Obra Directa", lmDirLabor
        lblSelect(1).Caption = LoadResString(gLanguageOffset + 301)
        gEntry.Add "S", "Peso de Desperdicio", lmScrap
        lblSelect(2).Caption = LoadResString(gLanguageOffset + 302)
         gEntry.Add "E", "Regresal al Menu Principal", lmEsc
         lblSelect(3).Caption = LoadResString(gLanguageOffset + 303)
        lblSelect(4).Caption = ""
        lblSelect(5).Caption = ""
    ElseIf div = singapore Then
        gEntry.Add "I", "Indirect Labor", lmIndirEnt
        lblSelect(0).Caption = "I for Indirect Labor"
        gEntry.Add "D", "Direct Labor", lmDirLabor
        lblSelect(1).Caption = "D for Direct Labor"
        gEntry.Add "S", "Scrap counts", lmScrap
        lblSelect(2).Caption = "S for additional Scrap counts"
        gEntry.Add "E", "E/Esc - Escape to Main Menu", lmEsc
        lblSelect(3).Caption = "E/Esc to Escape to Main Menu"
        lblSelect(4).Caption = ""
        lblSelect(5).Caption = ""
    ElseIf div = pv_singapore Then
        gEntry.Add "I", "Indirect Labor", lmIndirEnt
        lblSelect(0).Caption = "I for Indirect Labor"
        gEntry.Add "D", "Direct Labor", lmDirLabor
        lblSelect(1).Caption = "D for Direct Labor"
        gEntry.Add "S", "Scrap counts", lmScrap
        lblSelect(2).Caption = "S for additional Scrap counts"
        gEntry.Add "E", "E/Esc - Escape to Main Menu", lmEsc
        lblSelect(3).Caption = "E/Esc to Escape to Main Menu"
        lblSelect(4).Caption = ""
        lblSelect(5).Caption = ""
    
    ElseIf div = hyson_products Then
      Select Case frmLabor.Tag
         Case 0
             gEntry.Add "I", "Indirect Labor - Hours Calculated", lmIndirCalc
             lblSelect(0).Caption = "I   for Indirect Labor - Calculated Time"
             gEntry.Add "J", "Indirect Labor - Hours Entered", lmIndirEnt
             lblSelect(1).Caption = "J   for Indirect Labor - Entered Time"
            gEntry.Add "C", "Direct Labor - Calculated Time", lmDirLabor
            lblSelect(2).Caption = "C   for Direct Labor - Calculated Time"
            gEntry.Add "D", "Direct Labor - Entered Time", lmDirBatch
            lblSelect(3).Caption = "D   for Direct Labor - Entered Time"
            gEntry.Add "E", "E/Esc - Escape to Main Menu", lmEsc
            lblSelect(4).Caption = LoadResString(gLanguageOffset + 303)
            lblSelect(5).Caption = ""
'           lblSelect(6).Caption = ""
         Case 1
            gEntry.Add "D", "Engineering Labor", lmDirBatch
            lblSelect(0).Caption = "D  for Engineering Labor"
            gEntry.Add "E", "E/Esc - Escape to Main Menu", lmEsc
            lblSelect(1).Caption = LoadResString(gLanguageOffset + 303)
            lblSelect(2).Caption = ""
            lblSelect(3).Caption = ""
            lblSelect(4).Caption = ""
            lblSelect(5).Caption = ""
      End Select
    ElseIf div = stroms Then
      Select Case frmLabor.Tag
         Case 0
            gEntry.Add "1", LoadResString(gLanguageOffset + 202), lmStartIndir
            lblSelect(0).Caption = LoadResString(gLanguageOffset + 140)
            gEntry.Add "2", LoadResString(gLanguageOffset + 203), lmStartDir
            lblSelect(1).Caption = LoadResString(gLanguageOffset + 141)
            gEntry.Add "3", LoadResString(gLanguageOffset + 204), lmEndIndir
            lblSelect(2).Caption = LoadResString(gLanguageOffset + 142)
            gEntry.Add "4", LoadResString(gLanguageOffset + 205), lmEndDir
            lblSelect(3).Caption = LoadResString(gLanguageOffset + 143)
            gEntry.Add "5", LoadResString(gLanguageOffset + 287), lmOSSendOut
            lblSelect(4).Caption = LoadResString(gLanguageOffset + 290)
            gEntry.Add "6", " ", lmStdLbr
            lblSelect(5).Caption = LoadResString(gLanguageOffset + 296)
            
            gEntry.Add "E", LoadResString(gLanguageOffset + 144), lmEsc
            lblSelect(6).Caption = LoadResString(gLanguageOffset + 144)
     
'            lblSelect(7).Caption = ""
'            lblSelect(6).Caption = ""
         Case 1
            gEntry.Add "D", "Engineering Labor", lmDirBatch
            lblSelect(0).Caption = "D  for Engineering Labor"
            gEntry.Add "E", LoadResString(gLanguageOffset + 144), lmEsc
            lblSelect(1).Caption = LoadResString(gLanguageOffset + 144)
            lblSelect(2).Caption = ""
            lblSelect(3).Caption = ""
            lblSelect(4).Caption = ""
            lblSelect(5).Caption = ""
      End Select
    Else
        gEntry.Add "E", LoadResString(gLanguageOffset + 144), lmEsc
        lblSelect(0).Caption = LoadResString(gLanguageOffset + 144)
        lblSelect(1).Caption = ""
        lblSelect(2).Caption = ""
        lblSelect(3).Caption = ""
        lblSelect(4).Caption = ""
        lblSelect(5).Caption = ""
    End If
       
    vEntrySub = basGetEntryValue()
    If gEntry.Escaped Then
        vEntrySub = lmEsc
    End If
    'JC    Set gEntry = Nothing
    Call GetSysDateTime     'save current time
  
    '==================================================================
    '   JC 01/25/99
    '   What type of Labor to do
    '==================================================================
    Select Case vEntrySub
        Case lmDirLabor
            'Direct Labor
            If div = corry Then     'Corry Direct Labor
                gbNumbMach = True
                gbNoLoads = False
                gbInpSetOrMach = False   'Elapsed time calculation
                gbAnotherRec = False
                gbMachRecY = False
            ElseIf div = milw Or div = saline Or div = remesa Or _
                   div = syracuse Or div = singapore Or div = pv_singapore Then
                Rem no changes for this selection
           ElseIf div = hyson_products Then
                gbInpSetOrMach = False   'Elapsed time calculation
                gbMachRecY = False
                gbAnotherRec = False
            ElseIf div = stroms Then
                gbInpSetOrMach = False   'Elapsed time calculation
                gbMachRecY = False
                gbAnotherRec = False
            Else
                MsgBox LoadResString(gLanguageOffset + 56)
            End If
            okToPrompt = False
            'Call checkClockOut
            frmDirLabor.Show 1
        Case lmStdLbr
            gbInpSetOrMach = False   'Elapsed time calculation
            gbMachRecY = True
            gbAnotherRec = False
            gbNumbMach = False
            gbNoLoads = False
            okToPrompt = False
            frmStdLbr.Show 1
        Case lmReworkLabor
            'Rework Labor
            If div = bristol Then
                MsgBox "Bristol Rework Not Coded Yet"
                gDone = True
                Unload Me
                Exit Sub
            Else
                MsgBox LoadResString(gLanguageOffset + 56)
            End If
            okToPrompt = False
            'Call checkClockOut
            frmDirLabor.lblFormTitle = "Rework Labor"
            frmDirLabor.Show 1
            If gbRecTypeR = True Then
                gbRecTypeR = False
            End If
            
            
        Case lmMachLabor
            If div = bristol Then        'Bristol Single Direct
                gbInpSetOrMach = False   'Elapsed time calculation
                gbMachRecY = False
            Else
                MsgBox LoadResString(gLanguageOffset + 56)
            End If
            okToPrompt = False
            'Call checkClockOut
            frmDirLabor.Show 1
            
        Case lmDirBatch
            If div = bristol Then       'Bristol Direct Batch
                gbInpSetOrMach = True   'hours are entered
                gbMachRecY = True
            ElseIf div = corry Then
                gbNumbMach = False
                gbInpSetOrMach = True   'hours are entered
                gbNoLoads = True
                gbAnotherRec = True
                gbMachRecY = True
            ElseIf div = hyson_products Then
                gbInpSetOrMach = True
                gbMachRecY = False
                gbAnotherRec = False
            ElseIf div = stroms Then
                gbInpSetOrMach = True
                gbMachRecY = False
                gbAnotherRec = False
            Else
                MsgBox LoadResString(gLanguageOffset + 56)
            End If
            okToPrompt = True
            'Call checkClockOut
            frmDirLabor.Show 1
            
        Case lmIndirEnt
            'Call checkClockOut
            If div = milw Then
               frmIndLabor.lblVarbFld1.Caption = "Workcenter"
               frmIndLabor.lblVarbFld2.Caption = "Maint Reason"
               frmIndLabor.lblVarbFld3.Caption = ""
            ElseIf div = bristol Then
                 frmIndLabor.lblVarbFld1.Caption = "PM Job"
                 frmIndLabor.lblVarbFld2.Caption = ""
                 frmIndLabor.lblVarbFld3.Caption = ""
            ElseIf div = corry Then
                 frmIndLabor.lblVarbFld1.Caption = "Incen Type"
                 frmIndLabor.lblVarbFld2.Caption = "Status"
                 frmIndLabor.lblVarbFld3.Caption = "EO Job"
            ElseIf div = hyson_products Then
                 frmIndLabor.lblVarbFld1.Caption = ""
'                 frmIndLabor.lblVarbFld1.Caption = "JBA Work Order"
                 frmIndLabor.lblVarbFld2.Caption = ""
                 frmIndLabor.lblVarbFld3.Caption = ""
                 gbInpIndirHrs = True
            Else
                 frmIndLabor.lblVarbFld1.Caption = ""
                 frmIndLabor.lblVarbFld2.Caption = ""
                 frmIndLabor.lblVarbFld3.Caption = ""
            End If
            frmIndLabor.Show 1
        Case lmIndirCalc
            'Call checkClockOut
            If div = corry Then
                 frmIndLabor.lblVarbFld1.Caption = "Incen Type"
                 frmIndLabor.lblVarbFld2.Caption = "Status"
                 frmIndLabor.lblVarbFld3.Caption = "EO Job"
            ElseIf div = hyson_products Then
                 frmIndLabor.lblVarbFld1.Caption = ""
'                 frmIndLabor.lblVarbFld1.Caption = "JBA Work Order"
                 frmIndLabor.lblVarbFld2.Caption = ""
                 frmIndLabor.lblVarbFld3.Caption = ""
                 gbInpIndirHrs = False
            Else
                 frmIndLabor.lblVarbFld1.Caption = ""
                 frmIndLabor.lblVarbFld2.Caption = ""
                 frmIndLabor.lblVarbFld3.Caption = ""
            End If
            frmIndLabor.Show 1
            
        Case lmScrap
            frmScrap.Show 1
            
        Case lmDirCC
            'Call checkClockOut
            frmDCCLabor.Show 1
            
        Case lmStartIndir       'on 6/6/2000, used only by stromsholmen
           If OpenIndirectLabor(gEmployee.EmplNo, gEmployee.Actdate) Then
                  ErrBox LoadResString(gLanguageOffset + 18), 2
           Else
                  frmStartIndir.Show 1
            End If
         
         Case lmStartDir      'on 6/6/2000, used only by stromsholmen
               gbAnotherRec = False
               frmStartDirect.Show 1
          
          Case lmEndDir      'on 6/6/2000, used only by stromsholmen
               gbAnotherRec = False
               frmEndDirect.Show 1
         
         Case lmEndIndir     'on 6/6/2000, used only by stromsholmen
            If Not OpenIndirectLabor(gEmployee.EmplNo, gEmployee.Actdate) Then
                  ErrBox LoadResString(gLanguageOffset + 71), 2
            Else
                 frmEndIndirect.Show 1
            End If
        Case lmOSSendOut
            frmOSSend.Show 1
            
        Case lmEsc
            gDone = True
            Unload Me
            Exit Sub
            
        Case 999
'            SqlClose (gSQLConn)
    End Select
    
    
    ' Setting gDone to false here allows  Labor menu to remain active
    'even if ESC was pressed during data entry
    Select Case div
        'Some Divisions want to go back,
        'some want to stay...some don't want to do anything.
        Case corry
            'Go Back
            gDone = True
            Unload Me
        Case stroms
            'Go Back
            gDone = True
            Unload Me
        Case Else
            'Stay
            gDone = False
    End Select


End Sub



Private Sub Form_Load()
lblFormTitle.Caption = LoadResString(gLanguageOffset + 138)
lblPressCaption.Caption = LoadResString(gLanguageOffset + 139)
End Sub


