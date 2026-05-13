VERSION 5.00
Begin VB.Form frmDirLabor 
   Caption         =   "Data Collection - Labor Entry"
   ClientHeight    =   7215
   ClientLeft      =   1875
   ClientTop       =   165
   ClientWidth     =   13725
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7215
   ScaleWidth      =   13725
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Label lblMchStdType 
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   9120
      TabIndex        =   54
      Top             =   4320
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   2790
      Picture         =   "Dcdirlbr.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   870
   End
   Begin VB.Label lblUnplanned 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00C000C0&
      Height          =   270
      Left            =   2835
      TabIndex        =   53
      Top             =   3540
      Width           =   60
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   270
      Left            =   540
      TabIndex        =   52
      Top             =   5940
      Width           =   60
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   270
      Left            =   540
      TabIndex        =   51
      Top             =   6300
      Width           =   60
   End
   Begin VB.Label lblBatchMethod 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   8055
      TabIndex        =   50
      Top             =   5745
      Width           =   60
   End
   Begin VB.Label lblCycleTime 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   4980
      TabIndex        =   49
      Top             =   5745
      Width           =   60
   End
   Begin VB.Label lblNumbOfLoads 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   1860
      TabIndex        =   48
      Top             =   5745
      Width           =   60
   End
   Begin VB.Label lblMethodCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Method"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   7110
      TabIndex        =   47
      Top             =   5745
      Width           =   810
   End
   Begin VB.Label lblCycleTimeCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cycle Time"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3600
      TabIndex        =   46
      Top             =   5745
      Width           =   1170
   End
   Begin VB.Label lblNoLoadsCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "No. of Loads"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   375
      TabIndex        =   45
      Top             =   5745
      Width           =   1335
   End
   Begin VB.Label lblScrapQty 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   8055
      TabIndex        =   44
      Top             =   5100
      Width           =   60
   End
   Begin VB.Label lblScrapQtyCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Scrap Quantity"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   6345
      TabIndex        =   43
      Top             =   5100
      Width           =   1560
   End
   Begin VB.Label lblProdQty 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   1860
      TabIndex        =   42
      Top             =   5100
      Width           =   60
   End
   Begin VB.Label lblQuantityCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Quantity"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   825
      TabIndex        =   41
      Top             =   5100
      Width           =   885
   End
   Begin VB.Label lblMachHrs 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   8160
      TabIndex        =   40
      Top             =   4680
      Width           =   60
   End
   Begin VB.Label lblAddonHrs 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   4920
      TabIndex        =   39
      Top             =   4680
      Width           =   15
   End
   Begin VB.Label lblSetupHrs 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   1920
      TabIndex        =   38
      Top             =   4680
      Width           =   60
   End
   Begin VB.Label lblMachineHrsCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Machine Hrs."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   6600
      TabIndex        =   37
      Top             =   4680
      Width           =   1395
   End
   Begin VB.Label lblAddonHrsCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Addon Hrs."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3600
      TabIndex        =   36
      Top             =   4680
      Width           =   1170
   End
   Begin VB.Label lblSetupHrsCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Setup Hrs."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   600
      TabIndex        =   35
      Top             =   4680
      Width           =   1125
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   10545
      Y1              =   4600
      Y2              =   4600
   End
   Begin VB.Label lblEorS 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   5175
      TabIndex        =   34
      Top             =   4320
      Width           =   15
   End
   Begin VB.Label lblMachStd 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   8880
      TabIndex        =   33
      Top             =   4320
      Width           =   45
   End
   Begin VB.Label lblFixStd 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   2760
      TabIndex        =   32
      Top             =   4320
      Width           =   15
   End
   Begin VB.Label lblMachineStdCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Machine Std."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   6360
      TabIndex        =   31
      Top             =   4320
      Width           =   1380
   End
   Begin VB.Label lblESCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "E/S"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4590
      TabIndex        =   30
      Top             =   4320
      Width           =   360
   End
   Begin VB.Label lblFixedStdCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fixed Std."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1080
      TabIndex        =   29
      Top             =   4320
      Width           =   1560
   End
   Begin VB.Label lblCCID 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   5835
      TabIndex        =   28
      Top             =   3900
      Width           =   60
   End
   Begin VB.Label lblOperStatus 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   7875
      TabIndex        =   27
      Top             =   3900
      Width           =   60
   End
   Begin VB.Label lblCCIDCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cost Center"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4395
      TabIndex        =   26
      Top             =   3900
      Width           =   1290
   End
   Begin VB.Label lblStatusCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Status"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   7080
      TabIndex        =   25
      Top             =   3900
      Width           =   675
   End
   Begin VB.Label lblUnplnReas 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   4950
      TabIndex        =   24
      Top             =   5340
      Width           =   60
   End
   Begin VB.Label lblOperCd 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   5685
      TabIndex        =   23
      Top             =   3540
      Width           =   60
   End
   Begin VB.Label lblUnplanReasCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Unplan Reas."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3375
      TabIndex        =   22
      Top             =   5340
      Width           =   1410
   End
   Begin VB.Label lblOperCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Oper."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4935
      TabIndex        =   21
      Top             =   3540
      Width           =   600
   End
   Begin VB.Label lblWCID 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   2760
      TabIndex        =   20
      Top             =   3900
      Width           =   60
   End
   Begin VB.Label lblSeqNo 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   2970
      TabIndex        =   19
      Top             =   3540
      Width           =   60
   End
   Begin VB.Label lblWCIDCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Workcenter ID"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1080
      TabIndex        =   18
      Top             =   3900
      Width           =   1575
   End
   Begin VB.Label lblSeqNoCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Seq No."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1200
      TabIndex        =   17
      Top             =   3540
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   135
      X2              =   10440
      Y1              =   3465
      Y2              =   3465
   End
   Begin VB.Label lblOrdQty 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   9420
      TabIndex        =   16
      Top             =   2700
      Width           =   60
   End
   Begin VB.Label lblOrderQtyCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Order Qty"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   8160
      TabIndex        =   15
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label lblCustName 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   4200
      TabIndex        =   14
      Top             =   3075
      Width           =   60
   End
   Begin VB.Label lblCustomerCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3045
      TabIndex        =   13
      Top             =   3075
      Width           =   1035
   End
   Begin VB.Label lblPartRev 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   9420
      TabIndex        =   12
      Top             =   3075
      Width           =   60
   End
   Begin VB.Label lblRevCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Rev."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   8790
      TabIndex        =   11
      Top             =   3075
      Width           =   495
   End
   Begin VB.Label lblOrdType 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   2115
      TabIndex        =   10
      Top             =   3075
      Width           =   60
   End
   Begin VB.Label lblTypeCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Type"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1485
      TabIndex        =   9
      Top             =   3075
      Width           =   525
   End
   Begin VB.Label lblPartNo 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   4185
      TabIndex        =   8
      Top             =   2700
      Width           =   60
   End
   Begin VB.Label lblPartCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Part"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3390
      TabIndex        =   7
      Top             =   2700
      Width           =   690
   End
   Begin VB.Label lblJobNo 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   2115
      TabIndex        =   6
      Top             =   2700
      Width           =   60
   End
   Begin VB.Label lblJobNoCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Job No."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   840
      TabIndex        =   5
      Top             =   2700
      Width           =   1170
   End
   Begin VB.Label lblEmplName 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   4950
      TabIndex        =   4
      Top             =   2160
      Width           =   60
   End
   Begin VB.Label lblNameCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3915
      TabIndex        =   3
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label lblEmplNo 
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H00004040&
      Height          =   270
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   60
   End
   Begin VB.Label lblEmplNoCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Empl No."
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   675
      TabIndex        =   1
      Top             =   2160
      Width           =   1710
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      Caption         =   "Direct Labor"
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
      Left            =   3840
      TabIndex        =   0
      Top             =   1485
      Width           =   3525
   End
End
Attribute VB_Name = "frmDirLabor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mintLoaded As Integer
Dim bScrapColl As Boolean

Const vbColorWhite = &H8000005

Public Function inpLabJobSeq(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDirLabor.inpLabJobSeq(i)
'Description      : Routine for the entry of Job and Seq No, for a single
'                   machine entry for direct labor for all divisions. Boolean
'                   values are globally set by division and type of
'                   Labor entry, to determine which options are used for the
'                   labor process.
'Parameters       : None
'Called From      : frmDirLabor.DirLaborProcess
'Procedures Used  : GetJobMaster, getLabUnplnData, GetJobOper, GetWorkCenter,
'                   ValidNumber, basGetEntry, ErrorBeep, ErrBox
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    Dim strEntry As String
    Dim strDataId As String
    Dim strSeqNo As String
    
    inpLabJobSeq = False        'to force loop back till done
    
    Do Until inpLabJobSeq Or gDone
        Do Until inpLabJobSeq Or gDone       'Extra loop for exit do stmts
         '   ErrBox "If labor is unplanned, key in only job number and Enter"
            
            Set gEntry = New clsEntry
            gEntry.Title = LoadResString(gLanguageOffset + 165)
            gEntry.EntryType = 0        'Any Printable Characters
            strEntry = basGetEntry()
            If gEntry.Escaped Then
                Set gEntry = Nothing
                gDone = True
                Exit Function
            End If
            Set gEntry = Nothing
            
            strDataId = UCase(Left(strEntry, 2))
            If strDataId = "ZZ" Then
                gDone = True
                Exit Function
            End If
            
            If strDataId = "NU" Then
                ErrBox LoadResString(gLanguageOffset + 28), 2
                ErrorBeep
                Exit Do           'Loop Again
            End If
            
            If Len(strDataId) < 1 Then
                ErrBox LoadResString(gLanguageOffset + 78), 2
                ErrorBeep
                Exit Do           'Loop Again
            End If
               
            If Len(strEntry) > 11 Then
                If strDataId = "JS" Then
                    gLaborIn(iEntry).OrderNo = UCase(Mid(strEntry, 3, 7))
                    strSeqNo = Mid(strEntry, 10, 4)
                Else
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            ElseIf Len(strEntry) = 11 Then
                gLaborIn(iEntry).OrderNo = UCase(Left(strEntry, 7))
                'gLaborIn(iEntry).OrderNo = UCase(Mid(strEntry, 3, 7))
                strSeqNo = Mid(strEntry, 8, 4)
                'strSeqNo = Mid(strEntry, 10, 4)
           ElseIf Len(strEntry) = 10 Then
                gLaborIn(iEntry).OrderNo = UCase(Left(strEntry, 7))
                strSeqNo = Mid(strEntry, 8, 3)
            ElseIf Len(strEntry) = 9 Then   'likely work order number preceded by JS, and no seq #
                If strDataId = "JS" Then
                    gLaborIn(iEntry).OrderNo = UCase(Mid(strEntry, 3, 7))
                    strSeqNo = "000"
                Else
                   gLaborIn(iEntry).OrderNo = UCase(Left(strEntry, 7))
                   strSeqNo = Mid(strEntry, 8, 2)
                End If
            Else
                gLaborIn(iEntry).OrderNo = UCase(Left(strEntry, 7))  'likely an order number with no data ident.
                strSeqNo = "000"
            End If
            
            If Not GetJobMaster(gLaborIn(iEntry).OrderNo) Then
                If Left(gLaborIn(iEntry).OrderNo, 2) = "09" Then
                    Screen.ActiveForm.lblErrMsg.Caption = LoadResString(gLanguageOffset + 236)
                    Screen.ActiveForm.lblErrMsg.BackColor = vbColorWhite
                    strSeqNo = "000 "
                    gLaborIn(iEntry).MatlWgt = 0
                    gJobMaster.Actdate = 0
                    gJobMaster.OrderNo = gLaborIn(iEntry).OrderNo
                    gJobMaster.OrderType = "SO"
                Else
                    ErrBox LoadResString(gLanguageOffset + 61), 2
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            Else
                If gJobMaster.OrderStat = "C" Then
                       ErrBox "Order status is complete", 2
                       ErrorBeep
                       Exit Do           'Loop Again
                Else
                       gLaborIn(iEntry).MatlWgt = gJobMaster.MatlWgt
                End If
            End If
            
            If Len(strSeqNo) = 0 Then
                strSeqNo = "000 "
            Else
                If Not ValidNumber(strSeqNo) Then
                    ErrBox LoadResString(gLanguageOffset + 111), 2
                    ErrorBeep
                    Exit Do           'Loop Again
                End If
            End If
            
            gLaborIn(iEntry).SeqNo = CInt(strSeqNo)
            If gLaborIn(iEntry).SeqNo = 990 _
               Or gLaborIn(iEntry).SeqNo = 0 Then
                   gLaborIn(iEntry).SeqNo = 0
                   gLaborIn(iEntry).EOrS = "S"
                   lblErrMsg.Caption = ""
                   lblUnplanned = LoadResString(gLanguageOffset + 240)
                   If Not getLabUnplnData(iEntry) Then
                      Exit Do           'Loop Again
                   End If
                   inpLabJobSeq = True     'Done - Unplanned
            Else
                lblUnplanned = ""
                If Not GetJobOper(gLaborIn(iEntry).OrderNo, _
                        gLaborIn(iEntry).SeqNo) Then
                    ErrBox LoadResString(gLanguageOffset + 87), 2
                    ErrorBeep
                    Exit Do
                ElseIf gJobOpers.WCID = "-----" Then
                        ErrBox LoadResString(gLanguageOffset + 38), 2
                        ErrorBeep
                        Exit Do
                ElseIf gJobOpers.OperStatus = "C" Then
                        ErrBox LoadResString(gLanguageOffset + 80), 2
                        ErrorBeep
                        Exit Do
                ElseIf Not GetWorkCenter(gJobOpers.WCID) Then
                        ErrBox LoadResString(gLanguageOffset + 103), 2
                        ErrorBeep
                        Exit Do
                ElseIf gJobOpers.WCID = "OSSRV" And div = bristol Then
                       ErrBox "Cannot enter labor for outside service", 2
                       ErrorBeep
                       Exit Do
                ElseIf Left(gJobOpers.WCID, 1) = "V" And div = bristol Then
                       ErrBox "Cannot enter labor for outside service", 2
                       ErrorBeep
                       Exit Do
                End If
                gLaborIn(iEntry).UnplnReas = "  "
                gLaborIn(iEntry).WCID = gJobOpers.WCID
                gLaborIn(iEntry).Oper = gJobOpers.Oper
                gLaborIn(iEntry).EOrS = gJobOpers.EorSStd
                gLaborIn(iEntry).SetStd = gJobOpers.SetStd
                gLaborIn(iEntry).MchStd = gJobOpers.MchStd
                gLaborIn(iEntry).CCID = gWorkCenter.CCID
                inpLabJobSeq = True     'Done - Planned
            End If
        Loop
    Loop
 lblErrMsg.Caption = ""
End Function

Public Function getLabUnplnData(iEntry As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDirLabor.getLabUnplnData(i)
'Description      : Routine for the entry of Unplanned Data, for a single
'                   machine entry for direct labor for all divisions. Boolean
'                   values are globally set by division and type of
'                   Labor entry, to determine which options are used for the
'                   labor process.
'Parameters       : None
'Called From      : frmDirLabor.DirLaborProcess
'Procedures Used  : ErrorChk, inpJobAndSeq, dispValues, inpWCID, GetCostCenter,
'                   inpProdQty, getScrapCounts, inpSetHrs, inpMchHours, inpAddonHours,
'                   inpEndOfShift, CallCmpl, inpNoOfLoads, CalcBtErnHrs, CalcBatchHrs,
'                   inpCycleTime, CalcBatchHrs, getBatchIncn, ErrorBeep, ErrBox
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    getLabUnplnData = False
    Call initJobOpers
    
    Screen.ActiveForm.lblJobNo.Caption = gJobMaster.OrderNo
    Do Until gDone Or getLabUnplnData
        'gDone = True        'Default for exit to Main Menu
        If Not inpWCID(iEntry) Then
            gDone = True
            Exit Function
        End If
        Screen.ActiveForm.lblWCID.Caption = gLaborIn(iEntry).WCID
        'JC New Here
        gLaborIn(iEntry).CCID = gWorkCenter.CCID
        If Not GetCostCenter(gWorkCenter.CCID) Then
            gCostCenter.PaidLunch = "N"
        Else
            Screen.ActiveForm.lblCCID.Caption = gWorkCenter.CCID
        End If
        
        
        If Not GetWorkCenter(gLaborIn(iEntry).WCID) Then
            gCostCenter.PaidLunch = "N"
        End If
        
        If Not inpOperNumb(iEntry) Then
            gDone = True
            Exit Function
        End If

        Screen.ActiveForm.lblOperCd.Caption = gLaborIn(iEntry).Oper
        gLaborIn(iEntry).MchStd = 0
        gLaborIn(iEntry).SetStd = 0
        
        If gbInpStds Then
          If Not inpFixStd(iEntry) Then
              gDone = True
              Exit Function
          End If
          Screen.ActiveForm.lblFixStd.Caption = gLaborIn(iEntry).SetStd
          If gLaborIn(iEntry).SetStd = 0 Then
              If Not inpMchStd(iEntry) Then
                  gDone = True
                  Exit Function
              End If
          End If
          Screen.ActiveForm.lblMachStd.Caption = gLaborIn(iEntry).MchStd
        End If
        If Not inpUnplnReas(iEntry) Then Exit Function
        Screen.ActiveForm.lblUnplnReas.Caption = gLaborIn(iEntry).UnplnReas
        
        getLabUnplnData = True 'Exit loop
    Loop
End Function
Private Sub Form_Activate()
    lblEmplNo = gEmployee.EmplNo
    lblEmplName = gEmployee.EmplName
    
    If gDone Then
        Unload Me
    End If
    
    If mintLoaded <> 1 Then
        mintLoaded = 1
        Call DirLaborProcess
        mintLoaded = 0
        Unload Me
    End If
End Sub
Public Sub DirLaborProcess()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDirLabor.DirLaborProcess
'Description      : Routine that processes direct labor for all divisions.
'                   Boolean values are globally set by division and type of
'                   Labor entry, to determine which options are used for the
'                   labor process.
'Parameters       : None
'Called From      : frmDirLabor.Form_Activate
'Procedures Used  : inpEndOfShift, inpNumbMachs, inpLabJobSeq, inpSetHours,
'                   inpMchHours, inpAddonHours, inpProdQty, getScrapCounts,
'                   CallCmpl, inpNoOfLoads, inpCycleTime, CalcBatchHrs,
'                   getBatchIncn, DirLaborReEnter, CalcActHrs, AnotherRec,
'                   Build_Labor_Recs, ClockOut, Upd_Empl_Attd, ErrorBeep,
'                   ErrBox
'Return Parameters: None
'Tables Updates   : Updates done in Build_Labor_Recs, Upd_Empl_Attd
'==============================================================================
    
    Dim strEntry As String
    
    Dim I As Integer
        
    Dim lScrapTot As Long
    Dim CalcHours As Currency
    Dim strRecType As String
    Dim j As Integer
    Dim bEnter As Boolean
    Dim bZeroIt As Boolean
    Dim bNoQty As Boolean
    
'gEOSFlag is set to true by function CheckClockOut
'when the current time is within a window before
'and after the shift Clock-Out time (from the shift
'parameter record). The upper and lower bounds of
'the window are defined by the values in EOSStart
'and EOSEnd fields, also in the shift param rec.
'If EOSStart and EOSEnd are zero, gEOSFlag is set
'to false.

'InpEndOfShift is set to True when the user answers
'the prompt "Clocking Out?". Additionally, the
'boolean variable empClkOutOK is set to true if
'the user answers Yes to the prompt, or false if
'the user answers No to the prompt.
   Call checkClockOut
   
    If gbEOShift And gEosFlag Then
        If Not inpEndOfShift() Then
            gDone = True
            Unload Me
            Exit Sub
        End If
    End If
       
    If gbAnotherRec Then
        gLaborIn(1).NumbMachs = 15
    ElseIf gbNumbMach Then
        If Not inpNumbMachs() Then
            gDone = True
            Unload Me
            Exit Sub
        End If
    Else
        gLaborIn(1).NumbMachs = 1
    End If
    
    If gbMachRecY Then
        gLaborIn(1).MachRec = "Y"
    Else
        gLaborIn(1).MachRec = "N"
    End If
    
    'local variable may need to be reset...adjusting global would mess up future trans.
     bScrapColl = gScrapColl
    
    bEnter = True
    I = 1
    LoopCnt = I     'Keep global variable in sync.
    While (I <= gLaborIn(1).NumbMachs) And bEnter
        
        If gbNumbMach Then
            lblInfo.Caption = LoadResString(gLanguageOffset + 237) & CStr(I) _
                & " of " & CStr(gLaborIn(1).NumbMachs)
        End If
        
        lblEmplNo.Caption = gEmployee.EmplNo
        lblEmplName.Caption = gEmployee.EmplName
        
        If Not inpLabJobSeq(I) Then Exit Sub
        lblPartNo.Caption = gJobMaster.PartNo
        lblCustName.Caption = gJobMaster.CustName
        lblJobNo.Caption = gJobMaster.OrderNo
        lblOrdQty.Caption = gJobMaster.OrderQty
        lblPartRev.Caption = gJobMaster.PartRev
        lblOrdType.Caption = gJobMaster.OrderType
        
        lblSeqNo.Caption = gLaborIn(I).SeqNo
        lblWCID.Caption = gLaborIn(I).WCID
        lblCCID.Caption = gLaborIn(I).CCID
        lblOperCd.Caption = gLaborIn(I).Oper
        lblUnplnReas.Caption = gLaborIn(I).UnplnReas
        
        'reset boolean value gbscrapounds
        'only divs allowing entry by weight are saline and milwaukee
        If div = saline And gJobMaster.MatlWgt <> 0 Then
            gbScrapPounds = True
        ElseIf div = milw Then
            gbScrapPounds = True
        Else
             gbScrapPounds = False
        End If
        
        If div = bristol And gLaborIn(I).SeqNo = 0 Then
             bScrapColl = False
        End If
        
        If div <> burl Then
         'and  div <> bristol Then
            lblEorS.Caption = gLaborIn(I).EOrS
            If gLaborIn(I).SetStd > 0 Then
                lblFixStd.Caption = Format(gLaborIn(I).SetStd, gOneDecPos)
           Else
               lblFixStd.Caption = ""
           End If
           If gLaborIn(I).MchStd > 0 Then
              If gJobOpers.TimeBasisCd = "1" Then
                lblMachStd.Caption = Format(gLaborIn(I).MchStd, gOneDecPos)
                lblMchStdType.Caption = "hrs/pc"
              ElseIf gJobOpers.TimeBasisCd = "4" Then
                lblMachStd.Caption = Format(gLaborIn(I).MchStd, "#######")
                lblMchStdType.Caption = "pcs/hr"
              ElseIf gJobOpers.TimeBasisCd = "5" Then
                lblMachStd.Caption = Format(gLaborIn(I).MchStd, gOneDecPos)
                lblMchStdType.Caption = "hrs"
             Else
                lblMachStd.Caption = Format(gLaborIn(I).MchStd, "#######")
                lblMchStdType.Caption = ""
             End If
           Else
              lblMachStd.Caption = ""
              lblMchStdType.Caption = ""
           End If
       End If
       
       bNoQty = False
        
      If gLaborIn(I).SeqNo <> 0 And _
          gLaborIn(I).SeqNo <> 990 Then
                     If Not inpWcAns(I) Then Exit Sub
      End If
     lblCCID.Caption = gLaborIn(I).CCID
        
        If Left(gLaborIn(I).Oper, 3) = "645" Then
            gLaborIn(I).RecType = "F"       'Setup hours rec type
            If gbInpSetOrMach Then
                If Not inpSetHours(I) Then Exit Sub
                lblSetupHrs.Caption = gLaborIn(I).setupHours
                gLaborIn(I).machHours = 0
            End If
        Else
            gLaborIn(I).RecType = "D"       'Mach hrs rec type?
            If gbInpSetOrMach Then
                If Not inpMchHours(I) Then Exit Sub
                lblMachHrs.Caption = gLaborIn(I).machHours
                gLaborIn(I).setupHours = 0
            End If
        End If
        
        If gbInpSetAndMach Then
          If Not gbRecTypeR Then
            If Not inpSetHours(I) Then Exit Sub
            lblSetupHrs.Caption = gLaborIn(I).setupHours
          Else:
            gLaborIn(I).setupHours = 0
          End If
            If Not inpMchHours(I) Then Exit Sub
            lblMachHrs.Caption = gLaborIn(I).machHours
            If gLaborIn(I).setupHours = 0 And gLaborIn(I).machHours = 0 Then
                ErrBox LoadResString(gLanguageOffset + 102), 2
            End If
        End If
       
        If gbAddonHrs Then
            If Not inpAddonHours(I) Then Exit Sub
            lblAddonHrs = gLaborIn(I).addonHours
        End If
         
        bZeroIt = False
       
        If gbZeroQty Then
           If div = milw Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670, 770, 900, 920, 960
                       bZeroIt = True
               End Select
           End If
           If div = bristol Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670
                       bZeroIt = True
               End Select
           End If
           If div = syracuse Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670
                       bZeroIt = True
               End Select
           End If
           If div = saline Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670, 900, 910, 920, 930, 960
                       bZeroIt = True
               End Select
           End If
           If div = monterrey Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670, 900, 910, 920, 930, 960
                       bZeroIt = True
               End Select
           End If
           If div = singapore Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670, 900, 910, 930, 960
                       bZeroIt = True
               End Select
           End If
           If div = pv_singapore Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670, 900, 910, 930, 960
                       bZeroIt = True
               End Select
           End If
        
        End If
        
        If bZeroIt Then
            gLaborIn(I).ActQty = 0
            lblProdQty = gLaborIn(I).ActQty
        Else
            If div = saline Then
                If Not inpProdQtyReqd(I) Then Exit Sub
                lblProdQty = gLaborIn(I).ActQty
            ElseIf div = milw Then
                If Not inpProdQtyReqd(I) Then Exit Sub
                lblProdQty = gLaborIn(I).ActQty
            ElseIf div = monterrey Then
                If Not inpProdQtyReqd(I) Then Exit Sub
                lblProdQty = gLaborIn(I).ActQty
            Else
                If Not inpProdQty(I) Then Exit Sub
                lblProdQty = gLaborIn(I).ActQty
            End If
        End If
        
        lScrapTot = 0
        If bScrapColl Then
             If gJobMaster.OrderType <> "RS" Then
                 If Not getScrapCounts(I) Then Exit Sub
                 For j = 1 To 8
                     lScrapTot = lScrapTot _
                         + gLaborIn(I).Scrap(j).Qty
                 Next
             End If
        End If
        lblScrapQty.Caption = CStr(lScrapTot)
        
        If div = saline Then
            If Not getBoxNumbers(I) Then Exit Sub
        End If
        
        
        If gbCallCompl Then
            If Not CallCmpl(I) Then Exit Sub
            lblOperStatus = gLaborIn(I).OperStatus
        Else
            Rem Don't ask Operation Complete Question
            gLaborIn(I).OperStatus = "P"
        End If
        
        If gbNoLoads And LdBatchIncn(gLaborIn(LoopCnt).WCID) <> 0 Then
            If Not inpNoOfLoads(LoopCnt) Then Exit Sub
            lblNumbOfLoads.Caption = gLaborIn(LoopCnt).NoOfLoads
            If gLaborIn(LoopCnt).NoOfLoads > 0 Then
                If Not inpCycleTime(LoopCnt) Then Exit Sub
                lblCycleTime.Caption = gLaborIn(LoopCnt).CycleTime
                Call CalcBatchHrs(LoopCnt)
                'MsgBox "Is Ld_batch_incn the same as LoadBatIncnAry ?"
                'Call LdBatchIncn(gLaborIn(LoopCnt).WCID)
                If Not getBatchIncn(LoopCnt) Then Exit Sub
                lblBatchMethod.Caption = gLaborIn(LoopCnt).biMethod
                Call CalcBtErnHrs(LoopCnt)
            End If
        Else
Rem Coded to not ask info if no Batch Incentive Codes for WCID
            gLaborIn(LoopCnt).NoOfLoads = 0
            lblNumbOfLoads.Caption = ""
            gLaborIn(LoopCnt).CycleTime = 0
            lblCycleTime.Caption = ""
            gLaborIn(LoopCnt).biMethod = 0
            lblBatchMethod.Caption = ""
            Call CalcBtErnHrs(LoopCnt)
        End If
        
        If (Not gbInpSetOrMach) And (Not gbInpSetAndMach) Then
            Rem Calculate Hours
            If I = 1 Then
                CalcHours = CalcActHrs()
            End If
            If gLaborIn(I).RecType = "F" Then
                gLaborIn(I).setupHours = CalcHours
                gLaborIn(I).machHours = 0
            Else
                gLaborIn(I).machHours = CalcHours
                gLaborIn(I).setupHours = 0
            End If
            gLaborIn(I).inHours = 0
            lblSetupHrs.Caption = gLaborIn(I).setupHours
            lblMachHrs.Caption = gLaborIn(I).machHours
        End If
        
        If Not DirLaborReEnter(I) Then Exit Sub
        
        If I <> 1 Then
            If gbNumbMach Then
                gLaborIn(I).MachRec = "Y"
                gLaborIn(I).NumbMachs = gLaborIn(1).NumbMachs
                gLaborIn(I).incenRec = gLaborIn(1).incenRec
            ElseIf gbAnotherRec Then
                gLaborIn(I).MachRec = "Y"
                gLaborIn(I).NumbMachs = 1
                gLaborIn(I).incenRec = 0
            End If
        End If
        
        If (Left(gLaborIn(I).Oper, 3) = "640" Or Left(gLaborIn(I).Oper, 3) = "635") And _
           div = corry And _
           gLaborIn(I).incenRec = 1 Then
              gLaborIn(I).incenRec = 0
              'Bob K. said that Reset or sample time is NEVER paid incentive
              'therefore incentive code is set to 0.
              'IF this is among multimachine ringings, a problem may occur
              
              'After a phone conversation on Sept 9, 03, the Corry div. (Kensinger, Bailey)
            'requested that this code be removed. The program will no longer substitute
            'an incen code 0 for a 2 when the operation is Sample or Reset.  -MWP
            
            'Condition has been reactivated per Ron Bailey. Condition expanded to include
            'incenRec = 1. Incentive codes greater than 1 will not be touched, as it screws up
            'the incentive pay calculation. - MWP 09/25/2009
            
        End If
      
        
        'If rework, then clobber the RecType with R before update
        If gbRecTypeR Then
            gLaborIn(I).RecType = "R"
            Call RecordError("R")
        End If
        
        I = I + 1
        LoopCnt = I     'keep global variable in sync
        
        Rem if gbInpSetAndMach - make sure only one entry
        Rem (Just in Case)
        If gbInpSetAndMach Then
            gbAnotherRec = False
            gbNumbMach = False
        End If
        
        If gbAnotherRec Then
           Select Case AnotherRec()
              Case 0
                      'selected No when asked 'enter another record?'
                       bEnter = False
             Case 1
                      'selected Yes when asked 'enter another record?'
                      bEnter = True
             Case Else
                     ' escaped from anotherRec
                     Exit Sub
          End Select
       ElseIf Not gbNumbMach Then
            Rem Only enter one machine at a time
            bEnter = False
        End If
    
    ClearForm
    Wend
    
    'LoopCnt = gLaborIn(1).NumbMachs
    LoopCnt = LoopCnt - 1
    
    Rem Reset NumbMachs to 1 once looping is done
    If gbAnotherRec Then
        gLaborIn(1).NumbMachs = 1
    End If
    
'    If gbInpSetAndMach Then
'        Rem First write F rec with only setup hours
'        gLaborIn(i).RecType = "F"
'        CalcHours = gLaborIn(i).machHours   'save it
'        gLaborIn(i).machHours = 0
'        Call Build_Labor_Recs
'        Rem Next - Prepare to write D rec with mach hours
'        gLaborIn(i).RecType = "D"
'        gLaborIn(i).machHours = CalcHours   'restore it
'        gLaborIn(i).setupHours = 0
'    End If
    Call Build_Labor_Recs
    
    If gbEOShift Then
        If empClkOutOK Then
            gLaborIn(1).ErrorLst = ""
            Call ClockOut
            Call Upd_Empl_Attd
            gDone = True
        End If
    End If
   
    Unload Me
    
End Sub

Public Function DirLaborReEnter(I As Integer) As Boolean
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDirLabor.DirLaborReEnter(i)
'Description      : Routine that handles the re-enter an item menu for a single
'                   machine entry for direct labor for all divisions. Boolean
'                   values are globally set by division and type of
'                   Labor entry, to determine which options are used for the
'                   labor process.
'Parameters       : None
'Called From      : frmDirLabor.DirLaborProcess
'Procedures Used  : ErrorChk, inpJobAndSeq, DispValues, inpWCID, GetCostCenter,
'                   inpProdQty, getScrapCounts, inpSetHrs, inpMchHours, inpAddonHours,
'                   inpEndOfShift, CallCmpl, inpNoOfLoads, CalcBtErnHrs, CalcBatchHrs,
'                   inpCycleTime, getBatchIncn, ErrorBeep, ErrBox
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
    Dim strEntrySub As String
    Dim strEntryEmpNo As String
    Dim bNoQtyWC As Boolean
    
    
    DirLaborReEnter = False
    gDone = False
    
    Do Until gDone Or DirLaborReEnter
        Set gEntry = New clsEntry
        gEntry.Title = LoadResString(gLanguageOffset + 192)
        gEntry.Add "A", LoadResString(gLanguageOffset + 193)
        gEntry.Add "E", LoadResString(gLanguageOffset + 194)
        gEntry.Add "J", LoadResString(gLanguageOffset + 195)
       
        bNoQtyWC = False

        If bNoQtyWC = False Then
        gEntry.Add "W", LoadResString(gLanguageOffset + 196)
        End If
        If gbZeroQty = False And bNoQtyWC = False Then
            gEntry.Add "Q", LoadResString(gLanguageOffset + 197)
        Else
           If div = milw Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670, 770, 900, 920, 960
                  Case Else
                       gEntry.Add "Q", LoadResString(gLanguageOffset + 197)
               End Select
           End If
           If div = bristol Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670
                  Case Else
                       gEntry.Add "Q", LoadResString(gLanguageOffset + 197)
               End Select
           End If
           If div = syracuse Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670
                  Case Else
                       gEntry.Add "Q", LoadResString(gLanguageOffset + 197)
               End Select
           End If
           If div = saline Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670, 900, 910, 920, 930, 960
                  Case Else
                       gEntry.Add "Q", LoadResString(gLanguageOffset + 197)
               End Select
           End If
           If div = monterrey Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670, 900, 910, 920, 930, 960
                  Case Else
                       gEntry.Add "Q", LoadResString(gLanguageOffset + 197)
               End Select
           End If
           If div = singapore Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670, 900, 910, 930, 960
                  Case Else
                       gEntry.Add "Q", LoadResString(gLanguageOffset + 197)
               End Select
           End If
           If div = pv_singapore Then
               Select Case Left(gLaborIn(I).Oper, 3)
                  Case 635, 645, 660, 665, 670, 900, 910, 930, 960
                  Case Else
                       gEntry.Add "Q", LoadResString(gLanguageOffset + 197)
               End Select
           End If
        End If
        If bNoQtyWC = False Then
              If bScrapColl And gJobMaster.OrderType <> "RS" Then
                       gEntry.Add "S", LoadResString(gLanguageOffset + 198)
              End If
        End If
        
        If gbInpSetAndMach Or (gbInpSetOrMach _
                And gLaborIn(I).RecType = "F") _
                Or (gLaborIn(I).RecType = "F" And div = bristol) Then
            gEntry.Add "F", LoadResString(gLanguageOffset + 299)
        End If
        If gbInpSetAndMach Or (gbInpSetOrMach _
                And gLaborIn(I).RecType = "D") _
                Or (gLaborIn(I).RecType = "D" And div = bristol) Then
            gEntry.Add "M", LoadResString(gLanguageOffset + 298)
        End If
        If gbAddonHrs Then
            gEntry.Add "B", "change addon hours"
        End If
        If div = saline Then
           gEntry.Add "B", "change box numbers"
        End If
        If gbEOShift Then
            gEntry.Add "C", "change Clock-out status"
        End If
        If gbCallCompl Then
            gEntry.Add "O", LoadResString(gLanguageOffset + 199)
        End If
        If gbNoLoads Then
            gEntry.Add "L", "change number of Loads"
        End If
        If gbNoLoads And gLaborIn(I).NoOfLoads > 0 Then
            gEntry.Add "T", "change cycle Time"
            gEntry.Add "H", "change batcH method"
        End If
        strEntrySub = basGetEntry()
        If gEntry.Escaped Then
            Set gEntry = Nothing
            gDone = True
            Exit Function
        End If
        Set gEntry = Nothing
        Select Case strEntrySub
        Case "A"
            Call ErrorChk(I)
            DirLaborReEnter = True
        Case "E"
            gDone = True
            Exit Function
        Case "J"
            If Not inpLabJobSeq(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Call DispValues(I)
                        End If
        Case "W"
            If Not inpWCID(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Screen.ActiveForm.lblWCID.Caption = gLaborIn(I).WCID
                gLaborIn(I).CCID = gWorkCenter.CCID
                If Not GetCostCenter(gWorkCenter.CCID) Then
                    gCostCenter.PaidLunch = "N"
                End If
                Screen.ActiveForm.lblCCID.Caption = gLaborIn(I).CCID
            End If
        Case "Q"
            If div = saline Then 'quantity must be entered
                If Not inpProdQtyReqd(I) Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                Else
                    Screen.ActiveForm.lblProdQty.Caption = gLaborIn(I).ActQty
                End If
            ElseIf div = monterrey Then 'quantity must be entered
                If Not inpProdQtyReqd(I) Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                Else
                    Screen.ActiveForm.lblProdQty.Caption = gLaborIn(I).ActQty
                End If
            Else
                If Not inpProdQty(I) Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                Else
                    Screen.ActiveForm.lblProdQty.Caption = gLaborIn(I).ActQty
                End If
            End If
            
        Case "S"
            If Not bScrapColl Then
                ErrorBeep
            Else
                If Not getScrapCounts(I) Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                Else
                    Dim iScrap As Long
                    Dim j As Integer
                    iScrap = 0
                    For j = 1 To 8
                        iScrap = iScrap + gLaborIn(I).Scrap(j).Qty
                    Next j
                    Screen.ActiveForm.lblScrapQty.Caption = CStr(iScrap)
                End If
            End If
        Case "F"
            If gLaborIn(I).NoOfLoads > 0 _
                    Or gLaborIn(I).CycleTime > 0 Then
                ErrBox LoadResString(gLanguageOffset + 8), 2
                ErrorBeep
            Else
                If gLaborIn(I).RecType <> "F" And _
                   Not gbInpSetAndMach Then
                    ErrorBeep
                Else
                    If Not inpSetHours(I) Then
                        gDone = False       'be sure to loop back
                        ErrorBeep
                    Else
                        Screen.ActiveForm.lblSetupHrs.Caption = gLaborIn(I).setupHours
                    End If
                End If
            End If
        Case "M"
            If gLaborIn(I).NoOfLoads > 0 _
                    Or gLaborIn(I).CycleTime > 0 Then
                ErrBox LoadResString(gLanguageOffset + 8), 2
                ErrorBeep
            Else
                If gLaborIn(I).RecType <> "D" And _
                   Not gbInpSetAndMach Then
                    ErrorBeep
                Else
                    If Not inpMchHours(I) Then
                        gDone = False       'be sure to loop back
                        ErrorBeep
                    Else
                        Screen.ActiveForm.lblMachHrs.Caption = gLaborIn(I).machHours
                    End If
                End If
            End If
        Case "B"
            If gbAddonHrs Then
                If Not inpAddonHours(I) Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                Else
                    Screen.ActiveForm.lblAddonHrs.Caption = gLaborIn(I).addonHours
                End If
            End If
            If div = saline Then
                If Not getBoxNumbers(I) Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                End If
            End If
                
        Case "C"
            If gEosFlag = 0 Then
                ErrBox LoadResString(gLanguageOffset + 86), 2
                ErrorBeep
            Else
                If Not inpEndOfShift() Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                End If
            End If
        Case "O"
            If Not CallCmpl(I) Then
                gDone = False       'be sure to loop back
                ErrorBeep
            Else
                Screen.ActiveForm.lblOperStatus.Caption = gLaborIn(I).OperStatus
            End If
        Case "L"
            If Not okToPrompt Then
                ErrorBeep
            Else
                If Not inpNoOfLoads(I) Then
                    gDone = False       'be sure to loop back
                    ErrorBeep
                Else
                    Screen.ActiveForm.lblNumbOfLoads.Caption = gLaborIn(I).NoOfLoads
                    Call CalcBtErnHrs(I)
                    Call CalcBatchHrs(I)
                End If
            End If
        Case "T"
            If Not okToPrompt Then
                ErrorBeep
            Else
                If gLaborIn(I).NoOfLoads = 0 Then
                    ErrorBeep
                    ErrBox LoadResString(gLanguageOffset + 79), 2
                Else
                    If Not inpCycleTime(I) Then
                        gDone = False       'be sure to loop back
                        ErrorBeep
                    Else
                        Screen.ActiveForm.lblCycleTime.Caption = gLaborIn(I).CycleTime
                        Call CalcBatchHrs(I)
                    End If
                End If
            End If
        Case "H"
            If Not okToPrompt Then
                ErrorBeep
            Else
                If gLaborIn(I).NoOfLoads = 0 Then
                    ErrorBeep
                    ErrBox LoadResString(gLanguageOffset + 79), 2
                Else
                    LdBatchIncn (gLaborIn(I).WCID)
                    If Not getBatchIncn(I) Then
                        gDone = False       'be sure to loop back
                        ErrorBeep
                    Else
                        Screen.ActiveForm.lblBatchMethod.Caption = gLaborIn(I).biMethod
                        Call CalcBtErnHrs(I)
                    End If
                End If
            End If
        End Select
    Loop

End Function



Public Sub ClearForm()
'==============================================================================
'Author           : David A. Taylor, PRT Group Inc.
'Date             : June, 1998
'Procedure Name   : frmDirLabor.ClearForm
'Description      : Clears the label fields that display data
'Parameters       : None
'Called From      : DirLaborProcess (after an entry is done)
'Procedures Used  : None
'Return Parameters: None
'Tables Updates   : None
'==============================================================================
        lblCCID.Caption = ""
        lblPartNo.Caption = ""
        lblCustName.Caption = ""
        lblJobNo.Caption = ""
        lblOrdQty.Caption = ""
        lblPartRev.Caption = ""
        lblOrdType.Caption = ""
        
        lblSeqNo.Caption = ""
        lblWCID.Caption = ""
        lblOperCd.Caption = ""
        lblUnplnReas.Caption = ""
        lblFixStd.Caption = ""
        lblMachStd.Caption = ""
        lblMchStdType.Caption = ""
        lblEorS.Caption = ""
        lblSetupHrs.Caption = ""
        lblMachHrs.Caption = ""
        lblAddonHrs = ""
        lblProdQty = ""
        lblScrapQty.Caption = ""
        lblOperStatus = ""
        lblNumbOfLoads.Caption = ""
        lblCycleTime.Caption = ""
        lblBatchMethod.Caption = ""

End Sub

Private Sub Form_Load()
lblFormTitle.Caption = LoadResString(gLanguageOffset + 297)
lblEmplNoCaption.Caption = LoadResString(gLanguageOffset + 136)
lblNameCaption.Caption = LoadResString(gLanguageOffset + 137)
lblJobNoCaption.Caption = LoadResString(gLanguageOffset + 150)
lblPartCaption.Caption = LoadResString(gLanguageOffset + 151)
lblOrderQtyCaption.Caption = LoadResString(gLanguageOffset + 152)
lblTypeCaption.Caption = LoadResString(gLanguageOffset + 153)
lblCustomerCaption.Caption = LoadResString(gLanguageOffset + 154)
lblRevCaption.Caption = LoadResString(gLanguageOffset + 155)
lblSeqNoCaption.Caption = LoadResString(gLanguageOffset + 156)
lblOperCaption.Caption = LoadResString(gLanguageOffset + 157)
lblWCIDCaption.Caption = LoadResString(gLanguageOffset + 158)
lblCCIDCaption.Caption = LoadResString(gLanguageOffset + 159)
lblStatusCaption.Caption = LoadResString(gLanguageOffset + 160)
lblUnplanReasCaption.Caption = LoadResString(gLanguageOffset + 161)
lblFixedStdCaption.Caption = LoadResString(gLanguageOffset + 162)
lblESCaption.Caption = LoadResString(gLanguageOffset + 163)
lblMachineStdCaption.Caption = LoadResString(gLanguageOffset + 164)
lblSetupHrsCaption.Caption = LoadResString(gLanguageOffset + 167)
lblAddonHrsCaption.Caption = LoadResString(gLanguageOffset + 169)
lblMachineHrsCaption.Caption = LoadResString(gLanguageOffset + 168)
lblQuantityCaption.Caption = LoadResString(gLanguageOffset + 170)
lblScrapQtyCaption.Caption = LoadResString(gLanguageOffset + 171)
If div = corry Then
  lblNoLoadsCaption.Visible = True
  lblCycleTimeCaption.Visible = True
  lblMethodCaption.Visible = True
  lblNoLoadsCaption.Caption = LoadResString(gLanguageOffset + 172)
  lblCycleTimeCaption.Caption = LoadResString(gLanguageOffset + 173)
  lblMethodCaption.Caption = LoadResString(gLanguageOffset + 174)
Else
  lblNoLoadsCaption.Visible = False
  lblCycleTimeCaption.Visible = False
  lblMethodCaption.Visible = False
End If


End Sub


