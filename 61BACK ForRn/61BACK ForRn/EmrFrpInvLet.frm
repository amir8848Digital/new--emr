VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpInvLet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11760
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9570
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5925
      TabIndex        =   17
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   18
         Top             =   135
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         MaxLength       =   0
         DataField       =   "UOUSRCD"
         IdName          =   "UOUSRCD"
      End
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   1
         Left            =   1860
         TabIndex        =   19
         Top             =   150
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         MaxLength       =   0
         DataField       =   "UOMNUCD"
         IdName          =   "UOMNUCD"
      End
   End
   Begin VB.Frame FRA_REPOBJ 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7140
      TabIndex        =   16
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   -165
         Visible         =   0   'False
         Width           =   495
         lastProp        =   500
         _cx             =   5080
         _cy             =   5080
         DisplayGroupTree=   -1  'True
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   -1  'True
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   0   'False
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
         LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      End
   End
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   60
      TabIndex        =   9
      Top             =   9660
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   979
      AllowDelete     =   0   'False
      Automate        =   -1  'True
      EntType         =   1
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame fra_nkeys 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9735
      Left            =   0
      TabIndex        =   12
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   11
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpInvLet.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpInvLet.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   15
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   10
               Top             =   360
               Width           =   8145
               _ExtentX        =   14367
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   14
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1965
               TabIndex        =   5
               ToolTipText     =   "Enter From Export Number"
               Top             =   1290
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOFR"
               IdName          =   "UOINEXPNOFR"
               CmpStr          =   "Ih.InExpNo ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2745
               TabIndex        =   2
               ToolTipText     =   "Enter Invoice Year"
               Top             =   1005
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "Ih.InYy ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1965
               TabIndex        =   1
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   1005
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "Ih.InTc ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3165
               TabIndex        =   3
               ToolTipText     =   "Enter Invoice Character"
               Top             =   1005
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "Ih.InChr ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   3945
               TabIndex        =   4
               ToolTipText     =   "Enter Invoice Number"
               Top             =   1005
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "Ih.InNo ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1965
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   715
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "Ih.InCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   1965
               TabIndex        =   6
               ToolTipText     =   "Enter Export Numbers"
               Top             =   1290
               Width           =   6825
               _ExtentX        =   12039
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOINEXPNOSEL"
               IdName          =   "UOINEXPNOSEL"
            End
            Begin VB.Label LblExpNo 
               BackStyle       =   0  'Transparent
               Caption         =   "Exp No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   165
               TabIndex        =   25
               Top             =   1290
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Company Cd"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   15
               Left            =   165
               TabIndex        =   24
               Top             =   715
               Width           =   1335
            End
            Begin VB.Label LblChrSlash 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3870
               TabIndex        =   23
               Top             =   1005
               Width           =   105
            End
            Begin VB.Label LblInvNo 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Tc/Yy/Chr/No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   165
               TabIndex        =   22
               Top             =   1005
               Width           =   1815
            End
            Begin VB.Label LblYySlash 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3105
               TabIndex        =   21
               Top             =   1005
               Width           =   105
            End
            Begin VB.Label LblTcSlash 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2685
               TabIndex        =   20
               Top             =   1005
               Width           =   105
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpInvLet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepDeclnLet
Dim Rep1 As New EmrRepGSPCert
Dim Rep2 As New EmrRepGSPDetLet
Dim Rep3 As New EmrRepGRFORMBack
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Dim mb_MultiInv As Boolean    '*** (Jen 2.14)
Private Sub SetGroupSort()
  ''GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub
Private Sub adc_setreprecsource()
'*** (Jen 2.12)
On Error GoTo RepErr
'*** (Jen 2.12)

  '*** Report Sql for GR Form Back (Jen 2.14) ***
'      Select InTc, InYy, InChr, InNo, InExpNo, InCmCd, InDt, InLmgCst, InLmgDt, InLmgSal,
'      InVaDC, InVaX, InCnvRt, InPrdDesc, InPkgWt, InGrNo, InGrDt, InSbNo, InSbDt, InAwbNo,
'      InAwbDt, InHAwbNo, InHAwbDt, InHBinNo, InGspNo, InGspDt, InChqNo, InChqAmt, InPoNo,
'      InPoDt, InCmPayTerms, InFOB, InFOBRs, InComm, InCommRs, InFrt, InFrtRs, InIns,
'      InInsRs, InDsc, InDscRs, InDeduct, InDeductRs, InTot, InTotRs, InTotAs,
'      InDiaCost, InCSCost, InAdv, InDiaCostDesc, InCSCostDesc, InAdvDesc, InCls, InHBkName,
'      InHBkAdd1, InHBkAdd2, InHBkAdd3, InHBkAc, InHFrgnBkName, InHFrgnBkAc, InHExpOff,
'      InHRbiNo, InHIecNo, InHPreCarriage, InHPreCarrier, InHVessel, InHPortOfLd, InHConOfOrg,
'      InHMarks, InHPkgDesc, InHItcNo, InHEtcZone, InHEtcSection, InHConsignee, InCmInsBy as InHInsBy,
'      InHInsNo, InHAgName, InHAgAdd1, InHAgAdd2, InHAgAdd3, InCmName, InCmAdd1, InCmAdd2,
'      InCmAdd3, InCmTel, InCmFax, InCmEMail, InCmCpName, InCmPortOfDisch, InCmFinDest,
'      InCmBuyer, InSelfName, InSelfAdd1, InSelfAdd2, InSelfAdd3, InSelfTel, InSelfFax,
'      InSelfEMail, InhAuthSgn
'      From InvHd, TxndZ Where 1 = 2
  '*** Report Sql for GR Form Back (Jen 2.14) ***

  '*** Report Sql ***
'      Select InTc, InYy, InChr, InNo, InExpNo, InCmCd, InDt, InLmgCst, InLmgDt, InLmgSal,
'      InVaDC, InVaX, InCnvRt, InPrdDesc, InPkgWt, InGrNo, InGrDt, InSbNo, InSbDt, InAwbNo,
'      InAwbDt, InHAwbNo, InHAwbDt, InHBinNo, InGspNo, InGspDt, InChqNo, InChqAmt, InPoNo,
'      InPoDt, InCmPayTerms, InFOB, InFOBRs, InComm, InCommRs, InFrt, InFrtRs, InIns,
'      InInsRs, InDsc, InDscRs, InDeduct, InDeductRs, InTot, InTotRs, InTotAs,
'      InDiaCost, InCSCost, InAdv, InDiaCostDesc, InCSCostDesc, InAdvDesc, InCls, InHBkName,
'      InHBkAdd1, InHBkAdd2, InHBkAdd3, InHBkAc, InHFrgnBkName, InHFrgnBkAc, InHExpOff,
'      InHRbiNo, InHIecNo, InHPreCarriage, InHPreCarrier, InHVessel, InHPortOfLd, InHConOfOrg,
'      InHMarks, InHPkgDesc, InHItcNo, InHEtcZone, InHEtcSection, InHConsignee, InCmInsBy as InHInsBy,
'      InHInsNo, InHAgName, InHAgAdd1, InHAgAdd2, InHAgAdd3, InCmName, InCmAdd1, InCmAdd2,
'      InCmAdd3, InCmTel, InCmFax, InCmEMail, InCmCpName, InCmPortOfDisch, InCmFinDest,
'      InCmBuyer, InSelfName, InSelfAdd1, InSelfAdd2, InSelfAdd3, InSelfTel, InSelfFax,
'      InSelfEMail, InhAuthSgn, TzPurAmt as qDiaVal, TzPurAmt as qCSVal, TzPurAmt as qGldVal,
'      TzPurAmt As qPlVal, TzPurAmt As qSlVal, TzPurAmt As qPdVal, TzPurAmt As qAccsVal
'      From InvHd, TxndZ Where 1 = 2
  '*** Report Sql ***
  
  
  Dim wrepcnd As String, wCnd As String, wStr1 As String, wStr2 As String, wSqlStrg As String
  Dim ws_CurCd As String, ws_CurDesc As String  '*** (Jen 2.13)
  
  '*** (Jen 3.01)
  Dim wHdStr1 As String, wHdStr2 As String, wHdStr3 As String, wHdStr4 As String, wStr3 As String, wStr4 As String
  Dim wFinalStr1 As String, wFinalStr2 As String, wFinalStr3 As String
  '*** (Jen 3.01)
  
  '*** (Jen 2.14)
  Dim wAndCnd As String
  Dim ws_MinCnd As String, ws_MinWhereCnd As String, ws_MinAndCnd As String
  Dim ws_InExpNo As String, ws_InTotAs As String, ws_PrdDesc As String, ws_MinExpNo As String
    
  Dim wExpLocal As String   ' **** Manali 3.6.0 - InterCarat Change
    
  wExpLocal = moCn.GetFldVal("Select HExpLocal from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
    
  If mb_MultiInv = False Then
    wrepcnd = adc.RepCond
    wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, " ")
    wAndCnd = IIF(wrepcnd <> "", " and " + wrepcnd, " ")
    If adc("UoInNoFr") <> 0 Then
      ws_MinCnd = " Ih1.InCoCd= '" + adc("UoCoCdFr") + "' and Ih1.InTc= '" + adc("UoInTcFr") + "' and " + _
                  "Ih1.InYy= '" + adc("UoInYyFr") + "' and Ih1.InChr= '" + adc("UoInChrFr") + "' and " + _
                  "Ih1.InNo= " + CStr(adc("UoInNoFr")) + " "
    Else
      ws_MinCnd = " Ih1.InCoCd= '" + adc("UoCoCdFr") + "' and Ih1.InExpNo= '" + adc("UoInExpNoFr") + "' "
    End If
    
    ws_MinWhereCnd = " where " + ws_MinCnd
    ws_MinAndCnd = " and " + ws_MinCnd
    ws_InExpNo = " max(Ih1.InExpNo) as InExpNo "
    ws_InTotAs = " max(Ih1.InTotAsTxt) as InTotAs "
    ws_PrdDesc = " max(Ih1.InPrdDesc) as InPrdDesc "
  Else
    mb_MultiInv = True
    
    '*** (Jen 3.1.3)
    wCnd = " where " + IIF(gs_Partition = ctCurrPrtn, " Ih.InPrtKey='" + ctCurrPrtn + "' and ", "") + " Ih.InCoCd= '" + adc("UoCoCdFr") + "' and Ih.InExpNo in (" + adc("UoInExpNoSel") + ") "
    wAndCnd = " and Ih.InCoCd= '" + adc("UoCoCdFr") + "' and Ih.InExpNo in (" + adc("UoInExpNoSel") + ") "
    '*** (Jen 3.1.3)
    
    '*** (Bef 3.1.3)
    'wCnd = " where Ih.InExpNo in (" + adc("UoInExpNoSel") + ") "
    'wAndCnd = " and Ih.InExpNo in (" + adc("UoInExpNoSel") + ") "
    '*** (Bef 3.1.3)
    
    ws_MinExpNo = moCn.GetFldVal("Select min(InExpNo) from InvHd Ih " + wCnd)
    ws_MinCnd = " Ih1.InCoCd= Ih.InCoCd and Ih1.InExpNo = '" + ws_MinExpNo + "' "     '*** (Jen 3.1.3)
    'ws_MinCnd = " Ih1.InExpNo = '" + ws_MinExpNo + "' "    '*** (Bef 3.1.3)
    ws_MinWhereCnd = " where " + ws_MinCnd
    ws_MinAndCnd = " and " + ws_MinCnd
    ws_InExpNo = " '" + moCn.GetStrConcatVal("InExpNo", "InvHd Ih ", wCnd, ", ") + "' as InExpNo "
    
    Dim ws_MinCurCd As String, ws_MaxCurCd As String
    Dim ws_MinCnvFct As String, ws_MaxCnvFct As String
    Dim ws_MinCnvRt As String, ws_MaxCnvRt As String
    
    '*** (Jen 3.1.3) added wCnd
    ws_MinCurCd = moCn.GetFldVal("Select min(CmCurCd) from InvHd Ih join CustMst on CmCtg= 'C' and CmCd= InCmCd /* where  InExpNo in (" + adc("UoInExpNoSel") + ") */ " + wCnd)
    ws_MaxCurCd = moCn.GetFldVal("Select max(CmCurCd) from InvHd Ih join CustMst on CmCtg= 'C' and CmCd= InCmCd /* where  InExpNo in (" + adc("UoInExpNoSel") + ") */ " + wCnd)
    If ws_MinCurCd <> ws_MaxCurCd Then DispMsg "Invoice Export Number(s) Have Different Currency Codes", etError: wCnd = " and 1= 2": wAndCnd = " and 1= 2"
    
    ws_MinCnvFct = Format(moCn.GetFldVal("Select min(RTrim (LTrim (InMulDiv))+ Str(convert (decimal (16, 4), InCnvFct))) from InvHd Ih /* where  InExpNo in (" + adc("UoInExpNoSel") + ") */ " + wCnd), "####0.0000")
    ws_MaxCnvFct = Format(moCn.GetFldVal("Select max(RTrim (LTrim (InMulDiv))+ Str(convert (decimal (16, 4), InCnvFct))) from InvHd Ih /* where  InExpNo in (" + adc("UoInExpNoSel") + ") */ " + wCnd), "####0.0000")
    If ws_MinCnvFct <> ws_MaxCnvFct Then DispMsg "Invoice Export Number(s) Have Different Currency Conversion Factors", etError: wCnd = " and 1= 2": wAndCnd = " and 1= 2"
    
    ws_MinCnvRt = Format(moCn.GetFldVal("Select min(RTrim (LTrim (InRsMulDiv))+ Str(convert (decimal (16, 4), InCnvRt))) from InvHd Ih /* where  InExpNo in (" + adc("UoInExpNoSel") + ") */ " + wCnd), "####0.0000")
    ws_MaxCnvRt = Format(moCn.GetFldVal("Select max(RTrim (LTrim (InRsMulDiv))+ Str(convert (decimal (16, 4), InCnvRt))) from InvHd Ih /* where  InExpNo in (" + adc("UoInExpNoSel") + ") */ " + wCnd), "####0.0000")
    If ws_MinCnvRt <> ws_MaxCnvRt Then DispMsg "Invoice Export Number(s) Have Different Rupee Conversion Factors", etError: wCnd = " and 1= 2": wAndCnd = " and 1= 2"
    '*** (Jen 3.1.3) added wCnd

    ws_InTotAs = " '" + moCn.GetFldVal("Select (case when (sum(Ih.InFrt)<> 0 or sum(Ih.InFrtRs)<> 0) AND (sum(Ih.InIns)= 0 and sum(Ih.InInsRs)= 0) then 'C&F' " + _
                 " when (sum(Ih.InFrt)= 0 and sum(Ih.InFrtRs)= 0) AND (sum(Ih.InIns)<> 0 or sum(Ih.InInsRs)<> 0) then 'C&I' " + _
                 " when (sum(Ih.InFrt)<> 0 or sum(Ih.InFrtRs)<> 0) AND (sum(Ih.InIns)<> 0 or sum(Ih.InInsRs)<> 0) then 'CIF' " + _
                 " when sum(Ih.InFrt)= 0 and sum(Ih.InFrtRs)= 0 and sum(Ih.InIns)= 0 and sum(Ih.InInsRs)= 0 then 'FOB' else '' end) " + _
                 "from InvHd Ih " + wCnd) + "' as InTotAs "
    
    '###
    'If Trim(TxtInPrdDesc.text) <> "" Then
    '  ws_PrdDesc = " '" + Replace(Trim(TxtInPrdDesc.text), "'", "''") + "' as InPrdDesc "
    'Else
    '###
      ws_PrdDesc = " max(Ih1.InPrdDesc) as InPrdDesc "
    'End If
  End If
  '*** (Jen 2.14)
  
  Set Rep = Nothing
  Set Rep1 = Nothing
  Set Rep2 = Nothing
  Set Rep3 = Nothing

  Select Case UCase(adc.MenuCd)
  '****** (Jen 3.01) added RepDeclnLetMulti
  Case UCase("RepDeclnLet"), UCase("RepDeclnLetMulti")
    Set adc.RepSource = Rep
    Set mRep = Rep
  '****** (Jen 3.01) added RepGSPCertMulti
  Case UCase("RepGSPCert"), UCase("RepGSPCertMulti")
    Set adc.RepSource = Rep1
    Set mRep = Rep1
  '****** (Jen 3.01) added RepGSPDetLetMulti
  Case UCase("RepGSPDetLet"), UCase("RepGSPDetLetMulti")
    Set adc.RepSource = Rep2
    Set mRep = Rep2
  'Case UCase("RepGRFormBack")  '*** (Bef 2.14)
  '*** Jay 3.3(FG)
  '*** (Jen 2.14)
  ' ***** Manali Trading Module
  Case UCase("RepGRFormBack"), UCase("RepGRFormBackMulti"), UCase("RepFgGrBk"), UCase("RepJtGrBk")
  '*** (Jen 2.14)
  '*** Jay 3.3(FG)
    Set adc.RepSource = Rep3
    Set mRep = Rep3
  End Select
  'Set adc.RepSource = Rep
  
  Call DispCoNm
  ''Call GRP_REP.Gen3LRep
  ''grpflds = GRP_REP.GrpFldLst
  
  ''wCommStr = "Select " + grpflds + ", "
  
  
  '****** (Bef 2.13) max(InCnvRt) as InCnvRt,
  '****** (Bef 2.13) max(case when InRsMulDiv= 'D' and convert (decimal (16, 4), InCnvRt)> 0 then 1/ InCnvRt else InCnvRt end) as InCnvRt,
  
  '*** (Bef 3.01)
  ''*** (Jen 2.14)
  'If UCase(adc.MenuCd) = UCase("RepGRFormBack") Or UCase(adc.MenuCd) = UCase("RepGRFormBackMulti") Then
  '  wStr1 = "Select max(Ih1.InTc) as InTc, max(Ih1.InYy) as InYy, max(Ih1.InChr) as InChr, max(Ih1.InNo) as InNo, " + _
  '          " " + ws_InExpNo + ", max(Ih1.InCmCd) as InCmCd, max(Ih1.InPrnDt) as InDt, max(Ih1.InLmgCst) as InLmgCst, " + _
  '          "max(Ih1.InLmgDt) as InLmgDt, max(Ih1.InLmgSal) as InLmgSal, max(Ih1.InVaDC) as InVaDC, max(Ih1.InVaX) as InVaX, " + _
  '          "max(case when Ih1.InRsMulDiv= 'D' and convert (decimal (16, 4), Ih1.InCnvRt)> 0 then 1/ Ih1.InCnvRt else Ih1.InCnvRt end) as InCnvRt, " + _
  '          ws_PrdDesc + ", sum(Ih.InPkgWt) as InPkgWt, max(Ih1.InGrNo) as InGrNo, max(Ih1.InGrDt) as InGrDt, " + _
  '          "max(Ih1.InSbNo) as InSbNo, max(Ih1.InSbDt) as InSbDt, max(Ih1.InAwbNo) as InAwbNo, max(Ih1.InAwbDt) as InAwbDt, " + _
  '          "max(Ih1.InHAwbNo) as InHAwbNo, max(Ih1.InHAwbDt) as InHAwbDt, max(Ih1.InHBinNo) as InHBinNo, " + _
  '          "max(Ih1.InGspNo) as InGspNo, max(Ih1.InGspDt) as InGspDt, max(Ih1.InChqNo) as InChqNo, " + _
  '          "sum(Ih.InChqAmt) as InChqAmt, max(Ih1.InPoNo) as InPoNo, max(Ih1.InPoDt) as InPoDt, " + _
  '          "max(Ih1.InCmPayTerms) as InCmPayTerms, sum(Ih.InFOB) as InFOB, sum(Ih.InFOBRs) as InFOBRs, " + _
  '          "sum(Ih.InComm) as InComm, sum(Ih.InCommRs) as InCommRs, sum(Ih.InFrt) as InFrt, sum(Ih.InFrtRs) as InFrtRs, " + _
  '          "sum(Ih.InIns) as InIns, sum(Ih.InInsRs) as InInsRs, sum(Ih.InTot) as InTot, sum(Ih.InTotRs) as InTotRs, " + _
  '          ws_InTotAs + ", sum(Ih.InDiaCost) as InDiaCost, sum(Ih.InCSCost) as InCSCost, sum(Ih.InAdv) as InAdv, " + _
  '          "max(Ih.InDiaCostDesc) as InDiaCostDesc, max(Ih.InCSCostDesc) as InCSCostDesc, max(Ih.InAdvDesc) as InAdvDesc, "
  '  wStr2 = " max(Ih1.InCls) as InCls, max(Ih1.InHBkName) as InHBkName, max(Ih1.InHBkAdd1) as InHBkAdd1, " + _
  '          "max(Ih1.InHBkAdd2) as InHBkAdd2, max(Ih1.InHBkAdd3) as InHBkAdd3, max(Ih1.InHBkAc) as InHBkAc, " + _
  '          "max(Ih1.InHFrgnBkName) as InHFrgnBkName, max(Ih1.InHFrgnBkAc) as InHFrgnBkAc, max(Ih1.InHExpOff) as InHExpOff, " + _
  '          "max(Ih1.InHRbiNo) as InHRbiNo, max(Ih1.InHIecNo) as InHIecNo, max(Ih1.InHPreCarriage) as InHPreCarriage, " + _
  '          "max(Ih1.InHPreCarrier) as InHPreCarrier, max(Ih1.InHVessel) as InHVessel, max(Ih1.InHPortOfLd) as InHPortOfLd, " + _
  '          "max(Ih1.InHConOfOrg) as InHConOfOrg, max(Ih1.InHMarks) as InHMarks, max(Ih1.InHPkgDesc) as InHPkgDesc, " + _
  '          "max(Ih1.InHItcNo) as InHItcNo, max(Ih1.InHEtcZone) as InHEtcZone, max(Ih1.InHEtcSection) as InHEtcSection, " + _
  '          "max(Ih1.InHConsignee) as InHConsignee, max(Ih1.InCmInsBy) as InHInsBy, max(Ih1.InHInsNo) as InHInsNo, " + _
  '          "max(Ih1.InHAgName) as InHAgName, max(Ih1.InHAgAdd1) as InHAgAdd1, max(Ih1.InHAgAdd2) as InHAgAdd2, " + _
  '          "max(Ih1.InHAgAdd3) as InHAgAdd3, max(Ih1.InCmName) as InCmName, max(Ih1.InCmAdd1) as InCmAdd1, " + _
  '          "max(Ih1.InCmAdd2) as InCmAdd2, max(Ih1.InCmAdd3) as InCmAdd3, max(Ih1.InCmTel) as InCmTel, " + _
  '          "max(Ih1.InCmFax) as InCmFax, max(Ih1.InCmEMail) as InCmEMail, max(Ih1.InCmCpName) as InCmCpName, " + _
  '          "max(Ih1.InCmPortOfDisch) as InCmPortOfDisch, max(Ih1.InCmFinDest) as InCmFinDest, " + _
  '          "max(Ih1.InCmBuyer) as InCmBuyer, max(Ih1.InSelfName) as InSelfName, max(Ih1.InSelfAdd1) as InSelfAdd1, " + _
  '          "max(Ih1.InSelfAdd2) as InSelfAdd2, max(Ih1.InSelfAdd3) as InSelfAdd3, max(Ih1.InSelfTel) as InSelfTel, " + _
  '          "max(Ih1.InSelfFax) as InSelfFax, max(Ih1.InSelfEMail) as InSelfEMail, max(Ih1.InhAuthSgn) as InhAuthSgn " + _
  '          "from InvHd Ih " + _
  '          "join InvHd Ih1 on " + ws_MinCnd + wCnd
  '
  '  'wStr2 = wStr2 + " group by InTc, InYy, InChr, InNo "
  '
  '  wSqlStrg = wStr1 + wStr2 + " Order By Ih.InTc, Ih.InYy, Ih.InChr, Ih.InNo "
  'Else
  ''*** (Jen 2.14)
  '
  '  wrepcnd = adc.RepCond
  '  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  '
  '  '****** introduced the alias Ih in 2.14
  '  wStr1 = "Select InTc, InYy, InChr, InNo, max(InExpNo) as InExpNo, max(InCmCd) as InCmCd, max(InPrnDt) as InDt, " + _
  '          "max(InLmgCst) as InLmgCst, max(InLmgDt) as InLmgDt, max(InLmgSal) as InLmgSal, max(InVaDC) as InVaDC, max(InVaX) as InVaX, " + _
  '          "max(case when InRsMulDiv= 'D' and convert (decimal (16, 4), InCnvRt)> 0 then 1/ InCnvRt else InCnvRt end) as InCnvRt, " + _
  '          "max(InPrdDesc) as InPrdDesc, max(InPkgWt) as InPkgWt, max(InGrNo) as InGrNo, max(InGrDt) as InGrDt, max(InSbNo) as InSbNo, " + _
  '          "max(InSbDt) as InSbDt, max(InAwbNo) as InAwbNo, max(InAwbDt) as InAwbDt, max(InHAwbNo) as InHAwbNo, max(InHAwbDt) as InHAwbDt, max(InHBinNo) as InHBinNo, " + _
  '          "max(InGspNo) as InGspNo, max(InGspDt) as InGspDt, max(InChqNo) as InChqNo, " + _
  '          "max(InChqAmt) as InChqAmt, max(InPoNo) as InPoNo, max(InPoDt) as InPoDt, max(InCmPayTerms) as InCmPayTerms, max(InFOB) as InFOB, " + _
  '          "max(InFOBRs) as InFOBRs, max(InComm) as InComm, max(InCommRs) as InCommRs, max(InFrt) as InFrt, max(InFrtRs) as InFrtRs, max(InIns) as InIns, " + _
  '          "max(InInsRs) as InInsRs, max(InTot) as InTot, max(InTotRs) as InTotRs, max(InTotAs) as InTotAs, max(InDiaCost) as InDiaCost, " + _
  '          "max(InCSCost) as InCSCost, max(InAdv) as InAdv, max(InDiaCostDesc) as InDiaCostDesc, max(InCSCostDesc) as InCSCostDesc, max(InAdvDesc) as InAdvDesc, " + _
  '          "max(InCls) as InCls, max(InHBkName) as InHBkName, max(InHBkAdd1) as InHBkAdd1, max(InHBkAdd2) as InHBkAdd2, max(InHBkAdd3) as InHBkAdd3, " + _
  '          "max(InHBkAc) as InHBkAc, max(InHFrgnBkName) as InHFrgnBkName, max(InHFrgnBkAc) as InHFrgnBkAc, max(InHExpOff) as InHExpOff, max(InHRbiNo) as InHRbiNo, " + _
  '          "max(InHIecNo) as InHIecNo, max(InHPreCarriage) as InHPreCarriage, max(InHPreCarrier) as InHPreCarrier, max(InHVessel) as InHVessel, " + _
  '          "max(InHPortOfLd) as InHPortOfLd, max(InHConOfOrg) as InHConOfOrg, max(InHMarks) as InHMarks, max(InHPkgDesc) as InHPkgDesc, " + _
  '          "max(InHItcNo) as InHItcNo, max(InHEtcZone) as InHEtcZone, max(InHEtcSection) as InHEtcSection, max(InHConsignee) as InHConsignee, " + _
  '          "max(InCmInsBy) as InHInsBy, max(InHInsNo) as InHInsNo, max(InHAgName) as InHAgName, max(InHAgAdd1) as InHAgAdd1, max(InHAgAdd2) as InHAgAdd2, " + _
  '          "max(InHAgAdd3) as InHAgAdd3, max(InCmName) as InCmName, max(InCmAdd1) as InCmAdd1, max(InCmAdd2) as InCmAdd2, max(InCmAdd3) as InCmAdd3, " + _
  '          "max(InCmTel) as InCmTel, max(InCmFax) as InCmFax, max(InCmEMail) as InCmEMail, max(InCmCpName) as InCmCpName, max(InCmPortOfDisch) as InCmPortOfDisch, " + _
  '          "max(InCmFinDest) as InCmFinDest, max(InCmBuyer) as InCmBuyer, max(InSelfName) as InSelfName, max(InSelfAdd1) as InSelfAdd1, " + _
  '          "max(InSelfAdd2) as InSelfAdd2, max(InSelfAdd3) as InSelfAdd3, max(InSelfTel) as InSelfTel, max(InSelfFax) as InSelfFax, max(InSelfEMail) as InSelfEMail, max(InhAuthSgn) as InhAuthSgn, "
  '
  '  wStr2 = "sum(IsNull(TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('D') then 1 else 0 end)) as qDiaVal, " + _
  '          "sum(IsNull(TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('C') then 1 else 0 end)) as qCSVal, " + _
  '          "sum(IsNull(TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('G') then 1 else 0 end)) as qGldVal, " + _
  '          "sum(IsNull(TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('P') then 1 else 0 end)) as qPlVal, " + _
  '          "sum(IsNull(TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('S') then 1 else 0 end)) as qSlVal, " + _
  '          "sum(IsNull(TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('X') then 1 else 0 end)) as qAccsVal " + _
  '          "from InvHd Ih Left Outer Join TxndZ on TzCoCd = InCoCd and TzTc = InTc " + _
  '          "and TzYy= InYy and TzChr= InChr and TzNo= InNo " + _
  '          "Left Outer Join RmMst on RmCd = TzRmZCd " + _
  '           wCnd + " group by InTc, InYy, InChr, InNo "
  '
  '  wSqlStrg = wStr1 + wStr2 + " Order By InTc, InYy, InChr, InNo "
  '
  ''*** (Jen 2.14)
  'End If
  ''*** (Jen 2.14)
  '*** (Bef 3.01)
  
  '*** (Jen 3.01)
  ' **** Manali 3.6.0 - InCircularNo, InRegNo,InGspAdd1, InGspAdd2, InGspAdd3 added [InterCarat]
  ' **** Manali 3.8.0 - InHHideCommYN added
  ' **** Manali 3.8.1 - Self Payable In Invoice - InDsc, InDscRs, InDeduct, InDeductRs, InHSelfDscYN, InHSelfDeductYN added
  wHdStr1 = "Select max(Ih1.InTc) as InTc, max(Ih1.InYy) as InYy, max(Ih1.InChr) as InChr, max(Ih1.InNo) as InNo, " + _
            " " + ws_InExpNo + ", max(Ih1.InCmCd) as InCmCd, max(Ih1.InPrnDt) as InDt, max(Ih1.InLmgCst) as InLmgCst, " + _
            "max(Ih1.InLmgDt) as InLmgDt, max(Ih1.InLmgSal) as InLmgSal, max(Ih1.InVaDC) as InVaDC, max(Ih1.InVaX) as InVaX, " + _
            "max(case when Ih1.InRsMulDiv= 'D' and convert (decimal (16, 4), Ih1.InCnvRt)> 0 then 1/ Ih1.InCnvRt else Ih1.InCnvRt end) as InCnvRt, " + _
            ws_PrdDesc + ", sum(Ih.InPkgWt) as InPkgWt, max(Ih1.InGrNo) as InGrNo, max(Ih1.InGrDt) as InGrDt, " + _
            "max(Ih1.InSbNo) as InSbNo, max(Ih1.InSbDt) as InSbDt, max(Ih1.InAwbNo) as InAwbNo, max(Ih1.InAwbDt) as InAwbDt, " + _
            "max(Ih1.InHAwbNo) as InHAwbNo, max(Ih1.InHAwbDt) as InHAwbDt, max(Ih1.InHBinNo) as InHBinNo, " + _
            "max(Ih1.InGspNo) as InGspNo, max(Ih1.InGspDt) as InGspDt, max(Ih1.InChqNo) as InChqNo, " + _
            "sum(Ih.InChqAmt) as InChqAmt, max(Ih1.InPoNo) as InPoNo, max(Ih1.InPoDt) as InPoDt, " + _
            "max(Ih1.InCmPayTerms) as InCmPayTerms, sum(Ih.InFOB) as InFOB, sum(Ih.InFOBRs) as InFOBRs, " + _
            "sum(Ih.InComm) as InComm, sum(Ih.InCommRs) as InCommRs, sum(Ih.InFrt) as InFrt, sum(Ih.InFrtRs) as InFrtRs, " + _
            "sum(Ih.InIns) as InIns, sum(Ih.InInsRs) as InInsRs, sum(Ih.InDsc) as InDsc, sum(Ih.InDscRs) as InDscRs, sum(Ih.InDeduct) as InDeduct, sum(Ih.InDeductRs) as InDeductRs, " + _
            "sum(Ih.InTot) as InTot, sum(Ih.InTotRs) as InTotRs, " + _
            ws_InTotAs + ", sum(Ih.InDiaCost) as InDiaCost, sum(Ih.InCSCost) as InCSCost, sum(Ih.InAdv) as InAdv, " + _
            "max(Ih.InDiaCostDesc) as InDiaCostDesc, max(Ih.InCSCostDesc) as InCSCostDesc, max(Ih.InAdvDesc) as InAdvDesc, "
  '6.1
  wHdStr2 = " max(Ih1.InCls) as InCls, max(Ih1.InHBkName) as InHBkName, max(Ih1.InHBkAdd1) as InHBkAdd1, " + _
            "max(Ih1.InHBkAdd2) as InHBkAdd2, max(Ih1.InHBkAdd3) as InHBkAdd3, max(Ih1.InHBkAc) as InHBkAc, " + _
            "max(Ih1.InHFrgnBkName) as InHFrgnBkName, max(Ih1.InHFrgnBkAc) as InHFrgnBkAc, max(Ih1.InHExpOff) as InHExpOff, " + _
            "max(Ih1.InHRbiNo) as InHRbiNo, max(Ih1.InHIecNo) as InHIecNo, max(Ih1.InHPreCarriage) as InHPreCarriage, " + _
            "max(Ih1.InHPreCarrier) as InHPreCarrier, max(Ih1.InHVessel) as InHVessel, max(Ih1.InHPortOfLd) as InHPortOfLd, " + _
            "max(Ih1.InHConOfOrg) as InHConOfOrg, max(Ih1.InHMarks) as InHMarks, max(Ih1.InHPkgDesc) as InHPkgDesc, " + _
            "max(Ih1.InHItcNo) as InHItcNo, max(Ih1.InHEtcZone) as InHEtcZone, max(Ih1.InHEtcSection) as InHEtcSection, " + _
            "max(Ih1.InHConsignee) as InHConsignee, max(Ih1.InCmInsBy) as InHInsBy, max(Ih1.InHInsNo) as InHInsNo, " + _
            "max(Ih1.InHAgName) as InHAgName, max(Ih1.InHAgAdd1) as InHAgAdd1, max(Ih1.InHAgAdd2) as InHAgAdd2, " + _
            "max(Ih1.InHAgAdd3) as InHAgAdd3, max(Ih1.InCmName) as InCmName, max(Ih1.InCmAdd1) as InCmAdd1, " + _
            "max(Ih1.InCmAdd2) as InCmAdd2, max(Ih1.InCmAdd3) as InCmAdd3, max(Ih1.InCmTel) as InCmTel, " + _
            "max(Ih1.InCmFax) as InCmFax, max(Ih1.InCmEMail) as InCmEMail, max(Ih1.InCmCpName) as InCmCpName, " + _
            "max(Ih1.InCmPortOfDisch) as InCmPortOfDisch, max(Ih1.InCmFinDest) as InCmFinDest, " + _
            "max(Ih1.InCmBuyer) as InCmBuyer, max(Ih1.InSelfName) as InSelfName, max(Ih1.InSelfAdd1) as InSelfAdd1, " + _
            "max(Ih1.InSelfAdd2) as InSelfAdd2, max(Ih1.InSelfAdd3) as InSelfAdd3, max(Ih1.InSelfTel) as InSelfTel, " + _
            "max(Ih1.InSelfFax) as InSelfFax, max(Ih1.InSelfEMail) as InSelfEMail, max(Ih1.InhAuthSgn) as InhAuthSgn, " + _
            "max(Ih1.InCircularNo) as InCircularNo, max(Ih1.InRegNo) as InRegNo, max(Ih1.InGspAdd1) as InGspAdd1, max(Ih1.InGspAdd2) as InGspAdd2, max(Ih1.InGspAdd3) as InGspAdd3, " + _
            "max(Ih1.InHHideCommYN) as InHHideCommYN, max(Ih1.InHSelfDscYN) as InHSelfDscYN, max(Ih1.InHSelfDeductYN) as InHSelfDeductYN, max(Ih1.InHSCode) as InHSCode  " + _
            ",'" + UCase(MWLib.DtSetting) + "' as qDtsetting "
    
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  If UCase(adc.MenuCd) = UCase("RepGRFormBack") Or UCase(adc.MenuCd) = UCase("RepGRFormBackMulti") Or _
      UCase(adc.MenuCd) = UCase("RepFgGrBk") Or UCase(adc.MenuCd) = UCase("RepJtGrBk") Then
  '*** Jay 3.3(FG)
      wHdStr3 = " "
    wHdStr4 = " from InvHd Ih " + _
              "join InvHd Ih1 on " + ws_MinCnd + wCnd
    
    wSqlStrg = wHdStr1 + wHdStr2 + wHdStr3 + wHdStr4
  Else
    ' ***** Manali 3.8.0 - Value Categorised based on TImpCtg For GspDetLet
    If Not (UCase(adc.MenuCd) = UCase("RepGSPDetLet") Or UCase(adc.MenuCd) = UCase("RepGSPDetLetMulti")) Then
      wHdStr3 = ", 0 as qDiaVal, 0 as qCSVal, 0 as qGldVal, 0 as qPlVal, 0 as qSlVal, 0 as qPdVal, 0 as qAccsVal "
    Else
      wHdStr3 = ", 0 as qDiaVal, 0 as qCSVal, 0 as qGldVal, 0 as qPlVal, 0 as qSlVal, 0 as qPdVal, 0 as qAccsVal " + _
                ", 0 as qDiaIndVal, 0 as qCSIndVal, 0 as qGldIndVal, 0 as qPlIndVal, 0 as qSlIndVal, 0 as qPdIndVal, 0 as qAccsIndVal "
    End If
    
    wHdStr4 = " from InvHd Ih " + _
              "join InvHd Ih1 on " + ws_MinCnd + wCnd + " "
    
    wStr1 = " Select max(Ih1.InTc) as InTc, max(Ih1.InYy) as InYy, max(Ih1.InChr) as InChr, max(Ih1.InNo) as InNo, " + _
            " " + ws_InExpNo + ", max(Ih1.InCmCd) as InCmCd, max(Ih1.InPrnDt) as InDt, max(Ih1.InLmgCst) as InLmgCst, " + _
            "max(Ih1.InLmgDt) as InLmgDt, max(Ih1.InLmgSal) as InLmgSal, max(Ih1.InVaDC) as InVaDC, max(Ih1.InVaX) as InVaX, " + _
            "max(case when Ih1.InRsMulDiv= 'D' and convert (decimal (16, 4), Ih1.InCnvRt)> 0 then 1/ Ih1.InCnvRt else Ih1.InCnvRt end) as InCnvRt, " + _
            ws_PrdDesc + ", 0 as InPkgWt, max(Ih1.InGrNo) as InGrNo, max(Ih1.InGrDt) as InGrDt, " + _
            "max(Ih1.InSbNo) as InSbNo, max(Ih1.InSbDt) as InSbDt, max(Ih1.InAwbNo) as InAwbNo, max(Ih1.InAwbDt) as InAwbDt, " + _
            "max(Ih1.InHAwbNo) as InHAwbNo, max(Ih1.InHAwbDt) as InHAwbDt, max(Ih1.InHBinNo) as InHBinNo, " + _
            "max(Ih1.InGspNo) as InGspNo, max(Ih1.InGspDt) as InGspDt, max(Ih1.InChqNo) as InChqNo, " + _
            "0 as InChqAmt, max(Ih1.InPoNo) as InPoNo, max(Ih1.InPoDt) as InPoDt, " + _
            "max(Ih1.InCmPayTerms) as InCmPayTerms, 0 as InFOB, 0 as InFOBRs, " + _
            "0 as InComm, 0 as InCommRs, 0 as InFrt, 0 as InFrtRs, " + _
            "0 as InIns, 0 as InInsRs, 0 as InDsc, 0 as InDscRs, 0 as InDeduct, 0 as InDeductRs, 0 as InTot, 0 as InTotRs, " + _
            ws_InTotAs + ", 0 as InDiaCost, 0 as InCSCost, 0 as InAdv, " + _
            "max(Ih.InDiaCostDesc) as InDiaCostDesc, max(Ih.InCSCostDesc) as InCSCostDesc, max(Ih.InAdvDesc) as InAdvDesc, "
    '6.1
    wStr2 = " max(Ih1.InCls) as InCls, max(Ih1.InHBkName) as InHBkName, max(Ih1.InHBkAdd1) as InHBkAdd1, " + _
            "max(Ih1.InHBkAdd2) as InHBkAdd2, max(Ih1.InHBkAdd3) as InHBkAdd3, max(Ih1.InHBkAc) as InHBkAc, " + _
            "max(Ih1.InHFrgnBkName) as InHFrgnBkName, max(Ih1.InHFrgnBkAc) as InHFrgnBkAc, max(Ih1.InHExpOff) as InHExpOff, " + _
            "max(Ih1.InHRbiNo) as InHRbiNo, max(Ih1.InHIecNo) as InHIecNo, max(Ih1.InHPreCarriage) as InHPreCarriage, " + _
            "max(Ih1.InHPreCarrier) as InHPreCarrier, max(Ih1.InHVessel) as InHVessel, max(Ih1.InHPortOfLd) as InHPortOfLd, " + _
            "max(Ih1.InHConOfOrg) as InHConOfOrg, max(Ih1.InHMarks) as InHMarks, max(Ih1.InHPkgDesc) as InHPkgDesc, " + _
            "max(Ih1.InHItcNo) as InHItcNo, max(Ih1.InHEtcZone) as InHEtcZone, max(Ih1.InHEtcSection) as InHEtcSection, " + _
            "max(Ih1.InHConsignee) as InHConsignee, max(Ih1.InCmInsBy) as InHInsBy, max(Ih1.InHInsNo) as InHInsNo, " + _
            "max(Ih1.InHAgName) as InHAgName, max(Ih1.InHAgAdd1) as InHAgAdd1, max(Ih1.InHAgAdd2) as InHAgAdd2, " + _
            "max(Ih1.InHAgAdd3) as InHAgAdd3, max(Ih1.InCmName) as InCmName, max(Ih1.InCmAdd1) as InCmAdd1, " + _
            "max(Ih1.InCmAdd2) as InCmAdd2, max(Ih1.InCmAdd3) as InCmAdd3, max(Ih1.InCmTel) as InCmTel, " + _
            "max(Ih1.InCmFax) as InCmFax, max(Ih1.InCmEMail) as InCmEMail, max(Ih1.InCmCpName) as InCmCpName, " + _
            "max(Ih1.InCmPortOfDisch) as InCmPortOfDisch, max(Ih1.InCmFinDest) as InCmFinDest, " + _
            "max(Ih1.InCmBuyer) as InCmBuyer, max(Ih1.InSelfName) as InSelfName, max(Ih1.InSelfAdd1) as InSelfAdd1, " + _
            "max(Ih1.InSelfAdd2) as InSelfAdd2, max(Ih1.InSelfAdd3) as InSelfAdd3, max(Ih1.InSelfTel) as InSelfTel, " + _
            "max(Ih1.InSelfFax) as InSelfFax, max(Ih1.InSelfEMail) as InSelfEMail, max(Ih1.InhAuthSgn) as InhAuthSgn, " + _
            "max(Ih1.InCircularNo) as InCircularNo, max(Ih1.InRegNo) as InRegNo, max(Ih1.InGspAdd1) as InGspAdd1, max(Ih1.InGspAdd2) as InGspAdd2, max(Ih1.InGspAdd3) as InGspAdd3, " + _
            "max(Ih1.InHHideCommYN) as InHHideCommYN, max(Ih1.InHSelfDscYN) as InHSelfDscYN, max(Ih1.InHSelfDeductYN) as InHSelfDeductYN, max(Ih1.InHSCode) as InHSCode  " + _
            ",'" + UCase(MWLib.DtSetting) + "' as qDtsetting "
            
    
    ' ***** Manali 3.03 - 05/07/08 - Pd Mod - 'L' added
    ' ***** Manali 3.8.0 - Value Categorised based on TImpCtg For GspDetLet
    If Not (UCase(adc.MenuCd) = UCase("RepGSPDetLet") Or UCase(adc.MenuCd) = UCase("RepGSPDetLetMulti")) Then
      wStr3 = ", sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('D') then 1 else 0 end)) as qDiaVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('C') then 1 else 0 end)) as qCSVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('G') then 1 else 0 end)) as qGldVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('P') then 1 else 0 end)) as qPlVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('S') then 1 else 0 end)) as qSlVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('L','M') then 1 else 0 end)) as qPdVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(RmCtg, '') in ('X') then 1 else 0 end)) as qAccsVal "

    wStr4 = " from InvHd Ih " + _
            "join InvHd Ih1 on " + ws_MinCnd + _
            " Left Outer Join TxndZ M on TzInIdNo=Ih.InIdNo and TzCoCd= Ih.InCoCd and TzTc= Ih.InTc and TzYy= Ih.InYy and TzChr= Ih.InChr and TzNo= Ih.InNo " + IIF(gs_Partition = ctCurrPrtn, " and TzPrtKey=Ih.InPrtKey ", "") + _
            "Left Outer Join RmMst on RmCd = TzRmZCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=TzPrtKey ", "") + _
            wCnd + IIF(gs_Partition = ctCurrPrtn, " and Ih.InPrtKey='" + ctCurrPrtn + "' ", "")
    
    
    Else
      wStr3 = ", sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IMP' and IsNull(RmCtg, '') in ('D') then 1 else 0 end)) as qDiaVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IMP' and IsNull(RmCtg, '') in ('C') then 1 else 0 end)) as qCSVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IMP' and IsNull(RmCtg, '') in ('G') then 1 else 0 end)) as qGldVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IMP' and IsNull(RmCtg, '') in ('P') then 1 else 0 end)) as qPlVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IMP' and IsNull(RmCtg, '') in ('S') then 1 else 0 end)) as qSlVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IMP' and IsNull(RmCtg, '') in ('L','M') then 1 else 0 end)) as qPdVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IMP' and IsNull(RmCtg, '') in ('X') then 1 else 0 end)) as qAccsVal " + _
              ", sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IND' and IsNull(RmCtg, '') in ('D') then 1 else 0 end)) as qDiaIndVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IND' and IsNull(RmCtg, '') in ('C') then 1 else 0 end)) as qCSIndVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IND' and IsNull(RmCtg, '') in ('G') then 1 else 0 end)) as qGldIndVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IND' and IsNull(RmCtg, '') in ('P') then 1 else 0 end)) as qPlIndVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IND' and IsNull(RmCtg, '') in ('S') then 1 else 0 end)) as qSlIndVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IND' and IsNull(RmCtg, '') in ('L','M') then 1 else 0 end)) as qPdIndVal, " + _
              "sum(IsNull(M.TzPurAmt, 0) * (case when IsNull(TImpCtg, '')='IND' and IsNull(RmCtg, '') in ('X') then 1 else 0 end)) as qAccsIndVal "
    
    '****** Sachin 3.02 - Id fields In Joins
    wStr4 = " from InvHd Ih " + _
            "join InvHd Ih1 on " + ws_MinCnd + _
            " Left Outer Join TxndZ M on TzInIdNo=Ih.InIdNo and TzCoCd= Ih.InCoCd and TzTc= Ih.InTc and TzYy= Ih.InYy and TzChr= Ih.InChr and TzNo= Ih.InNo " + IIF(gs_Partition = ctCurrPrtn, " and TzPrtKey=Ih.InPrtKey ", "") + _
            " Left Outer Join TxndZ B On M.TzCoCd=B.TzCoCd and M.TzRefYy=B.TzYy " + IIF(gs_Partition = ctCurrPrtn, " and M.TzRefkey=B.TzKey ", "") + _
            " Left Outer Join Txn On TIdNo=B.TzTIdNo and TCoCd=B.TzCoCd and TTc=B.TzTc and TYy=B.TzYy and TChr=B.TzChr and TNo=B.TzNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=B.TzPrtKey ", "") + _
            "Left Outer Join RmMst on RmCd = M.TzRmZCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=M.TzPrtKey ", "") + _
            wCnd + IIF(gs_Partition = ctCurrPrtn, " and Ih.InPrtKey='" + ctCurrPrtn + "' ", "")
    
    End If
    
    
    
    wFinalStr1 = "Select max(InTc) as InTc, max(InYy) as InYy, max(InChr) as InChr, max(InNo) as InNo, " + _
                 " max(InExpNo) as InExpNo, max(InCmCd) as InCmCd, max(InDt) as InDt, max(InLmgCst) as InLmgCst, " + _
                 "max(InLmgDt) as InLmgDt, max(InLmgSal) as InLmgSal, max(InVaDC) as InVaDC, max(InVaX) as InVaX, " + _
                 "max(InCnvRt) as InCnvRt, " + _
                 "max(InPrdDesc) as InPrdDesc, sum(InPkgWt) as InPkgWt, max(InGrNo) as InGrNo, max(InGrDt) as InGrDt, " + _
                 "max(InSbNo) as InSbNo, max(InSbDt) as InSbDt, max(InAwbNo) as InAwbNo, max(InAwbDt) as InAwbDt, " + _
                 "max(InHAwbNo) as InHAwbNo, max(InHAwbDt) as InHAwbDt, max(InHBinNo) as InHBinNo, " + _
                 "max(InGspNo) as InGspNo, max(InGspDt) as InGspDt, max(InChqNo) as InChqNo, " + _
                 "sum(InChqAmt) as InChqAmt, max(InPoNo) as InPoNo, max(InPoDt) as InPoDt, " + _
                 "max(InCmPayTerms) as InCmPayTerms, sum(InFOB) as InFOB, sum(InFOBRs) as InFOBRs, " + _
                 "sum(InComm) as InComm, sum(InCommRs) as InCommRs, sum(InFrt) as InFrt, sum(InFrtRs) as InFrtRs, " + _
                 "sum(InIns) as InIns, sum(InInsRs) as InInsRs, " + _
                 "Sum(InDsc) as InDsc, Sum(InDscRs) as InDscRs, Sum(InDeduct) as InDeduct, Sum(InDeductRs) as InDeductRs,  " + _
                 "sum(InTot) as InTot, sum(InTotRs) as InTotRs, " + _
                 "max(InTotAs) as InTotAs, sum(InDiaCost) as InDiaCost, sum(InCSCost) as InCSCost, sum(InAdv) as InAdv, " + _
                 "max(InDiaCostDesc) as InDiaCostDesc, max(InCSCostDesc) as InCSCostDesc, max(InAdvDesc) as InAdvDesc, "
  '6.1
    wFinalStr2 = " max(InCls) as InCls, max(InHBkName) as InHBkName, max(InHBkAdd1) as InHBkAdd1, " + _
                 "max(InHBkAdd2) as InHBkAdd2, max(InHBkAdd3) as InHBkAdd3, max(InHBkAc) as InHBkAc, " + _
                 "max(InHFrgnBkName) as InHFrgnBkName, max(InHFrgnBkAc) as InHFrgnBkAc, max(InHExpOff) as InHExpOff, " + _
                 "max(InHRbiNo) as InHRbiNo, max(InHIecNo) as InHIecNo, max(InHPreCarriage) as InHPreCarriage, " + _
                 "max(InHPreCarrier) as InHPreCarrier, max(InHVessel) as InHVessel, max(InHPortOfLd) as InHPortOfLd, " + _
                 "max(InHConOfOrg) as InHConOfOrg, max(InHMarks) as InHMarks, max(InHPkgDesc) as InHPkgDesc, " + _
                 "max(InHItcNo) as InHItcNo, max(InHEtcZone) as InHEtcZone, max(InHEtcSection) as InHEtcSection, " + _
                 "max(InHConsignee) as InHConsignee, max(InHInsBy) as InHInsBy, max(InHInsNo) as InHInsNo, " + _
                 "max(InHAgName) as InHAgName, max(InHAgAdd1) as InHAgAdd1, max(InHAgAdd2) as InHAgAdd2, " + _
                 "max(InHAgAdd3) as InHAgAdd3, max(InCmName) as InCmName, max(InCmAdd1) as InCmAdd1, " + _
                 "max(InCmAdd2) as InCmAdd2, max(InCmAdd3) as InCmAdd3, max(InCmTel) as InCmTel, " + _
                 "max(InCmFax) as InCmFax, max(InCmEMail) as InCmEMail, max(InCmCpName) as InCmCpName, " + _
                 "max(InCmPortOfDisch) as InCmPortOfDisch, max(InCmFinDest) as InCmFinDest, " + _
                 "max(InCmBuyer) as InCmBuyer, max(InSelfName) as InSelfName, max(InSelfAdd1) as InSelfAdd1, " + _
                 "max(InSelfAdd2) as InSelfAdd2, max(InSelfAdd3) as InSelfAdd3, max(InSelfTel) as InSelfTel, " + _
                 "max(InSelfFax) as InSelfFax, max(InSelfEMail) as InSelfEMail, max(InhAuthSgn) as InhAuthSgn, " + _
                 "max(InCircularNo) as InCircularNo, max(InRegNo) as InRegNo, max(InGspAdd1) as InGspAdd1, max(InGspAdd2) as InGspAdd2, max(InGspAdd3) as InGspAdd3, " + _
                 "max(InHHideCommYN) as InHHideCommYN, max(InHSelfDscYN) as InHSelfDscYN, max(InHSelfDeductYN) as InHSelfDeductYN, max(InHSCode) as InHSCode  " + _
                 ",max(qDtsetting) as qDtsetting "


    ' **** Manali 3.03 - 05/07/08 - Pd Mod - Pd added ('L')
    If Not (UCase(adc.MenuCd) = UCase("RepGSPDetLet") Or UCase(adc.MenuCd) = UCase("RepGSPDetLetMulti")) Then
      wFinalStr3 = ", sum(qDiaVal) as qDiaVal, sum(qCSVal) as qCSVal, " + _
                   "sum(qGldVal) as qGldVal, sum(qPlVal) as qPlVal, " + _
                   "sum(qSlVal) as qSlVal, sum(qPdVal) as qPDVal, sum(qAccsVal) as qAccsVal "
    Else
      wFinalStr3 = ", sum(qDiaVal) as qDiaVal, sum(qCSVal) as qCSVal, " + _
                   "sum(qGldVal) as qGldVal, sum(qPlVal) as qPlVal, " + _
                   "sum(qSlVal) as qSlVal, sum(qPdVal) as qPDVal, sum(qAccsVal) as qAccsVal " + _
                   ", sum(qDiaIndVal) as qDiaIndVal, sum(qCSIndVal) as qCSIndVal, " + _
                   "sum(qGldIndVal) as qGldIndVal, sum(qPlIndVal) as qPlIndVal, " + _
                   "sum(qSlIndVal) as qSlIndVal, sum(qPdIndVal) as qPDIndVal, sum(qAccsIndVal) as qAccsIndVal "
    End If
    
    wSqlStrg = wFinalStr1 + wFinalStr2 + wFinalStr3 + _
               "From ( " + _
               wHdStr1 + wHdStr2 + wHdStr3 + wHdStr4 + _
               " Union All " + _
               wStr1 + wStr2 + wStr3 + wStr4 + _
               ") TmpTbl "
  End If
  '*** (Jen 3.01)
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt

  
  Select Case UCase(adc.MenuCd)
  Case Is = UCase("RepDeclnLet"), UCase("RepDeclnLetMulti")
    '****** (Jen 3.01 added RepDeclnLetMulti)
    Dim wRsInDsg As MDORowSet
    
    '*** (Bef 3.01)
    'Set wRsInDsg = moCn.OpenRes("Select sum(case when DmUom= 'PCS' then IdQty else 0 end) as qPcs, " + _
    '               "sum(case when DmUom= 'PRS' then IdQty else 0 end) as qPrs, " + _
    '               "sum(case when DmUom= 'SET' then IdQty else 0 end) as qSet " + _
    '               "from InvDsg join DsgMst on DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + _
    '               "where IdCoCd= '" + adc("UoCoCdFr") + "' and " + _
    '               "IdTc= '" + adc("UoInTcFr") + "' and IdYy= '" + adc("UoInYyFr") + "' and " + _
    '               "IdChr= '" + adc("UoInChrFr") + "' and IdNo= " + CStr(adc("UoInNoFr")) + " " + _
    '               "Group By IdTc, IdYy, IdChr, IdNo ")
    '*** (Bef 3.01)
    
    '*** (Jen 3.01)
    '****** Sachin 3.02 - Id fields In Joins
    Set wRsInDsg = moCn.OpenRes("Select sum(case when DmUom= 'PCS' then IdQty else 0 end) as qPcs, " + _
                   "sum(case when DmUom= 'PRS' then IdQty else 0 end) as qPrs, " + _
                   "sum(case when DmUom= 'SET' then IdQty else 0 end) as qSet " + _
                   "from InvHd Ih " + _
                   "join InvDsg on IdInIdNo=Ih.InIdNo and IdCoCd= Ih.InCoCd and IdTc= Ih.InTc and IdYy= Ih.InYy and IdChr= Ih.InChr and IdNo= Ih.InNo " + IIF(gs_Partition = ctCurrPrtn, " and IdPrtKey=Ih.InPrtKey ", "") + _
                   "join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=IdPrtKey ", "") + wCnd + IIF(gs_Partition = ctCurrPrtn, " and Ih.InPrtKey='" + ctCurrPrtn + "' ", ""))
    '*** (Jen 3.01)
    
    If wRsInDsg.RecCount > 0 Then
      GRP_REP.SetFormula mRep, "wTotPcs", wRsInDsg!qPcs
      GRP_REP.SetFormula mRep, "wTotPrs", wRsInDsg!qPrs
      GRP_REP.SetFormula mRep, "wTotSets", wRsInDsg!qSet
    End If
  
    '*** (Jen 2.12)
    '*** (Bef 3.01)
    'Dim ws_SubRep As String, ws_SubRepCnd
    'ws_SubRepCnd = adc.RepCond
    'ws_SubRepCnd = IIF(ws_SubRepCnd <> "", " and " + ws_SubRepCnd, " ")
    '*** (Bef 3.01)
    
    '****** (Bef 2.13) Round(sum(TzPurAmt)* max(InCnvRt), 0) as qCSValRs
    '****** (Jen 2.13) Round(sum(TzPurAmt)* max(case when InRsMulDiv= 'D' and convert (decimal (16, 4), InCnvRt)> 0 then 1/ InCnvRt else InCnvRt end), 0) as qCSValRs
    
    '*** (Bef 3.01)
    ''****** introduced the alias Ih in 2.14
    'ws_SubRep = "Select IsNull(PSCd, 'ZClrStn') as qStnCls, " + _
    '            "IsNull(max(PDesc), 'ColStn')+ ' Value' as qStnClsDesc, " + _
    '            "sum(TzPurAmt) as qCSVal, Round(sum(TzPurAmt)* max(case when InRsMulDiv= 'D' and convert (decimal (16, 4), InCnvRt)> 0 then 1/ InCnvRt else InCnvRt end), 0) as qCSValRs " + _
    '            "from InvHd Ih Join TxndZ on TzCoCd = InCoCd and TzTc = InTc " + _
    '            "and TzYy= InYy and TzChr= InChr and TzNo= InNo " + _
    '            "join RmMst on RmCd= TzRmZCd " + _
    '            "left outer join Param on PTyp= 'STNCLS' and PmCd= RmCtg " + _
    '            "   and ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' " + _
    '            "where RmCtg= 'C' and RmZ= 'Y' " + ws_SubRepCnd + _
    '            " Group By PSCd Order By qStnCls "
    '*** (Bef 3.01)
    
    '*** (Jen 3.01)
    '****** Sachin 3.02 - Id fields In Joins
    'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
    Dim ws_SubRep As String
    ws_SubRep = "Select IsNull(PSCd, 'ZClrStn') as qStnCls, " + _
                "IsNull(max(PDesc), 'ColStn')+ ' Value' as qStnClsDesc, " + _
                "sum(TzPurAmt) as qCSVal, Round(sum(TzPurAmt)* max(case when InRsMulDiv= 'D' and convert (decimal (16, 4), InCnvRt)> 0 then 1/ InCnvRt else InCnvRt end), 0) as qCSValRs " + _
                "from InvHd Ih Join TxndZ on TzInIdNo=InIdNo and TzCoCd = InCoCd and TzTc = InTc " + _
                "and TzYy= InYy and TzChr= InChr and TzNo= InNo " + IIF(gs_Partition = ctCurrPrtn, " and TzPrtKey=InPrtKey ", "") + _
                "join RmMst on RmCd= TzRmZCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=TzPrtKey ", "") + _
                "Left Outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
                wCnd + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey='" + ctCurrPrtn + "' ", "") + " and RmCtg= 'C' and RmZ= 'Y' " + _
                " Group By PSCd Order By qStnCls "
    '*** (Jen 3.01)
    
    Call moCn.RepRes(Rep.Subreport1.OpenSubreport, ws_SubRep + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
    '*** (Jen 2.12)
    
    '*** (Bef 3.01)
    ''****** introduced the alias Ih in 2.14
    ''*** (Jen 2.13)
    'ws_CurCd = moCn.GetFldVal("Select CmCurCd " + _
    '           "from InvHd Ih join Param on PTyp= 'TC' and PMCd= InTc and PSCd= '' " + _
    '           "join CustMst on CmCtg= (case PValue when 'IN' then 'C' when 'CTB' then 'T' else '' end) and CmCd= InCmCd " + _
    '           "where InCoCd ='" + adc("UoCoCdFr") + "' " + _
    '           "And (  (InTc= '" + adc("UoInTcFr") + "' and InYy= '" + adc("UoInYyFr") + "' " + _
    '           "        and InChr= '" + adc("UoInChrFr") + "' and InNo= " + CStr(adc("UoInNoFr")) + " and InNo> 0) " + _
    '           "    or (InExpNo<> '' and InExpNo = '" + adc("UoInExpNoFr") + "' ) " + _
    '           "    )")
    '*** (Bef 3.01)
    
    '*** (Jen 3.01)
    ws_CurCd = moCn.GetFldVal("Select max(CmCurCd) as CmCurCd " + _
               "from InvHd Ih join Param on PTyp= 'TC' and PMCd= InTc and PSCd= '' " + _
               "join CustMst on CmCtg= (case PValue when 'IN' then 'C' when 'CTB' then 'T' else '' end) " + _
               "and CmCd= InCmCd " + wCnd)
    '*** (Jen 3.01)
    
    ws_CurDesc = moCn.GetFldVal("Select PValue from Param where PTyp= 'CURNCY' and PMCd= '" + ws_CurCd + "'")
    
    mRep.TxtTotVal.SetText "Value In " + ws_CurDesc
    mRep.TxtFOB.SetText "Value In " + ws_CurDesc
    mRep.TxtCurCd.SetText "In " + ws_CurCd
    '*** (Jen 2.13)
    
  '****** (Jen 3.01 added RepGSPCertMulti)
  Case Is = UCase("RepGSPCert"), UCase("RepGSPCertMulti")
  ' **** Manali 3.6.0
  '    If wExpLocal = "Y" Then
  '      Rep1.TxtGsp.Suppress = True
  '      Rep1.TxtGspDTA.Suppress = False
  '    Else
  '      Rep1.TxtGsp.Suppress = False
  '      Rep1.TxtGspDTA.Suppress = True
  '    End If
  ' **** Manali 3.6.0
  '****** (Jen 3.01 added RepGSPDetLetMulti)
  Case Is = UCase("RepGSPDetLet"), UCase("RepGSPDetLetMulti")
    '*** (Jen 2.12)
    Dim wRs_StnCls As MDORowSet, ws_DiaCsStr As String
    Set wRs_StnCls = moCn.OpenRes("Select PSCd, PDesc from Param where PTyp= 'STNCLS' Order By PSCd")
    ws_DiaCsStr = "Diamonds"
    Do While Not (wRs_StnCls.EOF Or wRs_StnCls.BOF)
      ws_DiaCsStr = ws_DiaCsStr + ", " + wRs_StnCls!PDesc
      wRs_StnCls.MoveNext
    Loop
    ws_DiaCsStr = ws_DiaCsStr + " & Col Stones"
    mRep.TxtDiaCS.SetText ws_DiaCsStr
    mRep.TxtDisCsInd.SetText ws_DiaCsStr   ' Manali 3.8.0 -
    '*** (Jen 2.12)
    If wExpLocal = "Y" Then
      Rep2.PHDTA.Suppress = False
    Else
      Rep2.PHDTA.Suppress = True
    End If
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  Case Is = UCase("RepGRFormBack"), UCase("RepGRFormBackMulti"), UCase("RepFgGrBk"), UCase("RepJtGrBk")
  '*** Jay 3.3(FG)
    Dim wRsInHd As MDORowSet
    
    '*** (Bef 2.14)
    'Set wRsInHd = moCn.OpenRes("Select InFOB, InFOBRs, InComm, InCommRs, InFrt, InFrtRs, " + _
                  "InIns, InInsRs, InTot, InTotRs, InTotAs " + _
                  "from InvHd where InCoCd= '" + adc("UoCoCdFr") + "' and " + _
                  "InTc= '" + adc("UoInTcFr") + "' and InYy= '" + adc("UoInYyFr") + "' and " + _
                  "InChr= '" + adc("UoInChrFr") + "' and InNo= " + CStr(adc("UoInNoFr")) + " ")
    '*** (Bef 2.14)
    
    '*** (Jen 2.14)
    Set wRsInHd = moCn.OpenRes("Select sum(Ih.InFOB) as InFOB, sum(Ih.InFOBRs) as InFOBRs, sum(Ih.InComm) as InComm, " + _
                  "sum(Ih.InCommRs) as InCommRs, sum(Ih.InFrt) as InFrt, sum(Ih.InFrtRs) as InFrtRs, " + _
                  "sum(Ih.InIns) as InIns, sum(Ih.InInsRs) as InInsRs, sum(Ih.InTot) as InTot, sum(Ih.InTotRs) as InTotRs " + _
                  "from InvHd Ih " + wCnd)
    '*** (Jen 2.14)
    
    '*** (Bef 2.13)
    'If wRsInHd.RecCount > 0 Then _
    'SetRepText mRep, "TxtAmtInWrds", "Total " + FnNumToWords(Format(wRsInHd!InFOBRs, "######0"), "INR")
    '*** (Bef 2.13)
    '*** (Jen 2.13)
    If wRsInHd.RecCount > 0 Then SetRepText mRep, "TxtAmtInWrds", "Total " + FnNumToWords(Format(wRsInHd!InFOBRs, "######0"), ctIndCurncy)
    '*** (Jen 2.13)
  End Select
  CRV_REP.DisplayGroupTree = False

Call DispReport(adc) '****** Sachin 3.02 - Display Report ******

'*** (Jen 2.12)
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
'*** (Jen 2.12)

End Sub

'''' std code not to be changed
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub ADC_Load()
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  
  '*** (Bef 2.14)
  'Set adc.FirNKeyCtl = adc("UoInTcFr") 'adc("UoCoCdFr")
  '*** (Bef 2.14)
  
  '*** (Jen 2.14)
  'If UCase(adc.MenuCd) = UCase("RepGRFormBackMulti") Then    '*** (Bef 3.01)
  '*** (Jen 3.01)
  If UCase(adc.MenuCd) = UCase("RepGRFormBackMulti") Or UCase(adc.MenuCd) = UCase("RepDeclnLetMulti") _
     Or UCase(adc.MenuCd) = UCase("RepGSPCertMulti") Or UCase(adc.MenuCd) = UCase("RepGSPDetLetMulti") Then
  '*** (Jen 3.01)
    Set adc.FirNKeyCtl = adc("UoInExpNoSel")
  Else
    Set adc.FirNKeyCtl = adc("UoInTcFr") 'adc("UoCoCdFr")
  End If
  '*** (Jen 2.14)
  
  Call InitProp(Me)
  
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  If UCase(adc.MenuCd) = UCase("RepJtGrBk") Then
    gs_InTcTyp = ctJtTcTyp
  ElseIf UCase(adc.MenuCd) = UCase("RepFgGrBk") Then
    gs_InTcTyp = "FG"
  Else
    gs_InTcTyp = "IN"
  End If
  
  ' Bef Trading Module
  'gs_InTcTyp = IIF(UCase(adc.MenuCd) <> UCase("RepFgGrBk"), "IN", "FG")
  '*** Jay 3.3(FG)
  
  '*** (Jen 2.14)
  'If UCase(adc.MenuCd) = UCase("RepGRFormBackMulti") Then    '*** (Bef 3.01)
  '*** (Jen 3.01)
  If UCase(adc.MenuCd) = UCase("RepGRFormBackMulti") Or UCase(adc.MenuCd) = UCase("RepDeclnLetMulti") _
     Or UCase(adc.MenuCd) = UCase("RepGSPCertMulti") Or UCase(adc.MenuCd) = UCase("RepGSPDetLetMulti") Then
  '*** (Jen 3.01)
  
    mb_MultiInv = True
    LblExpNo.Caption = "Exp No(s)": adc("UoInExpNoSel").Visible = True: adc("UoInExpNoSel").Enabled = True
    adc("UoInExpNoFr").Visible = False: adc("UoInExpNoFr").Enabled = False
    LblInvNo.Visible = False: LblTcSlash.Visible = False: LblYySlash.Visible = False: LblChrSlash.Visible = False
    adc("UoInTcFr").Visible = False: adc("UoInTcFr").Enabled = False
    adc("UoInYyFr").Visible = False: adc("UoInYyFr").Enabled = False
    adc("UoInChrFr").Visible = False: adc("UoInChrFr").Enabled = False
    adc("UoInNoFr").Visible = False: adc("UoInNoFr").Enabled = False
    '###
    'LblInPrdDesc.Visible = True: TxtInPrdDesc.Visible = True: TxtInPrdDesc.Enabled = True
    'LblPrdDescCd.Visible = True: adc("UoPMCdTo").Visible = True: adc("UoPMCdTo").Enabled = True
    '###
  Else
    mb_MultiInv = False
    LblExpNo.Caption = "Exp No": adc("UoInExpNoSel").Visible = False: adc("UoInExpNoSel").Enabled = False
    adc("UoInExpNoFr").Visible = True: adc("UoInExpNoFr").Enabled = True
    LblInvNo.Visible = True: LblTcSlash.Visible = True: LblYySlash.Visible = True: LblChrSlash.Visible = True
    adc("UoInTcFr").Visible = True: adc("UoInTcFr").Enabled = True
    adc("UoInYyFr").Visible = True: adc("UoInYyFr").Enabled = True
    adc("UoInChrFr").Visible = True: adc("UoInChrFr").Enabled = True
    adc("UoInNoFr").Visible = True: adc("UoInNoFr").Enabled = True
    '###
    'LblInPrdDesc.Visible = False: TxtInPrdDesc.Visible = False: TxtInPrdDesc.Enabled = False
    'LblPrdDescCd.Visible = False: adc("UoPMCdTo").Visible = False: adc("UoPMCdTo").Enabled = False
    '###
  End If
  '*** (Jen 2.14)
  
  Select Case UCase(adc.MenuCd)
  Case Is = UCase("RepDeclnLet")
    'Me.Caption = "Declaration Letter"
  Case Is = UCase("RepGSPCert")
    'Me.Caption = "GSP Certification"
  Case Is = UCase("RepGSPDetLet")
    'Me.Caption = "GSP Detail Letter"
  End Select
  Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  If UCase(adc.MenuCd) = UCase("RepJtGrBk") Then
    gs_InTcTyp = ctJtTcTyp
  ElseIf UCase(adc.MenuCd) = UCase("RepFgGrBk") Then
    gs_InTcTyp = "FG"
  Else
    gs_InTcTyp = "IN"
  End If
  
  ' Bef Trading Module
  '''gs_InTcTyp = IIF(UCase(adc.MenuCd) <> UCase("RepFgGrBk"), "IN", "FG")
  '*** Jay 3.3(FG)
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
  
  '*** (Jen 2.14)
  Select Case UCase(IdName)
  Case Is = UCase("UoInExpNoSel")
    If mb_MultiInv = True Then
      Dim ws_MinCurCd As String, ws_MaxCurCd As String
      Dim ws_MinCnvFct As String, ws_MaxCnvFct As String
      Dim ws_MinCnvRt As String, ws_MaxCnvRt As String

      If pv_NewValue <> "" Then
        '*** (Jen 3.1.3) added " InCoCd= '" + adc("UoCoCdFr") + "' and "
        Cancel = Not (ChkVal("Select 'x' from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")", ErrMsg))
        If Cancel = True Then ErrMsg = "Invalid Inv Exp Number(s) Scope": Exit Sub
      
        ws_MinCurCd = moCn.GetFldVal("Select min(CmCurCd) from InvHd join CustMst on CmCtg= 'C' and CmCd= InCmCd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")")
        ws_MaxCurCd = moCn.GetFldVal("Select max(CmCurCd) from InvHd join CustMst on CmCtg= 'C' and CmCd= InCmCd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")")
        If ws_MinCurCd <> ws_MaxCurCd Then Cancel = True: ErrMsg = "Invoice Export Number(s) Have Different Currency Codes": Exit Sub
        
        'min (InCnvFct)
        ws_MinCnvFct = Format(moCn.GetFldVal("Select min (RTrim (LTrim (InMulDiv))+ Str(convert (decimal (16, 4), InCnvFct))) from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")"), "####0.0000")
        ws_MaxCnvFct = Format(moCn.GetFldVal("Select max (RTrim (LTrim (InMulDiv))+ Str(convert (decimal (16, 4), InCnvFct))) from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")"), "####0.0000")
        If ws_MinCnvFct <> ws_MaxCnvFct Then Cancel = True: ErrMsg = "Invoice Export Number(s) Have Different Currency Conversion Factors": Exit Sub
      
        'min(InCnvRt)
        ws_MinCnvRt = Format(moCn.GetFldVal("Select min (RTrim (LTrim (InRsMulDiv))+ Str(convert (decimal (16, 4), InCnvRt))) from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")"), "####0.0000")
        ws_MaxCnvRt = Format(moCn.GetFldVal("Select max (RTrim (LTrim (InRsMulDiv))+ Str(convert (decimal (16, 4), InCnvRt))) from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")"), "####0.0000")
        If ws_MinCnvRt <> ws_MaxCnvRt Then Cancel = True: ErrMsg = "Invoice Export Number(s) Have Different Rupee Conversion Factors": Exit Sub
      
        Cancel = IIF(moCn.GetFldVal("Select count('x') from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")") > 9, True, False)
        If Cancel = True Then ErrMsg = "Not More Than 9 Inv Exp Number(s) Can Be Selected": Exit Sub
        '****** Manali 3.2.2 - Invoice exp No Limit extended to 9
        'Cancel = IIF(moCn.GetFldVal("Select count('x') from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")") > 4, True, False)
        'If Cancel = True Then ErrMsg = "Not More Than 4 Inv Exp Number(s) Can Be Selected": Exit Sub
        '*** (Jen 3.1.3) added " InCoCd= '" + adc("UoCoCdFr") + "' and "
      Else
        'Cancel = True: ErrMsg = "Inv Exp Number(s) Scope Cannot Be Blank": Exit Sub
      End If
      
    End If
  '###
  'Case Is = UCase("UoPMCdTo")
  '  TxtInPrdDesc.text = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INVDESC' and PMCd = '" + pv_NewValue + "'")
  '###
  End Select
  '*** (Jen 2.14)
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  ' ***** Manali Trading Module
  If UCase(adc.MenuCd) = UCase("RepJtGrBk") Then
    gs_InTcTyp = ctJtTcTyp
  ElseIf UCase(adc.MenuCd) = UCase("RepFgGrBk") Then
    gs_InTcTyp = "FG"
  Else
    gs_InTcTyp = "IN"
  End If
  
  '*** Jay 3.3(FG)
  ' Bef Trading Module
  '''gs_InTcTyp = IIF(UCase(adc.MenuCd) <> UCase("RepFgGrBk"), "IN", "FG")
  '*** Jay 3.3(FG)
  SetProp Me, IdName, When
End Sub

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    adc("UOUSRCD") = adc.UsrCd
    adc("UOMNUCD") = adc.MenuCd
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
 adc("UoCoCdFr") = gs_CoCd
End Sub
Private Sub DispCoNm()
'mRep.wCoCd.SetText gs_CoNm
    Dim ws_HName As String
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
    mRep.wCoCd.SetText ws_HName

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If adc.HotKeys(KeyCode, Shift) Then
        KeyCode = 0
        Shift = 0
    End If
End Sub
Private Sub Form_Activate()
    Call FrmActivate(Me)
End Sub
Private Sub Form_Deactivate()
    Call FrmDeActivate(Me)
End Sub
Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
    Cancel = atxt(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set mRep = Nothing
    Set Rep = Nothing
    Set Rep1 = Nothing
    Set Rep2 = Nothing
    Set Rep3 = Nothing
    '*** (09/08/05)
End Sub

