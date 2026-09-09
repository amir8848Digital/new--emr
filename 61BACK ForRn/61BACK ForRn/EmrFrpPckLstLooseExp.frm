VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpPckLstLooseExp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice BOE Listing"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   12960
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4410
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   7065
      TabIndex        =   16
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   17
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
         TabIndex        =   18
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
      Left            =   8280
      TabIndex        =   15
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   12
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
      TabIndex        =   6
      Top             =   9660
      WhatsThisHelpID =   15120
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
      Height          =   9795
      Left            =   0
      TabIndex        =   11
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   10
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpPckLstLooseExp.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpPckLstLooseExp.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   14
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   600
               TabIndex        =   7
               Top             =   480
               Width           =   8355
               _ExtentX        =   14737
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   13
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2700
               TabIndex        =   2
               ToolTipText     =   "Enter Invoice Year"
               Top             =   1080
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "InYy ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1920
               TabIndex        =   1
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   1080
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "InTc ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3120
               TabIndex        =   3
               ToolTipText     =   "Enter Invoice Character"
               Top             =   1080
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "InChr ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   3900
               TabIndex        =   4
               ToolTipText     =   "Enter Invoice Number"
               Top             =   1080
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "InNo ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1920
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   795
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "InCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1920
               TabIndex        =   5
               ToolTipText     =   "Enter From Export Number"
               Top             =   1365
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOFR"
               IdName          =   "UOINEXPNOFR"
               CmpStr          =   "InExpNo ="
            End
            Begin VB.Label ALBL 
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
               Index           =   5
               Left            =   120
               TabIndex        =   24
               Top             =   1365
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
               Left            =   120
               TabIndex        =   23
               Top             =   795
               Width           =   1335
            End
            Begin VB.Label ALBL 
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
               Index           =   27
               Left            =   2640
               TabIndex        =   22
               Top             =   1080
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Index           =   25
               Left            =   3060
               TabIndex        =   21
               Top             =   1080
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Index           =   9
               Left            =   120
               TabIndex        =   20
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label ALBL 
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
               Index           =   0
               Left            =   3825
               TabIndex        =   19
               Top             =   1080
               Width           =   105
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpPckLstLooseExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepPckLstLooseExp
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  GRP_REP.Add "VoucherNo", "B.TzTc+B.TzYy+B.TzChr+str(B.TzNo)", "B.TzTc+'/'+B.TzYy+'/'+B.TzChr+'/'+LTrim(Str(B.TzNo))", "", "", "", "", "", ""
  '6.1 Boeno ,Boe date
  GRP_REP.Add "BOENo", "TBOENo", "", "TBOENo", "hTBOENo", "BOE Date", "", "", "(case when TBOEDt= '01/01/80' then ' ' else convert(varchar , TBOEDt, " + CStr(DtFmtSlahYyyyVal()) + ") end)"
  GRP_REP.Add "BOE Date", "convert(char(8),TBOEDt,112)", "(case when TBOEDt= '01/01/80' then '' else convert(varchar, TBOEDt, " + CStr(DtFmtSlahYyyyVal()) + ") end)", "TBOEDt", "hTBOEDt", "", "", "", ""
  GRP_REP.Add "Rm Code", "RmCtg+RmSCtg+B.TzRmZCd", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ B.TzRmZCd", "qBRmCd", "hqBRmCd", "", "", "", "RmDesc"
  GRP_REP.Add "Imp Type", "TImpTyp", "TImpTyp", "", "", "", "", "", "(Select PDesc From Param Where Ptyp ='IMPTYP' and  PmCd= TImpTyp) "
  GRP_REP.Add "Rm Desc", "B.TzRmCdDesc", "B.TzRmCdDesc", "QbRmCd", "hTzRmZCd", "", "", "", "(Select RmDesc From RmMst Where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= B.TzRmCdDesc) "

'  GRP_REP.Add "Supplier", "TSuppCd", "", "", "", "", "", "", "(Select CmName From CustMst Where CmCtg= 'S' and CmCd= TSuppCd) "
'  GRP_REP.Add "Supp Bill No", "TBillNo", "", "", "", "Supplier", "", "", "(case when TBillDt='01/01/80' then '' else convert(varchar, TBillDt, 103) end)"
'  GRP_REP.Add "Vch Date", "convert(char(8),TDt,112)", "convert(varchar, TDt, 103)", "", "", "", "", "", ""
'  GRP_REP.Add "Purchase Rt", "str(B.TzPurRt,12, 3)", "", "TzPurRt", "hTzPurRt", "", "", "", ""
'  GRP_REP.Add "(RmCd) PurRt", "B.TzRmZCd+str(B.TzPurRt,9,2)", "'('+B.TzRmZCd+') '+LTrim(str(B.TzPurRt,9,2))", "TzRmZCd,TzPurRt", "hTzRmZCd,hTzPurRt", "Rm Code,Purchase Rt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  '  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  '  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  '  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  '  space(2) as qTag, '' as qInvNo, InDt, InExpNo, InGrNo, InGrDt,
  '  InPrdDesc, InCnvRt, space(150) as qAddr, InFOB, InComm, InFrt, InIns, InDiaCost,
  '  InCSCost, InAdv, InTot, InTotAs, TzTc as qBTc, TzYy as qBYy, TzChr as qBChr,
  '  TzNo as qBNo, TzSr as qBSr, TzRmZCd as qBRmCd, TzRmCdDesc as qBRmCdDesc,
  '  TzPcPerCt as qBPcPerCt, TzPurRt as qBPurRt, TzRmQty as qMRmQty, TzRmWt as qMRmWtInCts,
  '  TzRmWt as qMRmWtInGms, TzPurAmt as qMPurAmt, TDt, TSuppCd, TBillNo, TImpTyp, TBillDt,
  '  TBOENo, TBOEDt, TCnvRt from InvHd, TxndZ, Txn where 1 = 2
  '*** Report Sql ***
  
  Dim wrepcnd As String, wCnd As String, grpflds As String, wSqlStrg As String
  'Dim i As Integer, wDetPos As Integer
  
  '*** (Jen 2.13)
  Dim ws_CurCd As String, wRsInHd As MDORowSet
  '*** (Jen 2.13)
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
 
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
  
  wrepcnd = adc.RepCond
  
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")

  '****** (Bef 2.13) InCnvRt, ******
  '6.1
  wSqlStrg = wSqlStrg + "Set Dateformat " + DtFmtStr() + " Select '#XXX#' as Grp1, '#XXX#' as Grp2, '#XXX#' as Grp3, " + _
             "' ' as DGrp1, ' ' as DGrp2, ' ' as DGrp3, ' ' as G1Desc, ' ' as G2Desc, " + _
             "' ' as G3Desc, '1H' as qTag, InTc+ '/ ' + InYy+ '/ ' + InChr+ '/ ' + LTrim(Str(InNo)) as qInvNo, " + _
             "InDt, InExpNo, InGrNo, InGrDt, InPrdDesc, " + _
             "(case when convert (decimal (16, 4), InCnvRt)> 0 and InRsMulDiv= 'D' then (1/ InCnvRt) else InCnvRt end) InCnvRt, " + _
             "RTrim(LTrim(InSelfAdd1))+ ' ' + RTrim(LTrim(InSelfAdd2))+ ' ' + " + _
             "RTrim(LTrim(InSelfAdd3)) as qAddr, InFOB, InComm, InFrt, InIns, InDiaCost, " + _
             "InCSCost, InAdv, InTot, InTotAsTxt as InTotAs, InSelfName, InhAuthSgn, " + _
             "'' as qBTc, '' as qBYy, '' as qBChr, 0 as qBNo, 0 as qBSr, " + _
             "'' as qBRmCd, '' as qBRmCdDesc, 0 as qBPcPerCt, 0 as qBPurRt, 0 as qMRmQty, " + _
             "0.0 as qMRmWtInCts, 0.0 as qMRmWtInGms, 0.0 as qMPurAmt, '01/01/80' as TDt, " + _
             "'' as TSuppCd, '' as TBillNo, '' as TImpTyp, '' as TBillDt, '' as TBOENo, " + _
             "'01/01/80' as TBOEDt, 0.0 as TCnvRt from InvHd " + wCnd
  
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  '****** (Bef 2.13) B.TzPurRt as qBPurRt, ******
  '****** Sachin 3.02 - Id fields in Join
  wSqlStrg = wSqlStrg + " Union All Select " + grpflds + ", '2D' as qTag, " + _
             "'' as qInvNo, '01/01/80' as InDt, '' as InExpNo, '' as InGrNo, " + _
             "'01/01/80' as InGrDt, '' as InPrdDesc, 0.0 as InCnvRt, " + _
             "'' as qAddr, 0.0 as InFOB, 0.0 as InComm, 0.0 as InFrt, 0.0 as InIns, " + _
             "0.0 as InDiaCost, 0.0 as InCSCost, 0.0 as InAdv, 0.0 as InTot, '' as InTotAs, " + _
             "InSelfName, InhAuthSgn, " + _
             "B.TzTc as qBTc, B.TzYy as qBYy, B.TzChr as qBChr, B.TzNo as qBNo, " + _
             "B.TzSr as qBSr, B.TzRmZCd as qBRmCd, B.TzRmCdDesc as qBRmCdDesc, " + _
             "B.TzPcPerCt as qBPcPerCt, M.TzPurRt as qBPurRt, M.TzRmQty as qMRmQty, " + _
             "(case when RmCtg in ('C', 'D') then M.TzRmWt else 0 end) as qMRmWtInCts, " + _
             "(case when RmCtg not in ('C', 'D') then M.TzRmWt else 0 end) as qMRmWtInGms, M.TzPurAmt as qMPurAmt, " + _
             "TDt, TSuppCd, TBillNo, TImpTyp, TBillDt, TBOENo, TBOEDt, TCnvRt from TxndZ M " + _
             "join InvHd on InIdNo=M.TzInIdNo and InCoCd= M.TzCoCd and InTc= M.TzTc and InYy= M.TzYy and InChr= M.TzChr " + _
             "and InNo= M.TzNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=M.TzPrtKey ", "") + _
             "join TxndZ B on B.TzCoCd= M.TzCoCd and B.TzYy= M.TzRefYy and B.TzKey = M.TzRefKey " + IIF(gs_Partition = ctCurrPrtn, " and B.TzPrtKey=M.TzPrtKey ", "") + _
             "join Txn on TIdNo=B.TzTIdNo and TCoCd= B.TzCoCd and TTc= B.TzTc and TYy= B.TzYy " + _
             "and TChr= B.TzChr and TNo= B.TzNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=B.TzPrtKey ", "") + _
             "join RmMst on RmCd= B.TzRmZCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=B.TzPrtKey ", "") + _
             "where M.TzBM= 'M' " + wCnd + IIF(gs_Partition = ctCurrPrtn, " and M.TzPrtKey='" + ctCurrPrtn + "' ", "") + _
             " Order By TBOEDt, TBOENo, qBRmCdDesc "
             
             '*** Note thet M.TzRmQty is always zero as it is not an entered/ updated field

  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
    
  '*** (Jen 2.13)
  Set wRsInHd = moCn.OpenRes("Select InTc, InYy, InChr, InNo, CmCurCd " + _
                "from InvHd join Param on PTyp= 'TC' and PMCd= InTc and PSCd= '' " + _
                "join CustMst on CmCtg= (case PValue when 'IN' then 'C' when 'CTB' then 'T' else '' end) and CmCd= InCmCd " + _
                "where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd ='" + adc("UoCoCdFr") + "' " + _
                "And ((InTc= '" + adc("UoInTcFr") + "' and " + _
                "   InYy= '" + adc("UoInYyFr") + "' and InChr= '" + adc("UoInChrFr") + "' and " + _
                "   InNo= " + CStr(adc("UoInNoFr")) + " and InNo> 0) " + _
                "Or (InExpNo= '" + adc("UoInExpNoFr") + "' and InExpNo<> '')) ")
  If wRsInHd.RecCount > 0 Then
    ws_CurCd = wRsInHd!CmCurCd
  End If
  
  With Rep
    .hBPurRt.SetText "Cost Rt (" + ws_CurCd + ")"
    .hqMPurAmt.SetText "Amount (" + ws_CurCd + ")"
    .h1qMPurAmt.SetText "Amount (" + ws_CurCd + ")"
  End With
  '*** (Jen 2.13)
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.qBRmCd}", 8)
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.qBRmCdDesc}", 8)
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
End Sub

'''' std code not to be changed
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub ADC_Load()
    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    Set adc.FirNKeyCtl = adc("UoInTcFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_InTcTyp = "IN"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_InTcTyp = "IN"
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_InTcTyp = "IN"
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
  'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                            "HCd = '" + ctSelfCmCd + "' ")
  Rep.wCoCd.SetText ws_HName
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  Rep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    Rep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
' ********* Manali 3.2.2 - Before 3.2.2
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  Rep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  Rep.wSrvrTime.SetText ws_DtTm

' ******** Manali 3.2.2 - User date and time in reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  Rep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  
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
    
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set Rep = Nothing
End Sub

