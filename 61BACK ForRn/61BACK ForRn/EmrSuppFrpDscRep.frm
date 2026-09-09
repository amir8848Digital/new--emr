VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrSuppFrpDscRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discrepancy Report"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   Tag             =   "13"
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11070
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   9540
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1005
   End
   Begin VB.PictureBox fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5430
      ScaleHeight     =   465
      ScaleWidth      =   3555
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9570
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   12
         Top             =   105
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
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         MaxLength       =   0
         DataField       =   "UOMNUCD"
         IdName          =   "UOMNUCD"
      End
   End
   Begin VB.PictureBox FRA_REPOBJ 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   8730
      ScaleHeight     =   495
      ScaleWidth      =   855
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9660
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   7
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
      TabIndex        =   3
      Top             =   9660
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   979
      AllowDelete     =   0   'False
      Automate        =   -1  'True
      EntType         =   1
      BtnPicture      =   "EmrSuppFrpDscRep.frx":0000
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.PictureBox fra_nkeys 
      BackColor       =   &H00C0C0C0&
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
      Height          =   9615
      Left            =   0
      ScaleHeight     =   9615
      ScaleWidth      =   15225
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   60
      Width           =   15225
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   5
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrSuppFrpDscRep.frx":0B1E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrSuppFrpDscRep.frx":0B3A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.PictureBox fra_tabrep 
            Enabled         =   0   'False
            Height          =   9150
            Index           =   1
            Left            =   -75000
            ScaleHeight     =   9090
            ScaleWidth      =   15120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   300
            Width           =   15180
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   4
               Top             =   360
               Width           =   9765
               _ExtentX        =   17224
               _ExtentY        =   635
            End
         End
         Begin VB.PictureBox fra_tabrep 
            Height          =   9150
            Index           =   0
            Left            =   0
            ScaleHeight     =   9090
            ScaleWidth      =   15120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   300
            Width           =   15180
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1725
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   885
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1725
               TabIndex        =   16
               ToolTipText     =   "Enter Company Code"
               Top             =   1170
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "From Date"
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
               Left            =   60
               TabIndex        =   17
               Top             =   1170
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "From"
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
               Index           =   1
               Left            =   1725
               TabIndex        =   15
               Top             =   480
               Width           =   1335
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
               Index           =   11
               Left            =   60
               TabIndex        =   14
               Top             =   885
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrSuppFrpDscRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepDscRep
Dim mRep As CRAXDRT.Report
Dim moCn As MDOConnection 'MwfLib.MDOConnection
Private Sub SetGroupSort()
'Do Not Remove
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
 adc("UoCoCdFr") = gs_CoCd
End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  '*** Report Sqls (Used in the Report designer) ***
  '*** Report Sqls (Used in the Report designer) ***
  Dim wrepcnd As String, wCnd As String, wSqlStrg As String
  Dim ws_SubRep1 As String, ws_SubRep2 As String, ws_SubRep3 As String, ws_SubRep4 As String
  Dim ws_SubRep5 As String, ws_SubRep6 As String, ws_SubRep7 As String, ws_SubRep8 As String
  Dim ws_SubRep9 As String, ws_SubRep10 As String
  Dim ws_SubRep11 As String   ' ****** Manali 3.8.0 - IrDt<>InDt Descrepancy added
  Dim ws_SubRep12 As String   ' ***** Manali 3.9.0 - FgDt>InDt Discrepancy
  Dim ws_SubRep13 As String   ' ***** Manali 3.9.0 - FgDt>InDt Discrepancy
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Set mRep = Rep
  Call DispCoNm
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  wSqlStrg = "Select 1 As Col"
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt

  '****** Sql to find the Fg Bag Records where the Gross Bag Wt does not match with the
  '    total Wt in Gms of the Raw Materials actually present in the Bag ******
  
  adc("UOCOCDFR").CmpStr = "FgCoCd="
  adc("UoDtFr").CmpStr = "FgDt>="
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  
  ws_SubRep1 = "select * from " + _
             "( " + _
             "Select max(FdBYy) as BYy, max(FdBChr) as BChr, max(FdBNo) as BNo, FdCoCd, FdTc, FdYy, FdChr, " + _
             "  FdNo, FdSr, convert(decimal (16, 3), max(FdGrWt)) as qBagWt, " + _
             "  convert(decimal (16, 3), " + _
             "    Isnull( sum((case (case when FrFrRmLoc= 'W' then FrFrRmDc when FrToRmLoc= 'W' then FrToRmDc end) " + _
             "       When 'C' then 1 when 'D' then -1 else 0 end)* FrRmWt/ " + _
             "      (case when RmCtg in ('D', 'C') then 5 else 1 end)),0) " + _
             "  ) as qRmWt,Max(FrDt) as FdFgDt " + _
             "From Fgd Left Outer Join FgRm on FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr " + _
             " and FdNo= FrNo and FdSr= FrSr " + _
             " Join Fg On FdCoCd=FgCoCd and FdTc=FgTc and FdYy=FgYy and FdChr=FgChr and FdNo=FgNo " + _
             "Left Outer Join RmMst on RmCd= FrRmCd " + _
             "Where FdTc= 'FB' " + wCnd + _
             "Group By FdCoCd, FdTc, FdYy, FdChr, FdNo, FdSr " + _
             ") Tbl " + _
             "Where Abs(convert(decimal (16, 3), qBagWt)-convert(decimal (16, 3), qRmWt))>0.01 " + _
             "Order By FdCoCd, BYy, BChr, BNo, FdTc, FdYy, FdChr, FdNo, FdSr "
  Call moCn.RepRes(mRep.Subreport1.OpenSubreport, ws_SubRep1 + ctMaxDopOpt) '****** Sachin 3.01 added ctMaxDopOpt
    
  '****** Query to find 'F' recs in Txnd ******
  adc("UOCOCDFR").CmpStr = "TdCoCd="
  adc("UoDtFr").CmpStr = "TdDt>="
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  ws_SubRep2 = "Select TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdFrRmLoc, TdToRmLoc, " + _
                 "TdCoCd from Txnd where (TdFrRmLoc= 'FG' or TdToRmLoc= 'FG')  " + wCnd
  Call moCn.RepRes(mRep.Subreport2.OpenSubreport, ws_SubRep2 + ctMaxDopOpt) '****** Sachin 3.01 added ctMaxDopOpt
    
  '****** List of all Invoices where there still exists InvFgd records where IfSr= 0 and there exist other InvDsg records ******
  adc("UOCOCDFR").CmpStr = "InCoCd="
  adc("UoDtFr").CmpStr = "InDt>="
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " And " + wrepcnd, "")
  
  ws_SubRep3 = "Select IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr, IfSrNo from InvFgd " + _
               " Join InvHd On InCoCd=IfCoCd and InTc=IfTc and InYy=IfYy and InChr=IfChr and InNo=IfNo " + _
               " where IfSr= 0 " + wCnd + _
               "and exists (Select 'x' from InvDsg where IdCoCd= IfCoCd and IdTc= IfTc and " + _
               "IdYy= IfYy and IdChr= IfChr and IdNo= IfNo)"
  Call moCn.RepRes(mRep.Subreport3.OpenSubreport, ws_SubRep3 + ctMaxDopOpt) '****** Sachin 3.01 added ctMaxDopOpt

  '****** All the Invoice Design Records where the InvRm Weight and total FgRm Weight of all the bags of the Inv Design do not match at the RmCode level ******
  adc("UOCOCDFR").CmpStr = "InCoCd="
  adc("UoDtFr").CmpStr = "InDt>="
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " And " + wrepcnd, "")
  
  ws_SubRep4 = "Select IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr, FdExpOdTc, FdExpOdYy, FdExpOdChr, FdExpOdNo, FdExpOdSr, " + _
    "  FrRmCd, FrRmWt, IsNull(IrRmWt, 0) As IrRmWt, " + _
    "  convert(decimal (14, 3), FrRmWt- IsNull(IrRmWt, 0)) as BalWt, SPACE(15) As SetSCd, Space(12) As ZRmCd " + _
    "From " + _
    "(Select IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr, max(FdExpOdTc) as FdExpOdTc, max(FdExpOdYy) as FdExpOdYy, " + _
    "max(FdExpOdChr) as FdExpOdChr, max(FdExpOdNo) as FdExpOdNo, max(FdExpOdSr) as FdExpOdSr, FrRmCd, " + _
    "  Sum(FrRmWt* (Case When LFr.LocTyp= 'F' And FrFrRmDc= 'D' Then  1 " + _
    "                   When LFr.LocTyp= 'F' And FrFrRmDc= 'C' Then -1 " + _
    "       When LTo.LocTyp= 'F' And FrToRmDc= 'D' Then  1 " + _
    "       When LTo.LocTyp= 'F' And FrToRmDc= 'C' Then -1 Else 0 End))  As FrRmWt " + _
    "From FgRm " + _
    "  Join Loc LFr On LFr.LocCoCd= FrCoCd and LFr.LocCd= FrFrRmLoc and LFr.LocTyP In ('W', 'F') " + _
    "  Join Loc LTo On LTo.LocCoCd= FrCoCd and LTo.LocCd= FrToRmLoc and LTo.LocTyP In ('W', 'F') " + _
    "  Join InvFgd On FrCoCd= IfCoCd and FrTc= IfFdTC and FrYy= IfFdYy and FrChr= IfFdChr and FrNo= IfFdNo and FrSr= IfFdSr /*And IfSr> 0*/ and (Select PValue From Param Where PTyp='TC' and PMCd=IfTc)='" + ctInTcTyp + "' " + _
    "  Join Bag On BCoCd= IfCoCd And BYy= IfFdBYy And BChr= IfFdBChr And BNo= IfFdBNo " + _
    "  Join Fgd On FdCoCd= FrCoCd And FdTc= FrTc And FdYy= FrYy And FdChr= FrChr And FdNo= FrNo And FdSr= FrSr " + _
    "Group By IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr, FrRmCd) Fr "
    '"Group By IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr, FdPrdOdTc, FdPrdOdYy, FdPrdOdChr, FdPrdOdNo, FdPrdOdSr, FrRmCd) Fr "

  ws_SubRep4 = ws_SubRep4 + " Left Outer Join " + _
    "(Select IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr, IrRmCd, " + _
    "  Sum(IrRmAWt*(Case When LFr.LocTyp= 'F' And IrFrRmDc= 'D' Then  -1 " + _
    "        When LFr.LocTyp= 'F' And IrFrRmDc= 'C' Then 1 " + _
    "        When LTo.LocTyp= 'F' And IrToRmDc= 'D' Then  -1 " + _
    "        When LTo.LocTyp= 'F' And IrToRmDc= 'C' Then 1 Else 0 End)) As IrRmWt " + _
    "From InvRm " + _
    "  Join InvDsg On IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + _
    "  Join Loc LFr On LFr.LocCoCd= IrCoCd and LFr.LocCd= IrFrRmLoc and LFr.LocTyP In ('XS', 'F') " + _
    "  Join Loc LTo On LTo.LocCoCd= IrCoCd and LTo.LocCd= IrToRmLoc and LTo.LocTyP In ('XS', 'F') " + _
    "Group By IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr, IrRmCd) Ir " + _
    "On IfCoCd= IrCoCd and IfTc= IrTc and IfYy= IrYy and IfChr= IrChr and IfNo= IrNo and IfSr= IrSr and FrRmCd= IrRmCd " + _
    " Join InvHd On InCoCd=IfCoCd and InTc=IfTc and InYy=IfYy and InChr=IfChr and InNo=IfNo " + _
    "Where Abs(convert(decimal (16, 3), FrRmWt) - convert(decimal (16, 3), IsNull(IrRmWt,0))) > 0.01 " + _
    "    And IsNull((Select ','+SdCoCd+',' From Sysdb Where SysId='EMR'),','+IfCoCd+',') Like '%,'+IfCoCd+',%' " + _
    wCnd + "Order By IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr "
  
  Call moCn.RepRes(mRep.Subreport4.OpenSubreport, ws_SubRep4 + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
  
  '*** All Custom Purchase records where Custom Wt < Match Wt ***
    adc("UOCOCDFR").CmpStr = "TzCoCd="
  adc("UoDtFr").CmpStr = "TDt>="
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  ws_SubRep5 = "Select TzCoCd, TzTc, TzYy, TzChr, TzNo, TzSr,  convert (decimal (16, 3), TzRmWt) as qBWt, " + _
        "convert(decimal(16, 3), TzRmDespWt) As TzRmDespWt " + _
        "From Txndz " + _
        " Join Txn On TzCoCd=TCoCd and TzTc=TTc and TzYy=TYy and TzChr=TChr and TzNo=TNo " + _
        "where TzTc= 'TZ' and convert (decimal (16, 3), TzRmWt)< convert (decimal (16, 3), TzRmDespWt) " + wCnd

  Call moCn.RepRes(mRep.Subreport5.OpenSubreport, ws_SubRep5 + ctMaxDopOpt) '****** Sachin 3.01 added ctMaxDopOpt
  
  '*** All Order Design Records where OrdQty <> (Entered Qty + Amendment Qty) ***
  adc("UOCOCDFR").CmpStr = "OdCoCd="
  adc("UoDtFr").CmpStr = "OdOmDt>="
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  ws_SubRep6 = "Select OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, convert (decimal (10, 1), OdOrdQty) as OdOrdQty, " + _
          "convert (decimal (10, 1), OdOrdEnt) as OdOrdEnt, " + _
          "convert (decimal (10, 1), OdPrdQty) as OdPrdQty, " + _
          "convert (decimal (10, 1), OdPrdEnt) as OdPrdEnt, " + _
          "convert (decimal (10, 1), IsNull(qOrdQty, 0)) as qOrdAmd, " + _
          "convert(decimal(10, 1), IsNull(qPrdQty, 0)) As qPrdAmd " + _
          "From OrdDsg " + _
          "Left Outer Join (Select OaCoCd, OaTc, OaYy, OaChr, OaNo, OaSr, " + _
          "     Sum(OaOrdQty) as qOrdQty, Sum(OaPrdQty) as qPrdQty From OrdAmd " + _
          "    Group By OaCoCd, OaTc, OaYy, OaChr, OaNo, OaSr) Oa " + _
          "On OdCoCd= OaCocd and OdTc= OaTc and OdYy= OaYy and OdChr= OaChr " + _
          "  and OdNo= OaNo and OdSr= OaSr " + _
          "Where   ( convert (decimal (10, 1), OdOrdQty)<> convert (decimal (10, 1), OdOrdEnt+ IsNull(qOrdQty, 0)) " + _
          "    or convert (decimal (10, 1), OdPrdQty)<> convert (decimal (10, 1), OdPrdEnt+ IsNull(qPrdQty, 0)) " + wCnd + _
          "  ) " + _
          "Order By OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr"
  Call moCn.RepRes(mRep.Subreport6.OpenSubreport, ws_SubRep6 + ctMaxDopOpt) '****** Sachin 3.01 added ctMaxDopOpt
    
  '*** All Order Design Records where Order Has Been Closed But OrdQty <> ExpQty ***
  adc("UOCOCDFR").CmpStr = "OdCoCd="
  adc("UoDtFr").CmpStr = "OdOmDt>="
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  ws_SubRep7 = "select OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OdOrdQty, OdExpQty, OdCls " + _
                "From OrdDsg " + _
                "where OdCls= 'Y' and convert (decimal (16, 1), OdOrdQty)<> convert (decimal (16, 1), OdExpQty) " + wCnd
  Call moCn.RepRes(mRep.Subreport7.OpenSubreport, ws_SubRep7 + ctMaxDopOpt) '****** Sachin 3.01 added ctMaxDopOpt
  
  '*** RmLoc= 'W'  In Loose Rm Movement ***
  adc("UOCOCDFR").CmpStr = "TdCoCd="
  adc("UoDtFr").CmpStr = "TdDt>="
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  ws_SubRep8 = "Select * from Txnd where TdSrNo= 0 and (TdFrRmLoc= 'W' Or TdToRmLoc= 'W') " + wCnd
  Call moCn.RepRes(mRep.Subreport8.OpenSubreport, ws_SubRep8 + ctMaxDopOpt) '****** Sachin 3.01 added ctMaxDopOpt
  
   
  '*** Bag Split MisMatch ***
  adc("UOCOCDFR").CmpStr = "TdCoCd="
  adc("UoDtFr").CmpStr = "TdDt>="
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  ws_SubRep9 = "Select * From (" + _
        "  Select TdCoCd, TdTc, TdYy, TdChr, TdNo, " + _
        "    convert (decimal (14, 1), sum ((case when TdBQty= 0 then TTrfQty else TdBQty end)* " + _
        "      (case when TdBDc= 'C' then -1 when TdBDc= 'D' then 1 else 0 end))) as qBQtyDiff, " + _
        "    sum (case when TdBDc= 'C' then -TdPtQty when TdBDc= 'D' then TdPtQty else 0 end) as qPtQtyDiff, " + _
        "    sum (case when TdBDc= 'D' then 1 else 0 end) as qD, " + _
        "    sum (case when TdBDc= 'C' then 1 else 0 end) as qC, " + _
        "    convert (decimal (14, 3), sum((Case when TdFrRmLoc= 'W' and TdFrRmDc= 'D' then 1 " + _
        "      when TdFrRmLoc= 'W' and TdFrRmDc= 'C' then -1 " + _
        "      when TdToRmLoc= 'W' and TdToRmDc= 'D' then 1 " + _
        "      when TdToRmLoc= 'W' and TdToRmDc= 'C' then -1 " + _
        "      Else 0 end)* TdRmWt)) as qRmWtDiff, " + _
        "    sum((Case when TdFrRmLoc= 'W' and TdFrRmDc= 'D' then 1 " + _
        "      when TdFrRmLoc= 'W' and TdFrRmDc= 'C' then -1 " + _
        "      when TdToRmLoc= 'W' and TdToRmDc= 'D' then 1 " + _
        "      when TdToRmLoc= 'W' and TdToRmDc= 'C' then -1 " + _
        "      Else 0 end)* TdRmQty) as qRmQtyDiff " + _
        "  from Txnd join Txn on TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo " + _
        "  where (TdTc= 'BS' Or TdTc= 'BM') " + wCnd + _
        "  Group By TdCoCd, TdTc, TdYy, TdChr, TdNo " + _
        ") Tbl where qBQtyDiff <> 0 or qPtQtyDiff<> 0 or qRmWtDiff<> 0 or qRmQtyDiff<> 0 or ((qD= 0 or qC= 0) and qD<> qC) " + _
        "Order By TdCoCd, TdTc, TdYy, TdChr, TdNo "
  Call moCn.RepRes(mRep.Subreport9.OpenSubreport, ws_SubRep9 + ctMaxDopOpt) '****** Sachin 3.01 added ctMaxDopOpt
  
  '*** Sachin 3.2.2 - Invoice Bags Exists But Invoice Design absent ***
    adc("UOCOCDFR").CmpStr = "InCoCd="
    adc("UoDtFr").CmpStr = "InDt>="
    wrepcnd = adc.RepCond
    wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    
  
  ws_SubRep10 = " Select InCoCd, InTc, InYy, InChr, InNo from InvHd  " + _
                " Join InvFgd On IfPrtKey=InPrtKey And IfCoCd=InCoCd And IfTc=InTc And IfYy=InYy And IfChr=InChr And IfNo=InNo " + _
                " Where Not Exists (Select 'x' From InvDsg Where IdPrtKey=IfPrtKey and IdCoCd=InCoCd and IdTc=InTc And IdYy=InYy And IdChr=InChr And IdNo=IfNo) " + wCnd + _
                " Group By InCoCd, InTc, InYy, InChr, InNo"
  Call moCn.RepRes(mRep.Subreport10.OpenSubreport, ws_SubRep10 + ctMaxDopOpt)   '****** Sachin 3.01 added ctMaxDopOpt
    
  ' ****** Manali 3.8.0 - IrDt<>InDt Descrepancy added
  ws_SubRep11 = " Select InCoCd, InTc, InYy, InChr, InNo from InvHd  " + _
                " Where Exists (Select 'x' From InvRm Where IrPrtKey=InPrtKey and IrCoCd=InCoCd and IrTc=InTc And IrYy=InYy And IrChr=InChr And IrNo=InNo and IrDt<>InDt) " + _
                " and (Select PValue From Param Where PTyp='TC' and PMCd=InTc)='IN' " + wCnd
                
  Call moCn.RepRes(mRep.Subreport11.OpenSubreport, ws_SubRep11 + ctMaxDopOpt)
  
  ' ***** Manali 3.9.0 - FgDt>InDt Discrepancy
  ws_SubRep12 = " Select InCoCd, InTc, InYy, InChr, InNo, InExpNo, InDt from InvHd  " + _
                "Where InDt < " + _
                "(Select Max(FdDt) from Fgd " + _
                " Join InvFgd on IfprtKey=FdPrtKey and IfFdIdNo=FdidNo and IfCoCd=FdCoCd and IfFdTc=FdTc and IfFdYy=FdYy and IfFdChr=FdChr and IfFdNo=FdNo and IfFdSr=FdSr " + _
                " Where IfPrtKey = InPrtKey And IfInIdNo = InIdNo And IfTc = InTc And IfYY = InYy And IfChr = InChr And IfNo = InNo)  " + wCnd
  Call moCn.RepRes(mRep.Subreport12.OpenSubreport, ws_SubRep12 + ctMaxDopOpt)
  
  '*** BagGrWt >0 without RM records
  adc("UOCOCDFR").CmpStr = "TA.TdCoCd="
  adc("UoDtFr").CmpStr = "TA.TdDt>="
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  ws_SubRep13 = " Select TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr from Txnd TA " + _
                " where (TA.TdTc ='BM' Or TA.TdTc ='BS') and TA.TdBGrWt >0 and TA.TdBDc ='C' " + wCnd + _
                " and not exists (Select * from Txnd TB where TB.TdCoCd= TA.TdCoCd and TB.TdTc= TA.TdTc and TB.TdYy= TA.TdYy " + _
                " and TB.TdChr= TA.TdChr and TB.TdNo= TA.TdNo and TB.TdSr=TA.TdSr and TB.TdSrNo <>0 ) "
  Call moCn.RepRes(mRep.Subreport13.OpenSubreport, ws_SubRep13 + ctMaxDopOpt)
    
    
  CRV_REP.DisplayGroupTree = False
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub
'''' std code not to be changed
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub ADC_Load()
    '*** Set the first Non Key Control as UoCoCdFr
    Set moCn = adc.Connection
    Set adc.FirNKeyCtl = adc("UoCoCdFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_LocTyp = "'P'"
    gs_DmTcTyp = "DM"
    gs_OmTcTyp = "SO"
    gs_CmCtg = "C"
    gs_PTyp = "PCT"
    gs_Tbl = "Param"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
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
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UOCOCDFR") + "' and " + _
                                                      "HCd = '" + ctSelfCoCd + "' ")
  Rep.wCoCd.SetText ws_HName
  Dim ws_DtTm As String
  ws_DtTm = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
  Rep.wSrvrDate.SetText ws_DtTm
  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  Rep.wSrvrTime.SetText ws_DtTm
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
    Set moCn = Nothing
    Set mRep = Nothing
    Set Rep = Nothing
    'Set Rep2 = Nothing
    'Set Rep3 = Nothing
    '*** (09/08/05)
End Sub

