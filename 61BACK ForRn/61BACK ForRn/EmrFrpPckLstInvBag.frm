VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpPckLstInvBag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Packing List"
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
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11700
      TabIndex        =   14
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
      TabIndex        =   13
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
      Left            =   5865
      TabIndex        =   23
      Top             =   9570
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   24
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
         TabIndex        =   25
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
      Left            =   7080
      TabIndex        =   22
      Top             =   9585
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   19
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
      TabIndex        =   15
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
      TabIndex        =   18
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   17
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpPckLstInvBag.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpPckLstInvBag.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   21
            Top             =   360
            Width           =   15045
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   16
               Top             =   360
               Width           =   8745
               _ExtentX        =   15425
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   8900
            Index           =   0
            Left            =   90
            TabIndex        =   20
            Top             =   390
            Width           =   15045
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Sort by Order Cust"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   6480
               TabIndex        =   7
               Tag             =   "SortbyOrdCust"
               ToolTipText     =   "Check to Sort by Order Customer"
               Top             =   1680
               Width           =   2400
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Bag Details                     "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   9
               Left            =   120
               TabIndex        =   36
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   2760
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Cust PO, DsgCd               "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   120
               TabIndex        =   12
               Tag             =   "ShowCustPoDsg"
               ToolTipText     =   "Check To Show Customer PO Number, Customer Design Code "
               Top             =   2475
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Picture                                      "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   120
               TabIndex        =   10
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   2190
               Width           =   3240
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   5145
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
               Index           =   20
               Left            =   3945
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
               Left            =   3165
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
               Left            =   4365
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
               Index           =   2
               Left            =   3585
               TabIndex        =   9
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   2190
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   3165
               TabIndex        =   6
               ToolTipText     =   "Enter Value On Which Invoice Report Grouping Has To Be Done"
               Top             =   1575
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINGRP"
               IdName          =   "UOINGRP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   3165
               TabIndex        =   8
               ToolTipText     =   "Enter Detail Sorting To Be Shown In Report"
               Top             =   1860
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINDET"
               IdName          =   "UOINDET"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   3165
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   720
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
               Index           =   9
               Left            =   3165
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
               Index           =   3
               Left            =   3585
               TabIndex        =   11
               ToolTipText     =   "Show Customer PO Number, Customer Design Code (Yes / No)"
               Top             =   2475
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   3585
               TabIndex        =   37
               ToolTipText     =   "Show Bag(Yes / No)"
               Top             =   2760
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN9"
               IdName          =   "UOYN9"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   9120
               TabIndex        =   38
               ToolTipText     =   "Sort by Order Customer"
               Top             =   1680
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
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
               Left            =   3885
               TabIndex        =   35
               Top             =   1005
               Width           =   105
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
               Left            =   4305
               TabIndex        =   34
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
               TabIndex        =   33
               Top             =   1005
               Width           =   1845
            End
            Begin VB.Label LblPic 
               BackStyle       =   0  'Transparent
               Caption         =   "Picture"
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   165
               TabIndex        =   32
               Top             =   2430
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblGrpCombo 
               BackStyle       =   0  'Transparent
               Caption         =   "Grp Combination"
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
               TabIndex        =   31
               Top             =   1575
               Width           =   1845
            End
            Begin VB.Label LblDet 
               BackStyle       =   0  'Transparent
               Caption         =   "Detail Level"
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
               TabIndex        =   30
               Top             =   1860
               Width           =   1845
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
               TabIndex        =   29
               Top             =   720
               Width           =   1845
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
               TabIndex        =   28
               Top             =   1290
               Width           =   1845
            End
            Begin VB.Label lblRem 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Cust PO, DsgCd"
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   165
               TabIndex        =   27
               Top             =   2475
               Visible         =   0   'False
               Width           =   2355
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
               Left            =   5070
               TabIndex        =   26
               Top             =   1005
               Width           =   105
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpPckLstInvBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepPckLstInvBag

Dim wFlagLocalExp As String
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Dim mb_MultiInv As Boolean    '*** (Jen 2.13)

Private Sub SetGroupSort()
  ''GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub
  
Private Sub adc_setreprecsource()
'*** (Jen 2.12)
On Error GoTo RepErr
'*** (Jen 2.12)

  '    Select space(2) as qTag, IdTc, IdYy, IdChr, IdNo, IdSr, InExpNo, InCmCd, InDt, InLmgCst, InLmgDt, InLmgSal,
  '    InVaDC, InVaX, InCnvRt, InPrdDesc, InPkgWt, InGrNo, InGrDt, InSbNo, InSbDt, InAwbNo,
  '    InAwbDt, InHAwbNo, InHAwbDt, InHBinNo, InGspNo, InGspDt, InChqNo, InChqAmt, InPoNo,
  '    InPoDt, InCmPayTerms, InFOB,
  '    InFOBRs, InComm, InCommRs, InFrt, InFrtRs, InIns, InInsRs, InTot, InTotRs, InTotAs, InTot, InTotRs, InTotAs,
  '    InDiaCost, InCSCost, InAdv, InDiaCostDesc, InCSCostDesc, InAdvDesc, InCls, InHBkName,
  '    InHBkAdd1, InHBkAdd2, InHBkAdd3, InHBkAc, InHFrgnBkName, InHFrgnBkAc, InHExpOff,
  '    InHRbiNo, InHIecNo, InHPreCarriage, InHPreCarrier, InHVessel, InHPortOfLd, InHConOfOrg,
  '    InHMarks, InHPkgDesc, InHItcNo, InHEtcZone, InHEtcSection, InHConsignee, InCmInsBy as InHInsBy,
  '    InHInsNo, InHAgName, InHAgAdd1, InHAgAdd2, InHAgAdd3, InCmName, InCmAdd1, InCmAdd2,
  '    InCmAdd3, InCmTel, InCmFax, InCmEMail, InCmCpName, InCmPortOfDisch, InCmFinDest, InCmBuyer,
  '    InCmAgName, InCmAgAdd1, InCmAgAdd2, InCmAgAdd3,
  '    InSelfName, InSelfAdd1, InSelfAdd2, InSelfAdd3, InSelfTel, InSelfFax,
  '    InSelfEMail, InHLicNo, InHETCSvg, InHAR4No, InHQCertNo, InHGSPOrgCrit, InHStatCd,
  '    InHFtRem, space(700) as InHBnkInstr, InConAdd1, InConAdd2, InConAdd3, InBuyAdd1, InBuyAdd2,
  '    InBuyAdd3, space(750) as InHUnResolutions,InHCustRem,space(180) as InHEximPolicy, InHGspEligibility, InMetalUtil, InhAuthSgn,
  '    InHETCTxt,InHStatOfOrg,InHImpValTxt,InHShpBillSgnBy,InHMetalBOETxt, InHGspPkgDesc, InBOELooseExp, InHHideCommYN, InCircularNo,
  '    IdKt, DmCtg, DmSalCtg, IdDmCd, IdSfx, IdDmSz, DmUom, OmPONo, OdPoNo, OdPicNm, CdCustDmCd, OdDmCol,
  '    IdVaCtg, IdGldLs, IdQty, IrRmIWt as qInvGrWt, IrRmIWt as qInvGldWt,
  '    IrRmIWt as qInvPlWt, IrRmIWt as qInvSlWt, IrRmIWt as qInvPdWt, IrRmIWt as qInvAccWt, IrRmIWt as qInvDiaWt,
  '    IrRmIWt as qInvPureGldWt, IrRmIWt as qInvPurePlWt, IrRmIWt as qInvPureSlWt,  IrRmIWt as qInvPurePdWt,IrRmIWt as qInvPureGldLs,
  '    IrRmIWt as qInvPurePlLs, IrRmIWt as qInvPureSlLs,  IrRmIWt as qInvPurePdLs, IrRmIVal as qIGldVal, IrRmIVal as qIPlVal,
  '    IrRmIVal as qISlVal, IrRmIVal as qIPdVal, IrRmIVal as qIDCVal, IrSetIVal as qISetVal, iLIVal as qILabVal, IdiFOB,
  '    IrRmZWt as qZGrWt, IrRmZWt as qZGldWt, IrRmZWt as qZPlWt, IrRmZWt as qZSlWt, IrRmZWt as qZPdWt, IrRmZWt as qZAccWt,
  '    IrRmZWt As qZDiaWt, IrRmZWt as qZPureGldWt, IrRmZWt as qZPurePlWt, IrRmZWt as qZPureSlWt, IrRmZWt as qZPurePdWt,
  '    IrRmZWt as qZPureGldLs, IrRmZWt as qZPurePlLs, IrRmZWt as qZPureSlLs, IrRmZWt as qZPurepdLs, IrRmZVal as qZGldVal,
  '    IrRmZVal as qZPlVal, IrRmZVal as qZSlVal, IrRmZVal as qZPdVal, IrRmZVal as qZDCVal, IrSetZVal as qZSetVal, iLZVal as qZLabVal,
  '    IdZFOB, OrRmCtg, space(254) as qMainRmCtgDesc, space(254) as qDmCtgDesc, space(254) as qDmSalCtgDesc,
  '    space(254) as qDmColDesc, space(254) as qVaCtgDesc, InMMTC,InHGldFiness,InHPlFiness,InHSlFiness, InHPdFiness,
  '    InHGjepcNo,InHGjepcValidUpTo,InHRbiPolicy,Space(11) as RmCtg, RmKt, IrRmIRt as qRmRt,
  '    IrRmQty, IrRmIWt as qWtInCts, IrRmIWt as qWtInGms, IrRmIVal as qVal
  '    From InvRm, InvLab, RmMst, InvDsg, DsgMst, OrdMst, OrdDsg, OrdRm, InvHd, CustDsg
  '    Where 1 = 2
  '*** Report Sql ***
  
  '*** Loss Summary Sub Rep Query (Invoice Printing) '*** (Jen 2.13)
  'Select IdTc, IdYy, IdChr, IdNo, IdGldLs, IdKt, IrRmCd, RmCtg, RmKt, PDesc as RmCtgDesc,
  'IrRmIWt as qMetWt, IrRmIWt as qMetPureWt, IrRmIWt as qMetFineWt,
  'IrRmIWt As qInvLsWt, IrRmIWt As qInvPureLsWt, IrRmIWt as qInvFineLsWt,
  'IrRmIWt As qInvTotWt, IrRmIWt As qInvPureTotWt, IrRmIWt As qInvFineTotWt
  'from InvRm, RmMst, InvDsg, InvHd, Param where 1= 2
  '*** Loss Summary Sub Rep Query (Invoice Printing) '*** (Jen 2.13)
  
  
  '**** Invoice DsgDesc SubReport **For Local Invoice   ***Geeta****Emr208******
  ' Zubin 212 (Sales Ctg and qDmSalCtgDesc added)
  '   Select Space(2) as qTag, DmCtg,DmSalCtg,PDesc as DmCtgDesc,PDesc as DmSalCtgDesc,DmUom,IdKt,
  '   IdVaCtg,IdGldLs,OdDmCol,OrRmCtg,IdQty as qIdQty From InvDsg, DsgMst,OrdDsg,OrdRm,Param
  '   Where 1 = 2
  '********************************
    
    
    
  Dim wRmWtOpt As String, wrepcnd As String, wCnd As String, wAndCnd As String, wHdStr1 As String, wStr1 As String, wStr2 As String, wStr3 As String, wSqlStrg As String, wSubRepStr As String
  Dim wRsMetUtlsd As MDORowSet, wMetUtlsdStr As String
  Dim wShowWtFld As String, wShowMetLsSumm As String, wShowCnvRt As String, wShowRtFld As String
  Dim wSubRepStr1 As String, wRmValOpt As String
  Dim wILabValStr As String, wZLabValStr As String, wISetValStr As String, wZSetValStr As String
  Dim wRsFiness As MDORowSet, ws_FinessStr As String, RsCnt As Integer
  Dim wd_GldOunce As Double, wd_PlOunce As Double, wd_LmgGms As Double, ws_GldStr As String
  Dim ws_PlStr As String, wd_LmpGms As Double
  Dim wb_MLsRmCd As Boolean     '****** Sachin 2.14.0 Next ******
  
  ' ****** Sachin 2.12 - 06/12/05 - SJM
  Dim wd_SlOunce As Double, wd_LmsGms As Double, ws_SlvStr As String
  ' ****** Sachin 2.12 - 06/12/05 - SJM
  
  ' **** Manali 3.03 - 25/06/08 - Pd Mod
  Dim wd_PdOunce As Double, wd_LmlGms As Double, ws_PdStr As String, ws_PdOunceDsc As String
  ' **** Manali 3.03 - 25/06/08 - Pd Mod
     
  '*** (Jen 2.12)
  Dim wRmRtOpt As String
  Dim ws_RmSubRep As String, ws_RmGrpOpt As String
  Dim ws_RmCtg As String, ws_RmSCtg As String, ws_RmCd As String, ws_RmRt As String
  Dim ws_ShowRmSubRepCnd As String
  '*** (Jen 2.12)
  
  '*** (Jen 2.13)
  Dim ws_CurCd As String
  Dim ws_MinCnd As String, ws_MinWhereCnd As String, ws_MinAndCnd As String ',ws_MultiWhereCnd As String, ws_MultiAndCnd As String
  Dim ws_InExpNo As String, ws_InTotAs As String, ws_PrdDesc As String, ws_MinExpNo As String
  Dim ws_GldOunceDsc As String, ws_PlOunceDsc As String, ws_SlOunceDsc As String
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, " ")
  wAndCnd = IIF(wrepcnd <> "", " and " + wrepcnd, " ")
  
  ws_InExpNo = " max(Ih.InExpNo) as InExpNo "
  ws_InTotAs = " max(Ih.InTotAsTxt) as InTotAs "        '3.11.0 - Display field changed
  ws_PrdDesc = " max(Ih.InPrdDesc) as InPrdDesc "

  
                                                                 
                                                                    
     wILabValStr = " 0 "
    wZLabValStr = " 0 "
  
  Set Rep = Nothing

      
      Set adc.RepSource = Rep
      Set mRep = Rep
      
    
    
  
  Call DispCoNm
    
  '*** For the Detail & Group Formula of the report
  Dim wDetArr() As String, wGrpArr() As String, wDetFormula As String, wDetDescFormula As String, wDetDesc2Str As String, wDetDesc2Formula As String, wGrpFormula As String, wGrpDescFormula As String
  Dim wDmUomFlag As Boolean, qDmCtgDesc As String, qMainRmCtgDesc As String, qDmColDesc As String, qDmKtDesc As String, qVaCtgDesc As String, qMainMet As String, wOrJoin As String, i As Integer
  'Dim wGldLsFlag As Boolean  '*** (Bef 2.13) Jen
  Dim wSubRepGrpFormula As String
  Dim wStr4 As String, wStr5 As String, wStr6 As String
  
  Dim qDmSalCtgDesc As String
  
  Dim wGrpSubRep() As String, ws_IdDmCd As String, ws_IdSfx As String, ws_IdDmSz As String
  Dim ws_IdVaCtg As String, Ws_IdKt As String, ws_IdQty As String
  
  qDmCtgDesc = "'' as qDmCtgDesc": qMainRmCtgDesc = "'' as qMainRmCtgDesc"
  qDmColDesc = "'' as qDmColDesc": qVaCtgDesc = "'' as qVaCtgDesc": qDmKtDesc = "'' as qDmKtDesc"
  qMainMet = "'' as OrRmCtg": wOrJoin = ""
  
  qDmSalCtgDesc = "'' as qDmSalCtgDesc"
  
  wDmUomFlag = False
    
    ws_IdDmCd = "IdDmCd":  ws_IdSfx = "IdSfx":  ws_IdDmSz = "IdDmSz"
    ws_IdVaCtg = "IdVaCtg":  Ws_IdKt = "IdKt":  ws_IdQty = "IdQty"
'  Else
'    ws_IdDmCd = "OdDmCd":  ws_IdSfx = "OdSfx":  ws_IdDmSz = "OdDmSz"
'    ws_IdVaCtg = "OdVaCtg":  Ws_IdKt = "OdKt":  ws_IdQty = "OdOrdQty"
'  End If
  
  wDetArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INDET' and " + _
            "PMCd= '" + adc("UoInDet") + "' "), ",")
  For i = 0 To UBound(wDetArr)
    wDetFormula = wDetFormula + IIF(wDetFormula = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdNo}, '######0')+ '/ ' + ToText({rdo.IdSr}, '###0')", IIF(UCase(wDetArr(i)) = "IDTRAYNO", "ToText({rdo.IdTrayNo}, '######0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))"))
    ' bef  3.8.0
    'wDetFormula = wDetFormula + IIF(wDetFormula = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdNo}, '######0')+ '/ ' + ToText({rdo.IdSr}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))")
  Next i
  wDetFormula = IIF(wDetFormula = "", "", "{rdo.qTag}+' '+" + wDetFormula) '*** (Jen 2.12)
  
  wDetDesc2Str = " '' as OmPONo, '' as OdPoNo, '' as OdPicNm, '' as CdCustDmCd, "    '*** Jay 3.2.0 [OdPoNo] '*** Jay 3.2.0 [OdPicNm]
  
  If UCase(adc("UoInDet")) = "DMCTG" Then
    qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc "
    wDetDescFormula = "Trim({rdo.qDmCtgDesc})"
    wDetDesc2Formula = ""
  ElseIf UCase(adc("UoInDet")) = "PCS/PRS" Then
    wDetDescFormula = ""
    wDetDesc2Str = " '' as OmPONo, '' as OdPoNo, '' as OdPicNm, '' as CdCustDmCd, "    '*** Jay 3.2.0 [OdPoNo] '*** Jay 3.2.0 [OdPicNm]
    wDetDesc2Formula = ""
  ElseIf UCase(adc("UoInDet")) = "DMCD" Then
    wDetDescFormula = "{rdo.IdDMCd}"
    wDetDesc2Formula = ""
  ElseIf UCase(adc("UoInDet")) = "INVSR" Then
    wDetDescFormula = "Trim(ToText({rdo.IdSr}, '###0'))+'    '+ Trim({rdo.IdDmCd} )+ " + _
                      "(If {rdo.IdDmSz}= '' then '' else ' /   ') +Trim({rdo.IdDmSz})"
    wDetDesc2Str = " max(OmPONo) as OmPONo, max(OdPoNo) as OdPoNo, Max(OdPicNm) as OdPicNm, IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= ''), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= ''), '')))) as CdCustDmCd, "
    wDetDesc2Formula = "{rdo.OmPoNo}+" + "(If {rdo.OdPoNo}= '' then '' else ' - ') +Trim({rdo.OdPoNo})+" + "(If {rdo.CdCustDmCd}= '' then '' else ' /   ') +Trim({rdo.CdCustDmCd})"    '*** Jay 3.2.0 [OdPoNo]
  ElseIf UCase(adc("UoInDet")) = "DMCD/SR" Then
    wDetDescFormula = "Trim(ToText({@wRunSrNo}, '###0'))+'    '+Trim({rdo.IdDmCd})+ " + _
                      "(If {rdo.IdDmSz}= '' then '' else ' /   ') +Trim({rdo.IdDmSz})"
    wDetDesc2Str = " max(OmPONo) as OmPONo, max(OdPoNo) as OdPoNo, Max(OdPicNm) as OdPicNm, IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= ''), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= ''), '')))) as CdCustDmCd, "
    wDetDesc2Formula = "{rdo.OmPoNo}+" + "(If {rdo.OdPoNo}= '' then '' else ' - ') +Trim({rdo.OdPoNo})+" + "(If {rdo.CdCustDmCd}= '' then '' else ' /   ') +Trim({rdo.CdCustDmCd})"    '*** Jay 3.2.0 [OdPoNo]
  ElseIf UCase(adc("UoInDet")) = "INVSR/SR" Or UCase(adc("UoInDet")) = "TRNO/SR" Then
    wDetDescFormula = "Trim(ToText({@wRunSrNo}, '###0'))+'    '+ " + "Trim({rdo.IdDmCd} )+ " + _
                      "(If {rdo.IdDmSz}= '' then '' else ' /   ') +Trim({rdo.IdDmSz})"
    wDetDesc2Str = " max(OmPONo) as OmPONo, max(OdPoNo) as OdPoNo,Max(OdPicNm) as OdPicNm,  IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= ''), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= ''), '')))) as CdCustDmCd, "
    wDetDesc2Formula = "{rdo.OmPoNo}+" + "(If {rdo.OdPoNo}= '' then '' else ' - ') +Trim({rdo.OdPoNo})+" + "(If {rdo.CdCustDmCd}= '' then '' else ' /   ') +Trim({rdo.CdCustDmCd})"    '*** Jay 3.2.0 [OdPoNo]
  End If
  
  '4.1.2.0
  If adc("UOYN2") = "Y" Then
    wGrpFormula = "Trim(UpperCase(ToText({rdo.OmCmCd})))"
    wSubRepGrpFormula = "OdOmCmCd"
  End If
  
  wGrpArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
         "PMCd= '" + adc("UoInGrp") + "' "), ",")
  For i = 0 To UBound(wGrpArr)
    If UCase(wGrpArr(i)) = "RMCTG" Then
      qMainRmCtgDesc = "(Select PDesc from Param where PTyp= 'RMCTG' and PMCd= max(OrRmCtg)) as qMainRmCtgDesc"
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.qMainRmCtgDesc})"
      qMainMet = "max(OrRmCtg) as OrRmCtg "
          wOrJoin = " join OrdRm on " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OdPrtKey and ", "") + " OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy " + _
                        " and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr and OrMainMet= 'Y' "
    ElseIf UCase(wGrpArr(i)) = "KT" Then
      qDmKtDesc = "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(IdKt)) as qDmKtDesc"
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.qDmKtDesc})"
    ElseIf UCase(wGrpArr(i)) = "DMCOL" Then
      qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
    ElseIf UCase(wGrpArr(i)) = "VACTG" Then
        qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(" + ws_IdVaCtg + ")) as qVaCtgDesc"
        wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
    ElseIf UCase(wGrpArr(i)) = "GLDLS" Then
        wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "'@'+ToText({rdo.IdGldLs}, '##0.000')+ '%'"
    ElseIf UCase(wGrpArr(i)) = "PCS/PRS" Then
      wDmUomFlag = True
    ElseIf UCase(wGrpArr(i)) = "DMCTG" Then
      qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
    ElseIf UCase(wGrpArr(i)) = "SALCTG" Then
      qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
    'sv.53 OmPo option added
    ElseIf UCase(wGrpArr(i)) = "OMPONO" Then
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.OmPoNo})"
    End If
    wGrpFormula = wGrpFormula + IIF(wGrpFormula = "", "", "+") + "Trim(UpperCase(ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr(i) + "'") + "})))"
    
    wSubRepGrpFormula = wSubRepGrpFormula + IIF(wSubRepGrpFormula = "", "", ",") + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr(i) + "'")
  Next i
  
  
  
  wSubRepGrpFormula = IIF(wSubRepGrpFormula = "", "", wSubRepGrpFormula + IIF(wSubRepGrpFormula = "", "", ","))
  wGrpFormula = IIF(wGrpFormula = "", "", "{rdo.qTag}+' '+" + wGrpFormula) '*** New Code
  
    GRP_REP.SetFormula mRep, "wDetGrWt", "Sum ({rdo.qInvGrWt}, {@wDet})"
    GRP_REP.SetFormula mRep, "wTotGrWt", "Sum ({rdo.qInvGrWt})"
    GRP_REP.SetFormula mRep, "wDetAmt", "Sum ({rdo.IdiFOB}, {@wDet})"
      wRmWtOpt = "IrRmIWt"
      wRmRtOpt = "IrRmIRt"
      wRmValOpt = "IrRmIVal"
'    Else
'      wRmWtOpt = "OrWt"
'      wRmRtOpt = "OrSalRt"
'      wRmValOpt = "OrSalVal"
'    End If
      
    
  
  wHdStr1 = "Select '2H' as qTag, max(Ih.InTc) as IdTc, max(Ih.InYy) as IdYy, max(Ih.InChr) as IdChr, " + _
            "max(Ih.InNo) as IdNo, 0 as IdSr, " + ws_InExpNo + ", max(Ih.InCmCd) as InCmCd, max(Ih.InPrnDt) as InDt, " + _
            "max(Ih.InLmgCst) as InLmgCst, max(Ih.InLmgDt) as InLmgDt, max(Ih.InLmgSal) as InLmgSal, max(Ih.InVaDC) as InVaDC, " + _
            "max(Ih.InVaX) as InVaX, " + _
            "max(case when Ih.InRsMulDiv= 'D' and convert (decimal (16, 4), Ih.InCnvRt)> 0 then 1/ Ih.InCnvRt else Ih.InCnvRt end) as InCnvRt, " + _
            ws_PrdDesc + ", sum(Ih.InPkgWt) as InPkgWt, max(Ih.InGrNo) as InGrNo, " + _
            "max(Ih.InGrDt) as InGrDt, max(Ih.InSbNo) as InSbNo, max(Ih.InSbDt) as InSbDt, max(Ih.InAwbNo) as InAwbNo, max(Ih.InAwbDt) as InAwbDt, " + _
            "max(Ih.InHAwbNo) as InHAwbNo, max(Ih.InHAwbDt) as InHAwbDt, max(Ih.InHBinNo) as InHBinNo, max(Ih.InGspNo) as InGspNo, max(Ih.InGspDt) as InGspDt, " + _
            "max(Ih.InChqNo) as InChqNo, sum(Ih.InChqAmt) as InChqAmt, max(Ih.InPoNo) as InPoNo, max(Ih.InPoDt) as InPoDt, max(Ih.InCmPayTerms) as InCmPayTerms, " + _
            "sum(Ih.InFOB) as InFOB, sum(Ih.InFOBRs) as InFOBRs, sum(Ih.InComm) as InComm, sum(Ih.InCommRs) as InCommRs, sum(Ih.InFrt) as InFrt, " + _
            "sum(Ih.InFrtRs) as InFrtRs, sum(Ih.InIns) as InIns, sum(Ih.InInsRs) as InInsRs, sum(Ih.InDsc) as InDsc, " + _
            "sum(Ih.InDscRs) as InDscRs, max(Ih.InDscDesc) as InDscDesc, sum(Ih.InTot) as InTot, sum(Ih.InTotRs) as InTotRs, " + ws_InTotAs + ", " + _
            "sum(Ih.InDiaCost) as InDiaCost, sum(Ih.InCSCost) as InCSCost, sum(Ih.InAdv) as InAdv, max(Ih.InDiaCostDesc) as InDiaCostDesc, " + _
            "max(Ih.InCSCostDesc) as InCSCostDesc, max(Ih.InAdvDesc) as InAdvDesc, max(Ih.InCls) as InCls, max(Ih.InHBkName) as InHBkName, " + _
            "max(Ih.InHBkAdd1) as InHBkAdd1, max(Ih.InHBkAdd2) as InHBkAdd2, max(Ih.InHBkAdd3) as InHBkAdd3, max(Ih.InHBkAc) as InHBkAc, " + _
            "max(Ih.InHFrgnBkName) as InHFrgnBkName, max(Ih.InHFrgnBkAc) as InHFrgnBkAc, max(Ih.InHExpOff) as InHExpOff, max(Ih.InHRbiNo) as InHRbiNo, " + _
            "max(Ih.InHIecNo) as InHIecNo, max(Ih.InHPreCarriage) as InHPreCarriage, max(Ih.InHPreCarrier) as InHPreCarrier, max(Ih.InHVessel) as InHVessel, " + _
            "max(Ih.InHPortOfLd) as InHPortOfLd, max(Ih.InHConOfOrg) as InHConOfOrg, max(Ih.InHMarks) as InHMarks, max(Ih.InHPkgDesc) as InHPkgDesc, "
  wHdStr1 = wHdStr1 + "max(Ih.InHItcNo) as InHItcNo, max(Ih.InHEtcZone) as InHEtcZone, max(Ih.InHEtcSection) as InHEtcSection, max(Ih.InHConsignee) as InHConsignee, " + _
            "max(Ih.InCmInsBy) as InHInsBy, max(Ih.InHInsNo) as InHInsNo, max(Ih.InHAgName) as InHAgName, max(Ih.InHAgAdd1) as InHAgAdd1, " + _
            "max(Ih.InHAgAdd2) as InHAgAdd2, max(Ih.InHAgAdd3) as InHAgAdd3, max(Ih.InCmName) as InCmName, max(Ih.InCmAdd1) as InCmAdd1, " + _
            "max(Ih.InCmAdd2) as InCmAdd2, max(Ih.InCmAdd3) as InCmAdd3, max(Ih.InCmTel) as InCmTel, max(Ih.InCmFax) as InCmFax, max(Ih.InCmEMail) as InCmEMail, " + _
            "max(Ih.InCmCpName) as InCmCpName, max(Ih.InCmPortOfDisch) as InCmPortOfDisch, max(Ih.InCmFinDest) as InCmFinDest, max(Ih.InCmBuyer) as InCmBuyer, " + _
            "max(Ih.InCmAgName) as InCmAgName, max(Ih.InCmAgAdd1) as InCmAgAdd1, max(Ih.InCmAgAdd2) as InCmAgAdd2, max(Ih.InCmAgAdd3) as InCmAgAdd3, " + _
            "max(Ih.InSelfName) as InSelfName, max(Ih.InSelfAdd1) as InSelfAdd1, max(Ih.InSelfAdd2) as InSelfAdd2, max(Ih.InSelfAdd3) as InSelfAdd3, " + _
            "max(Ih.InSelfTel) as InSelfTel, max(Ih.InSelfFax) as InSelfFax, max(Ih.InSelfEMail) as InSelfEMail, max(Ih.InHLicNo) as InHLicNo, " + _
            "max(Ih.InHETCSvg) as InHETCSvg, max(Ih.InHAR4No) as InHAR4No, max(Ih.InHQCertNo) as InHQCertNo, max(Ih.InHGSPOrgCrit) as InHGSPOrgCrit, " + _
            "max(Ih.InHStatCd) as InHStatCd, max(Ih.InHFtRem) as InHFtRem, max(SubString(Ih.InHBnkInstr,1, 700)) as InHBnkInstr, " + _
            "max(Ih.InConAdd1) as InConAdd1, max(Ih.InConAdd2) as InConAdd2, max(Ih.InConAdd3) as InConAdd3, max(Ih.InBuyAdd1) as InBuyAdd1, " + _
            "max(Ih.InBuyAdd2) as InBuyAdd2, max(Ih.InBuyAdd3) as InBuyAdd3, max(SubString(Ih.InHUnResolutions, 1, 750)) as InHUnResolutions, " + _
            "max(Ih.InHCustRem) as InHCustRem, max(Ih.InHEximPolicy) as InHEximPolicy, max(Ih.InHGspEligibility) as InHGspEligibility, " + _
            "max(Ih.InMetalUtil) as InMetalUtil, max(Ih.InhAuthSgn) as InhAuthSgn, max(Ih.InHETCTxt) as InHETCTxt, max(Ih.InHStatOfOrg) as InHStatOfOrg, " + _
            "max(Ih.InHImpValTxt) as InHImpValTxt, max(Ih.InHShpBillSgnBy) as InHShpBillSgnBy, max(Ih.InHMetalBOETxt) as InHMetalBOETxt, " + _
            "max(Ih.InHGspPkgDesc) as InHGspPkgDesc, max(Ih.InBOELooseExp) as InBOELooseExp, max(Ih.InHHideCommYN) as InHHideCommYN, " + _
            "max(Ih.InCircularNo) as InCircularNo, "
  '6.1
  wHdStr1 = wHdStr1 + "'' as OmCmCd, '' as IdKt, '' as DmCtg, '' as DmSalCtg, '' as IdDmCd, '' as IdSfx, '' as IdDmSz, '' as DmUOM, " + _
            "'' as OmPONo, '' as OdPoNo, '' as OdPicNm, '' as CdCustDmCd, '' as OdDmCol, '' IdVaCtg, 0.0 as IdGldLs, " + _
            "0.0 as IdQty, 0.0 qInvGrWt, 0.0 as qInvGldWt, 0.0 as qInvPlWt, 0.0 as qInvSlWt, 0.0 as qInvPdWt, 0.0 as qInvAccWt, " + _
            "0.0 as qInvDiaWt, 0.0 as qInvPureGldWt,0.0 as qInvPurePlWt, 0.0 as qInvPureSlWt, 0.0 as qInvPurePdWt,0.0 as qInvPureGldLs, " + _
            "0.0 as qInvPurePlLs, 0.0 as qInvPureSlLs, 0.0 as qInvPurePdLs, 0.0 as qIGldVal, 0.0 as qIPlVal, 0.0 as qISlVal, 0.0 as qIPdVal, 0.0 as qIDCVal, 0.0 as qISetVal, 0.00 as qILabVal, " + _
            "0.0 as IdiFOB, 0.0 as qZGrWt, 0.0 as qZGldWt, 0.0 as qZPlWt, 0.0 as qZSlWt, 0.0 as qZPdWt, 0.0 as qZAccWt, 0.0 as qZDiaWt, " + _
            "0.0 as qZPureGldWt, 0.0 as qZPurePlWt, 0.0 as qZPureSlWt, 0.0 as qZPurePdWt, 0.0 as qZPureGldLs, 0.0 as qZPurePlLs, 0.0 as qZPureSlLs, 0.0 as qZPurePdLs, " + _
            "0.0 as qZGldVal, 0.0 as qZPlVal, 0.0 as qZSlVal, 0.0 as qZPdVal, 0.0 as qZDCVal, 0.0 as qZSetVal, 0.00 as qZLabVal, 0.0 as IdZFOB, " + _
            "'' as OrRmCtg, '' as qMainRmCtgDesc, '' as qDmCtgDesc, '' as qDmSalCtgDesc, '' as qDmColDesc, '' as qDmKtDesc, '' as qVaCtgDesc, " + _
            "max(Ih.InMMTC) as InMMTC, max(Ih.InHGldFiness) as InHGldFiness, max(Ih.InHPlFiness) as InHPlFiness, max(Ih.InHSlFiness) as InHSlFiness,  max(Ih.InHPdFiness) as InHPdFiness, " + _
            "max(Ih.InHGjepcNo) as InHGjepcNo, max(Ih.InHGjepcValidUpTo) as InHGjepcValidUpTo, max(Ih.InHRbiPolicy) as InHRbiPolicy, " + _
            "'' as RmCtg, '' as RmKt, 0 as qRmRt, 0 as IrRmQty, 0 as qWtInCts, 0 as qWtInGms, 0 as qVal, 0 as IdTrayNo " + _
            ",'" + UCase(MWLib.DtSetting) + "' as qDtsetting from InvHd Ih " + wCnd
  
   If UCase(adc.MenuCd) = UCase("RepPckLstInvBag") Then
  
    wStr1 = "Select '1D' as qTag, IdTc, IdYy, IdChr, IdNo, IdSr, " + ws_InExpNo + ", max(Ih.InCmCd) as InCmCd, max(Ih.InPrnDt) as InDt, " + _
            "max(Ih.InLmgCst) as InLmgCst, max(Ih.InLmgDt) as InLmgDt, max(Ih.InLmgSal) as InLmgSal, max(Ih.InVaDC) as InVaDC, max(Ih.InVaX) as InVaX, " + _
            "max(case when Ih.InRsMulDiv= 'D' and convert (decimal (16, 4), Ih.InCnvRt)> 0 then 1/ Ih.InCnvRt else Ih.InCnvRt end) as InCnvRt, " + _
            ws_PrdDesc + ", 0 as InPkgWt, max(Ih.InGrNo) as InGrNo, max(Ih.InGrDt) as InGrDt, max(Ih.InSbNo) as InSbNo, " + _
            "max(Ih.InSbDt) as InSbDt, max(Ih.InAwbNo) as InAwbNo, max(Ih.InAwbDt) as InAwbDt, max(Ih.InHAwbNo) as InHAwbNo, max(Ih.InHAwbDt) as InHAwbDt, max(Ih.InHBinNo) as InHBinNo, " + _
            "max(Ih.InGspNo) as InGspNo, max(Ih.InGspDt) as InGspDt, max(Ih.InChqNo) as InChqNo, " + _
            "max(Ih.InChqAmt) as InChqAmt, max(Ih.InPoNo) as InPoNo, max(Ih.InPoDt) as InPoDt, max(Ih.InCmPayTerms) as InCmPayTerms, 0 as InFOB, " + _
            "0 as InFOBRs, 0 as InComm, 0 as InCommRs, 0 as InFrt, 0 as InFrtRs, 0 as InIns, 0 as InInsRs, " + _
            "0 as InDsc, 0 as InDscRs, max(Ih.InDscDesc) as InDscDesc, 0 as InTot, 0 as InTotRs, " + ws_InTotAs + ", 0 as InDiaCost, " + _
            "0 as InCSCost, 0 as InAdv, max(Ih.InDiaCostDesc) as InDiaCostDesc, max(Ih.InCSCostDesc) as InCSCostDesc, max(Ih.InAdvDesc) as InAdvDesc, "
            
    wStr1 = wStr1 + "max(Ih.InCls) as InCls, max(Ih.InHBkName) as InHBkName, max(Ih.InHBkAdd1) as InHBkAdd1, max(Ih.InHBkAdd2) as InHBkAdd2, max(Ih.InHBkAdd3) as InHBkAdd3, " + _
            "max(Ih.InHBkAc) as InHBkAc, max(Ih.InHFrgnBkName) as InHFrgnBkName, max(Ih.InHFrgnBkAc) as InHFrgnBkAc, max(Ih.InHExpOff) as InHExpOff, max(Ih.InHRbiNo) as InHRbiNo, " + _
            "max(Ih.InHIecNo) as InHIecNo, max(Ih.InHPreCarriage) as InHPreCarriage, max(Ih.InHPreCarrier) as InHPreCarrier, max(Ih.InHVessel) as InHVessel, " + _
            "max(Ih.InHPortOfLd) as InHPortOfLd, max(Ih.InHConOfOrg) as InHConOfOrg, max(Ih.InHMarks) as InHMarks, max(Ih.InHPkgDesc) as InHPkgDesc, " + _
            "max(Ih.InHItcNo) as InHItcNo, max(Ih.InHEtcZone) as InHEtcZone, max(Ih.InHEtcSection) as InHEtcSection, max(Ih.InHConsignee) as InHConsignee, " + _
            "max(Ih.InCmInsBy) as InHInsBy, max(Ih.InHInsNo) as InHInsNo, max(Ih.InHAgName) as InHAgName, max(Ih.InHAgAdd1) as InHAgAdd1, max(Ih.InHAgAdd2) as InHAgAdd2, " + _
            "max(Ih.InHAgAdd3) as InHAgAdd3, max(Ih.InCmName) as InCmName, max(Ih.InCmAdd1) as InCmAdd1, max(Ih.InCmAdd2) as InCmAdd2, max(Ih.InCmAdd3) as InCmAdd3, " + _
            "max(Ih.InCmTel) as InCmTel, max(Ih.InCmFax) as InCmFax, max(Ih.InCmEMail) as InCmEMail, max(Ih.InCmCpName) as InCmCpName, max(Ih.InCmPortOfDisch) as InCmPortOfDisch, " + _
            "max(Ih.InCmFinDest) as InCmFinDest, max(Ih.InCmBuyer) as InCmBuyer, " + _
            "max(Ih.InCmAgName) as InCmAgName, max(Ih.InCmAgAdd1) as InCmAgAdd1, max(Ih.InCmAgAdd2) as InCmAgAdd2, max(Ih.InCmAgAdd3) as InCmAgAdd3, " + _
            "max(Ih.InSelfName) as InSelfName, max(Ih.InSelfAdd1) as InSelfAdd1, " + _
            "max(Ih.InSelfAdd2) as InSelfAdd2, max(Ih.InSelfAdd3) as InSelfAdd3, max(Ih.InSelfTel) as InSelfTel, max(Ih.InSelfFax) as InSelfFax, max(Ih.InSelfEMail) as InSelfEMail, " + _
            "max(Ih.InHLicNo) as InHLicNo, max(Ih.InHETCSvg) as InHETCSvg, max(Ih.InHAR4No) as InHAR4No, max(Ih.InHQCertNo) as InHQCertNo, max(Ih.InHGSPOrgCrit) as InHGSPOrgCrit, " + _
            "max(Ih.InHStatCd) as InHStatCd, max(Ih.InHFtRem) as InHFtRem, max(SubString(Ih.InHBnkInstr,1, 700)) as InHBnkInstr, max(Ih.InConAdd1) as InConAdd1, max(Ih.InConAdd2) as InConAdd2, " + _
            "max(Ih.InConAdd3) as InConAdd3, max(Ih.InBuyAdd1) as InBuyAdd1, max(Ih.InBuyAdd2) as InBuyAdd2, max(Ih.InBuyAdd3) as InBuyAdd3, max(SubString(Ih.InHUnResolutions, 1, 750)) as InHUnResolutions, max(Ih.InHCustRem) as InHCustRem, " + _
            "max(Ih.InHEximPolicy) as InHEximPolicy, max(Ih.InHGspEligibility) as InHGspEligibility, max(Ih.InMetalUtil) as InMetalUtil, max(Ih.InhAuthSgn) as InhAuthSgn, " + _
            "max(Ih.InHETCTxt) as InHETCTxt, max(Ih.InHStatOfOrg) as InHStatOfOrg, max(Ih.InHImpValTxt) as InHImpValTxt, max(Ih.InHShpBillSgnBy) as InHShpBillSgnBy, max(Ih.InHMetalBOETxt) as InHMetalBOETxt, max(Ih.InHGspPkgDesc) as InHGspPkgDesc, max(Ih.InBOELooseExp) as InBOELooseExp, " + _
            "max(Ih.InHHideCommYN) as InHHideCommYN, max(Ih.InCircularNo) as InCircularNo, "
    
    wStr2 = "Max(OdOmCmCd) as OmCmCd, max(" + Ws_IdKt + ") as IdKt, max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmCd) as IdDmCd, max(" + ws_IdSfx + ") as IdSfx, max(" + ws_IdDmSz + ") as IdDmSz, max(DmUom) as DmUom, " + wDetDesc2Str + _
            " max(OdDmCol) as OdDmCol, max(" + ws_IdVaCtg + ") as IdVaCtg, max(IdGldLs) as IdGldLs, max(" + ws_IdQty + ") as IdQty, " + _
            "sum(IrRmAWt / (case when R1.RmCtg in ('C', 'D') then 5 else 1 end)) as qInvGrWt, " + _
            "sum(IrRmAWt * (case R1.RmCtg when 'G' then 1 else 0 end)) as qInvGldWt, " + _
            "sum(IrRmAWt * (case R1.RmCtg when 'P' then 1 else 0 end)) as qInvPlWt, " + _
            "sum(IrRmAWt * (case R1.RmCtg when 'S' then 1 else 0 end)) as qInvSlWt, " + _
            "sum(IrRmAWt * (case R1.RmCtg when 'L' then 1 else 0 end)) as qInvPdWt, " + _
            "sum(IrRmAWt * (case R1.RmCtg when 'X' then 1 else 0 end)) as qInvAccWt, " + _
            "sum(IrRmAWt * (case when R1.RmCtg in ('C', 'D') then 1 else 0 end)) as qInvDiaWt, " + _
            "sum(IrRmAWt * R1.RmPurityZ * (case R1.RmCtg when 'G' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureGldWt, " + _
            "sum(IrRmAWt * R1.RmPurityZ * (case R1.RmCtg when 'P' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPurePlWt, " + _
            "sum(IrRmAWt * R1.RmPurityZ * (case R1.RmCtg when 'S' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureSlWt, " + _
            "sum(IrRmAWt * R1.RmPurityZ * (case R1.RmCtg when 'L' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPurePdWt, " + _
            "sum(IrRmAWt * R1.RmPurityZ * (case R1.RmCtg when 'G' then 1 else 0 end)* (IdGldLs/100)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureGldLs, " + _
            "sum(IrRmAWt * R1.RmPurityZ * (case R1.RmCtg when 'P' then 1 else 0 end)* (IdGldLs/100)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPurePlLs, " + _
            "sum(IrRmAWt * R1.RmPurityZ * (case R1.RmCtg when 'S' then 1 else 0 end)* (IdGldLs/100)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureSlLs, " + _
            "sum(IrRmAWt * R1.RmPurityZ * (case R1.RmCtg when 'L' then 1 else 0 end)* (IdGldLs/100)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPurePdLs, " + _
            "sum(IrRmIVal * (case R1.RmCtg when 'G' then 1 else 0 end)) as qIGldVal, " + _
            "sum(IrRmIVal * (case R1.RmCtg when 'P' then 1 else 0 end)) as qIPlVal, " + _
            "sum(IrRmIVal * (case R1.RmCtg when 'S' then 1 else 0 end)) as qISlVal, " + _
            "sum(IrRmIVal * (case R1.RmCtg when 'L' then 1 else 0 end)) as qIPdVal, " + _
            "sum(IrRmIVal * (case when R1.RmCtg in ('C', 'D') then 1 else 0 end)) as qIDCVal, " + _
            "sum(IrSetIVal) as qISetVal, " + wILabValStr + " as qILabVal, max(IdiFOB) as IdiFOB, "
    
     wStr3 = "sum(IrRmZWt / (case when R1.RmCtg in ('C', 'D') then 5 else 1 end)) as qZGrWt, " + _
            "sum(IrRmZWt * (case R1.RmCtg when 'G' then 1 else 0 end)) as qZGldWt, " + _
            "sum(IrRmZWt * (case R1.RmCtg when 'P' then 1 else 0 end)) as qZPlWt, " + _
            "sum(IrRmZWt * (case R1.RmCtg when 'S' then 1 else 0 end)) as qZSlWt, " + _
            "sum(IrRmZWt * (case R1.RmCtg when 'L' then 1 else 0 end)) as qZPdWt, " + _
            "sum(IrRmZWt * (case R1.RmCtg when 'X' then 1 else 0 end)) as qZAccWt, sum(IrRmZWt * (case when R1.RmCtg in ('C', 'D') then 1 else 0 end)) as qZDiaWt, " + _
            "sum(IrRmZWt * R1.RmPurityZ * (case R1.RmCtg when 'G' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPureGldWt, " + _
            "sum(IrRmZWt * R1.RmPurityZ * (case R1.RmCtg when 'P' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPurePlWt, " + _
            "sum(IrRmZWt * R1.RmPurityZ * (case R1.RmCtg when 'S' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPureSlWt, " + _
            "sum(IrRmZWt * R1.RmPurityZ * (case R1.RmCtg when 'L' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPurePdWt, " + _
            "sum(" + " (case when IrMainMet= 'Y' then IdCustmMtchWt- IrRmZWt else IrRmZWt* IdGldLs/100 end) " + "* R1.RmPurityZ * (case R1.RmCtg when 'G' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPureGldLs, " + _
            "sum(" + " (case when IrMainMet= 'Y' then IdCustmMtchWt- IrRmZWt else IrRmZWt* IdGldLs/100 end) " + "* R1.RmPurityZ * (case R1.RmCtg when 'P' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPurePlLs, " + _
            "sum(" + " (case when IrMainMet= 'Y' then IdCustmMtchWt- IrRmZWt else IrRmZWt* IdGldLs/100 end) " + "* R1.RmPurityZ * (case R1.RmCtg when 'S' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPureSlLs, " + _
            "sum(" + " (case when IrMainMet= 'Y' then IdCustmMtchWt- IrRmZWt else IrRmZWt* IdGldLs/100 end) " + "* R1.RmPurityZ * (case R1.RmCtg when 'L' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPurePdLs, " + _
            "sum(IrRmZVal * (case R1.RmCtg when 'G' then 1 else 0 end)) as qZGldVal, " + _
            "sum(IrRmZVal * (case R1.RmCtg when 'P' then 1 else 0 end)) as qZPlVal, " + _
            "sum(IrRmZVal * (case R1.RmCtg when 'S' then 1 else 0 end)) as qZSlVal, " + _
            "sum(IrRmZVal * (case R1.RmCtg when 'L' then 1 else 0 end)) as qZPdVal, " + _
            "sum(IrRmZVal * (case when R1.RmCtg in ('C', 'D') then 1 else 0 end)) as qZDCVal, "
    '6.1
    wStr3 = wStr3 + "sum(IrSetZVal) as qZSetVal, " + wZLabValStr + " as qZLabVal, max(IdZFOB) as IdZFOB, " + qMainMet + ", " + qMainRmCtgDesc + ", " + qDmCtgDesc + ", " + qDmSalCtgDesc + ", " + qDmColDesc + ", " + qDmKtDesc + ", " + qVaCtgDesc + ", max(Ih.InMMTC) as InMMTC," + _
            " max(Ih.InHGldFiness) as InHGldFiness, max(Ih.InHPlFiness) as InHPlFiness, max(Ih.InHSlFiness) as InHSlFiness,  max(Ih.InHPdFiness) as InHPdFiness, max(Ih.InHGjepcNo) as InHGjepcNo, max(Ih.InHGjepcValidUpTo) as InHGjepcValidUpTo, Max(Ih.InHRbiPolicy) as InHRbiPolicy, " + _
            " '' as RmCtg, '' as RmKt, 0 as qRmRt, 0 as IrRmQty, 0 as qWtInCts, 0 as qWtInGms, 0 as qVal, max(IdTrayNo) as IdTrayNo  " + _
            ",'" + UCase(MWLib.DtSetting) + "' as qDtsetting " + _
            " from InvRm join RmMst R1 on " + IIF(gs_Partition = ctCurrPrtn, " R1.RmPrtKey=IrPrtKey and ", "") + " R1.RmCd= IrRmCd join RmMst R2 on R2.RmCd= R1.RmZCd " + IIF(gs_Partition = ctCurrPrtn, " And R2.RmPrtKey=R1.RmPrtKey ", "") + _
            " join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr " + _
            " and IdNo= IrNo and IdSr= IrSr " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=IrPrtKey ", "") + _
            " join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and DmCd = IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + _
            " join OrdMst on OmCoCd= IdCoCd and OmTc= IdExpOdTc and OmYy= IdExpOdYy " + _
            " and OmChr= IdExpOdChr and OmNo= IdExpOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=IdPrtKey ", "") + _
            " join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
            " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey ", "") + wOrJoin + _
            " join InvHd Ih on Ih.InIdNo=IdInIdNo and Ih.InCoCd= IdCoCd and Ih.InTc = IdTc and Ih.InYy= IdYy and Ih.InChr= IdChr and Ih.InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=IdPrtKey ", "") + _
            wCnd + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey='" + ctCurrPrtn + "' ", "") + " Group by IdTc, IdYy, IdChr, IdNo, IdSr "
            
            
  Else
    wStr1 = "Select '1D' as qTag, OdTc as IdTc, OdYy as IdYy, OdChr as IdChr, OdNo as IdNo, OdSr as IdSr, " + ws_InExpNo + ", max(Ih1.InCmCd) as InCmCd, max(Ih1.InPrnDt) as InDt, " + _
            "max(Ih1.InLmgCst) as InLmgCst, max(Ih1.InLmgDt) as InLmgDt, max(Ih1.InLmgSal) as InLmgSal, max(Ih1.InVaDC) as InVaDC, max(Ih1.InVaX) as InVaX, " + _
            "max(case when Ih1.InRsMulDiv= 'D' and convert (decimal (16, 4), Ih1.InCnvRt)> 0 then 1/ Ih1.InCnvRt else Ih1.InCnvRt end) as InCnvRt, " + _
            ws_PrdDesc + ", 0 as InPkgWt, max(Ih1.InGrNo) as InGrNo, max(Ih1.InGrDt) as InGrDt, max(Ih1.InSbNo) as InSbNo, " + _
            "max(Ih1.InSbDt) as InSbDt, max(Ih1.InAwbNo) as InAwbNo, max(Ih1.InAwbDt) as InAwbDt, max(Ih1.InHAwbNo) as InHAwbNo, max(Ih1.InHAwbDt) as InHAwbDt, max(Ih1.InHBinNo) as InHBinNo, " + _
            "max(Ih1.InGspNo) as InGspNo, max(Ih1.InGspDt) as InGspDt, max(Ih1.InChqNo) as InChqNo, " + _
            "max(Ih1.InChqAmt) as InChqAmt, max(Ih1.InPoNo) as InPoNo, max(Ih1.InPoDt) as InPoDt, max(Ih1.InCmPayTerms) as InCmPayTerms, 0 as InFOB, " + _
            "0 as InFOBRs, 0 as InComm, 0 as InCommRs, 0 as InFrt, 0 as InFrtRs, 0 as InIns, 0 as InInsRs, " + _
            "0 as InDsc, 0 as InDscRs, max(Ih.InDscDesc) as InDscDesc, 0 as InTot, 0 as InTotRs, " + ws_InTotAs + ", 0 as InDiaCost, " + _
            "0 as InCSCost, 0 as InAdv, max(Ih.InDiaCostDesc) as InDiaCostDesc, max(Ih.InCSCostDesc) as InCSCostDesc, max(Ih.InAdvDesc) as InAdvDesc, "
    wStr1 = wStr1 + "max(Ih1.InCls) as InCls, max(Ih1.InHBkName) as InHBkName, max(Ih1.InHBkAdd1) as InHBkAdd1, max(Ih1.InHBkAdd2) as InHBkAdd2, max(Ih1.InHBkAdd3) as InHBkAdd3, " + _
            "max(Ih1.InHBkAc) as InHBkAc, max(Ih1.InHFrgnBkName) as InHFrgnBkName, max(Ih1.InHFrgnBkAc) as InHFrgnBkAc, max(Ih1.InHExpOff) as InHExpOff, max(Ih1.InHRbiNo) as InHRbiNo, " + _
            "max(Ih1.InHIecNo) as InHIecNo, max(Ih1.InHPreCarriage) as InHPreCarriage, max(Ih1.InHPreCarrier) as InHPreCarrier, max(Ih1.InHVessel) as InHVessel, " + _
            "max(Ih1.InHPortOfLd) as InHPortOfLd, max(Ih1.InHConOfOrg) as InHConOfOrg, max(Ih1.InHMarks) as InHMarks, max(Ih1.InHPkgDesc) as InHPkgDesc, " + _
            "max(Ih1.InHItcNo) as InHItcNo, max(Ih1.InHEtcZone) as InHEtcZone, max(Ih1.InHEtcSection) as InHEtcSection, max(Ih1.InHConsignee) as InHConsignee, " + _
            "max(Ih1.InCmInsBy) as InHInsBy, max(Ih1.InHInsNo) as InHInsNo, max(Ih1.InHAgName) as InHAgName, max(Ih1.InHAgAdd1) as InHAgAdd1, max(Ih1.InHAgAdd2) as InHAgAdd2, " + _
            "max(Ih1.InHAgAdd3) as InHAgAdd3, max(Ih1.InCmName) as InCmName, max(Ih1.InCmAdd1) as InCmAdd1, max(Ih1.InCmAdd2) as InCmAdd2, max(Ih1.InCmAdd3) as InCmAdd3, " + _
            "max(Ih1.InCmTel) as InCmTel, max(Ih1.InCmFax) as InCmFax, max(Ih1.InCmEMail) as InCmEMail, max(Ih1.InCmCpName) as InCmCpName, max(Ih1.InCmPortOfDisch) as InCmPortOfDisch, " + _
            "max(Ih1.InCmFinDest) as InCmFinDest, max(Ih1.InCmBuyer) as InCmBuyer, " + _
            "max(Ih1.InCmAgName) as InCmAgName, max(Ih1.InCmAgAdd1) as InCmAgAdd1, max(Ih1.InCmAgAdd2) as InCmAgAdd2, max(Ih1.InCmAgAdd3) as InCmAgAdd3, " + _
            "max(Ih1.InSelfName) as InSelfName, max(Ih1.InSelfAdd1) as InSelfAdd1, " + _
            "max(Ih1.InSelfAdd2) as InSelfAdd2, max(Ih1.InSelfAdd3) as InSelfAdd3, max(Ih1.InSelfTel) as InSelfTel, max(Ih1.InSelfFax) as InSelfFax, max(Ih1.InSelfEMail) as InSelfEMail, " + _
            "max(Ih1.InHLicNo) as InHLicNo, max(Ih1.InHETCSvg) as InHETCSvg, max(Ih1.InHAR4No) as InHAR4No, max(Ih1.InHQCertNo) as InHQCertNo, max(Ih1.InHGSPOrgCrit) as InHGSPOrgCrit, " + _
            "max(Ih1.InHStatCd) as InHStatCd, max(Ih1.InHFtRem) as InHFtRem, max(SubString(Ih1.InHBnkInstr,1, 700)) as InHBnkInstr, max(Ih1.InConAdd1) as InConAdd1, max(Ih1.InConAdd2) as InConAdd2, " + _
            "max(Ih1.InConAdd3) as InConAdd3, max(Ih1.InBuyAdd1) as InBuyAdd1, max(Ih1.InBuyAdd2) as InBuyAdd2, max(Ih1.InBuyAdd3) as InBuyAdd3, max(SubString(Ih1.InHUnResolutions, 1, 750)) as InHUnResolutions, max(Ih1.InHCustRem) as InHCustRem, " + _
            "max(Ih1.InHEximPolicy) as InHEximPolicy, max(Ih1.InHGspEligibility) as InHGspEligibility, max(Ih1.InMetalUtil) as InMetalUtil, max(Ih1.InhAuthSgn) as InhAuthSgn, " + _
            "max(Ih1.InHETCTxt) as InHETCTxt, max(Ih1.InHStatOfOrg) as InHStatOfOrg, max(Ih1.InHImpValTxt) as InHImpValTxt, max(Ih1.InHShpBillSgnBy) as InHShpBillSgnBy, max(Ih1.InHMetalBOETxt) as InHMetalBOETxt, max(Ih1.InHGspPkgDesc) as InHGspPkgDesc, max(Ih1.InBOELooseExp) as InBOELooseExp, " + _
            "max(Ih1.InHHideCommYN) as InHHideCommYN, max(Ih1.InCircularNo) as InCircularNo, "
    
      '*** Jay 3.3(FG)
      wStr2 = "Max(OdOmCmCd) as OmCmCd,max(" + Ws_IdKt + ") as IdKt, max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmCd) as IdDmCd, max(" + ws_IdSfx + ") as IdSfx, max(" + ws_IdDmSz + ") as IdDmSz, max(DmUom) as DmUom, " + wDetDesc2Str + _
            " max(OdDmCol) as OdDmCol, max(" + ws_IdVaCtg + ") as IdVaCtg, 0 as IdGldLs, max(" + ws_IdQty + ") as IdQty, " + _
            "sum(OrWt / (case when R1.RmCtg in ('C', 'D') then 5 else 1 end)) as qInvGrWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'G' then 1 else 0 end)) as qInvGldWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'P' then 1 else 0 end)) as qInvPlWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'S' then 1 else 0 end)) as qInvSlWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'L' then 1 else 0 end)) as qInvPdWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'X' then 1 else 0 end)) as qInvAccWt, " + _
            "sum(OrWt * (case when R1.RmCtg in ('C', 'D') then 1 else 0 end)) as qInvDiaWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'G' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureGldWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'P' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPurePlWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'S' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureSlWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'L' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPurePdWt, " + _
            "0 as qInvPureGldLs, " + _
            "0 as qInvPurePlLs, " + _
            "0 as qInvPureSlLs, " + _
            "0 as qInvPurePdLs, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'G' then 1 else 0 end)) as qIGldVal, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'P' then 1 else 0 end)) as qIPlVal, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'S' then 1 else 0 end)) as qISlVal, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'L' then 1 else 0 end)) as qIPdVal, " + _
            "sum(OrSalVal * (case when R1.RmCtg in ('C', 'D') then 1 else 0 end)) as qIDCVal, " + _
            "sum(OrSetSalVal) as qISetVal, " + wILabValStr + " as qILabVal, max(OdSalPrc) as IdiFOB, "
    
    wStr3 = "sum(OrWt / (case when R1.RmCtg in ('C', 'D') then 5 else 1 end)) as qZGrWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'G' then 1 else 0 end)) as qZGldWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'P' then 1 else 0 end)) as qZPlWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'S' then 1 else 0 end)) as qZSlWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'L' then 1 else 0 end)) as qZPdWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'X' then 1 else 0 end)) as qZAccWt, sum(OrWt * (case when R1.RmCtg in ('C', 'D') then 1 else 0 end)) as qZDiaWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'G' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPureGldWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'P' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPurePlWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'S' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPureSlWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'L' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPurePdWt, " + _
            "0 as qZPureGldLs, " + _
            "0 as qZPurePlLs, " + _
            "0 as qZPureSlLs, " + _
            "0 as qZPurePdLs, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'G' then 1 else 0 end)) as qZGldVal, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'P' then 1 else 0 end)) as qZPlVal, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'S' then 1 else 0 end)) as qZSlVal, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'S' then 1 else 0 end)) as qZPdVal, " + _
            "sum(OrSalVal * (case when R1.RmCtg in ('C', 'D') then 1 else 0 end)) as qZDCVal, "
    '6.1
    wStr3 = wStr3 + "sum(OrSetSalVal) as qZSetVal, " + wZLabValStr + " as qZLabVal, max(OdSalPrc) as IdZFOB, " + qMainMet + ", " + qMainRmCtgDesc + ", " + qDmCtgDesc + ", " + qDmSalCtgDesc + ", " + qDmColDesc + ", " + qDmKtDesc + ", " + qVaCtgDesc + ", max(Ih1.InMMTC) as InMMTC," + _
            " max(Ih1.InHGldFiness) as InHGldFiness, max(Ih1.InHPlFiness) as InHPlFiness, max(Ih1.InHSlFiness) as InHSlFiness,  max(Ih1.InHPdFiness) as InHPdFiness, max(Ih1.InHGjepcNo) as InHGjepcNo, max(Ih1.InHGjepcValidUpTo) as InHGjepcValidUpTo, Max(Ih1.InHRbiPolicy) as InHRbiPolicy, " + _
            " '' as RmCtg, '' as RmKt, 0 as qRmRt, 0 as IrRmQty, 0 as qWtInCts, 0 as qWtInGms, 0 as qVal " + _
            ",'" + UCase(MWLib.DtSetting) + "' as qDtsetting " + _
            " from OrdRm join RmMst R1 on " + IIF(gs_Partition = ctCurrPrtn, " R1.RmPrtKey=OrPrtKey and ", "") + " R1.RmCd= OrRmCd join RmMst R2 on R2.RmCd= R1.RmZCd " + IIF(gs_Partition = ctCurrPrtn, " And R2.RmPrtKey=R1.RmPrtKey ", "") + _
            " join OrdDsg on OdIdNo=OrOdIdNo and OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and OdChr= OrChr " + _
            " and OdNo= OrNo and OdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=OrPrtKey ", "") + _
            " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd = OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=OdPrtKey ", "") + _
            " join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy " + _
            " and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
            " join InvHd Ih on Ih.InCoCd= OdCoCd and Ih.InTc = OdTc and Ih.InYy= OdYy and Ih.InChr= OdChr and Ih.InNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=OdPrtKey ", "") + _
            " join InvHd Ih1 on " + ws_MinCnd + wCnd + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey='" + ctCurrPrtn + "' ", "") + " Group by OdTc, OdYy, OdChr, OdNo, OdSr "
  End If
    
  wSqlStrg = wHdStr1 + " Union All " + wStr1 + wStr2 + wStr3 + " Order By qTag, IdTc, IdYy, IdChr, IdNo, IdSr "

  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added

  
  Select Case UCase(adc.MenuCd)
  Case Is = UCase("RepPckLstInvBag")
    GRP_REP.SetFormula mRep, "wGrp1", "if {rdo.qTag}='1D' Then " + wGrpFormula
    GRP_REP.SetFormula mRep, "wGrp1Desc", wGrpDescFormula
    GRP_REP.SetFormula mRep, "wDet", wDetFormula
    GRP_REP.SetFormula mRep, "wDetDesc", wDetDescFormula
    GRP_REP.SetFormula mRep, "wDetDesc2", wDetDesc2Formula
    
    
    If UCase(adc("UoInDet")) = "" Then
        GRP_REP.SetFormula mRep, "wDet", wGrpFormula
    End If
    
    
    If UCase(adc("UoYN")) = "Y" Then
      mRep.PICSEC.Suppress = True: mRep.PicSec1.Suppress = True
      If UCase(adc("UoInDet")) = "DMCD" Or UCase(adc("UoInDet")) = "INVSR" Or UCase(adc("UoInDet")) = "INVSR/SR" Or _
         UCase(adc("UoInDet")) = "DMCD/SR" Then If UCase(adc("UoYN9")) = "Y" Then mRep.PicSec1.Suppress = False Else mRep.PICSEC.Suppress = False
    Else
        mRep.PICSEC.Suppress = True: mRep.PicSec1.Suppress = True
    End If
  

  End Select
  
        'sv.53 OrdMst join, max(OmPoNo) as OmPoNo added
        wSubRepStr = "Select '1D' as qTag,Max(OdOmCmCd) as OmCmCd, DmCtg, max(DmSalCtg) as DmSalCtg, (Select PDesc From Param where PTyp= 'DMCTG'" + _
                  " and PMCd= DmCtg and PSCd= '') as DmCtgDesc, " + _
                  " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
                  " PSCd= max(DmSalCtg)) as DmSalCtgDesc, max(DmUom) as DmUom,Max(" + Ws_IdKt + ") as IdKt," + _
                  "Max(" + ws_IdVaCtg + ") as IdVaCtg, Max(IdGldLs) as IdGldLs, Max(OdDmCol) as OdDmCol," + qMainMet + ", sum(" + ws_IdQty + ") as qIdQty, max(IdTrayNo) as IdTrayNo, max(OmPoNo) as OmPoNo " + _
                  " From InvDsg Join DsgMst  " + _
                  " On DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and IdDmCd= DmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=DmPrtKey", "") + _
                  " join InvHd Ih on Ih.InIdNo=IdInIdNo and Ih.InCoCd= IdCoCd and Ih.InTc= IdTc and Ih.InYy= IdYy and Ih.InChr= IdChr and Ih.InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=IdPrtKey ", "") + _
                  " join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
                  " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey", "") + _
                  " join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy  and OmChr= OdChr and OmNo= OdNo  And OmPrtKey=OdPrtKey " + wOrJoin + _
                  " " + wCnd + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey='" + ctCurrPrtn + "' ", "") + " group by " + wSubRepGrpFormula + "DmCtg, DmUom "
      
      Call moCn.RepRes(mRep.Subreport3.OpenSubreport, wSubRepStr + ctMaxDopOpt)
      
      SetSubFormula mRep.Subreport3, "wGrp", wGrpFormula
  
  
  Dim ws_PckLstSubRep As String
  If adc("UoYN9") = "Y" Then
    ws_PckLstSubRep = "Select '1D' as qTag, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, IfFdBYy, IfFdBChr, IfFdBNo, max(IfSrNo) as IfSrno, Sum(IfQty) as IfQty, Sum(IfGrsWt) as IfGrsWt, " + _
                      "Max(OdOmCmCd) as OmCmCd,max(IdKt) as IdKt, max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, " + _
                      "max(DmCd) as IdDmCd, max(IdSfx) as IdSfx, max(IdDmSz) as IdDmSz, " + _
                      "max(DmUom) as DmUom,  max(OmPONo) as OmPONo, max(OdPoNo) as OdPoNo, " + _
                      "max(OdDmCol) as OdDmCol, max(IdVaCtg) as IdVaCtg, max(IdGldLs) as IdGldLs, " + qMainMet + ", max(IdTrayNo) as IdTrayNo  " + _
                      " from InvFgd " + _
                      " join InvDsg on IdInIdNo=IfInIdNo and IdCoCd= IfCoCd and IdTc= IfTc and IdYy= IfYy and IdChr= IfChr " + _
                      " and IdNo= IfNo and IdSr= IfSr " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=IfPrtKey ", "") + _
                      " join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and DmCd = IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + _
                      " join OrdMst on OmCoCd= IdCoCd and OmTc= IdExpOdTc and OmYy= IdExpOdYy " + _
                      " and OmChr= IdExpOdChr and OmNo= IdExpOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=IdPrtKey ", "") + _
                      " join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
                      " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey ", "") + wOrJoin + _
                      " join InvHd Ih on Ih.InIdNo=IdInIdNo and Ih.InCoCd= IdCoCd and Ih.InTc = IdTc and Ih.InYy= IdYy and Ih.InChr= IdChr and Ih.InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=IdPrtKey ", "") + _
                      wCnd + IIF(gs_Partition = ctCurrPrtn, " And IfPrtKey='" + ctCurrPrtn + "' ", "") + " Group by IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, IfFdBYy, IfFdBChr, IfFdBNo Order By IdSr, IfFdBYy, IfFdBChr, IfFdBNo "
    Call moCn.RepRes(Rep.SubRepInvBag.OpenSubreport, ws_PckLstSubRep + ctMaxDopOpt)
    
        SetSubFormula Rep.SubRepInvBag, "wSubDet", wDetFormula
      SetSubFormula Rep.SubRepInvBag, "wSubGrp", wGrpFormula
      If UCase(adc("UoInDet")) = "" Then
          SetSubFormula Rep.SubRepInvBag, "wSubDet", wGrpFormula
      End If
  Else
    Rep.SecInvBag.Suppress = True
  End If
  If wFlagLocalExp = "N" Then Rep.SecDescSubRep.Suppress = True
  
    
     
  GRP_REP.SetFormula mRep, "wGrNoDt", " iif({rdo.InGrNo}<>'', {rdo.InGrNo}, '') " + _
                                        "+ iif({rdo.InGrNo}<>'' and {rdo.InGrDt}<> DateValue (1980,01,01), ' & ', '')  + " + _
                                        "iif({rdo.InGrDt}<> DateValue (1980,01,01), ToText({rdo.InGrDt}), '') "
  
  If wFlagLocalExp = "N" Then
      Rep.FldInMetalUtil.Suppress = True
  Else
      Rep.FldInCmAgAdd1.Suppress = True
      Rep.FldInCmAgAdd2.Suppress = True
      Rep.FldInCmAgAdd3.Suppress = True
  End If
  
  CRV_REP.DisplayGroupTree = False
  
  '****** DemarkSec is necessary for the lines to extend till the page footer ******
    If UCase(adc("UoYN1")) = "Y" Then
      mRep.RemSec.Suppress = True: mRep.RemSec1.Suppress = True
      If UCase(adc("UoInDet")) = "DMCD" Or UCase(adc("UoInDet")) = "INVSR" Or UCase(adc("UoInDet")) = "INVSR/SR" Or _
         UCase(adc("UoInDet")) = "DMCD/SR" Then If UCase(adc("UoYN9")) = "Y" Then mRep.RemSec1.Suppress = False Else mRep.RemSec.Suppress = False
    Else
        mRep.RemSec.Suppress = True: mRep.RemSec1.Suppress = True
    End If


  GRP_REP.SetFormula Rep, "wDetLineYN", "'Y'"
  
  If adc("UoYn9") = "Y" Then
    GRP_REP.SetFormula Rep, "wShowBag", "'Y'"
  Else
    GRP_REP.SetFormula Rep, "wShowBag", "'N'"
  End If
  
Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
'*** (Jen 2.12)
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
'*** (Jen 2.12)
End Sub

' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0       ' Tag = "ShowPic"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1      ' tag = "ShowCustPoDsg"
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  Case Is = 2      ' tag = "SortbyOrdCust"
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  Case Is = 9      ' Tag = "ShowbagYn"
    If .Value = Checked Then adc("UoYN9") = "Y" Else adc("UoYN9") = "N"
  End Select
End With
End Sub
' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added

' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added

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
    Set moCn = adc.Connection
      Set adc.FirNKeyCtl = adc("UoInTcFr")
    
    Call InitProp(Me)
    
      gs_InTcTyp = "IN"
    
    
    gs_RmGrp = "INV"
    gs_Tbl = "Param"
    gs_PTyp = "GRLSSUMM"
          
    wFlagLocalExp = moCn.GetFldVal("Select hExpLocal from Head where HCoCd= '" + gs_CoCd + "' and " + _
                                                                  "HCd= '" + ctSelfCmCd + "' ")
    
       mb_MultiInv = False
      
                                        

    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
 
    gs_InTcTyp = "IN"
    
  gb_InHlpFrOm = False

  gs_RmGrp = "INV"
  gs_Tbl = "Param"
  gs_PTyp = "GRLSSUMM"
  '*** (Jen 2.12)
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
    gs_InTcTyp = "IN"

  gb_InHlpFrOm = False ' **** Manali 3.7.1

  
  '*** (Jen 2.12)
  gs_RmGrp = "INV"
  gs_Tbl = "Param"
  gs_PTyp = "GRLSSUMM"
  '*** (Jen 2.12)

  SetProp Me, IdName, When


End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
''''''''
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
  
  If adc("UOINDET") = "" Then adc("UOINDET") = "INVSR/SR"
  
  If adc("UOYN") = "" Then adc("UOYN") = "N"
  If adc("UOYN1") = "" Then adc("UOYN1") = "N"
  If adc("UOYN2") = "" Then adc("UOYN2") = "N"
  If adc("UOYN9") = "" Then adc("UOYN9") = "N"

  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If adc("UoYN9") = "Y" Then ChkBoxArr(9).Value = Checked Else ChkBoxArr(9).Value = Unchecked
  
End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                            "HCd = '" + ctSelfCmCd + "' ")
  mRep.wCoCd.SetText ws_HName
  
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
  'Rep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  
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
End Sub



