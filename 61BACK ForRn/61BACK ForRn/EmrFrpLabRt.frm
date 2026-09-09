VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpLabRt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Labour Rate Master Listing"
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
      Left            =   11580
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9600
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
      Left            =   5745
      TabIndex        =   23
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   24
         Top             =   165
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
         Top             =   180
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
      Left            =   6960
      TabIndex        =   22
      Top             =   9615
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
      Height          =   9675
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
         TabPicture(0)   =   "EmrFrpLabRt.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpLabRt.frx":001C
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
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   16
               Top             =   360
               Width           =   9885
               _ExtentX        =   17436
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   360
            WhatsThisHelpID =   11760
            Width           =   15105
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   2955
               TabIndex        =   6
               ToolTipText     =   "Enter To Labour Sub Code"
               Top             =   1500
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOLSCDTO"
               IdName          =   "UOLSCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   4500
               TabIndex        =   4
               ToolTipText     =   "Enter Labour Main Code Selection"
               Top             =   1215
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPMCDSEL"
               IdName          =   "UOPMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1425
               TabIndex        =   5
               ToolTipText     =   "Enter From Labour Sub Code"
               Top             =   1500
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOLSCDFR"
               IdName          =   "UOLSCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1425
               TabIndex        =   2
               ToolTipText     =   "Enter From Labour Main Code"
               Top             =   1215
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   2955
               TabIndex        =   3
               ToolTipText     =   "Enter To Labour Main Code"
               Top             =   1215
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   2955
               TabIndex        =   8
               ToolTipText     =   "Enter To Customer"
               Top             =   1785
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   4500
               TabIndex        =   9
               ToolTipText     =   "Enter Customer Selection"
               Top             =   1785
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1425
               TabIndex        =   7
               ToolTipText     =   "Enter From Customer"
               Top             =   1785
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1425
               TabIndex        =   0
               ToolTipText     =   "Enter Customer Category"
               Top             =   930
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOCMCTGFR"
               IdName          =   "UOCMCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1425
               TabIndex        =   1
               ToolTipText     =   "Enter From YearMonth"
               Top             =   2565
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOYYMMFR"
               IdName          =   "UOYYMMFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   1425
               TabIndex        =   10
               ToolTipText     =   "Enter From Currency Code"
               Top             =   2070
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               ReCalcOn        =   "UOPMCDFR,UOPMCDTO"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   2955
               TabIndex        =   11
               ToolTipText     =   "Enter To Currency Code"
               Top             =   2070
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDTO"
               IdName          =   "UOCURCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   4500
               TabIndex        =   12
               ToolTipText     =   "Enter Currency Code Selection"
               Top             =   2070
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCURCDSEL"
               IdName          =   "UOCURCDSEL"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Currency"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   175
               Left            =   60
               TabIndex        =   34
               Top             =   2070
               Width           =   1005
            End
            Begin VB.Label LblLhYyyyMm 
               BackStyle       =   0  'Transparent
               Caption         =   "YyyyMm"
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
               Left            =   60
               TabIndex        =   33
               Top             =   2565
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Cust Ctg"
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
               Index           =   6
               Left            =   60
               TabIndex        =   32
               ToolTipText     =   "Location"
               Top             =   930
               Width           =   1395
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer"
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
               Index           =   17
               Left            =   60
               TabIndex        =   31
               Top             =   1785
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Lab Main Cd"
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
               Index           =   4
               Left            =   60
               TabIndex        =   30
               Top             =   1215
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Lab Sub Cd"
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
               TabIndex        =   29
               Top             =   1500
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Selection"
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
               Index           =   3
               Left            =   4500
               TabIndex        =   28
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "To"
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
               Index           =   2
               Left            =   2955
               TabIndex        =   27
               Top             =   480
               Width           =   1275
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
               Left            =   1425
               TabIndex        =   26
               Top             =   480
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpLabRt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepLabRt
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ws_Opt As String
Private Sub SetGroupSort()
'  If UCase(adc.MenuCd) = UCase("RepLabRt") Then
'    GRP_REP.Add "Lab Main Cd", "LrMCd", "", "LrMCd", "hLrMCd", "Lab Sub Cd", "", "", "(select PDesc from Param where PTyp= 'LABMCD' and PMCd= LrMCd) "
'    GRP_REP.Add "Lab Sub Cd", "LrSCd", "", "LrSCd", "hLrSCd", "Lab Main Cd,Customer", "", "", "(select PDesc from Param where PTyp= 'LABSCD' and PMCd= LrMCd and PSCd= LrSCd) "
'    GRP_REP.Add "Customer", "LrCmCd", "", "LrCmCd", "hLrCmCd", "Lab Sub Cd", "", "", "(select CmName from CustMst where CmCtg= 'C' and CmCd= LrCmCd) "
'  ElseIf UCase(adc.MenuCd) = UCase("RepLabRtHist") Then
'    GRP_REP.Add "Lab Main Cd", "LhMCd", "", "LrMCd", "hLrMCd", "", "", "", "(Select PDesc from Param where PTyp= 'LABMCD' and PMCd= LhMCd) "
'    GRP_REP.Add "Lab Sub Cd", "LhSCd", "", "LrSCd", "hLrSCd", "Lab Main Cd,Customer", "", "", "(Select PDesc from Param where PTyp= 'LABSCD' and PMCd= LhMCd and PSCd= LhSCd) "
'    GRP_REP.Add "Customer", "LhCmCd", "", "LrCmCd", "hLrCmCd", "", "", "", "(Select CmName from CustMst where CmCtg= 'C' and CmCd= LhCmCd) "
'  End If

  GRP_REP.Add "Lab Main Cd", Array("LrMCd", "LhMCd"), "", "LrMCd", "hLrMCd", "Lab Sub Cd", "", "", Array("(select PDesc from Param where PTyp= 'LABMCD' and PMCd= LrMCd) ", "(Select PDesc from Param where PTyp= 'LABMCD' and PMCd= LhMCd) ")
  GRP_REP.Add "Lab Sub Cd", Array("LrSCd", "LhSCd"), "", "LrSCd", "hLrSCd", "Lab Main Cd,Customer", "", "", Array("(select PDesc from Param where PTyp= 'LABSCD' and PMCd= LrMCd and PSCd= LrSCd) ", "(Select PDesc from Param where PTyp= 'LABSCD' and PMCd= LhMCd and PSCd= LhSCd) ")
    '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  ' ***** Manali 3.8.0 - EmrFrmMDI.TXT_UCD.text Replaced by gs_UsrCd
  GRP_REP.Add "Customer", Array("LrCmCd", "LhCmCd"), "", "LrCmCd", "hLrCmCd", "Lab Sub Cd", "", "", Array("(select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end ) from CustMst where CmCtg= 'C' and CmCd= LrCmCd) ", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end ) from CustMst where CmCtg= 'C' and CmCd= LhCmCd) ")
  'GRP_REP.Add "Customer", Array("LrCmCd", "LhCmCd"), "", "LrCmCd", "hLrCmCd", "Lab Sub Cd", "", "", Array("(select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end ) from CustMst where CmCtg= 'C' and CmCd= LrCmCd) ", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end ) from CustMst where CmCtg= 'C' and CmCd= LhCmCd) ")
  
  GRP_REP.Add "Currency", "LrCmCurCd", "", "LrCmCurCd", "hLrCmCurCd", "", "", "", "(Select PDesc from Param where PTyp= 'CURNCY' and PMCd= LrCmCurCd and PSCd= '')"
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'LrCmCd, LrCmCurCd, LrMCd, LrSCd, LrQw, LrCstRt, LrSalRt, LrSalMin, LrSalMax, LrFrWt, LrToWt From LabRt where 1= 2
  '*** Report Sql ***
  
  Dim grpflds0 As String, grpflds1 As String, wrepcnd As String, wCnd As String, wSqlStrg As String
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)

  'grpflds = GRP_REP.GrpFldLst
    
  ws_Opt = IIF(adc("UoYyMmFr") = 0, "RepLabRt", "RepLabRtHist")
      
  ' **** As Labour Rate History does not exist in the system from 206 (Done in 210)
  adc("UoCmCtgFr").CmpStr = "LrCmCtg = "
  adc("UoYyMmFr").CmpStr = " "
  adc("UoPmCdFr").CmpStr = "LrMCd >= ": adc("UoPmCdTo").CmpStr = "LrMCd <= ": adc("UoPmCdSel").CmpStr = "LrMCd In "
  adc("UoLSCdFr").CmpStr = "LrSCd >= ": adc("UoLSCdTo").CmpStr = "LrSCd <= "
  adc("UoCmCdFr").CmpStr = "LrCmCd >= ": adc("UoCmCdTo").CmpStr = "LrCmCd <= ": adc("UoCmCdSel").CmpStr = "LrCmCd In "
  'Uni.20-Currency code added
  adc("UoCurCdFr").CmpStr = "LrCmCurCd >= ": adc("UoCurCdTo").CmpStr = "LrCmCurCd <= ": adc("UoCurCdSel").CmpStr = "LrCmCurCd In "

'  If UCase(ws_Opt) = UCase("RepLabRt") Then
'    adc("UoCmCtgFr").CmpStr = "LrCmCtg = "
'    adc("UoYyMmFr").CmpStr = " "
'    adc("UoPmCdFr").CmpStr = "LrMCd >= ": adc("UoPmCdTo").CmpStr = "LrMCd <= ": adc("UoPmCdSel").CmpStr = "LrMCd In "
'    adc("UoLSCdFr").CmpStr = "LrSCd >= ": adc("UoLSCdTo").CmpStr = "LrSCd <= "
'    adc("UoCmCdFr").CmpStr = "LrCmCd >= ": adc("UoCmCdTo").CmpStr = "LrCmCd <= ": adc("UoCmCdSel").CmpStr = "LrCmCd In "
'  ElseIf UCase(ws_Opt) = UCase("RepLabRtHist") Then
'    adc("UoCmCtgFr").CmpStr = "LhCmCtg = "
'    adc("UoYyMmFr").CmpStr = "LhYyyyMm = "
'    adc("UoPmCdFr").CmpStr = "LhMCd >= ": adc("UoPmCdTo").CmpStr = "LhMCd <= ": adc("UoPmCdSel").CmpStr = "LhMCd In "
'    adc("UoLSCdFr").CmpStr = "LhSCd >= ": adc("UoLSCdTo").CmpStr = "LhSCd <= "
'    adc("UoCmCdFr").CmpStr = "LhCmCd >= ": adc("UoCmCdTo").CmpStr = "LhCmCd <= ": adc("UoCmCdSel").CmpStr = "LhCmCd In "
'  End If
  ' **** As Labour Rate History does not exist in the system from 206 (Done in 210)
    
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  
  ' **** As Labour Rate History does not exist in the system from 206 (Done in 210)
  wSqlStrg = "Select " + grpflds0 + ", " + _
             "LrCmCd, LrCmCurCd, LrMCd, LrSCd, LrQw, LrCstRt, LrSalRt, LrSalMin, LrSalMax, LrFrWt, LrToWt From LabRt " + wCnd
  Rep.TxtHead.SetText "Labour Rate Listing"
  Rep.TxtYyyyMm.SetText ""
  Rep.TxtYyyyMmLbl.Suppress = True

'  If UCase(ws_Opt) = UCase("RepLabRt") Then
'    wSqlStrg = "Select " + grpflds0 + ", " + _
'               "LrCmCd, LrMCd, LrSCd, LrQw, LrCstRt, LrSalRt, LrSalMin, LrSalMax From LabRt " + wCnd
'    Rep.TxtHead.SetText "Labour Rate Listing"
'    Rep.TxtYyyyMm.SetText ""
'    Rep.TxtYyyyMmLbl.Suppress = True
'  ElseIf UCase(ws_Opt) = UCase("RepLabRtHist") Then
'    wSqlStrg = "Select " + grpflds1 + ", " + _
'               "LhCmCd, LhMCd, LhSCd, LhQw, LhCstRt, LhSalRt, LhSalMin, LhSalMax From LabRtHist " + wCnd
'    Rep.TxtHead.SetText "Labour Rate History"
'    Rep.TxtYyyyMm.SetText adc("UoYyMmFr")
'    Rep.TxtYyyyMmLbl.Suppress = False
'  End If
  ' **** As Labour Rate History does not exist in the system from 206 (Done in 210)
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt
  
  If UCase(adc("UoCmCtgFr")) = "C" Then
    Rep.TxtHead1.SetText "(Customer)"
  ElseIf UCase(adc("UoCmCtgFr")) = "P" Then
    Rep.TxtHead1.SetText "(Price List)"
  End If
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
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub ADC_Load()
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  Set adc.FirNKeyCtl = adc("UoCmCtgFr")
  Call InitProp(Me)
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_PTyp = "LABMCD"
  gs_CmCtg = ""
  gs_Tbl = "Param"
  gb_ShowYyMmHlp = True     '*** Emr2-06
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  ' **** As Labour Rate History does not exist in the system from 206 (Done in 210)
  LblLhYyyyMm.Visible = False
  adc("UoYyMmFr").Visible = False
  ' **** As Labour Rate History does not exist in the system from 206 (Done in 210)
  
'*** Emr2-06
'  If UCase(adc.MenuCd) = UCase("RepLabRt") Then
'    adc("UoCmCtgFr").CmpStr = "LrCmCtg = "
'    adc("UoYyMmFr").CmpStr = " "
'    adc("UoPmCdFr").CmpStr = "LrMCd >= ": adc("UoPmCdTo").CmpStr = "LrMCd <= ": adc("UoPmCdSel").CmpStr = "LrMCd In "
'    adc("UoLSCdFr").CmpStr = "LrSCd >= ": adc("UoLSCdTo").CmpStr = "LrSCd <= "
'    adc("UoCmCdFr").CmpStr = "LrCmCd >= ": adc("UoCmCdTo").CmpStr = "LrCmCd <= ": adc("UoCmCdSel").CmpStr = "LrCmCd In "
'    adc("UoYyMmFr").Visible = False: adc("UoYyMmFr").Enabled = False
'    LblLhYyyyMm.Visible = False
'  ElseIf UCase(adc.MenuCd) = UCase("RepLabRtHist") Then
'    adc("UoCmCtgFr").CmpStr = "LhCmCtg = "
'    adc("UoYyMmFr").CmpStr = "LhYyyyMm = "
'    adc("UoPmCdFr").CmpStr = "LhMCd >= ": adc("UoPmCdTo").CmpStr = "LhMCd <= ": adc("UoPmCdSel").CmpStr = "LhMCd In "
'    adc("UoLSCdFr").CmpStr = "LhSCd >= ": adc("UoLSCdTo").CmpStr = "LhSCd <= "
'    adc("UoCmCdFr").CmpStr = "LhCmCd >= ": adc("UoCmCdTo").CmpStr = "LhCmCd <= ": adc("UoCmCdSel").CmpStr = "LhCmCd In "
'    adc("UoYyMmFr").Visible = True: adc("UoYyMmFr").Enabled = True
'    LblLhYyyyMm.Visible = True
'  End If
'*** Emr2-06

  Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr"), UCase("UoPMCdTo"), UCase("UoPMCdSel")
    gs_Tbl = "Param"
  Case Is = UCase("UoLSCdFr")
    gs_LabMCd = adc("UoPMCdFr")
  Case Is = UCase("UoLSCdTo")
    gs_LabMCd = adc("UoPMCdTo")
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr"), UCase("UoPMCdTo"), UCase("UoPMCdSel")
    gs_Tbl = "Param"
  Case Is = UCase("UoLSCdFr")
    gs_LabMCd = adc("UoPMCdFr")
  Case Is = UCase("UoLSCdTo")
    gs_LabMCd = adc("UoPMCdTo")
  End Select
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
  ' **** As Labour Rate History does not exist in the system from 206 (Done in 210)
  adc("UoYyMmFr") = 0
End Sub
Private Sub DispCoNm()
  Rep.wCoCd.SetText gs_CoNm
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  Rep.wCoCdLogo.SetText gs_CoNm
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
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set Rep = Nothing
    '*** (09/08/05)
End Sub
