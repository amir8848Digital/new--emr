VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpPrdPts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Production Points Master Listing"
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
      Left            =   10890
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9510
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4470
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4875
      TabIndex        =   18
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   19
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
         TabIndex        =   20
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
      Left            =   6090
      TabIndex        =   17
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   14
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
      TabIndex        =   10
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
      Height          =   9705
      Left            =   0
      TabIndex        =   13
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   12
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpPrdPts.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpPrdPts.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   16
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   11
               Top             =   360
               Width           =   10365
               _ExtentX        =   18283
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   3380
               TabIndex        =   5
               ToolTipText     =   "Enter To Sub Process"
               Top             =   1800
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSUBPRCTO"
               IdName          =   "UOSUBPRCTO"
               CmpStr          =   "PpSubPrc <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   5280
               TabIndex        =   4
               ToolTipText     =   "Enter  Location Code Selection"
               Top             =   1515
               Width           =   8430
               _ExtentX        =   14870
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
               CmpStr          =   "PpLoc in "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1425
               TabIndex        =   8
               ToolTipText     =   "Enter From Sub Process"
               Top             =   1800
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSUBPRCFR"
               IdName          =   "UOSUBPRCFR"
               CmpStr          =   "PpSubPrc >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1425
               TabIndex        =   2
               ToolTipText     =   "Enter From Location Code"
               Top             =   1515
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
               CmpStr          =   "PpLoc >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   3380
               TabIndex        =   3
               ToolTipText     =   "Enter To  Location Code"
               Top             =   1515
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
               CmpStr          =   "PpLoc <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   1425
               TabIndex        =   1
               ToolTipText     =   "Enter The Production Type"
               Top             =   1230
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOPMCD"
               IdName          =   "UOPMCD"
               CmpStr          =   "PpTyp = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   5280
               TabIndex        =   6
               ToolTipText     =   "Enter Sub Process Selection"
               Top             =   1800
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSUBPRCSEL"
               IdName          =   "UOSUBPRCSEL"
               CmpStr          =   "PpSubPrc in "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   315
               Index           =   6
               Left            =   1425
               TabIndex        =   0
               ToolTipText     =   "Company Code"
               Top             =   930
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   556
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               Index           =   6
               Left            =   60
               TabIndex        =   27
               Top             =   960
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Type"
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
               Left            =   60
               TabIndex        =   26
               Top             =   1230
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sub Process"
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
               TabIndex        =   25
               Top             =   1800
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Loc Code"
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
               TabIndex        =   24
               Top             =   1515
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               Left            =   4800
               TabIndex        =   23
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               Left            =   3380
               TabIndex        =   22
               Top             =   480
               Width           =   1275
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   21
               Top             =   480
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpPrdPts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ms_Tbl As String
Dim Rep As New EmrRepPrdPts
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  'Keep This Section
  '##### Groups for Sort Order
  GRP_REP.Add "Loc", "PpLoc", "", "PpLoc", "tLoc", "", "", "", "(Select LocDesc From Loc Where LocCoCd = '" + gs_CoCd + "' And LocCd = PpLoc)"
  GRP_REP.Add "Sub Prc", "PpLoc+PpSubPrc", "PpSubPrc", "PpSubPrc", "tSubPrc", "Loc", "PpLoc", "PpLoc", "(Select PDesc From Param Where PTyp = 'SUBPRC' And PMCd = PpLoc And PSCd = PpSubPrc)"
  GRP_REP.Add "Prd Ctg", "PpPrdCtg", "", "PpPrdCtg", "tPrdCtg", "Set Type,Rm Ctg,Rm SCtg", "", "", "(Select PDesc from Param Where PTyp ='PRDCTG' And PMCd = PpPrdCtg)"
  GRP_REP.Add "Design Cd", "PpDmCd", "", "PpDmCd", "tDmCd", "Set Type,Rm Ctg,Rm SCtg", "", "", "(Select DmDesc from DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' and ", "") + " DmTcTyp ='DM' And DmCd = PpDmCd And DmSz = '')"
  GRP_REP.Add "Set Type", "PpSetTyp", "", "PpSetTyp", "tSetTyp", "Prd Ctg,Design Cd,Rm Ctg,Rm SCtg", "", "", "(Select PDesc from Param Where PTyp ='LABSCD' And PMCd = 'SET' And PSCd = PpSetTyp)"
  GRP_REP.Add "Rm Ctg", "PpRmCtg", "", "PpRmCtg", "tRmCtg", "Prd Ctg,Set Type,Design Cd", "", "", "(Select PDesc from Param Where PTYP ='RMCTG' And PMCD = PpRmCtg)"
  GRP_REP.Add "Rm SCtg", "PpRmSCtg", "", "PpRmSCtg", "tRmSCtg,text3", "Prd Ctg,Set Type,Design Cd,Rm Ctg", "", "", "(Select PDesc from Param Where PTYP ='RMSCTG' And PMCD = PpRmCtg And PSCd = PpRmSCtg)"
End Sub

Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  Dim wMsg As String
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispFlds
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  wMsg = "For Type '" + adc("UoPMCd") + "' Groups "
  Select Case UCase$(adc("UoPMCd"))
    Case UCase$("GN")
      pr_Cancel = InStr(1, UCase$("," + GRP_REP.GrpOpts + ","), "," + UCase$("Set Type") + ",") >= 1 Or _
                  InStr(1, UCase$("," + GRP_REP.GrpOpts + ","), "," + UCase$("Rm Ctg") + ",") >= 1 Or _
                  InStr(1, UCase$("," + GRP_REP.GrpOpts + ","), "," + UCase$("Rm SCtg") + ",") >= 1
      pr_ErrMsg = wMsg + "'Set Type', 'Rm Ctg' & 'Rm SCtg' can not used"
    Case UCase$("GS"), ("GP")
      pr_Cancel = InStr(1, UCase$("," + GRP_REP.GrpOpts + ","), "," + UCase$("Prd Ctg") + ",") >= 1 Or _
                  InStr(1, UCase$("," + GRP_REP.GrpOpts + ","), "," + UCase$("Design Cd") + ",") >= 1 Or _
                  InStr(1, UCase$("," + GRP_REP.GrpOpts + ","), "," + UCase$("Rm Ctg") + ",") >= 1 Or _
                  InStr(1, UCase$("," + GRP_REP.GrpOpts + ","), "," + UCase$("Rm SCtg") + ",") >= 1
      pr_ErrMsg = wMsg + "'Prd Ctg', 'Design Cd', 'Rm Ctg' & 'Rm SCtg' can not used"
    Case UCase$("WS"), ("WP")
      pr_Cancel = InStr(1, UCase$("," + GRP_REP.GrpOpts + ","), "," + UCase$("Prd Ctg") + ",") >= 1 Or _
                  InStr(1, UCase$("," + GRP_REP.GrpOpts + ","), "," + UCase$("Design Cd") + ",") >= 1 Or _
                  InStr(1, UCase$("," + GRP_REP.GrpOpts + ","), "," + UCase$("Set Type") + ",") >= 1
      pr_ErrMsg = wMsg + "'Prd Ctg', 'Design Cd' & 'Set Type' can not used"
  End Select
End Sub

Private Sub adc_setreprecsource()
  '*** Report Sql ***
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  PpTyp, Space(30) as PDesc, PpLoc , PpSubPrc, PpPrdCtg, PpSetTyp,
'  PpDmCd, PpRmCtg, PpRmSCtg, PpFrRmWt, PpToRmWt, PpPts From PrdPts
  '*** Report Sql ***
  
  Dim wrepcnd As String, wSqlStrg As String, grpflds As String
  
'  Set Rep = Nothing
'  Set adc.RepSource = Rep
  Call DispCoNm
  
'  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
  
  wrepcnd = adc.RepCond
  wSqlStrg = "Select " + grpflds + ", PpTyp, (Select PDesc From Param Where PTyp = 'PRDTYP' " + _
              "And PMCD = '" + adc("UoPMCd") + "') As PDesc, " + _
              "PpLoc, PpSubPrc, PpPrdCtg, PpSetTyp, PpDmCd, PpRmCtg, PpRmSCtg, PpFrRmWt, PpToRmWt, PpPts, PpSkl " + _
              "From PrdPts Where " + IIF(gs_Partition = ctCurrPrtn, " PpPrtKey='" + ctCurrPrtn + "' and ", "") + wrepcnd
 
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
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
    Set adc.FirNKeyCtl = adc("UoPMCd")
   
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_PTyp1 = "PRDTYP"
    gs_PTyp = "SUBPRC"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  adc("UOCoCdFr") = gs_CoCd
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case UCase(IdName)
  Case UCase("UoPMCd")
    If Not moCn.RecSeek("Select 'x' from Param where PTyp= 'PRDTYP' " + _
           "and PMCd= '" + pv_NewValue + "' and PSCd= ''") Then Cancel = True: ErrMsg = "Enter Valid Production Type": Exit Sub
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
  Case adc("UoPMCd")
    gs_PTyp1 = "PRDTYP"
  Case UCase("UoIssLocFr"), UCase("UoIssLocTo"), UCase("UoIssLocSel")
    gs_LocTyp = "'R'"
  Case UCase$("UoSubPrcFr"), UCase$("UoSubPrcTo"), UCase$("UoSubPrcSel")
    gs_PTyp = "SUBPRC"
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

Public Sub DispFlds()
  Dim wTcStr As MwfLib.MDORowSet
  Dim wSuppress As Boolean
  
  Set wTcStr = moCn.OpenResultset("Select TsIdName, TsStyle From TcStr " + _
                      "Where TsTbl = 'PrdPts' And TsTyp = '" & adc("UoPMCd") & "'")
  With wTcStr
    If .RecCount >= 0 Then
      .MoveFirst
      While Not .EOF
        wSuppress = IIF(.FldValue("TsStyle") = "I", True, False)
        Select Case UCase$(.FldValue("TsIdName"))
          Case UCase$("PpLoc")
            Rep.tLoc.Suppress = wSuppress
            Rep.PpLoc.Suppress = wSuppress
          Case UCase$("PpSubPrc")
            Rep.tSubPrc.Suppress = wSuppress
            Rep.PpSubPrc.Suppress = wSuppress
          Case UCase$("PpPrdCtg")
            Rep.TPrdCtg.Suppress = wSuppress
            Rep.PpPrdCtg.Suppress = wSuppress
          Case UCase$("PpSetTyp")
            Rep.TSetTyp.Suppress = wSuppress
            Rep.PpSetTyp.Suppress = wSuppress
          Case UCase$("PpDmCd")
            Rep.TDmCd.Suppress = wSuppress
            Rep.PpDmCd.Suppress = wSuppress
          Case UCase$("PpRmCtg")
            Rep.tRmCtg.Suppress = wSuppress
            Rep.PpRmCtg.Suppress = wSuppress
          Case UCase$("PpRmSCtg")
            Rep.tRmSCtg.Suppress = wSuppress
            Rep.PpRmSCtg.Suppress = wSuppress
          Case UCase$("PpFrRmWt")
            Rep.tFrRmWt.Suppress = wSuppress
            Rep.PpFrRmWt.Suppress = wSuppress
          Case UCase$("PpToRmWt")
            Rep.tToRmWt.Suppress = wSuppress
            Rep.PpToRmWt.Suppress = wSuppress
          Case UCase$("PpPts")
            Rep.tPts.Suppress = wSuppress
            Rep.PpPts.Suppress = wSuppress
        End Select
        .MoveNext
      Wend
    End If
  End With
  Set wTcStr = Nothing
End Sub
