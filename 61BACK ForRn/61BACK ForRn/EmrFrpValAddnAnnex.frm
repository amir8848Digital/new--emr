VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.0#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrpValAddnAnnex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Annexure To Value Addition"
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9660
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4410
      TabIndex        =   9
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
      Left            =   5805
      TabIndex        =   19
      Top             =   9660
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   20
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
         TabIndex        =   21
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
      Left            =   7020
      TabIndex        =   18
      Top             =   9675
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   15
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
      TabIndex        =   11
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
      TabIndex        =   14
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   13
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpValAddnAnnex.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpValAddnAnnex.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   17
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   12
               Top             =   360
               Width           =   9435
               _ExtentX        =   16642
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   16
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   20
               Left            =   2700
               TabIndex        =   2
               ToolTipText     =   "Enter Invoice Year"
               Top             =   990
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "IdYy ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   43
               Left            =   1920
               TabIndex        =   1
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   990
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "IdTc ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   45
               Left            =   3120
               TabIndex        =   3
               ToolTipText     =   "Enter Invoice Character"
               Top             =   990
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "IdChr ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   26
               Left            =   3900
               TabIndex        =   4
               ToolTipText     =   "Enter Invoice Number"
               Top             =   990
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "IdNo ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   0
               Left            =   1920
               TabIndex        =   5
               ToolTipText     =   "Enter First Group"
               Top             =   1275
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINGRP"
               IdName          =   "UOINGRP"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   1
               Left            =   1920
               TabIndex        =   6
               ToolTipText     =   "Enter Second Group"
               Top             =   1560
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINGRP1"
               IdName          =   "UOINGRP1"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   2
               Left            =   1920
               TabIndex        =   8
               ToolTipText     =   "Show Running Serial Number or Invoice Serial  (Yes / No)"
               Top             =   2130
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   35
               Left            =   1920
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   705
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "IdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   3
               Left            =   1920
               TabIndex        =   7
               ToolTipText     =   "Enter Detail Level To Be Shown In Report"
               Top             =   1845
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINDET"
               IdName          =   "UOINDET"
            End
            Begin VB.Label LblDet 
               BackStyle       =   0  'Transparent
               Caption         =   "Det Level"
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
               TabIndex        =   30
               Top             =   1845
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
               Index           =   1
               Left            =   60
               TabIndex        =   29
               Top             =   705
               Width           =   1335
            End
            Begin VB.Label LblRunnSr 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Running Sr"
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
               TabIndex        =   28
               Top             =   2130
               Width           =   1845
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Grp Combination2"
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
               TabIndex        =   27
               ToolTipText     =   "Location"
               Top             =   1560
               Width           =   1875
            End
            Begin VB.Label LblGrpCombo 
               BackStyle       =   0  'Transparent
               Caption         =   "Grp Combination1"
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
               TabIndex        =   26
               ToolTipText     =   "Location"
               Top             =   1275
               Width           =   1875
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
               TabIndex        =   25
               Top             =   990
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
               TabIndex        =   24
               Top             =   990
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
               Left            =   60
               TabIndex        =   23
               Top             =   990
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
               TabIndex        =   22
               Top             =   990
               Width           =   105
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpValAddnAnnex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepValAddnAnnex
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  ''GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  ' Zubin 212 (Sales Ctg and qDmSalCtgDesc added)
  'Select 0 as qIrSrNo, IdSr, IdDmCd, IdDmSz, IdQty, IdGldLs, IrRmZWt as qIrWt,
  'IrRmZWt as qGrsWt, IrRmZWt as qNetWt, IrRmZWt as qXWt, IrRmZWt as qDCSWt,
  'RmCtg, RmSCtg, IrRmZCd, RmDesc as qTyp, IrRmQty, IrRmZRt as qSalRt,
  'space(22) as qIrZCdRt, IrRmZVal as qDCSSalVal, IdZFob as qAvgSalPrc,
  'IdZFob as qFOB, IdKt, IdVaCtg, DmCtg, DmSalCtg, DmUom, OdDmCol, OrRmCtg,
  'space(254) as qMainRmCtgDesc, space(254) as qDmCtgDesc, space(254) as qDmSalCtgDesc, space(254) as qDmColDesc,
  'space(254) as qVaCtgDesc, TBOENo, TBOEDt, TzPurRt as qCstRt, IrRmZWt as qApportnWt,0 as qCnt
  'from InvDsg, OrdDsg, DsgMst, InvRm, OrdRm, RmMst, Txn, Txndz where 1= 2
  '*** Report Sql ***

  Dim wrepcnd As String, wCndA As String, wCndB As String, wStr1 As String, wStr2 As String
  Dim wStr3 As String, wStr4 As String, wStr5 As String, wSqlStrg1 As String, wSqlStrg2 As String

  Dim ws_CurCd As String  '*** (Jen 2.13)

  Set Rep = Nothing
  Set ADC.RepSource = Rep
  Call DispCoNm
  ''Call GRP_REP.Gen3LRep
  ''grpflds = GRP_REP.GrpFldLst
  
  
  '*** For the Group Formula of the report
  Dim wGrpArr1() As String, wGrpFormula1 As String, wGrpDescFormula1 As String
  Dim wGrpArr2() As String, wGrpFormula2 As String, wGrpDescFormula2 As String
  Dim i As Integer, qDmCtgDesc As String, qMainRmCtgDesc As String, qDmColDesc As String, qDmKtDesc As String
  Dim qVaCtgDesc As String, qMainMet As String, wOrJoin As String
  
' **** Zubin 211 **** '
  Dim wDetArr() As String, wDetSrt As String
' **** Zubin 211 **** '
  ' Zubin 212
  Dim qDmSalCtgDesc As String
  
  Dim ws_qIrWt As String    '*** (Jen 2.13)
  
  ' **** Manali 3.6.0
  Dim ws_FlagLocalExp As String
  ws_FlagLocalExp = moCn.GetFldVal("Select hExpLocal from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'")
  ' **** Manali 3.6.0
  
  qDmCtgDesc = "'' as qDmCtgDesc": qMainRmCtgDesc = "'' as qMainRmCtgDesc"
  qDmColDesc = "'' as qDmColDesc": qVaCtgDesc = "'' as qVaCtgDesc": qDmKtDesc = "'' as qDmKtDesc"
  qMainMet = "space(2) as OrRmCtg": wOrJoin = ""
  ' Zubin 212
  qDmSalCtgDesc = "'' as qDmSalCtgDesc"
  
  wGrpArr1 = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
         "PMCd= '" + ADC("UoInGrp") + "' "), ",")
  For i = 0 To UBound(wGrpArr1)
    If UCase(wGrpArr1(i)) = "RMCTG" Then
      qMainRmCtgDesc = "(Select PDesc from Param where PTyp= 'RMCTG' and PMCd= max(OrRmCtg)) as qMainRmCtgDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qMainRmCtgDesc})"
      qMainMet = "max(OrRmCtg) as OrRmCtg"
      wOrJoin = " join OrdRm on " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OdPrtKey and ", "") + " OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr and OrMainMet= 'Y' "
    ElseIf UCase(wGrpArr1(i)) = "KT" Then
      qDmKtDesc = "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(IdKt)) as qDmKtDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmKtDesc})"
    ElseIf UCase(wGrpArr1(i)) = "DMCOL" Then
      qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
    ElseIf UCase(wGrpArr1(i)) = "VACTG" Then
      qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(IdVaCtg)) as qVaCtgDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
    ElseIf UCase(wGrpArr1(i)) = "GLDLS" Then
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim(ToText({rdo.IdGldLs}, '#0.00'))+ '%'"
      wGrpFormula1 = wGrpFormula1 + IIF(wGrpFormula1 = "", "", "+") + "ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'") + "}, '#0.00')"
    ElseIf UCase(wGrpArr1(i)) = "PCS/PRS" Then
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "{rdo.DmUom}"
    ElseIf UCase(wGrpArr1(i)) = "DMCTG" Then
      qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
    ' Zubin 212
    ElseIf UCase(wGrpArr1(i)) = "SALCTG" Then
      qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
    ' Zubin 212
    End If
    If wGrpArr1(i) <> "GLDLS" Then wGrpFormula1 = wGrpFormula1 + IIF(wGrpFormula1 = "", "", "+") + "Trim(UpperCase({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'") + "}))"
  Next i
  
  wGrpArr2 = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
         "PMCd= '" + ADC("UoInGrp1") + "' "), ",")
  For i = 0 To UBound(wGrpArr2)
    If UCase(wGrpArr2(i)) = "RMCTG" Then
      qMainRmCtgDesc = "(Select PDesc from Param where PTyp= 'RMCTG' and PMCd= max(OrRmCtg)) as qMainRmCtgDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qMainRmCtgDesc})"
      qMainMet = "max(OrRmCtg) as OrRmCtg"
      wOrJoin = " join OrdRm on " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OdPrtKey and ", "") + " OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr and OrMainMet= 'Y' "
    ElseIf UCase(wGrpArr2(i)) = "KT" Then
      qDmKtDesc = "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(IdKt)) as qDmKtDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmKtDesc})"
    ElseIf UCase(wGrpArr2(i)) = "DMCOL" Then
      qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
    ElseIf UCase(wGrpArr2(i)) = "VACTG" Then
      qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(IdVaCtg)) as qVaCtgDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
    ElseIf UCase(wGrpArr2(i)) = "GLDLS" Then
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim(ToText({rdo.IdGldLs}, '#0.00'))+ '%'"
      wGrpFormula2 = wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + "ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'") + "}, '#0.00')"
    ElseIf UCase(wGrpArr2(i)) = "PCS/PRS" Then
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "{rdo.DmUom}"
    ElseIf UCase(wGrpArr2(i)) = "DMCTG" Then
      qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
    ' Zubin 212
    ElseIf UCase(wGrpArr2(i)) = "SALCTG" Then
      qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
    ' Zubin 212
    End If
    If wGrpArr2(i) <> "GLDLS" Then wGrpFormula2 = wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + "Trim(UpperCase({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'") + "}))"
  Next i
  '*** For the Group Formula of the report
  
' **** Zubin 211 **** '
  wDetArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INDET' and " + _
            "PMCd= '" + ADC("UoInDet") + "' "), ",")
  For i = 0 To UBound(wDetArr)
    ' ***** Manali 3.8.0 - IdTrayNo in Sort
    wDetSrt = wDetSrt + IIF(wDetSrt = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdSr}, '###0')", IIF(UCase(wDetArr(i)) = "IDTRAYNO", "ToText({rdo.IdTrayNo}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))"))
    ' bef 3.8.0
    'wDetSrt = wDetSrt + IIF(wDetSrt = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdSr}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))")
  Next i
  
  If ADC("UOINGRP") = "" Then GRP_REP.SetFormula Rep, "wGrp1", wDetSrt + IIF(wDetSrt <> "", "+", "") + " ToText ({rdo.IdSr}, '###0')"
  If ADC("UOINGRP1") = "" Then GRP_REP.SetFormula Rep, "wGrp2", wDetSrt + IIF(wDetSrt <> "", "+", "") + " ToText ({rdo.IdSr}, '###0')"
  GRP_REP.SetFormula Rep, "wGrp3", wDetSrt + IIF(wDetSrt <> "", "+", "") + " ToText ({rdo.IdSr}, '###0')"
' **** Zubin 211 **** '
  
  ADC("UoCoCdFr").CmpStr = "M.TzCoCd = "
  ADC("UoInTcFr").CmpStr = "M.TzTc= "
  ADC("UoInYyFr").CmpStr = "M.TzYy= "
  ADC("UoInChrFr").CmpStr = "M.TzChr= "
  ADC("UoInNoFr").CmpStr = "M.TzNo= "
  wrepcnd = ADC.RepCond
  wCndA = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  
  '*** Creating Temporary Table #TMPA ***
  '****** Sachin 3.02 - Id fields in Joins
  wStr1 = "Select IDENTITY(INT, 1, 1) as qTzSrNo, max(TBOENo) as TBOENo, " + _
          "max(TBOEDt) as TBOEDt, M.TzRmZCd as TzRmZCd, M.TzPurRt as qCstRt, " + _
          "M.TzRmZCd+ Space(12-Len(M.TzRmZCd))+ Str(IzRmZRt, 10, 3) as qTzZCdRt, " + _
          "IzRmZRt, sum(M.TzRmWt/(case when RmCtg in ('D', 'C') then 5 else 1 end)) as qTzWt " + _
          "into #TMPA From TxndZ M " + _
          "join TxndZ B on B.TzCoCd= M.TzCoCd and B.TzYy= M.TzRefYy and B.TzKey= M.TzRefKey " + IIF(gs_Partition = ctCurrPrtn, " and B.TzPrtKey=M.TzPrtKey ", "") + _
          "and B.TzBM= 'B' join Txn on TIdNo=B.TzTIdNo and TCoCd= B.TzCoCd and TTc= B.TzTc and " + _
          "TYy= B.TzYy and TChr= B.TzChr and TNo= B.TzNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=B.TzPrtKey ", "") + " join InvZRm on " + _
          " IzCoCd= M.TzCoCd and IzTc= M.TzTc and IzYy= M.TzYy and IzChr= M.TzChr and " + _
          "IzNo= M.TzNo and IzSr= M.TzSr " + IIF(gs_Partition = ctCurrPrtn, " and IzPrtKey=M.TzPrtKey ", "") + " join RmMst on RmCd= M.TzRmZCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=M.TzPrtKey ", "") + _
          wCndA + IIF(gs_Partition = ctCurrPrtn, " and M.TzPrtKey='" + ctCurrPrtn + "' ", "") + _
          " Group By TBOENo, M.TzRmZCd, IzRmZRt, M.TzPurRt "
  moCn.CreateTmpTable "#TMPA", wStr1

  ADC("UoCoCdFr").CmpStr = "IdCoCd = "
  ADC("UoInTcFr").CmpStr = "IdTc= "
  ADC("UoInYyFr").CmpStr = "IdYy= "
  ADC("UoInChrFr").CmpStr = "IdChr= "
  ADC("UoInNoFr").CmpStr = "IdNo= "
  wrepcnd = ADC.RepCond
  wCndB = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
 
  '*** (Jen 2.13) ***
  ' ***** Manali 3.03 - 30/06/08 - Pd Mod - 'L' added
  If moCn.GetFldVal("Select hExpLocal from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'") = "Y" Then
    ws_qIrWt = " sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt end) " + _
               "    / (case when Rm.RmCtg in ('D', 'C') then 5 else 1 end) " + _
               "    * (case when Rm.RmCtg in ('G', 'P', 'S', 'L','M') then Rm.RmPurityZ" + _
               "            * (1 + (case when IrMainMet= 'Y' then 0 else IdGldLs end)/ 100) " + _
               "            / (case when Z.RmPurityZ> 0 then Z.RmPurityZ else 1 end) else 1 end) " + _
               "    / (case when (Rm.RmCtg= 'G' and InHGldFiness> 0) then InHGldFiness " + _
               "            when (Rm.RmCtg= 'P' and InHPlFiness> 0) then InHPlFiness " + _
               "            when (Rm.RmCtg= 'S' and InHSlFiness> 0) then InHSlFiness " + _
               "            when (Rm.RmCtg= 'L' and InHPdFiness> 0) then InHPdFiness else 1 end) " + _
               "   ) as qIrWt "
  Else
    ' ***** Manali 3.03 - 30/06/08 - Pd Mod - 'L' added
    ws_qIrWt = " sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt end) " + _
               "    / (case when Rm.RmCtg in ('D', 'C') then 5 else 1 end) " + _
               "    * (case when Rm.RmCtg in ('G', 'P', 'S', 'L','M') then Rm.RmPurityZ" + _
               "            * (1 + (case when IrMainMet= 'Y' then 0 else IdGldLs end)/100) " + _
               "            / (case when Z.RmPurityZ> 0 then Z.RmPurityZ else 1 end) else 1 end) " + _
               "   ) as qIrWt "
  End If
  '*** (Jen 2.13) ***
  
  '*** (Bef 2.13)
  '"sum(IrRmZWt/(case when Rm.RmCtg in ('D', 'C') then 5 else 1 end)* " + _
   "(case when Rm.RmCtg in ('G', 'P', 'S') then Rm.RmPurityZ* (1 + IdGldLs/100) else 1 end)) as qIrWt "
  '*** (Bef 2.13)
  
  '****** (Jen 2.13) InvHd join added ******'
  ' Zubin 212 (DmSalCtg added)
  '****** Sachin 3.02 - Id fields in Join
  ' ***** Manali 3.03 - 30/06/08 - Pd Mod - 'L' added
  ' ***** Manali 3.8.0 - IdTrayNo added in Sort
  wStr2 = "Select IDENTITY(INT, 1, 1) AS qIrSrNo, IdSr, max(IdDmCd) as IdDmCd, " + _
          "max(IdDmSz) as IdDmSz, max(IdQty) as IdQty, max(IdGldLs) as IdGldLs, " + ws_qIrWt + ", " + _
          "sum(IrRmZWt/(case when Rm.RmCtg in ('D', 'C') then 5 else 1 end)) as qGrsWt, " + _
          "sum(IrRmZWt*(case when Rm.RmCtg in ('M', 'G', 'P', 'S', 'L') then 1 else 0 end)) as qNetWt, " + _
          "sum(IrRmZWt*(case when Rm.RmCtg in ('X') then 1 else 0 end)) as qXWt, " + _
          "sum(IrRmZWt*(case when Rm.RmCtg in ('D', 'C') then 1 else 0 end))  as qDCSWt, " + _
          "max(Rm.RmCtg) as RmCtg, max(Rm.RmSCtg) as RmSCtg, IrRmZCd, max(Z.RmDesc) as qTyp, " + _
          "sum(IrRmQty* (case when Rm.RmCtg in ('D', 'C') then 1 else 0 end)) as IrRmQty, IrRmZCd+ Space(12-Len(IrRmZCd))+ Str(IrRmZPureRt, 10, 3) as qIrZCdRt, " + _
          "IrRmZPureRt as qSalRt, " + _
          "sum(IrRmZRt*IrRmZWt*(case when Rm.RmCtg in ('D', 'C') then 1 else 0 end)) as qDCSSalVal, " + _
          "max(case when IdQty> 0 then IdZFob/IdQty else 0 end) as qAvgSalPrc, " + _
          "max(IdZFob) as qFOB, max(IdKt) as IdKt, max(IdVaCtg) as IdVaCtg, " + _
          "max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmUom) as DmUom, max(OdDmCol) as OdDmCol, " + qMainMet + ", max(IdTrayNo) as IdTrayNo " + _
          "into #TMPB " + _
          "From InvRm join RmMst Rm on Rm.RmCd= IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and Rm.RmPrtKey=IrPrtKey ", "") + " join RmMst Z on Z.RmCd= IrRmZCd " + IIF(gs_Partition = ctCurrPrtn, " and Z.RmPrtKey=IrPrtKey ", "") + _
          "join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + IIF(gs_Partition = ctCurrPrtn, " and IdPrtKey=IrPrtKey ", "") + _
          "join InvHd on InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IdPrtKey ", "") + _
          "join OrdDsg on OdIdNo=IdOdIdNo and OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=IdPrtKey ", "") + _
          "join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + wOrJoin + _
          "Where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' and ", "") + " Round(IrRmZWt, 3) > 0 " + wCndB + _
          " Group By IdSr, IrRmZCd, IrRmZPureRt Order By IdSr, IrRmZCd, qSalRt "

  moCn.CreateTmpTable "#TMPB", wStr2
  
  '*** Creating Temporary Table #TMPC ***
  ' Zubin 212 (DmSalCtg added)
  ' ***** Manali 3.8.0 - IdTrayNo added in Sort
  wStr3 = "Select qIrSrNo, IdSr, IdDmCd, IdDmSz, IdQty, IdGldLs, qIrWt, qGrsWt, qNetWt, qXWt, " + _
          "qDCSWt, RmCtg, RmSCtg, IrRmZCd, qTyp, IrRmQty, qIrZCdRt, qSalRt, qDCSSalVal, " + _
          "qAvgSalPrc, qFOB, IdKt, IdVaCtg, DmCtg, DmSalCtg, DmUom, OdDmCol, OrRmCtg, IdTrayNo, " + _
          "qTzSrNo, TBOENo, TBOEDt, qCstRt, qTzWt into #TMPC " + _
          "From #TMPB join #TMPA on qTzZCdRt= qIrZCdRt order By qIrSrNo, qTzWt, qTzSrNo "
  moCn.CreateTmpTable "#TMPC", wStr3

  Call ApportnTbl

  '*** Inserting Recs in Temporary Table #TMPD which is the main Apportioning Sql ******
  ' Zubin 212 (DmSalCtg added)
  ' ***** Manali 3.8.0 - IdTrayNo added in Sort
  wStr4 = "Insert into #TMPD Select qIrSrNo, IdSr, IdDmCd, IdDmSz, IdQty, IdGldLs, qIrWt, " + _
          "qGrsWt, qNetWt, qXWt, qDCSWt, RmCtg, RmSCtg, IrRmZCd, qTyp, IrRmQty, qSalRt, " + _
          "qIrZCdRt, qDCSSalVal, qAvgSalPrc, qFOB, IdKt, IdVaCtg, DmCtg, DmSalCtg, DmUom, OdDmCol, " + _
          "OrRmCtg, IdTrayNo, TBOENo, TBOEDt, qCstRt, " + _
          "qIrWt- (case when IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo< c.qTzSrNo " + _
          "                          and a.qTzZCdRt= c.qIrZCdRt), 0) - " + _
          "                  IsNull((Select sum(qIrWt) from #TMPB b where b.qIrSrNo< c.qIrSrNo " + _
          "                          and b.qIrZCdRt= c.qIrZCdRt), 0) > 0 " + _
          "             then IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo< c.qTzSrNo " + _
          "                          and a.qTzZCdRt= c.qIrZCdRt), 0) - " + _
          "                  IsNull((Select sum(qIrWt) from #TMPB b where b.qIrSrNo< c.qIrSrNo " + _
          "                          and b.qIrZCdRt= c.qIrZCdRt), 0) " + _
          "             else 0 end) as qFinalReq, "
  wStr5 = "qTzWt- (case when IsNull((Select sum(qIrWt) from #TMPB b where b.qIrSrNo< c.qIrSrNo " + _
          "                          and b.qIrZCdRt= c.qIrZCdRt), 0) - " + _
          "                  IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo< c.qTzSrNo " + _
          "                          and a.qTzZCdRt= c.qIrZCdRt), 0) > 0 " + _
          "             then IsNull((Select sum(qIrWt) from #TMPB b where b.qIrSrNo< c.qIrSrNo " + _
          "                          and b.qIrZCdRt= c.qIrZCdRt), 0) - " + _
          "                  IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo< c.qTzSrNo " + _
          "                          and a.qTzZCdRt= c.qIrZCdRt), 0) " + _
          "             else 0 end) as qFinalAvl from #TMPC c "
  moCn.Execute (wStr4 + wStr5)
  
  '*** This part of the Sql is used for group & other fields ***
  ' Zubin 212 (DmSalCtg and qDmSalCtgDesc added)
  ' ***** Manali 3.8.0 - IdTrayNo added in Sort
  wSqlStrg1 = "Select 0 As qIrSrNo, IdSr, max(IdDmCd) as IdDmCd, max(IdDmSz) as IdDmSz, " + _
              "max(IdQty) as IdQty, max(IdGldLs) as IdGldLs, " + _
              "0 as qIrWt, sum(qGrsWt) as qGrsWt, sum(qNetWt) as qNetWt, sum(qXWt) as qXWt, " + _
              "sum(qDCSWt) as qDCSWt, '' as RmCtg, '' as RmSCtg, '' as IrRmZCd, " + _
              "'' as qTyp, sum(IrRmQty) as IrRmQty, 0 as qSalRt, '' as qIrZCdRt, " + _
              "sum(qDCSSalVal) as qDCSSalVal, max(qAvgSalPrc) as qAvgSalPrc, " + _
              "max(qFOB) as qFOB, " + _
              "max(IdKt) as IdKt, max(IdVaCtg) as IdVaCtg, max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, " + _
              "max(DmUom) as DmUom, max(OdDmCol) as OdDmCol, max(OrRmCtg) as OrRmCtg, max(IdTrayNo) as IdTrayNo, " + _
              qMainRmCtgDesc + ", " + qDmCtgDesc + ", " + qDmSalCtgDesc + ", " + qDmColDesc + ", " + qDmKtDesc + ", " + qVaCtgDesc + ", " + _
              "'' as TBOENo, '' as TBOEDt, 0 as qCstRt, 0 as qApportnWt, 0 as qCnt " + _
              "From #TMPB Group By IdSr "
  
  '*** This part of the Sql contains the apportioned weights ***
  ' Zubin 212 (DmSalCtg and qDmSalCtgDesc added)
  ' ***** Manali 3.8.0 - IdTrayNo added in Sort
  wSqlStrg2 = "Select qIrSrNo, IdSr, IdDmCd, IdDmSz, IdQty, IdGldLs, qIrWt, qGrsWt, qNetWt, " + _
             "qXWt, qDCSWt, RmCtg, RmSCtg, IrRmZCd, qTyp, IrRmQty, qSalRt, qIrZCdRt, " + _
             "qDCSSalVal, qAvgSalPrc, qFOB, IdKt, IdVaCtg, DmCtg, DmSalCtg, DmUom, OdDmCol, OrRmCtg, IdTrayNo, " + _
             "'' as qMainRmCtgDesc, '' as qDmCtgDesc, '' as qDmSalCtgDesc, '' as qDmColDesc, '' as qDmKtDesc, '' as qVaCtgDesc, " + _
             "TBOENo, TBOEDt, qCstRt, " + _
              "(case when qFinalReq>0 and qFinalAvl> 0 " + _
              "      then (case when qFinalReq>= qFinalAvl then qFinalAvl else qFinalReq end) " + _
              "      else 0 end) as qApportnWt, (Select Count('x') From #TmpD D1 Where D1.qIrSrNo<>0 And D1.IdSr=D.IdSr And D.qIrZCdRt=D1.qIrZCdRt " + _
              "   And D1.qFinalReq>0 and D1.qFinalAvl> 0 /*(case when D1.qFinalReq>0 and D1.qFinalAvl> 0 " + _
              "      then (case when D1.qFinalReq>= D1.qFinalAvl then D1.qFinalAvl else D1.qFinalReq end) " + _
              "      else 0 end)>0*/) qCnt" + _
              " from #TMPD D Where qFinalReq>0 and qFinalAvl>0 "
  ADC.RepRecSource = wSqlStrg1 + " Union All " + wSqlStrg2 + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  Dim wRsInHd As MDORowSet
  Dim w_InhValBOEDetYN As String
  Set wRsInHd = moCn.OpenRes("Select InTc, InYy, InChr, InNo, InPrnDt, InExpNo, InSelfName, InSelfAdd1, " + _
                "InSelfAdd2, InSelfAdd3, InSelfTel, InSelfFax, InGrNo, InGrDt, InhAuthSgn, InHValBOEDetYN, CmCurCd, InCircularNo, InFob " + _
                "from InvHd join Param on PTyp= 'TC' and PMCd= InTc and PSCd= '' " + _
                "join CustMst on CmCtg= (case PValue when 'IN' then 'C' when 'CTB' then 'T' else '' end) and CmCd= InCmCd " + _
                "where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd ='" + ADC("UoCoCdFr") + "' and InTc= '" + ADC("UoInTcFr") + "' and " + _
                "InYy= '" + ADC("UoInYyFr") + "' and InChr= '" + ADC("UoInChrFr") + "' and " + _
                "InNo= " + CStr(ADC("UoInNoFr")) + " ")
  Dim wrs_CompanyRecord As MDORowSet
  Set wrs_CompanyRecord = moCn.OpenRes("Select CmBuyAdd1, CmBuyAdd2, CmBuyAdd3 from CustMst where CmCtg='Z' and CmCd='" + ADC("UoCoCdFr") + "' ")
  If wrs_CompanyRecord.RecCount > 0 Then
    SetRepText Rep, "TxtRegOff", Trim(wrs_CompanyRecord!CmBuyAdd1) + " " + Trim(wrs_CompanyRecord!CmBuyAdd2) + _
                                 " " + Trim(wrs_CompanyRecord!CmBuyAdd3)
  End If
                
  If wRsInHd.RecCount > 0 Then
'    SetRepText Rep, "TxtRegOff", Trim(wRsInHd!InSelfAdd1) + " " + Trim(wRsInHd!InSelfAdd2) + _
                                 " " + Trim(wRsInHd!InSelfAdd3)
    SetRepText Rep, "TxtFact", Trim(wRsInHd!InSelfAdd1) + " " + Trim(wRsInHd!InSelfAdd2) + _
                                 " " + Trim(wRsInHd!InSelfAdd3)
    SetRepText Rep, "TxtRegOffTel", wRsInHd!InSelfTel
    SetRepText Rep, "TxtRegOffFax", wRsInHd!InSelfFax
    SetRepText Rep, "TxtFactTel", wRsInHd!InSelfTel
    SetRepText Rep, "TxtFactFax", wRsInHd!InSelfFax
    SetRepText Rep, "TxtInvVchNo", Trim(wRsInHd!InTc) + "/" + Trim(wRsInHd!InYy) + "/" + Trim(wRsInHd!InChr) + "/" + Trim(CStr(wRsInHd!InNo))
    SetRepText Rep, "TxtInvDt", Format(wRsInHd!InPrnDt, "dd") + "/" + Format(wRsInHd!InPrnDt, "mm") + "/" + Format(wRsInHd!InPrnDt, "yyyy")
    SetRepText Rep, "TxtInvExpNo", wRsInHd!InExpNo
    
    ' **** Manali 3.6.0 - 17/11/09 - Circular No Printed instead of GR No (InFob < 25000 and DTA )
    SetRepText Rep, "TxtGrNo", IIF(wRsInHd!InFob < 25000 And ws_FlagLocalExp = "Y", wRsInHd!InCircularNo, wRsInHd!InGrNo)
    '''SetRepText Rep, "TxtGrNo", wRsInHd!InGrNo
    Rep.LblGrNo.Suppress = wRsInHd!InFob < 25000 And ws_FlagLocalExp = "Y"
    Rep.LblGrDt.Suppress = wRsInHd!InFob < 25000 And ws_FlagLocalExp = "Y"
    
    If wRsInHd!InGrDt <> "01/01/80" And Not (wRsInHd!InFob < 25000 And ws_FlagLocalExp = "Y") Then _
    SetRepText Rep, "TxtGrDt", Format(wRsInHd!InGrDt, "dd") + "/" + Format(wRsInHd!InGrDt, "mm") + "/" + Format(wRsInHd!InGrDt, "yyyy")
    ' **** Manali 3.6.0 - 17/11/09 - Circular No Printed instead of GR No (InFob < 25000 and DTA )
    
    SetRepText Rep, "TxtAuthSgn", wRsInHd!InhAuthSgn
    w_InhValBOEDetYN = wRsInHd!InHValBOEDetYN
    ws_CurCd = wRsInHd!CmCurCd    '*** (Jen 2.13)
  End If
'???  Rep.wCoCd1.SetText gs_CoNm
  
' ******* Manali - This part is shifted to DispCoNm
''  Dim ws_HName As String
''  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + adc("UoCoCdFr") + "' and " + _
''                                                    "HCd = '" + ctSelfCmCd + "' ")
' ******* Manali - This part is shifted to DispCoNm
  
  With Rep
''    .wCoCd1.SetText ws_HName      ' ******* Manali - This part is shifted to DispCoNm

      
    If w_InhValBOEDetYN = "Y" Then
      .hLaBeDet.Suppress = False
      .hLABENo.Suppress = False
      .FldTBOEDt.Suppress = False
      .FldTBOEno.Suppress = False
    ElseIf w_InhValBOEDetYN = "N" Then
      .hLaBeDet.Suppress = True
      .hLABENo.Suppress = True
      .FldTBOEDt.Suppress = True
      .FldTBOEno.Suppress = True
    End If
  
    '*** (Jen 2.13)
    .hRt.SetText "Rate" + Chr(13) + ws_CurCd
    .hVal.SetText "Value" + Chr(13) + ws_CurCd
    .hAvgPrc.SetText "Avg Prc" + Chr(13) + ws_CurCd
    .hFOBVal.SetText "FOB Value" + Chr(13) + ws_CurCd
    '*** (Jen 2.13)
  
    '****** In the Designer the formula for wGrp1 and wGrp2 are written only for making the
           'group total fields in the designer. The actual formula for the groups are written
           'in this form
           'In the designer the wGrp1 and wGrp2 formula is written as :
           '      Trim({rdo.IdDmCd})+' -'+ToText ({rdo.IdSr}, '###0')
           'This is if the user does not select the second group option or for that
           'matter does not select any group option, the report has to be sorted on the
           'Design Code & IdSr respectively in that order
    If Trim(ADC("UoInGrp")) <> "" Then
      GRP_REP.SetFormula Rep, "wGrp1", wGrpFormula1
      GRP_REP.SetFormula Rep, "wGrp1Desc", wGrpDescFormula1
    End If
    If Trim(ADC("UoInGrp1")) <> "" Then
      GRP_REP.SetFormula Rep, "wGrp2", wGrpFormula2
      GRP_REP.SetFormula Rep, "wGrp2Desc", wGrpDescFormula2
    End If
    If Trim(ADC("UoInGrp")) = "" Then .G1H.Suppress = True: .G1F.Suppress = True
    If Trim(ADC("UoInGrp1")) = "" Then .G2H.Suppress = True: .G2F.Suppress = True
    
    ' **** Zubin 211 **** '
    '''''  If adc("UoYN") = "Y" Then
    If UCase(ADC("UoInDet")) = "INVSR/SR" Or ADC("UoInDet") = "DMCD/SR" Then
    ' **** Zubin 211 **** '
      .FldRunningSrNo.Suppress = False
      .FldIdSr.Suppress = True
    Else
      .FldRunningSrNo.Suppress = True
      .FldIdSr.Suppress = False
    End If
    CRV_REP.DisplayGroupTree = False
    
    Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******
    
    '*** DemarkSec is necessary for the lines to extend till the page footer
  End With
End Sub
Private Sub ApportnTbl()
  Dim wStr1 As String, wStr2 As String
  '*** Creating Temporary Table #TMPD which is the main Apportioning Sql ******
  ' ***** Manali 3.8.0 - IdTrayNo added in Sort
  wStr1 = "Create table #TMPD " + _
          "(  qIrSrNo     integer         null, " + _
          "   IdSr        integer         null, " + _
          "   IdDmCd      varchar(15)     null, " + _
          "   IdDmSz      varchar(5)      null, " + _
          "   IdQty       float           null, " + _
          "   IdGldLs     float           null,  " + _
          "   qIrWt       float           null,  " + _
          "   qGrsWt      float           null,  " + _
          "   qNetWt      float           null,  " + _
          "   qXWt        float           null,  " + _
          "   qDCSWt      float           null,  " + _
          "   RmCtg       varchar(2)      null, " + _
          "   RmSCtg      varchar(5)      null, " + _
          "   IrRmZCd     varchar(12)     null, " + _
          "   qTyp        varchar(50)     null, " + _
          "   IrRmQty     integer         null, " + _
          "   qSalRt      float           null, " + _
          "   qIrZCdRt    varchar(22)     null, " + _
          "   qDCSSalVal  float           null, "
          
          '*** (Bef 2.14)
          '"   qTyp        varchar(30)     null, "
          '*** (Bef 2.14)
          
          '*** (Jen 2.14)
          '"   qTyp        varchar(50)     null, "
          '*** (Jen 2.14)
          
          
  ' Zubin 212 (DmSalCtg added)
  wStr2 = "   qAvgSalPrc  float           null, " + _
          "   qFOB        float           null, " + _
          "   IdKt        varchar(5)      null, " + _
          "   IdVaCtg     varchar(5)      null, " + _
          "   DmCtg       varchar(5)      null, " + _
          "   DmSalCtg    varchar(8)      null, " + _
          "   DmUom       varchar(5)      null, " + _
          "   OdDmCol     varchar(5)      null, " + _
          "   OrRmCtg     varchar(2)      null, " + _
          "   IdTrayNo    int             null, " + _
          "   TBOENo      varchar(20)     null, " + _
          "   TBOEDt      smalldatetime   null, " + _
          "   qCstRt      float           null, " + _
          "   qFinalReq   float           null, " + _
          "   qFinalAvl   float           null " + _
          ")"
   moCn.CreateTmpTable "#TMPD", wStr1 + wStr2
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
    Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
    Set ADC.FirNKeyCtl = ADC("UoInTcFr") 'adc("UoCoCdFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
 ' **** Zubin 211 **** '
    LblRunnSr.Visible = False
    ADC("UOYN").Visible = False
 ' **** Zubin 211 **** '
 
    Call SetGroupSort
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
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
    ADC("UOUSRCD") = ADC.UsrCd
    ADC("UOMNUCD") = ADC.MenuCd
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
 ADC("UoCoCdFr") = gs_CoCd
 
 ' **** Zubin 211 **** '
 If ADC("UoInDet") = "" Then ADC("UoInDet") = "INVSR"
 ' **** Zubin 211 **** '
End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
    Dim ws_HName As String
    ws_HName = moCn.GetFldVal("Select CmName from CustMst where CmCtg='Z' and CmCd='" + ADC("UoCoCdFr") + "' ")
    'ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + adc("UoCoCdFr") + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
    Rep.wCoCd.SetText ws_HName
    Rep.wCoCd1.SetText "For  " + ws_HName
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If ADC.HotKeys(KeyCode, Shift) Then
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
    Cancel = ATXT(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set Rep = Nothing
    '*** (09/08/05)
End Sub
