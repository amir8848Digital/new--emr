VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpCustMst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Master"
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
      Left            =   11100
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4455
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5205
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
      Left            =   6420
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
      Height          =   9705
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
         TabPicture(0)   =   "EmrFrpCustMst.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpCustMst.frx":001C
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
               Width           =   8025
               _ExtentX        =   14155
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
            Begin VB.Frame FraOptBag 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1440
               TabIndex        =   26
               ToolTipText     =   "Show Open Or Closed Or All Bags"
               Top             =   1560
               Width           =   3465
               Begin VB.OptionButton OptBag 
                  Caption         =   "All"
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
                  Index           =   2
                  Left            =   2700
                  TabIndex        =   6
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptBag 
                  Caption         =   "No"
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
                  Index           =   1
                  Left            =   1260
                  TabIndex        =   5
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   945
               End
               Begin VB.OptionButton OptBag 
                  Caption         =   "Yes"
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
                  Index           =   0
                  Left            =   0
                  TabIndex        =   4
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   2910
               TabIndex        =   2
               ToolTipText     =   "Enter To Customer"
               Top             =   1245
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "CmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   4380
               TabIndex        =   3
               ToolTipText     =   "Enter Customer Selection"
               Top             =   1245
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "CmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1455
               TabIndex        =   1
               ToolTipText     =   "Enter From Customer"
               Top             =   1245
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "CmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   1455
               TabIndex        =   0
               ToolTipText     =   "Enter Customer Category"
               Top             =   960
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCD"
               IdName          =   "UOPMCD"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   5160
               TabIndex        =   27
               ToolTipText     =   "Show Open Bags (Yes / No / All)"
               Top             =   1560
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Valid YN"
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
               Left            =   30
               TabIndex        =   25
               Top             =   1560
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Ctg"
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
               Left            =   30
               TabIndex        =   24
               Top             =   960
               Width           =   1425
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
               Left            =   30
               TabIndex        =   23
               Top             =   1245
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
               Left            =   4380
               TabIndex        =   22
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
               Left            =   2910
               TabIndex        =   21
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
               Left            =   1455
               TabIndex        =   20
               Top             =   480
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpCustMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepCustMst
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
'******Geeta*****************Emr206****05/04/04*******
'**********If UserLevel>6 then don't display the Customer Name
' ***** Manali 3.8.0 - EmrFrmMDI.TXT_UCD.text Replaced by gs_UsrCd
  GRP_REP.Add "Customer", "CmCd", "", "CmCd", "hCmCd", "(Detail)", "", "", "case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end"
  'GRP_REP.Add "Customer", "CmCd", "", "CmCd", "hCmCd", "(Detail)", "", "", "case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end"
  
End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  '*** Report Sql ***
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  CmCtg, CmCd, CmName, CmMulBy, CmFixPrc, CmGldAs, CmIWtEqOrd, CmIWtFrOrd, CmIGldRtEqOrd,
'  CmDiaRtFA, CmCSRtFA, CmPrdInst, CmStmpInst, CmPayTerms, CmDelTerms, CmAdd1, CmAdd2,
'  CmAdd3 , CmTel, CmFax, CmEMail, CmCpName, CmPortofDisch, CmFinDest, CmBuyer, CmSalPer,
'  CmBkCd, CmAgCd, CmCurCd, CmValidYN from CustMst where 1= 2
'*** Report Sql ***
   
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  If Not (moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y") Then
    Rep.HdValidCoCd.Suppress = True
  End If
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wStr As String
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  '*** To check if Detail is selected ***
  Dim wGrp(3) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 3
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
  Next i
  '*** If Detail not selected then suppress Section DetSec2 of Detail ***
  If wDetPos = 0 Then Rep.DetSec2.Suppress = True
    
  'If adc.MenuCd = UCase("RepCustLst") Then
  Rep.Text17.Suppress = False
  Rep.Field26.Suppress = False
  If adc("UoPMCd") = "C" Then
    Rep.TxtHead.SetText "Customer Listing"
    wStr = " CmCtg= 'C' "
  ElseIf adc("UoPMCd") = "S" Then
    Rep.TxtHead.SetText "Supplier Listing"
    wStr = " CmCtg= 'S' "
  ElseIf adc("UoPMCd") = "P" Then
    Rep.TxtHead.SetText "Price List Customer Listing"
    wStr = " CmCtg= 'P' "
    Rep.Text17.Suppress = True
    Rep.Field26.Suppress = True
  ElseIf adc("UoPMCd") = "Z" Then
    Rep.TxtHead.SetText "Company Listing"
    wStr = " CmCtg= 'Z' "
    Rep.Text17.Suppress = True
    Rep.Field26.Suppress = True
    Rep.HdValidCoCd.Suppress = True
  ElseIf adc("UoPMCd") = "T" Then
    Rep.TxtHead.SetText "Contractor Listing"
    wStr = " CmCtg= 'T' "
    Rep.HdValidCoCd.Suppress = True
  End If
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")

  '4.1.0.0
  If adc("UOYN") = "Y" Then
    wCnd = wCnd + " and CmValidYN='Y' "
  ElseIf adc("UOYN") = "N" Then
    wCnd = wCnd + " and CmValidYN='N' "
  End If

' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
  'Sachin - 4.1.0.0 - New fields added
  wSqlStrg = " Select " + grpflds + ", " + _
              " CmCtg, CmCd, (case when " + CStr(GetUsrLevel()) + " <= 5 then CmName else CmCd end) as CmName, " + _
              " CmMulBy, CmFixPrc, CmGldAs, CmIWtEqOrd, CmIWtFrOrd, " + _
              " CmIGldRtEqOrd, CmDiaRtFA, CmCSRtFA, CmPrdInst, CmStmpInst, CmPayTerms, " + _
              " CmDelTerms, CmAdd1, CmAdd2, CmAdd3, CmTel, CmFax, CmEMail, CmCpName, " + _
              " CmPortofDisch, CmFinDest, CmBuyer, CmSalPer, CmBkCd, CmAgCd, CmCurCd, " + _
              " CmValidYN, CmGrp, CmTyp, CmValidCoCd from CustMst where " + wStr + wCnd
' ###########################################  Manoj  ###########################################
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added
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
Private Sub OptBag_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN") = "Y"
Case Is = 1
  adc("UoYN") = "N"
Case Is = 2
  adc("UoYN") = ""
End Select
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
    Set adc.FirNKeyCtl = adc("UoPmCd")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_PTyp1 = "CMCTG"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  If adc("UoPmCd") = "C" Then
    gs_CmCtg = "C"
  ElseIf adc("UoPmCd") = "S" Then
    gs_CmCtg = "S"
  ElseIf adc("UoPmCd") = "P" Then
    gs_CmCtg = "P"
  ElseIf adc("UoPmCd") = "Z" Then
    gs_CmCtg = "Z"
  ElseIf adc("UoPmCd") = "T" Then
    gs_CmCtg = "T"
  End If
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  '4.1.0.0
  If adc("UoYN") = "Y" Then
    OptBag(0).Value = True
  ElseIf adc("UoYN") = "N" Then
    OptBag(1).Value = True
  ElseIf adc("UoYN") = "" Then
    OptBag(2).Value = True
  End If
  
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_PTyp1 = "CMCTG"
  
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCd")
    'If Not moCn.RecSeek("Select 'x' from Param where PTyp= 'CMCTG' " + _
    '       "and PMCd= '" + pv_NewValue + "' and PSCd= ''") Then Cancel = True: ErrMsg = "Enter Valid Customer Category": Exit Sub
           
    If pv_NewValue = "C" Then gs_CmCtg = "C"
    If pv_NewValue = "S" Then gs_CmCtg = "S"
    If pv_NewValue = "P" Then gs_CmCtg = "P"
    If pv_NewValue = "Z" Then gs_CmCtg = "Z"
    If pv_NewValue = "T" Then gs_CmCtg = "T"
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_PTyp1 = "CMCTG"
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
