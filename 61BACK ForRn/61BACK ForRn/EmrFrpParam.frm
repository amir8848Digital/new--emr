VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpParam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parameter Listing"
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
      Left            =   13110
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9660
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4440
      TabIndex        =   6
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
      Left            =   7095
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
      Left            =   8310
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
      TabIndex        =   8
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
         TabPicture(0)   =   "EmrFrpParam.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpParam.frx":001C
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
               Left            =   240
               TabIndex        =   9
               Top             =   360
               Width           =   9825
               _ExtentX        =   17330
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
               Index           =   1
               Left            =   2955
               TabIndex        =   5
               ToolTipText     =   "Enter To Parameter Code"
               Top             =   1800
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   4500
               TabIndex        =   3
               ToolTipText     =   "Enter Parameter Type Selection"
               Top             =   1515
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPTYPSEL"
               IdName          =   "UOPTYPSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1425
               TabIndex        =   4
               ToolTipText     =   "Enter From Parameter Code"
               Top             =   1800
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1425
               TabIndex        =   1
               ToolTipText     =   "Enter From Parameter Type"
               Top             =   1515
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPTYPFR"
               IdName          =   "UOPTYPFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   2955
               TabIndex        =   2
               ToolTipText     =   "Enter To Parameter Type"
               Top             =   1515
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPTYPTO"
               IdName          =   "UOPTYPTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   1425
               TabIndex        =   24
               ToolTipText     =   "Enter Company Code"
               Top             =   940
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
               Index           =   4
               Left            =   1425
               TabIndex        =   0
               ToolTipText     =   "Enter The Parameter Definition Type  ('M'icro/ 'S'ystem/ 'O'ther)"
               Top             =   1230
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCD"
               IdName          =   "UOPMCD"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Entry Type"
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
               TabIndex        =   26
               Top             =   1230
               Width           =   1425
            End
            Begin VB.Label LblCoCd 
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
               Left            =   30
               TabIndex        =   25
               Top             =   945
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Main Code"
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
               TabIndex        =   23
               Top             =   1800
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Param Type"
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
               Left            =   30
               TabIndex        =   22
               Top             =   1515
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
               TabIndex        =   21
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
               TabIndex        =   20
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
               TabIndex        =   19
               Top             =   480
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ms_Tbl As String
Dim Rep As New EmrRepParam
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  'Keep This Section
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
'  Select space(1) as pTag, space(20) as qHdPMCd, space(20) as qHdPSCd, space(20) as qHdPDesc,
'  space(20) as qHdPDesc225, space(20) as qHdPValue, space(20) as qHdPNum,
'  space(20) as qHdPValue1, space(20) as qHdPNum1, space(20) as qHdPValue2,space(20) as qHdPValidYn,
'  space(20) as qMskPNum, space(20) as qMskPNum1, PTyp, PMCd, PSCd, PDesc, PDesc225,
'  PValue, PNum, PValue1, PNum1, PValue2, PValidYN From Param where 1= 2
  '*** Report Sql ***
  
  Dim wrepcnd As String, wCnd As String, wSqlStrg As String, wsql1 As String, wsql2 As String, wMask As String
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  'Call GRP_REP.Gen3LRep
  'grpflds = GRP_REP.GrpFldLst

  Dim ws_PTyp As String, ws_PMCd As String, ws_PSCd As String
  Dim ws_PDesc As String, ws_PDesc225 As String, ws_PValue As String, ws_PNum As String
  Dim ws_PValue1 As String, ws_PNum1 As String, ws_PValue2 As String, ws_PValue3 As String
  Dim ws_Join As String
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:  21/01/2004
  Dim ws_PValidYN As String

  If UCase(ms_Tbl) = UCase("vParam") Then
    ws_PTyp = "vPTyp": ws_PMCd = "vPMCd": ws_PSCd = "vPSCd": ws_PDesc = "vPDesc"
    ws_PValue = "vPValue": ws_PDesc225 = "vPDesc225": ws_PNum = "vPNum"
    ws_PValue1 = "vPValue1": ws_PNum1 = "vPNum1": ws_PValue2 = "vPValue2": ws_PValue3 = "vPValue3"
    ws_PValidYN = "vPValidYN"
  ElseIf UCase(ms_Tbl) = UCase("Param") Then
    ws_PTyp = "PTyp": ws_PMCd = "PMCd": ws_PSCd = "PSCd": ws_PDesc = "PDesc"
    ws_PValue = "PValue": ws_PDesc225 = "PDesc225": ws_PNum = "PNum":
    ws_PValue1 = "PValue1": ws_PNum1 = "PNum1": ws_PValue2 = "PValue2": ws_PValue3 = "PValue3"
    ws_PValidYN = "PValidYN"
  End If
' ###########################################  Manoj  ###########################################
  
  adc("UoPTypFr").CmpStr = " P1." + ws_PTyp + " >= "
  adc("UoPTypTo").CmpStr = " P1." + ws_PTyp + " <= "
  adc("UoPTypSel").CmpStr = " P1." + ws_PTyp + " In "
  adc("UoPMCdFr").CmpStr = "P1." + ws_PMCd + " >= "
  adc("UoPMCdTo").CmpStr = "P1." + ws_PMCd + " <="
 'adc("UoPMCdTo").CmpStr = "P1." + ws_PMCd + " In "

  wrepcnd = adc.RepCond
  If ms_Tbl = "Param" Then
    ws_Join = " "
    wCnd = IIF(wrepcnd <> "", " where ", "") + wrepcnd
  ElseIf ms_Tbl = "vParam" Then
    ws_Join = " and P2.vPCoCd= '" + ctSelfCoCd + "' "
    wCnd = "Where P1.vPCoCd = '" + adc("UoCoCdFr") + "'"
    wCnd = wCnd + IIF(wrepcnd <> "", " and ", "") + wrepcnd
  End If
  
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:  21/01/2004
  wsql1 = "Select 'P' as pTag, '' as qHdPMCd, '' as qHdPSCd, '' as qHdPDesc, " + _
          "'' as qHdPDesc225, '' as qHdPValue, '' as qHdPNum, '' as qHdPValue1, " + _
          "'' as qHdPNum1, '' as qHdPValue2, '' as qHdPValidYN, '' as qMskPNum, '' as qMskPNum1, " + _
          "Upper(P1." + ws_PTyp + ") as PTyp , P1." + ws_PMCd + "  as PMCd , P1." + ws_PSCd + " as PSCd , " + _
          "P1." + ws_PDesc + " as PDesc , P1." + ws_PDesc225 + " as PDesc225, P1." + ws_PValue + " as PValue , " + _
          "P1." + ws_PNum + " as PNum , P1." + ws_PValue1 + " as PValue1 , P1." + ws_PNum1 + " as PNum1 , P1." + ws_PValue2 + " as PValue2 , " + _
          "P1." + ws_PValidYN + " as PValidYN " + _
          "From " + ms_Tbl + " P1 join " + ms_Tbl + " P2 on P2." + ws_PTyp + "= 'PTYP' " + ws_Join + " and " + _
          "P2." + ws_PMCd + "= P1." + ws_PTyp + " " + _
          "and P2." + ws_PValue + "= '" + gs_ParamOpt + "' " + wCnd
' ###########################################  Manoj  ###########################################
  
  adc("UoPTypFr").CmpStr = " " + ws_PMCd + "  >= "
  adc("UoPTypTo").CmpStr = " " + ws_PMCd + "  <= "
  adc("UoPTypSel").CmpStr = " " + ws_PMCd + "  In "
  adc("UoPMCdFr").CmpStr = ""
  adc("UoPMCdTo").CmpStr = ""
  
  wrepcnd = adc.RepCond
  If ms_Tbl = "Param" Then
    wCnd = IIF(wrepcnd <> "", " and ", "") + wrepcnd
  ElseIf ms_Tbl = "vParam" Then
    wCnd = "and vPCoCd = '" + ctSelfCoCd + "'"
    wCnd = wCnd + IIF(wrepcnd <> "", " and ", "") + wrepcnd
  End If
 
  '*** wMask might be helpful in setting the mask for the PNum and PNum1 fields ***
  'wMask = " '1'+ (case when CHARINDEX('.', tsmask, 1)> 0 then '.' else '' end)+ " + _
          "replicate(0, case when CHARINDEX('.', tsmask, 1)> 0 then " + _
          "len(tsmask)-CHARINDEX('.', tsmask, 1) else 0 end) "
  
  wMask = " (case when CHARINDEX('.', tsmask, 1)> 0 then " + _
          "len(tsmask)-CHARINDEX('.', tsmask, 1) else 0 end) "
  
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:  21/01/2004
  wsql2 = " Union All " + _
         "Select 'H' as pTag, " + _
         "(Select TsLbl From TcStr where TsTbl= '" + ms_Tbl + "' and TsTyp= " + ws_PMCd + "  and TsIdName= '" + ws_PMCd + "') as qHdPMCd, " + _
         "(Select TsLbl From TcStr where TsTbl= '" + ms_Tbl + "' and TsTyp= " + ws_PMCd + "  and TsIdName= '" + ws_PSCd + "') as qHdPSCd, " + _
         "(Select TsLbl From TcStr where TsTbl= '" + ms_Tbl + "' and TsTyp= " + ws_PMCd + "  and TsIdName= '" + ws_PDesc + "') as qHdPDesc, " + _
         "(Select TsLbl From TcStr where TsTbl= '" + ms_Tbl + "' and TsTyp= " + ws_PMCd + "  and TsIdName= '" + ws_PDesc225 + "') as qHdPDesc225, " + _
         "(Select TsLbl From TcStr where TsTbl= '" + ms_Tbl + "' and TsTyp= " + ws_PMCd + "  and TsIdName= '" + ws_PValue + "') as qHdPValue, " + _
         "(Select TsLbl From TcStr where TsTbl= '" + ms_Tbl + "' and TsTyp= " + ws_PMCd + "  and TsIdName= '" + ws_PNum + "') as qHdPNum, " + _
         "(Select TsLbl From TcStr where TsTbl= '" + ms_Tbl + "' and TsTyp= " + ws_PMCd + "  and TsIdName= '" + ws_PValue1 + "' ) as qHdPValue1, " + _
         "(Select TsLbl From TcStr where TsTbl= '" + ms_Tbl + "' and TsTyp= " + ws_PMCd + "  and TsIdName= '" + ws_PNum1 + "') as qHdPNum1, " + _
         "(Select TsLbl From TcStr where TsTbl= '" + ms_Tbl + "' and TsTyp= " + ws_PMCd + "  and TsIdName= '" + ws_PValue2 + "') as qHdPValue2, " + _
         "(Select TsLbl From TcStr where TsTbl= '" + ms_Tbl + "' and TsTyp= " + ws_PMCd + "  and TsIdName= '" + ws_PValidYN + "') as qHdPValidYN, " + _
         "(Select " + wMask + " From TcStr where TsTbl= '" + ms_Tbl + "' and TsTyp= " + ws_PMCd + "  and TsIdName= '" + ws_PNum + "') as qMskPNum, " + _
         "(Select " + wMask + " From TcStr where TsTbl= '" + ms_Tbl + "' and TsTyp= " + ws_PMCd + "  and TsIdName= '" + ws_PNum1 + "') as qMskPNum1, " + _
         "Upper(" + ws_PMCd + " ) as PTyp , '' as PMCd , '' as PSCd , " + _
         "" + ws_PDesc + " as PDesc , '' as PDesc225 , " + ws_PValue + " as PValue , 0 as PNum , '' as PValue1 , " + _
         "0 as PNum1 , '' as PValue2 , '' as PValidYN " + _
         "from " + ms_Tbl + " where " + ws_PTyp + " = 'PTYP' and " + ws_PValue + " = '" + gs_ParamOpt + "' " + _
         wCnd
' ###########################################  Manoj  ###########################################
    
  wSqlStrg = wsql1 + wsql2
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt
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
    
    'If UCase(adc.MenuCd) = UCase("RepVarSysPrm") Or UCase(adc.MenuCd) = UCase("RepVarOthPrm") Or UCase(adc.MenuCd) = UCase("RepVarMicPrm") Then
    If UCase(adc.MenuCd) = UCase("RepVarPrm") Then
      ms_Tbl = "vParam"
      gs_Tbl = "vParam"
    Else
      ms_Tbl = "Param"
      gs_Tbl = "Param"
      adc("UoCoCdFr").Visible = False
      adc("UoCoCdFr").Enabled = False
      LblCoCd.Visible = False
    End If
    
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_PTyp1 = "SYSDFN"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  adc("UoCoCdFr") = gs_CoCd
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  If adc("UoPMCd") = "Y" Then
    gs_ParamOpt = "Y"
  ElseIf adc("UoPMCd") = "N" Then
    gs_ParamOpt = "N"
  ElseIf adc("UoPMCd") = "M" Then
    gs_ParamOpt = "M"
  End If
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_PTyp1 = "SYSDFN"
  
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCd")
    'If Not moCn.RecSeek("Select 'x' from Param where PTyp= 'SYSDFN' " + _
    '       "and PMCd= '" + pv_NewValue + "' and PSCd= ''") Then Cancel = True: ErrMsg = "Enter Valid Entry Type": Exit Sub
           
    If pv_NewValue = "Y" Then gs_ParamOpt = "Y"
    If pv_NewValue = "N" Then gs_ParamOpt = "N"
    If pv_NewValue = "M" Then gs_ParamOpt = "M"
  End Select
  
  'If UCase(adc.MenuCd) = UCase("RepVarSysPrm") Or UCase(adc.MenuCd) = UCase("RepVarOthPrm") Or UCase(adc.MenuCd) = UCase("RepVarMicPrm") Then
  If UCase(adc.MenuCd) = UCase("RepVarPrm") Then
    gs_Tbl = "vParam"
  Else
    gs_Tbl = "Param"
  End If
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  'If UCase(adc.MenuCd) = UCase("RepVarSysPrm") Or UCase(adc.MenuCd) = UCase("RepVarOthPrm") Or UCase(adc.MenuCd) = UCase("RepVarMicPrm") Then
  If UCase(adc.MenuCd) = UCase("RepVarPrm") Then
    gs_Tbl = "vParam"
  Else
    gs_Tbl = "Param"
  End If
  
  Select Case UCase(IdName)
  Case Is = adc("UoPMCd")
    gs_PTyp1 = "SYSDFN"
  Case Is = UCase("UoPMCdFr")
    gs_PTyp = adc("UoPTypFr")
  Case Is = UCase("UoPMCdTo")
    gs_PTyp = adc("UoPTypTo")
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
  If ms_Tbl = "Param" Then
    Rep.wCoCd.SetText gs_CoNm
    Rep.wCoCdLogo.SetText gs_CoNm
  ElseIf ms_Tbl = "vParam" Then
    Dim ws_HName As String
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                                                        "HCd = '" + ctSelfCmCd + "' ")
    Rep.wCoCd.SetText ws_HName
    Rep.wCoCdLogo.SetText ws_HName
  End If
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
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
