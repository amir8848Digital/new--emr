VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpShpBillFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice"
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   22
      Top             =   9570
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   23
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
         TabIndex        =   24
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
      TabIndex        =   21
      Top             =   9585
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   18
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
      TabIndex        =   14
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
      TabIndex        =   17
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   16
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpShpBillFile.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpShpBillFile.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   20
            Top             =   360
            Width           =   15045
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   15
               Top             =   360
               Width           =   8745
               _ExtentX        =   15425
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9015
            Index           =   0
            Left            =   90
            TabIndex        =   19
            Top             =   390
            Width           =   15045
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   2295
               TabIndex        =   5
               ToolTipText     =   "Enter From Export Number"
               Top             =   1290
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOFR"
               IdName          =   "UOINEXPNOFR"
               CmpStr          =   "InExpNo ="
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Unit Price          "
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
               Height          =   270
               Index           =   0
               Left            =   120
               TabIndex        =   11
               Tag             =   "ShowUnitPrc"
               ToolTipText     =   "Check To Display Unit Price"
               Top             =   2805
               Visible         =   0   'False
               Width           =   2370
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   2295
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
               CmpStr          =   "InCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   315
               Index           =   14
               Left            =   2295
               TabIndex        =   6
               ToolTipText     =   "Enter Export Numbers"
               Top             =   1290
               Width           =   6825
               _ExtentX        =   12039
               _ExtentY        =   556
               MaxLength       =   285
               DataField       =   "UOINEXPNOSEL"
               IdName          =   "UOINEXPNOSEL"
               CmpStr          =   "InExpNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   2295
               TabIndex        =   7
               Tag             =   "WICC"
               ToolTipText     =   "Enter Consignee Country"
               Top             =   1605
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOSBICC"
               IdName          =   "UOSBICC"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   2295
               TabIndex        =   8
               Tag             =   "WIPOD"
               ToolTipText     =   "Enter From Export Number"
               Top             =   1905
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               MaxLength       =   6
               DataType        =   4
               DataField       =   "UOSBPOD"
               IdName          =   "UOSBPOD"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   2295
               TabIndex        =   9
               Tag             =   "WICOD"
               ToolTipText     =   "Enter Country of Destination"
               Top             =   2190
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOSBCOD"
               IdName          =   "UOSBCOD"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   2295
               TabIndex        =   10
               Tag             =   "WIAF"
               ToolTipText     =   "Enter Ad Freight"
               Top             =   2475
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOSBAF"
               IdName          =   "UOSBAF"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   2955
               TabIndex        =   30
               ToolTipText     =   "Show Invoice Or Order Customer ('Y' for Invoice / 'N' For Order)"
               Top             =   2745
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
               Index           =   26
               Left            =   4290
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
               CmpStr          =   "InNo ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   3090
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
               CmpStr          =   "InYy ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   2295
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
               CmpStr          =   "InTc ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3510
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
               CmpStr          =   "InChr ="
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
               Left            =   4215
               TabIndex        =   35
               Top             =   990
               Width           =   105
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
               Left            =   180
               TabIndex        =   34
               Top             =   1335
               Width           =   1845
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
               Left            =   180
               TabIndex        =   33
               Top             =   1050
               Width           =   1845
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
               Left            =   3450
               TabIndex        =   32
               Top             =   990
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
               Left            =   3030
               TabIndex        =   31
               Top             =   990
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Consignee Country"
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
               Left            =   165
               TabIndex        =   29
               Top             =   1620
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Port Of Dest"
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
               Index           =   7
               Left            =   165
               TabIndex        =   28
               Top             =   1905
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Country Of Dest"
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
               Index           =   8
               Left            =   165
               TabIndex        =   27
               Top             =   2190
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Add Freight"
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
               Index           =   13
               Left            =   165
               TabIndex        =   26
               Top             =   2475
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
               TabIndex        =   25
               Top             =   720
               Width           =   1845
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpShpBillFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rep As New EmrRepDummy

Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Dim mb_MultiInv As Boolean    '*** (Jen 2.13)
Dim ms_ErrMsg As String

Private Sub SetGroupSort()
  ''GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub
  
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  
Set Rep = Nothing
Set adc.RepSource = Rep

ms_ErrMsg = ""
Call Inserts
adc.RepRecSource = "Select 'x' from Param where 1= 2"

If ms_ErrMsg <> "" Then
  DispMsg ms_ErrMsg, etError
End If

Exit Sub
RepErr:
  DispMsg ms_ErrMsg, etError
End Sub
Private Sub Inserts()
On Error GoTo ErrHldIns
Dim wCnd As String
Dim wInExpNo As String
Dim ws_MinExpNo As String, ws_MinCnd As String
Dim ws_CurCd As String
Dim ws_CoUsrMnuStr As String, ws_CoUsrMnuFmt As String, ws_CoUsrMnuOut As String

ws_CoUsrMnuStr = "'" + UCase(adc("UoCoCdFr")) + "', '" + UCase(adc.UsrCd) + "', '" + UCase(adc.MenuCd) + "' "
ws_CoUsrMnuOut = " qCoCd='" + UCase(adc("UoCoCdFr")) + "' and qUsrCd='" + UCase(adc.UsrCd) + "' and qMnuCd='" + UCase(adc.MenuCd) + "' "
ws_CoUsrMnuFmt = " CoCd='" + UCase(adc("UoCoCdFr")) + "' and UsrCd='" + UCase(adc.UsrCd) + "' and MnuCd='" + UCase(adc.MenuCd) + "' "

If mb_MultiInv Then
  If adc("UoInExpNoSel") = "" Then ms_ErrMsg = "Export No Selection Cannot Be Blank": Exit Sub
Else
    If (adc("UoInTcFr") = "" Or adc("UoInYyFr") = "" Or adc("UoInChrFr") = "" Or adc("UoInNoFr") = 0) _
        And adc("UoInExpNoFr") = "" Then
       ms_ErrMsg = "Enter Voucher Tc/YY/Chr/No or Exp No": Exit Sub
    End If
End If

'wInExpNo = "EX-174/09-10"
wCnd = adc.RepCond
'wCnd = " inExpNo = '" + wInExpNo + "' "

ws_MinExpNo = moCn.GetFldVal("Select min(InExpNo) from InvHd Where " + wCnd)
ws_MinCnd = " InCoCd='" + adc("UoCoCdFr") + "' and InExpNo = '" + ws_MinExpNo + "' "

If Not moCn.RecSeek("Select 'x' from InvHd Where " + wCnd) Then _
  ms_ErrMsg = "Invoice Not Found, Check Export No(s)": Exit Sub
  
If Not moCn.RecSeek("Select 'x' From InvDsg " + _
                    "Join InvHd On InCoCd=IdCoCd And InTc=IdTc And InYy=IdYy And InChr=IdChr And InNo=IdNo " + _
                    "Where " + wCnd + " ") Then _
  ms_ErrMsg = "Invoice Design Not Generated": Exit Sub
  

moCn.Execute ("Delete From SBOutput Where " + ws_CoUsrMnuOut + " ")
moCn.Execute ("Delete From ShpBillFmt Where " + ws_CoUsrMnuFmt + "  ")
moCn.Execute ("Insert Into ShpBillFmt Select " + ws_CoUsrMnuStr + ", SbSTbl, SbsSr, SbsFldTyp, SbsFldNm, SbsScpFld from ShpBillStr ")

Dim wsUpdStr As String
wsUpdStr = "Update ShpBillFmt " + _
          "Set FldName = (Case When (UPPER(ScpFld)='UNITCODE') Then ''''+Cast(HUnitCd as varChar(20))+'''' " + _
                     " When (UPPER(ScpFld)='CUSTOMHOUSECODE') Then ''''+Cast(HCustomHouseCd as varChar(20))+'''' " + _
                     " When (UPPER(ScpFld)='CHALICENSECD') Then ''''+Cast(HCHALicenseCd as varChar(20))+'''' " + _
                     " When (UPPER(ScpFld)='STATEOFORIGINEXP') Then ''''+Cast(InHStatOfOrg as varChar(2))+'''' " + _
                     " When (UPPER(ScpFld)='TYPEOFEXPORTER') Then ''''+Cast(HTypOfExporter as varChar(20))+'''' " + _
                     " When (UPPER(ScpFld)='CONSIGNEECOUNTRY') Then ''''+Cast('" + adc("UoSbICC") + "' as varChar(2))+'''' " + _
                     " When (UPPER(ScpFld)='PORTOFLOADING') Then ''''+Cast(HPoL as varChar(20))+'''' " + _
                     " When (UPPER(ScpFld)='PORTOFDESTINATION') Then ''''+Cast('" + adc("UoSbPOD") + "' as varChar(6))+'''' " + _
                     " When (UPPER(ScpFld)='COUNTRYOFDESTINATION') Then ''''+Cast('" + adc("UoSbCOD") + "' as varChar(2))+'''' " + _
                     " When (UPPER(ScpFld)='SBUOM') Then ''''+Cast(HSbUom as varChar(20))+'''' " + _
                     " When (UPPER(ScpFld)='CURRENCY') Then ''''+Cast(Cur.PValue2 as varChar(5))+'''' " + _
                     " When (UPPER(ScpFld)='NATUREOFCONTRACT') Then ''''+Cast(Noc.PValue as varChar(5))+'''' " + _
                     " When (UPPER(ScpFld)='ADDFREIGHT') Then ''''+Cast('" + adc("UoSbAF") + "' as varChar(20))+'''' " + _
                     " When (UPPER(ScpFld)='TOTALNOOFPACKAGES') Then ''''+Cast(HTotNoOfPack as varChar(20))+'''' " + _
                     " Else FldName End) " + _
                     " From ShpBillFmt Join Head On HCoCd=CoCd and HCd='" + ctSelfCmCd + "' " + _
                     " Join InvHd on HCoCd=InCoCd and InExpNo = '" + ws_MinExpNo + "' " + _
                     " Join CustMst On CmCtg='C' and CmCd=InCmCd " + _
                     " Join Param Cur On Cur.PTyp='CURNCY' and Cur.PMCd=CmCurCd " + _
                     " Join Param Noc On Noc.PTyp='INVTOTAS' and Noc.PMCd=InTotAs " + _
                     " Where " + ws_CoUsrMnuFmt + " "
                        
moCn.Execute (wsUpdStr)

'moCn.Execute ("Update ShpBillFmt " + _
             "Set FldName = (Case When (UPPER(ScpFld)='UNITCODE') Then '''" + ADC("UOXTANA01FR") + "''' " + _
                        " When (UPPER(ScpFld)='CUSTOMHOUSECODE') Then '''" + ADC("UOXTANA01TO") + "''' " + _
                        " When (UPPER(ScpFld)='CHALICENSECD') Then '''" + ADC("UOSETCDFR") + "''' " + _
                        " When (UPPER(ScpFld)='STATEOFORIGINEXP') Then '''" + ADC("UOXTANA02TO") + "''' " + _
                        " When (UPPER(ScpFld)='TYPEOFEXPORTER') Then '''" + ADC("UOXTANA02FR") + "''' " + _
                        " When (UPPER(ScpFld)='CONSIGNEECOUNTRY') Then '''" + ADC("UOXTANA03FR") + "''' " + _
                        " When (UPPER(ScpFld)='PORTOFLOADING') Then '''" + ADC("UOXTANA03TO") + "''' " + _
                        " When (UPPER(ScpFld)='PORTOFDESTINATION') Then '''" + ADC("UOXTANA04FR") + "''' " + _
                        " When (UPPER(ScpFld)='COUNTRYOFDESTINATION') Then '''" + ADC("UOXTANA04TO") + "''' " + _
                        " When (UPPER(ScpFld)='SBUOM') Then '''" + ADC("UOXTANA05FR") + "''' " + _
                        " When (UPPER(ScpFld)='CURRENCY') Then '''" + ADC("UOXTANA05TO") + "''' " + _
                        " When (UPPER(ScpFld)='NATUREOFCONTRACT') Then '''" + ADC("UOXTANA06FR") + "''' " + _
                        " When (UPPER(ScpFld)='ADDFREIGHT') Then '''" + ADC("UOXTANA06TO") + "''' " + _
                        " When (UPPER(ScpFld)='CURRENCY') Then '''" + ADC("UOXTANA05TO") + "''' " + _
                        " When (UPPER(ScpFld)='TOTALNOOFPACKAGES') Then '''" + CStr(ADC("UOTRNOFR")) + "''' " + _
                        " When TblNm='ITEM' and SrNo=16 and '" + ADC("UoYn") + "'='N' Then '''''' " + _
                        " Else FldName End) " + _
                        " From ShpBillFmt  ")

Dim ws_HdStr As String, ws_SBStr As String, ws_INVStr As String, ws_ITMStr As String, ws_FTStr As String
Dim ws_Dlmtr As String

ws_Dlmtr = "'" + Chr(29) + "'"

Dim wRs_ShpBill As MwfLib.MDORowSet
Set wRs_ShpBill = moCn.OpenRes("Select * from ShpBillFmt Where " + ws_CoUsrMnuFmt + "  Order By TblNm, SrNo ")

With wRs_ShpBill
  .MoveFirst
  Do While Not (.EOF Or .BOF)
    If UCase(!TblNm) = UCase("HEAD") Then
      ws_HdStr = ws_HdStr + !FldName + IIF(Trim(!FldName) = "", "", "+") + ws_Dlmtr + "+"
    End If
  
    If UCase(!TblNm) = UCase("SB") Then
      ws_SBStr = ws_SBStr + !FldName + IIF(Trim(!FldName) = "", "", "+") + ws_Dlmtr + "+"
    End If

    If UCase(!TblNm) = UCase("INV") Then
      ws_INVStr = ws_INVStr + !FldName + IIF(Trim(!FldName) = "", "", "+") + ws_Dlmtr + "+"
    End If

    If UCase(!TblNm) = UCase("ITEM") Then
      ws_ITMStr = ws_ITMStr + !FldName + IIF(Trim(!FldName) = "", "", "+") + ws_Dlmtr + "+"
    End If
  
    If UCase(!TblNm) = UCase("FOOTER") Then
      ws_FTStr = ws_FTStr + !FldName + IIF(Trim(!FldName) = "", "", "+") + ws_Dlmtr + "+"
    End If
  
    .MoveNext
  Loop
End With

ws_HdStr = Left(ws_HdStr, Len(ws_HdStr) - 5)
ws_SBStr = Left(ws_SBStr, Len(ws_SBStr) - 5)
ws_INVStr = Left(ws_INVStr, Len(ws_INVStr) - 5)
ws_ITMStr = Left(ws_ITMStr, Len(ws_ITMStr) - 5)
ws_FTStr = Left(ws_FTStr, Len(ws_FTStr) - 5)

moCn.Execute ("Insert Into SBOutput " + _
              "Select " + ws_CoUsrMnuStr + ", 'HEAD',1,'" + wInExpNo + "',0 , " + ws_HdStr + " " + _
              "From InvHD Where " + ws_MinCnd)

If Not moCn.RecSeek("Select 'x' from SbOutPut Where qTblNm='HEAD' and " + ws_CoUsrMnuOut + " ") Then
  ms_ErrMsg = "'HEAD' Section Not Generated in SB File"
  GoTo ErrHldIns
  Exit Sub
End If

moCn.Execute ("Insert Into SBOutput Values (" + ws_CoUsrMnuStr + ", 'TAG', 2,'',0,'<TABLE>SB') ")

moCn.Execute ("Insert Into SBOutput " + _
              "Select " + ws_CoUsrMnuStr + ", 'SB',3,'',0, " + ws_SBStr + " From InvDsg " + _
              " Join InvHd On InCoCd=IdCoCd And InTc=IdTc And InYy=IdYy And InChr=IdChr And InNo=IdNo " + _
              " Join CustMst On CmCtg='C' And CmCd=InCmCd " + _
              " Join Head On HCd='" + ctSelfCmCd + "' And HCoCd=InCoCd " + _
              " Join (Select IrCoCd, IrIdIdNo, Sum(IrRmZWt/Case When IrRmCtg In ('D','C') Then 5 Else 1 End) as qNetWt " + _
              "From InvRm Group By IrCoCd,IrIdIdNo ) IR On IrCoCd=IdCoCd And IrIdIdNo=IdIdNo " + _
              " Where " + wCnd)
            'InCoCd='" + gs_CoCd + "' And '," + wInExpNo + ",' Like '%,'+InExpNo+',%' ")

If Not moCn.RecSeek("Select 'x' from SbOutPut Nolock Where qTblNm='SB' and " + ws_CoUsrMnuOut + " ") Then
  ms_ErrMsg = "'SB' Section Not Generated in SB File"
  GoTo ErrHldIns
  Exit Sub
End If

moCn.Execute ("Insert Into SBOutput Values (" + ws_CoUsrMnuStr + ", 'TAG',4,'',0,'<TABLE>INVOICE')")

moCn.Execute ("Insert Into SBOutput " + _
             "Select " + ws_CoUsrMnuStr + ", 'INV',5,INEXPNO,ROW_NUMBER()OVER(ORDER BY INEXPNO) , " + ws_INVStr + " From InvHD " + _
             "Join CustMst On CmCtg='C' And CmCd=InCmCd " + _
             "Where " + wCnd)
             'InCoCd='" + gs_CoCd + "' And '," + wInExpNo + ",' Like '%,'+InExpNo+',%' ")

If Not moCn.RecSeek("Select 'x' from SbOutPut Where qTblNm='INV' and " + ws_CoUsrMnuOut + " ") Then
  ms_ErrMsg = "'INV' Section Not Generated in SB File"
  GoTo ErrHldIns
  Exit Sub
End If

moCn.Execute ("Insert Into SBOutput Values (" + ws_CoUsrMnuStr + ", 'TAG',6,'',0,'<TABLE>ITEM')")

moCn.Execute ("Insert Into SBOutput " + _
             "Select " + ws_CoUsrMnuStr + ", 'ITEM',7,'', 0, " + ws_ITMStr + " From InvDsg " + _
             "Join InvHd On InCoCd=IdCoCd And InTc=IdTc And InYy=IdYy And InChr=IdChr And InNo=IdNo " + _
             "Join DsgMst On DmTcTyp='DM' And DmCd=IdDmCd And DmSz='' " + _
             "Join SBOutput On qTblNm='INV' And qExpNo=InExpNo " + _
             " Join CustMst On CmCtg='C' and CmCd=InCmCd " + _
             "Where " + wCnd + " and " + ws_CoUsrMnuOut + " Group By InExpNo, DmUOM ")
             
             'InCoCd='" + gs_CoCd + "' And '," + wInExpNo + ",' Like '%,'+InExpNo+',%'

If Not moCn.RecSeek("Select 'x' from SbOutPut Where qTblNm='ITEM' and " + ws_CoUsrMnuOut + " ") Then
  ms_ErrMsg = "'ITEM' Section Not Generated in SB File"
  GoTo ErrHldIns
  Exit Sub
End If

moCn.Execute ("Insert Into SBOutput Values (" + ws_CoUsrMnuStr + ", 'TAG',8,'',0,'<END-SB>')")

moCn.Execute ("Insert Into SBOutput " + _
             "Select " + ws_CoUsrMnuStr + ", 'FOOTER',9,'',0," + ws_FTStr + " From InvHD " + _
             "Where " + ws_MinCnd)
             ''InCoCd='" + gs_CoCd + "' And InExpNo='" + wInExpNo + "' ")

If Not moCn.RecSeek("Select 'x' from SbOutPut Where qTblNm='FOOTER' and " + ws_CoUsrMnuOut + " ") Then
  ms_ErrMsg = "'FOOTER' Section Not Generated in SB File"
  GoTo ErrHldIns
  Exit Sub
End If

CreateFile
If ms_ErrMsg <> "" Then GoTo ErrHldIns:  Exit Sub


Set wRs_ShpBill = Nothing

Exit Sub
ErrHldIns:
  Set wRs_ShpBill = Nothing
  moCn.Execute ("Delete From SBOutput Where " + ws_CoUsrMnuOut + "  ")
  moCn.Execute ("Delete From ShpBillFmt Where " + ws_CoUsrMnuFmt + " ")
  ms_ErrMsg = ms_ErrMsg + Err.Description
  Err.Clear
End Sub

Private Sub CreateFile()
On Error GoTo ErrHldFile
Dim wFileNm As String, wRsSbOutPut As MwfLib.MDORowSet
Dim i As Integer

Dim wsOutPut As String

' ***** Manali 3.9.0 - Output folder Path From Param
Dim wOutputPath As String

wOutputPath = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='OUTPUT' and PMCd='OUTPUT' ")
' ***** Manali 3.9.0 - Output folder Path From Param

wFileNm = wOutputPath + "\SB_" + UCase(adc("UoCoCdFr")) + UCase(adc.UsrCd) + UCase(adc.MenuCd) + Format(Now(), "YYYYmmDD_HHMMss") + ".sb"

'wFileNm = App.Path + "\Output\" + "SB_" + UCase(adc("UoCoCdFr")) + UCase(adc.UsrCd) + UCase(adc.MenuCd) + Format(Now(), "YYYYmmDD_HHMMss") + ".sb"

Set wRsSbOutPut = moCn.OpenRes("Select qOutput From SBOutput Where qCoCd='" + UCase(adc("UoCoCdFr")) + "' and qUsrCd='" + UCase(adc.UsrCd) + "' and qMnuCd='" + UCase(adc.MenuCd) + "'  Order by qSeqNo")
If wRsSbOutPut.RecCount > 0 Then
  Open wFileNm For Output As #1
  wRsSbOutPut.MoveFirst
  For i = 1 To wRsSbOutPut.RecCount
    'wsOutPut = IIF(adc("UoYn") = "N" And qTblNm = "ITEM", Replace(wRsSbOutPut!qOutput, , ), wRsSbOutPut!qOutput)
    Print #1, wRsSbOutPut!qOutput
    wRsSbOutPut.MoveNext
  Next
  Close #1
  Shell "Notepad " & wFileNm, vbNormalFocus
  
Else
  DispMsg "Output Data Not Found..", etError: Exit Sub
End If

Set wRsSbOutPut = Nothing

Exit Sub
ErrHldFile:
    Close #1
    Set wRsSbOutPut = Nothing
    ms_ErrMsg = Err.Description
     
End Sub


'Private Sub ChkBoxArr_Click(Index As Integer)
'With ChkBoxArr(Index)
'  Select Case Index
'  Case Is = 0     'Tag = "ShowRepPrc"
'    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
'  End Select
'End With
'
'End Sub
'
'Private Sub ChkBoxArr_GotFocus(Index As Integer)
'  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
'
'End Sub

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
    
    mb_MultiInv = moCn.RecSeek("Select 'x' from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' ")
    
    If mb_MultiInv Then
      Set adc.FirNKeyCtl = adc("UoInExpNoSel")
    Else
      Set adc.FirNKeyCtl = adc("UoInTcFr")
    End If
     
    Call InitProp(Me)
    
    If mb_MultiInv Then
      adc("UoInTcFr").Enabled = False: adc("UoInTcFr").Visible = False
      adc("UoInYyFr").Enabled = False: adc("UoInYyFr").Visible = False
      adc("UoInChrFr").Enabled = False: adc("UoInChrFr").Visible = False
      adc("UoInNoFr").Enabled = False: adc("UoInNoFr").Visible = False
      adc("UoInExpNoFr").Enabled = False: adc("UoInExpNoFr").Visible = False
      LblInvNo.Visible = False: LblChrSlash.Visible = False
      LblTcSlash.Visible = False: LblYySlash.Visible = False
      LblExpNo.Caption = "Exp No (Multiple)"
    Else
      adc("UoInExpNoSel").Enabled = False: adc("UoInExpNoSel").Visible = False
    End If
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_InTcTyp = "IN"
  
  Select Case UCase(IdName)
  Case Is = UCase("UoInExpNoSel")
    If mb_MultiInv = True Then
      Dim ws_MinCurCd As String, ws_MaxCurCd As String
      Dim ws_MinCnvFct As String, ws_MaxCnvFct As String
      Dim ws_MinCnvRt As String, ws_MaxCnvRt As String
      If pv_NewValue <> "" Then
        Cancel = Not (ChkVal("Select 'x' from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InExpNo in (" + pv_NewValue + ")", ErrMsg))
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
      
          
      End If
    End If
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
    gs_InTcTyp = "IN"
    SetProp Me, IdName, When
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
''''
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
  
  If mb_MultiInv Then
    adc("UoInTcFr") = "": adc("UoInYyFr") = "": adc("UoInChrFr") = "": adc("UoInNoFr") = 0
    adc("UoInExpNoFr") = ""
  Else
    adc("UoInExpNoSel") = ""
  End If
  
  
'  adc("WHCUSTOMHOUSECD") = "INBOM4"
'  adc("WHUNITCD") = "SDL8"
'  adc("WHCHALICENSECD") = "SELF"
'  adc("WHSBUOM") = "GMS"
'  adc("WHTYPEXP") = "F"
'  adc("WHSTATEORIGINEXP") = "60"
'  adc("WIPOL") = "INBOM4"
'  adc("WICC") = "US"
'  adc("WIPOD") = "USKWT"
'  adc("WICOD") = "US"
'  adc("WITYPSHP") = "OS"
'  adc("WIMODTRNSPORT") = "A"
'  adc("WIPKJUNIT") = "TIN"
'  adc("WITOTNOPKG") = 1
'  'adc("WIAF") = "B"
'  adc("WINOC") = "CIF"
'  adc("WICURCD") = "USD"
  
'  If adc("UoYn") = "" Then adc("UoYn") = "N"
'  If adc("UoYn") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  
End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                            "HCd = '" + ctSelfCmCd + "' ")
  'Rep.wCoCd.SetText ws_HName
  
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
    Set Rep = Nothing

End Sub


