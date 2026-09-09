VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpPckLstInv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing List From Invoice"
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
      Left            =   11010
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
      Left            =   5175
      TabIndex        =   23
      Top             =   9600
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
      Left            =   6390
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
         TabPicture(0)   =   "EmrFrpPckLstInv.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpPckLstInv.frx":001C
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
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show PO No      "
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
               Index           =   3
               Left            =   120
               TabIndex        =   42
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show PO No "
               Top             =   3600
               Width           =   2130
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Cust Dsg      "
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
               Left            =   120
               TabIndex        =   12
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture (when Detail Sort is DMCD / INVSR)"
               Top             =   3315
               Width           =   2130
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Picture      "
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
               TabIndex        =   11
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture (when Detail Sort is DMCD / INVSR)"
               Top             =   3075
               Width           =   2130
            End
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
               Index           =   1
               Left            =   120
               TabIndex        =   10
               Tag             =   "SortbyOrdCust"
               ToolTipText     =   "Check to Sort by Order Customer"
               Top             =   2805
               Width           =   2130
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2835
               TabIndex        =   1
               ToolTipText     =   "Enter Invoice Year"
               Top             =   765
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
               Left            =   2055
               TabIndex        =   0
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   765
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
               Left            =   3255
               TabIndex        =   2
               ToolTipText     =   "Enter Invoice Character"
               Top             =   765
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
               Left            =   4035
               TabIndex        =   3
               ToolTipText     =   "Enter Invoice Number"
               Top             =   765
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
               Index           =   0
               Left            =   2055
               TabIndex        =   5
               ToolTipText     =   "Enter Value On Which Invoice Report Grouping Has To Be Done"
               Top             =   1335
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
               Index           =   4
               Left            =   2055
               TabIndex        =   8
               ToolTipText     =   "Specify Whether To Show Invoice Or Customs Weight & Amount"
               Top             =   2190
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOINWT"
               IdName          =   "UOINWT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   2055
               TabIndex        =   32
               ToolTipText     =   "Enter Company Code"
               Top             =   480
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
               Index           =   9
               Left            =   2055
               TabIndex        =   4
               ToolTipText     =   "Enter From Export Number"
               Top             =   1050
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOFR"
               IdName          =   "UOINEXPNOFR"
               CmpStr          =   "InExpNo ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   2055
               TabIndex        =   6
               ToolTipText     =   "Enter Detail Sorting To Be Shown In Report"
               Top             =   1620
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
               Index           =   23
               Left            =   2055
               TabIndex        =   9
               ToolTipText     =   "Enter Currency Code Whose Invoice You Wish To View Or Leave It Blank If You Wish To View All Invoices In Base Currency"
               Top             =   2475
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   2640
               TabIndex        =   37
               ToolTipText     =   "Sort by Order Customer"
               Top             =   2785
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
               Index           =   2
               Left            =   2640
               TabIndex        =   38
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   3075
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
               Index           =   3
               Left            =   2640
               TabIndex        =   39
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   3315
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   2055
               TabIndex        =   7
               ToolTipText     =   "Enter Level Of Raw Material Detail To Be Shown"
               Top             =   1905
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UORMGRP"
               IdName          =   "UORMGRP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   2640
               TabIndex        =   41
               ToolTipText     =   "Show PO No (Yes / No)"
               Top             =   3600
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               IdName          =   "UOYN3"
            End
            Begin VB.Label LblRmGrp 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "RM Det Level"
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
               Height          =   240
               Left            =   165
               TabIndex        =   40
               ToolTipText     =   "Location"
               Top             =   1905
               Width           =   1260
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
               Left            =   165
               TabIndex        =   36
               Top             =   2475
               Width           =   1005
            End
            Begin VB.Label LblDet 
               BackStyle       =   0  'Transparent
               Caption         =   "Detail Sort"
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
               TabIndex        =   35
               Top             =   1620
               Width           =   1845
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
               Left            =   165
               TabIndex        =   34
               Top             =   1050
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
               TabIndex        =   33
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label LblWtAs 
               BackStyle       =   0  'Transparent
               Caption         =   "Weight/ Value As"
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
               Top             =   2190
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
               TabIndex        =   30
               Top             =   1335
               Width           =   1845
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
               Left            =   3960
               TabIndex        =   29
               Top             =   765
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
               Left            =   165
               TabIndex        =   28
               Top             =   765
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
               Index           =   25
               Left            =   3195
               TabIndex        =   27
               Top             =   765
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
               Index           =   27
               Left            =   2775
               TabIndex        =   26
               Top             =   765
               Width           =   105
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpPckLstInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepPckLstInv
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  ''GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub
Private Sub adc_setreprecsource()
'*** Report Sql ***
'*** Jay 2.14 *** [Finding Err] [qAccType,qMinAccRt, qMaxAccRt,qAccWt, qAccVal added]
'****** (Jen 2.13) added qGldLsWt, qPltLsWt, qSlvLsWt ******
'***Geeta***Additional  qMinGldRt,qMaxGldRt,qMinPltRt,qMaxPlt
' Zubin 212 (Sales Ctg and qDmSalCtgDesc added)
' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' fields added
' ***** Manali 3.6.0 - 12/11/09 - IdTrayNo added
'
'select space(4) as qTag, IdTc, IdYy, IdChr, IdNo, IdSr,OmCmCd, IdTrayNo, IdIAdd as qIdAdd,
'RmCtg, RmSCtg, IrRmQty as qRmQty, IrRmIWt as qRmWt, IrRmIRt,
'IrRmIVal as qRmVal, IdDmCd, IdQty as qIdQty, IdKt, IdVaCtg, OdDmCol, DmCtg, DmSalCtg,
'IrRmIWt as qGrsWt, IdGldLs, space(1) as qGldType, space(1) as qPltType, space(1) as qSlvType, space(1) as qAccType,
'IrRmIRt as qMinGldRt,IrRmIRt as qMaxGldRt,IrRmIRt as qMinPltRt, IrRmIRt as qMaxPltRt,
'IrRmIRt as qMinSlvRt, IrRmIRt as qMaxSlvRt, IrRmIRt as qMinPdRt, IrRmIRt as qMaxPdRt, IrRmIRt  as qMinAccRt, IrRmIRt  as qMaxAccRt,
'IrRmIWt as qGldWt, IrRmIWt as qGldLsWt, IrRmIVal as qGldVal,
'IrRmIWt as qPltWt, IrRmIWt as qPltLsWt, IrRmIVal As qPltVal,
'IrRmIWt as qSlvWt, IrRmIWt as qSlvLsWt, IrRmIVal As qSlvVal,
'IrRmIWt as qPdWt, IrRmIWt as qPdLsWt, IrRmIVal As qPdVal,
'IrRmIWt as qAccWt, IrRmIVal As qAccVal,
'iLiVal As qLabVal, IrSetIVal As qSetVal, InPrnDt As qInvDt, IdIFOB,
'InExpNo, DmUom, RmCtg as OrRmCtg, space(254) as qMainRmCtgDesc, space(254) as qDmCtgDesc, space(254) as qDmSalCtgDesc,
'space(254) as qDmColDesc, space(254) as qVaCtgDesc
'from InvDsg, InvHd, InvLab, InvRm, OrdDsg, DsgMst, RmMst where 1=2

'***Geeta***Additional  qMinGldRt,qMaxGldRt,qMinPltRt,qMaxPlt
'*** Jay 2.14 *** [Finding Err]
'*** Report Sql ***
   
  Dim wSqlStrg1, wSqlStrg2, wSqlStrg3, wSqlStrg4, wrepcnd, wCnd, wGrpArr, wGrpFormula, wGrpDescFormula As String
  Dim wOdJoin As String
  Dim wIrRmWt, wIrRmRt, wIrRmVal, wiLVal, wIrSetVal, wIdFOB As String
  Dim qDmCtgDesc As String, qMainRmCtgDesc As String, qDmColDesc As String, qDmKtDesc As String
  Dim qVaCtgDesc As String, qMainMet As String, i As Integer, wOdDmCol As String
  'dim wDmCtg As String
  Dim wAdd As String, wFlagLocalExp As String
  ' Zubin 212
  Dim qDmSalCtgDesc As String
  '*** Jay 3.3(FG)
  Dim ws_qIdVaCtg  As String, ws_qIdiFOB As String, ws_IdIAdd As String
  Dim ws_qIrRmIWt As String, ws_qIrRmIRt As String, ws_qIrRmIVal As String, ws_qIrSetIVal As String
  Dim ws_qiLIVal As String, wCustDsgCd As String, wPoNo As String
  '*** Jay 3.3(FG)
  'Hunnar.1 Rm detail level scope should not blank
  If UCase(adc.MenuCd) <> UCase("RepFgPckLst") And adc("UoRmGrp") = "" Then DispMsg "RM Detail Level Should Not Blank", etError: Exit Sub
  If UCase(adc.MenuCd) <> UCase("RepFgPckLst") And UCase(adc("UoRmGrp")) = "RMINTQLY" Then DispMsg "RM Detail Level Internatonal Quality is not supported for this report", etError: Exit Sub
    
  'wDmJoin, wDmUomFlag As Boolean, wOrJoin As String,
  qDmCtgDesc = "'' as qDmCtgDesc": qMainRmCtgDesc = "'' as qMainRmCtgDesc"
  qDmColDesc = "'' as qDmColDesc": qVaCtgDesc = "'' as qVaCtgDesc": qDmKtDesc = "'' as qDmKtDesc"
  qMainMet = "'' as OrRmCtg"
  wOdDmCol = "'' as OdDmCol": '(Bef 20/06/05) wDmCtg = "'' as DmCtg"
  ' Zubin 212
  qDmSalCtgDesc = "'' as qDmSalCtgDesc"
    
  wOdJoin = ""
  'wDmUomFlag = False
  
  Set Rep = Nothing
  'Set mRep = Rep
  
  ' Zubin 213
  Dim ws_BaseCurCd As String, wb_InBaseCur As Boolean, ws_CurJoin As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  If adc("UoCurCdFr") = "" Then
    wb_InBaseCur = True
    Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
  Else
    wb_InBaseCur = False
    Rep.TxtValIn.SetText "Value In " + adc("UoCurCdFr")
    ws_CurJoin = " Join CustMst On CmCtg= 'C' and CmCd= InCmCd "
  End If
  ' Zubin 213
  '*** Jay 3.3(FG)
  If UCase(adc.MenuCd) <> UCase("RepFgPckLst") Then
    ws_qIdiFOB = "IdiFOB": ws_IdIAdd = "IdIAdd": ws_qIdVaCtg = "IdVaCtg"
    ws_qIrRmIWt = "IrRmIWt": ws_qIrRmIRt = "IrRmIRt": ws_qIrRmIVal = "IrRmIVal": ws_qIrSetIVal = "IrSetIVal"
    ws_qiLIVal = "iLIVal"
    'Hunnar.1 - Show Shape/RmCd in report
    If Trim(UCase(adc("UoRmGrp"))) = "RMSCTG" Then
      Rep.hShpCd.SetText ("Type")
      Rep.RmCd1.Suppress = True
    Else
        Rep.hShpCd.SetText ("Rm Code")
        Rep.RmSCtg1.Suppress = True
    End If
  Else
    ws_qIdiFOB = "OdSalPrc": ws_IdIAdd = "0": ws_qIdVaCtg = "OdVaCtg"
    ws_qIrRmIWt = "OrWt": ws_qIrRmIRt = "OrSalRt": ws_qIrRmIVal = "OrSalVal": ws_qIrSetIVal = "OrSetSalVal"
    ws_qiLIVal = "OlSalVal"
    
    'Hunnar.1 - Suppress RmCd and PoNo for Fg Packing List report
    Rep.RmCd1.Suppress = True
    Rep.FldPoNo.Suppress = True
  End If
  '*** Jay 3.3(FG)
  If adc("UOINWT") = "I" Then
    '*** Jay 3.3(FG)
    wIrRmWt = ws_qIrRmIWt
    wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then " + ws_qIrRmIRt + "/InCnvFct else " + ws_qIrRmIRt + "*InCnvFct end) ", ws_qIrRmIRt)
    wIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then " + ws_qIrRmIVal + "/InCnvFct else " + ws_qIrRmIVal + "*InCnvFct end) ", ws_qIrRmIVal)
    wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then " + ws_qiLIVal + "/InCnvFct else " + ws_qiLIVal + "*InCnvFct end) ", ws_qiLIVal)
    wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then " + ws_qIrSetIVal + "/InCnvFct else " + ws_qIrSetIVal + "*InCnvFct end) ", ws_qIrSetIVal)
    wIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then " + ws_qIdiFOB + "/InCnvFct else " + ws_qIdiFOB + "*InCnvFct end) ", ws_qIdiFOB)
    wAdd = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then " + ws_IdIAdd + "/InCnvFct else " + ws_IdIAdd + "*InCnvFct end) ", ws_IdIAdd)
    '*** Jay 3.3(FG)
  ElseIf adc("UOINWT") = "Z" Then
    ' Zubin 213
    wIrRmWt = "IrRmZWt"
    wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZRt/InCnvFct else IrRmZRt*InCnvFct end) ", "IrRmZRt")
    wIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal*InCnvFct end) ", "IrRmZVal")
    wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZVal/InCnvFct else iLZVal*InCnvFct end) ", "iLZVal")
    wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZVal/InCnvFct else IrSetZVal*InCnvFct end) ", "IrSetZVal")
    wIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdZFOB/InCnvFct else IdZFOB*InCnvFct end) ", "IdZFOB")
    wAdd = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdZAdd/InCnvFct else IdZAdd*InCnvFct end) ", "IdZAdd")
  End If
  
  Set adc.RepSource = Rep
  Call DispCoNm
  ''Call GRP_REP.Gen3LRep
  ''grpflds = GRP_REP.GrpFldLst
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  ' Zubin 213
  wCnd = wCnd + IIF(wb_InBaseCur = False, " and CmCurCd= '" + adc("UoCurCdFr") + "'", "")

  wFlagLocalExp = moCn.GetFldVal("Select hExpLocal from Head where HCoCd= '" + adc("UoCoCdFr") + "' and " + _
                   "HCd= '" + ctSelfCmCd + "' ")
    
  wGrpArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
         "PMCd= '" + adc("UoInGrp") + "' "), ",")
    
  '4.1.2
  If adc("UOYN1") = "Y" Then
    wGrpFormula = "Trim(UpperCase(ToText({ado.OmCmCd})))"  '4.1.2
    'wSubRepGrpFormula = "OdOmCmCd"
  End If
     
  If UCase(adc("UoYN")) = "Y" Then
'        Rep.PICSEC.Suppress = False
  Else
        Rep.PICSEC.Suppress = True
 '       Rep.hPrtCustDsg.Suppress = True
  End If
   
  If UCase(adc("UoYN2")) = "Y" Then
'        Rep.CustDmCdSec.Suppress = "Trim({ado.CdCustDmCd})=''"
  Else
        Rep.wPrtCustDsg.Suppress = True
        Rep.hPrtCustDsg.Suppress = True
  End If
  
  'Hunnar.1 - Show PO No. based on selection.
  If UCase(adc("UoYN3")) = "Y" Then
        Rep.FldPoNo.Suppress = False
  Else
        Rep.FldPoNo.Suppress = True
  End If
   
  If UCase(adc("UoYN2")) = "Y" Then
        wCustDsgCd = "IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(IdDmCd) and CdOurSfx= max(IdSfx) and CdOurDmSz= max(IdDmSz)), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(IdDmCd) and CdOurSfx= max(IdSfx) and CdOurDmSz= ''), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(IdDmCd) and CdOurSfx= '' and CdOurDmSz= max(IdDmSz)), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(IdDmCd) and CdOurSfx= '' and CdOurDmSz= ''), '')))) as CdCustDmCd, "
  Else
        wCustDsgCd = "'' as CdCustDmCd, "
  End If
  'wGrpFormula = "Trim(UpperCase(ToText({ado.OmCmCd})))"  '4.1.2
     
  For i = 0 To UBound(wGrpArr)
    If UCase(wGrpArr(i)) = "RMCTG" Then
      'qMainRmCtgDesc = "(Select PDesc from Param where PTyp= 'RMCTG' and PMCd= max(RmCtg)) as qMainRmCtgDesc"
      qMainRmCtgDesc = "'' as qMainRmCtgDesc"
      'wGrpDescFormula = wGrpDescFormula + IIf(wGrpDescFormula = "", "", "+', '+") + "Trim({ado.qMainRmCtgDesc})"
      '*** Jay 3.3(FG)
      qMainMet = IIF(UCase(adc.MenuCd) = UCase("RepFgPckLst"), _
                  "IsNull((Select OrRmCtg from OrdRm Or1 where " + IIF(gs_Partition = ctCurrPrtn, " Or1.OrPrtKey=OdPrtKey and ", "") + " Or1.OrCoCd= OdCoCd and and Or1.OrTc= OdTc and Or1.OrYy= OdYy and Or1.OrChr= OdChr and Or1.OrNo= OdNo and Or1.OrSr= OdSr and Or1.OrMainMet= 'Y'),'')", _
                  " IsNull((Select RmCtg from InvRm join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=IrPrtKey and ", "") + " RmCd= IrRmCd where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey=IdPrtKey and ", "") + " IrCoCd= IdCoCd " + _
                  "and IrTc= IdTc and IrYy= IdYy and IrChr= IdChr and IrNo= IdNo and " + _
                  "IrSr= IdSr and IrMainMet= 'Y'), '' )") + " as OrRmCtg "
      '*** Jay 3.3(FG)
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({ado.OrRmCtg})"
    ElseIf UCase(wGrpArr(i)) = "KT" Then
      qDmKtDesc = "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(IdKt)) as qDmKtDesc"
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({ado.qDmKtDesc})"
    ElseIf UCase(wGrpArr(i)) = "DMCOL" Then
      qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({ado.qDmColDesc})"
      wOdDmCol = "max(OdDmCol) as OdDmCol"
      wOdJoin = " join OrdDsg on " + IIF(gs_Partition = ctCurrPrtn, " IdPrtKey=OdPrtKey and ", "") + " IdOdIdNo=OdIdNo and IdExpOdTc = OdTc and IdExpOdYy = OdYy and IdExpOdChr = OdChr and  " + _
                  " IdExpOdNo = OdNo and IdExpOdSr = OdSr "
    ElseIf UCase(wGrpArr(i)) = "VACTG" Then
      '*** Jay 3.3(FG)
      qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(" + ws_qIdVaCtg + ")) as qVaCtgDesc"
      '*** Jay 3.3(FG)
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({ado.qVaCtgDesc})"
    ElseIf UCase(wGrpArr(i)) = "GLDLS" Then
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "'@'+ToText({ado.IdGldLs}, '##0.000')+ '%'"
    ElseIf UCase(wGrpArr(i)) = "PCS/PRS" Then
      ''wDmUomFlag = True
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "{ado.DmUom}"
    ElseIf UCase(wGrpArr(i)) = "DMCTG" Then
      qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({ado.qDmCtgDesc})"
      '(Bef 20/06/05) wDmCtg = "max(DmCtg) as DmCtg"
      'wDmJoin = " join DsgMst on DmTcTyp = 'DM' and DmCd = IdDmCd and DmSz = '' "
    ElseIf UCase(wGrpArr(i)) = "SALCTG" Then
      qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({ado.qDmSalCtgDesc})"
    End If
    wGrpFormula = wGrpFormula + IIF(wGrpFormula = "", "", "+") + "Trim(UpperCase(ToText({ado." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr(i) + "'") + "})))"
  Next i
  
' **** Zubin 211 **** '
  '*** For the Detail Formula of the report
  Dim wDetArr() As String, wDetSrt As String
  
  wDetArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INDET' and " + _
            "PMCd= '" + adc("UoInDet") + "' "), ",")
  For i = 0 To UBound(wDetArr)
    ' ***** Manali 3.6.0 - 12/11/09 - IdTrayNo added
    wDetSrt = wDetSrt + IIF(wDetSrt = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({ado.IdSr}, '###0')", IIF(UCase(wDetArr(i)) = "IDTRAYNO", "ToText({ado.IdTrayNo}, '###0')", "Trim(UpperCase({ado." + wDetArr(i) + "}))"))
    ' Bef 3.6.0
    'wDetSrt = wDetSrt + IIF(wDetSrt = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({ado.IdSr}, '###0')", "Trim(UpperCase({ado." + wDetArr(i) + "}))")
  Next i
  
  ' ***** Manali 3.6.0 - 12/11/09 - IdTrayNo ["TRNO/SR"] added
  If UCase(adc("UoInDet")) = "INVSR/SR" Or adc("UoInDet") = "DMCD/SR" Or adc("UoInDet") = "TRNO/SR" Then
    Rep.FldIdSr.Suppress = True
    Rep.FldRunSr.Suppress = False
  Else
    Rep.FldIdSr.Suppress = False
    Rep.FldRunSr.Suppress = True
  End If
  
  ' ***** Manali 3.6.0 - 12/11/09 - IdTrayNo ["TRNO/SR"] added
  If adc("UoInDet") <> "TRNO/SR" Then Rep.FldTrayNo.Suppress = True: Rep.HTrayNo.Suppress = True

  
  GRP_REP.SetFormula Rep, "wInvSr", wDetSrt + IIF(wDetSrt <> "", "+", "") + " {ado.IdTc}  +'/'+ {ado.IdYy} +'/'+ {ado.IdChr} +'/'+ Trim(ToText ({ado.IdNo}, '####0'))+ '/'+ ToText ({ado.IdSr}, '###0')"
' **** Zubin 211 **** '
 
  GRP_REP.SetFormula Rep, "wGrp1", wGrpFormula
  GRP_REP.SetFormula Rep, "wGrp1Desc", wGrpDescFormula
  
' **** Zubin 211 (NOTE) **** '
' In wSqlStrg1 below, "'' as IdDmCd" has been changed to "max(IdDmCd) as IdDmCd"
' **** Zubin 211 **** '

  'Geeta - Emr 2.10.0
'***Geeta***Additional  qMinGldRt,qMaxGldRt,qMinPltRt,qMaxPlt
'  wSqlStrg1 = "Select '2Det' as qTag, IdTc, IdYy, IdChr, IdNo, IdSr, 0 as qIdAdd, " + _
            "RmCtg, RmSCtg, sum(IrRmQty) as qRmQty, sum(" + wIrRmWt + ") as qRmWt, " + wIrRmRt + " as IrRmIRt, " + _
            "sum(" + wIrRmVal + ") as qRmVal, " + _
            "max(IdDmCd) as IdDmCd, 0 as qIdQty, max(IdKt) as IdKt, max(IdVaCtg) as IdVaCtg, " + wOdDmCol + ", " + wDmCtg + ", 0 as qGrsWt, " + _
            "max(IdGldLs) as IdGldLs, '' as qGldType, '' as qPltType, 0 as qGldWt, 0 as qGldVal, " + _
            "0 as qPltWt, 0 as qPltVal, 0 as qLabVal, 0 as qSetVal, " + _
            "max(InPrnDt) as qInvDt, 0 as IdIFOB, max(InExpNo) as InExpNo, max(DmUom) as DmUom, " + qMainMet + ", " + qMainRmCtgDesc + ", " + _
            qDmCtgDesc + ", " + qDmColDesc + ", " + qVaCtgDesc + " From InvDsg " + _
            "join InvHd on InCoCd = IdCoCd and InTc = IdTc and InYy = IdYy and " + _
            "InChr = IdChr And InNo = IdNo " + _
            "left outer join InvRm Ir on IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy and " + _
            "IrChr = IdChr And IrNo = IdNo And IrSr = IdSr " + _
            "left outer Join RmMst Rm on RmCd = IrRmCd " + _
            "join DsgMst on DmTcTyp = 'DM' and DmCd = IdDmCd and DmSz = '' " + wOdJoin + " " + _
            "where (Select PValue from Param where Ptyp= 'TC' and PMCd= IdTc)= 'IN' and " + _
            "IsNull(RmCtg, '') not in ('G', 'P', '') " + wCnd + _
            " group by IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, RmCtg, RmSCtg, " + wIrRmRt + "  "
  
  '*** Jay 3.3(FG)
  If UCase(adc.MenuCd) <> UCase("RepFgPckLst") Then
  '*** Jay 3.3(FG)
    '*** Jay 3.3(FG) [IdCoCd fld added]
    '*** Jay 2.14 *** [Finding Err]
    '****** (Jen 2.13) added qGldLsWt, qPltLsWt, qSlvLsWt ******
    ' Zubin 213 (ws_CurJoin added)
    ' Zubin 212 (DmSalCtg and qDmSalCtgDesc added)
    '****** Sachin 3.02 - Id Fields In Joins
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    ' ***** Manali 3.6.0 - 12/11/09 - IdTrayNo added
    
    'Hunnar.1 - RmCd, OmPoNo,OdPoNo,OdSubCust added.  Changed group by based on Rm Det Level
    wSqlStrg1 = "Select '2Det' as qTag, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr,Max(OmCmCd) as OmCmCd, max(IdTrayNo) as IdTrayNo, 0 as qIdAdd, " + _
            "RmCtg, RmSCtg, max(RmCd) as RmCd, max(OmPoNo) as OmPoNo,max(OdSubCust) as OdSubCust,max(OdPoNo) as OdPoNo, sum(IrRmQty) as qRmQty, sum(" + wIrRmWt + ") as qRmWt, " + wIrRmRt + " as IrRmIRt, " + _
            "sum(" + wIrRmVal + ") as qRmVal, max(OdDmSz) as OdDmSz, max(OdPicNm) as OdPicNm," + wCustDsgCd + _
            "max(IdDmCd) as IdDmCd, max(IdSfx) as IdSfx, 0 as qIdQty, max(IdKt) as IdKt, max(" + ws_qIdVaCtg + ") as IdVaCtg, " + wOdDmCol + ", max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, 0 as qGrsWt, " + _
            "max(IdGldLs) as IdGldLs, '' as qGldType, '' as qPltType, '' as qSlvType, '' as qPdType, '' as qAccType, " + _
            "0 as qMinGldRt, 0 as qMaxGldRt,0 as qMinPltRt,0 as qMaxPltRt, 0 as qMinSlvRt,0 as qMaxSlvRt, 0 as qMinPdRt,0 as qMaxPdRt, 0 as qMinAccRt, 0 as qMaxAccRt, " + _
            "0 as qGldWt, 0 as qGldLsWt, 0 as qGldVal, 0 as qPltWt, 0 as qPltLsWt, 0 as qPltVal, " + _
            "0 as qSlvWt, 0 as qSlvLsWt, 0 as qSlvVal, 0 as qPdWt, 0 as qPdLsWt, 0 as qPdVal, 0 as qAccWt, 0 as qAccVal, " + _
            "0 as qLabVal, 0 as qSetVal, " + _
            "max(InPrnDt) as qInvDt, 0 as IdIFOB, max(InExpNo) as InExpNo, max(DmUom) as DmUom, " + qMainMet + ", " + qMainRmCtgDesc + ", " + _
            qDmCtgDesc + ", " + qDmSalCtgDesc + ", " + qDmColDesc + ", " + qDmKtDesc + ", " + qVaCtgDesc + " From InvDsg " + _
            "join InvHd on InIdNo=IdInIdNo and InCoCd = IdCoCd and InTc = IdTc and InYy = IdYy and " + _
            "InChr = IdChr And InNo = IdNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IdPrtKey ", "") + ws_CurJoin + _
            " join OrdMst on OmCoCd=IdCoCd  and OmTc= IdExpOdTc and OmYy= IdExpOdYy " + _
            " and OmChr=IdExpOdChr and OmNo= IdExpOdNo " + _
            " join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
            " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey ", "") + _
            "join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp = 'DM' and DmCd = IdDmCd and DmSz = '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=IdPrtKey ", "") + wOdJoin + " " + _
            "left outer join InvRm Ir on IrIdIdNo=IdIdNo and IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy and " + _
            "IrChr = IdChr And IrNo = IdNo And IrSr = IdSr " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=IdPrtKey ", "") + _
            "left outer Join RmMst Rm on RmCd = IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=IrPrtKey ", "") + _
            "where (Select PValue from Param where Ptyp= 'TC' and PMCd= IdTc)= 'IN' and " + IIF(gs_Partition = ctCurrPrtn, " IdPrtKey='" + ctCurrPrtn + "' and ", "") + _
            "IsNull(RmCtg, '') not in ('M', 'G', 'P', 'S', 'L', 'X', '') " + wCnd + _
            " group by IdPrtKey, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, RmCtg, " + IIF(UCase(Trim(adc("UoRmGrp"))) = "RMCD", " RmSCtg,RmCd,  ", " RmSCtg, ") + "" + wIrRmRt + "  "
            
  '***Geeta***Additional  qMinGldRt,qMaxGldRt,qMinPltRt,qMaxPlt
  '  wSqlStrg2 = " Union All " + _
  '            "Select '1Hd' as qTag, IdTc, IdYy, IdChr, IdNo, IdSr, max(" + wAdd + ") as qIdAdd, " + _
  '            "'' as RmCtg, '' as RmSCtg, 0 as qRmQty, 0 as qRmWt, 0 as IrRmIRt, 0 as qRmVal, " + _
  '            "max(IdDmCd) as IdDmCd, max(IdQty) as qIdQty, max(IdKt) as IdKt, max(IdVaCtg) as IdVaCtg, " + wOdDmCol + ", " + wDmCtg + ", " + _
  '            "sum((case when RmCtg in ('D', 'C') then IsNull(" + wIrRmWt + ", 0)/5 " + _
  '            "else IsNull(" + wIrRmWt + ", 0) end)) as qGrsWt, max(IdGldLs) as IdGldLs, " + _
  '            "max(Case when RmCtg= 'G' then 'G' else '' end) as qGldType, " + _
  '            "max(Case when RmCtg= 'P' then 'P' else '' end) as qPltType, "
  
  '*** Jay 2.14 *** [Finding Err]
  ' ****** Sachin 2.12 - 05/12/05 - SJM - Silver added to query
    ' Zubin 212 (DmSalCtg added)
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    '*** Jay 3.3(FG)
    ' ***** Manali 3.6.0 - 12/11/09 - IdTrayNo added
    'Hunnar.1 - RmCd, OmPoNo,OdSubCust,OdPoNo added
    wSqlStrg2 = " Union All " + _
              "Select '1Hd' as qTag, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr,max(OmcmCd) as OmCmCd, max(IdTrayNo) as IdTrayNo, max(" + wAdd + ") as qIdAdd, " + _
              "'' as RmCtg, '' as RmSCtg, '' RmCd, max(OmPoNo) as OmPoNo,max(OdSubCust) as OdSubCust,max(OdPoNo) as OdPoNo, 0 as qRmQty, 0 as qRmWt, 0 as IrRmIRt, 0 as qRmVal, max(OdDmSz) as OdDmSz, max(OdPicNm) as OdPicNm, " + wCustDsgCd + _
              "max(IdDmCd) as IdDmCd, max(IdSfx) as IdSfx, max(IdQty) as qIdQty, max(IdKt) as IdKt, max(" + ws_qIdVaCtg + ") as IdVaCtg, " + wOdDmCol + ", max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, " + _
              "sum((case when RmCtg in ('D', 'C') then IsNull(" + wIrRmWt + ", 0)/5 " + _
              "else IsNull(" + wIrRmWt + ", 0) end)) as qGrsWt, max(IdGldLs) as IdGldLs, " + _
              "max(Case when RmCtg= 'G' then 'G' else '' end) as qGldType, " + _
              "max(Case when RmCtg= 'P' then 'P' else '' end) as qPltType, " + _
              "max(Case when RmCtg= 'S' then 'S' else '' end) as qSlvType, " + _
              "max(Case when RmCtg= 'L' or RmCtg = 'M' then 'L+M' else '' end) as qPdType, " + _
              "max(Case when RmCtg= 'X' then 'X' else '' end) as qAccType, " + _
              "Min(Case when RmCtg= 'G' then IsNull(" + wIrRmRt + ", 0) else 9999 end) as qMinGldRt," + _
              "Max(Case when RmCtg= 'G' then IsNull(" + wIrRmRt + ", 0) else 0 end) as qMaxGldRt," + _
              "Min(Case when RmCtg= 'P' then IsNull(" + wIrRmRt + ", 0) else 9999 end) as qMinPltRt, " + _
              "Max(Case when RmCtg= 'P' then IsNull(" + wIrRmRt + ", 0) else 0 end) as qMaxPltRt, " + _
              "Min(Case when RmCtg= 'S' then IsNull(" + wIrRmRt + ", 0) else 9999 end) as qMinSlvRt, " + _
              "Max(Case when RmCtg= 'S' then IsNull(" + wIrRmRt + ", 0) else 0 end) as qMaxSlvRt, " + _
              "Min(Case when RmCtg= 'L' or RmCtg = 'M' then IsNull(" + wIrRmRt + ", 0) else 9999 end) as qMinPdRt, " + _
              "Max(Case when RmCtg= 'L' or RmCtg = 'M' then IsNull(" + wIrRmRt + ", 0) else 0 end) as qMaxPdRt, " + _
              "Min(Case when RmCtg= 'X' then IsNull(" + wIrRmRt + ", 0) else 9999 end) as qMinAccRt, " + _
              "Max(Case when RmCtg= 'X' then IsNull(" + wIrRmRt + ", 0) else 0 end) as qMaxAccRt, "
              
    'Geeta - Emr 2.10.0
    ' Zubin 212 (qDmSalCtgDesc added)
    ' Zubin 213 (InvHd Join added in sub query)
    
    '*** (Bef 2.13)
    'wSqlStrg3 = " sum(Case when RmCtg= 'G' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qGldWt, " + _
    '          "sum(Case when RmCtg= 'G' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qGldVal, " + _
    '          "sum(Case when RmCtg= 'P' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qPltWt, " + _
    '          "sum(Case when RmCtg= 'P' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qPltVal, " + _
    '          "sum(Case when RmCtg= 'S' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qSlvWt, " + _
    '          "sum(Case when RmCtg= 'S' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qSlvVal, "
    '*** (Bef 2.13)
    '*** Jay 2.14 *** [Finding Err]
    '*** (Jen 2.13)
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    wSqlStrg3 = " sum(Case when RmCtg= 'G' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qGldWt, " + _
                "sum(Case when RmCtg= 'G' then IsNull(" + IIF(UCase(adc("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt- " + wIrRmWt + " else " + wIrRmWt + " * IdGldLs/100 end) ", wIrRmWt + " * IdGldLs/100 ") + ", 0) else 0 end) as qGldLsWt, " + _
                "sum(Case when RmCtg= 'G' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qGldVal, " + _
                "sum(Case when RmCtg= 'P' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qPltWt, " + _
                "sum(Case when RmCtg= 'P' then IsNull(" + IIF(UCase(adc("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt- " + wIrRmWt + " else " + wIrRmWt + " * IdGldLs/100 end) ", wIrRmWt + " * IdGldLs/100 ") + ", 0) else 0 end) as qPltLsWt, " + _
                "sum(Case when RmCtg= 'P' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qPltVal, " + _
                "sum(Case when RmCtg= 'S' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qSlvWt, " + _
                "sum(Case when RmCtg= 'S' then IsNull(" + IIF(UCase(adc("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt- " + wIrRmWt + " else " + wIrRmWt + " * IdGldLs/100 end) ", wIrRmWt + " * IdGldLs/100 ") + ", 0) else 0 end) as qSlvLsWt, " + _
                "sum(Case when RmCtg= 'S' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qSlvVal, " + _
                "sum(Case when RmCtg= 'L' or RmCtg ='M' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qPdWt, " + _
                "sum(Case when RmCtg= 'L' or RmCtg ='M' then IsNull(" + IIF(UCase(adc("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt- " + wIrRmWt + " else " + wIrRmWt + " * IdGldLs/100 end) ", wIrRmWt + " * IdGldLs/100 ") + ", 0) else 0 end) as qPdLsWt, " + _
                "sum(Case when RmCtg= 'L' or RmCtg ='M' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qPdVal, " + _
                "sum(Case when RmCtg= 'X' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qAccWt, " + _
                "sum(Case when RmCtg= 'X' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qAccVal, "
    '*** (Jen 2.13)
    ' Zubin 213 (wb_InBaseCur cnd for InvHd join and ws_CurJoin added)
    '*** Jay 2.14 *** [Finding Err]
    '****** Sachin 3.02 - Id fields in Joins
    wSqlStrg4 = "IsNull((Select Sum(" + wiLVal + ") from InvLab " + _
              IIF(wb_InBaseCur, " Join InvHd on " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey=IlPrtKey and ", "") + " InCoCd = IlCoCd and InTc = IlTc and InYy = IlYy and InChr = IlChr And InNo = IlNo ", "") + _
              "where " + IIF(gs_Partition = ctCurrPrtn, " IlPrtKey=IdPrtKey and ", "") + " iLCoCd = IdCoCd and " + _
              "iLTc = IdTc And iLYy = IdYy And iLChr = IdChr and iLNo= IdNo and iLSr= IdSr),0) as qLabVal, " + _
              "IsNull((Select Sum(" + wIrSetVal + ") from InvRm SetRm " + _
              IIF(wb_InBaseCur, " Join InvHd on " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey=SetRm.IrPrtKey and ", "") + " InCoCd = SetRm.IrCoCd and InTc = SetRm.IrTc and InYy = SetRm.IrYy and InChr = SetRm.IrChr And InNo = SetRm.IrNo ", "") + _
              "Where " + IIF(gs_Partition = ctCurrPrtn, " SetRm.IrPrtKey='" + ctCurrPrtn + "' and ", "") + " SetRm.IrCoCd=IdCoCd and SetRm.IrTc = IdTc And SetRm.IrYy = IdYy " + _
              "and SetRm.IrChr = IdChr and SetRm.IrNo = IdNo and SetRm.IrSr = IdSr), 0) as qSetVal, " + _
              "max(InPrnDt) as qInvDt, max(" + wIdFOB + ") as IdIFOB, max(InExpNo) as InExpNo, max(DmUom) as DmUom, " + qMainMet + ", " + _
              qMainRmCtgDesc + ", " + qDmCtgDesc + ", " + qDmSalCtgDesc + ", " + qDmColDesc + ", " + qDmKtDesc + ", " + qVaCtgDesc + _
              " From InvDsg " + _
              "join InvHd on InIdNo=IdInIdNo and InCoCd = IdCoCd and InTc = IdTc and InYy = IdYy and InChr = IdChr And InNo = IdNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IdPrtKey ", "") + ws_CurJoin + _
              " join OrdMst on OmCoCd=IdCoCd  and OmTc= IdExpOdTc and OmYy= IdExpOdYy " + _
              " and OmChr=IdExpOdChr and OmNo= IdExpOdNo " + _
              "left outer join InvRm Ir on IrIdIdNo=IdIdNo and IrCoCd = IdCoCd and IrTc = IdTc and IrYy = IdYy and " + _
              "IrChr = IdChr And IrNo = IdNo And IrSr = IdSr " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=IdPrtKey ", "") + _
              "left outer Join RmMst Rm on RmCd = IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=IrPrtKey ", "") + _
              " join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
              " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey ", "") + _
              "join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp = 'DM' and DmCd = IdDmCd and DmSz = '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=IdPrtKey ", "") + wOdJoin + " " + _
              "where (Select PValue from Param where Ptyp= 'TC' and PMCd= IdTc)= 'IN' " + IIF(gs_Partition = ctCurrPrtn, " and IdPrtKey='" + ctCurrPrtn + "' ", "") + wCnd + _
              " group by IdPrtKey, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr " + _
              " order by qTag, IdTc, IdChr, IdNo, IdSr "
  '*** Jay 3.3(FG)
  Else
    ' ***** Manali 3.6.0 - 12/11/09 - IdTrayNo added
    wSqlStrg1 = "Select '2Det' as qTag, OdCoCd as IdCoCd, OdTc as IdTc, OdYy as IdYy, OdChr as IdChr, OdNo as IdNo, OdSr as IdSr,Max(OdOmCmCd) as OmCmCd,0 as IdTrayNo, 0 as qIdAdd, " + _
              "RmCtg, RmSCtg,max(RmCd) as RmCd, '' as OmPoNo,'' as OdSubCust,'' as OdPoNo, sum(OrQty) as qRmQty, sum(" + wIrRmWt + ") as qRmWt, " + wIrRmRt + " as IrRmIRt, " + _
              "sum(" + wIrRmVal + ") as qRmVal, " + _
              "max(OdDmCd) as IdDmCd, 0 as qIdQty, max(OdKt) as IdKt, max(" + ws_qIdVaCtg + ") as IdVaCtg, " + wOdDmCol + ", max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, 0 as qGrsWt, " + _
              "0 as IdGldLs, '' as qGldType, '' as qPltType, '' as qSlvType, '' as qPdType, '' as qAccType, " + _
              "0 as qMinGldRt, 0 as qMaxGldRt,0 as qMinPltRt,0 as qMaxPltRt, 0 as qMinSlvRt,0 as qMaxSlvRt, 0 as qMinPdRt,0 as qMaxPdRt, 0 as qMinAccRt, 0 as qMaxAccRt, " + _
              "0 as qGldWt, 0 as qGldLsWt, 0 as qGldVal, 0 as qPltWt, 0 as qPltLsWt, 0 as qPltVal, " + _
              "0 as qSlvWt, 0 as qSlvLsWt, 0 as qSlvVal, 0 as qPdWt, 0 as qPdLsWt, 0 as qPdVal, 0 as qAccWt, 0 as qAccVal, " + _
              "0 as qLabVal, 0 as qSetVal, " + _
              "max(InPrnDt) as qInvDt, 0 as IdIFOB, max(InExpNo) as InExpNo, max(DmUom) as DmUom, " + qMainMet + ", " + qMainRmCtgDesc + ", " + _
              qDmCtgDesc + ", " + qDmSalCtgDesc + ", " + qDmColDesc + ", " + qDmKtDesc + ", " + qVaCtgDesc + " From OrdDsg " + _
              "join InvHd on InCoCd =OdCoCd and InTc = OdTc and InYy = OdYy and " + _
              "InChr = OdChr And InNo = OdNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=OdPrtKey ", "") + ws_CurJoin + _
              "join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp = OdDmTcTyp and DmCd = OdDmCd and DmSz = '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + " " + _
              "left outer join OrdRm Ir on OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and " + _
              "OrChr = OdChr And OrNo = OdNo And OrSr = OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OdPrtKey ", "") + _
              "left outer Join RmMst Rm on RmCd = OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + _
              "where (Select PValue from Param where Ptyp= 'TC' and PMCd= OdTc)= 'FG' and " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey='" + ctCurrPrtn + "' and ", "") + _
              "IsNull(RmCtg, '') not in ('G', 'P', 'S', 'L', 'X','M', '') " + wCnd + _
              " group by OdPrtKey, OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, RmCtg, RmSCtg, " + wIrRmRt + "  "
    wSqlStrg2 = " Union All " + _
              "Select '1Hd' as qTag, OdCoCd as IdCoCd, OdTc as IdTc, OdYy as IdYy, OdChr as IdChr, OdNo as IdNo, " + _
              "OdSr as IdSr,Max(OdOmCmCd) as OmCmCd, 0 as IdTrayNo, 0 as qIdAdd, " + _
              "'' as RmCtg, '' as RmSCtg,'' as RmCd, '' as OmPoNo,'' as OdSubCust,'' as OdPoNo, 0 as qRmQty, 0 as qRmWt, 0 as IrRmIRt, 0 as qRmVal, " + _
              "max(OdDmCd) as IdDmCd, max(OdOrdQty) as qIdQty, max(OdKt) as IdKt, max(" + ws_qIdVaCtg + ") as IdVaCtg, " + wOdDmCol + ", max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, " + _
              "sum((case when RmCtg in ('D', 'C') then IsNull(" + wIrRmWt + ", 0)/5 " + _
              "else IsNull(" + wIrRmWt + ", 0) end)) as qGrsWt, 0 as IdGldLs, " + _
              "max(Case when RmCtg= 'G' then 'G' else '' end) as qGldType, " + _
              "max(Case when RmCtg= 'P' then 'P' else '' end) as qPltType, " + _
              "max(Case when RmCtg= 'S' then 'S' else '' end) as qSlvType, " + _
              "max(Case when RmCtg= 'L' or RmCtg = 'M' then 'L+M' else '' end) as qPdType, " + _
              "max(Case when RmCtg= 'X' then 'X' else '' end) as qAccType, " + _
              "Min(Case when RmCtg= 'G' then IsNull(" + wIrRmRt + ", 0) else 9999 end) as qMinGldRt," + _
              "Max(Case when RmCtg= 'G' then IsNull(" + wIrRmRt + ", 0) else 0 end) as qMaxGldRt," + _
              "Min(Case when RmCtg= 'P' then IsNull(" + wIrRmRt + ", 0) else 9999 end) as qMinPltRt, " + _
              "Max(Case when RmCtg= 'P' then IsNull(" + wIrRmRt + ", 0) else 0 end) as qMaxPltRt, " + _
              "Min(Case when RmCtg= 'S' then IsNull(" + wIrRmRt + ", 0) else 9999 end) as qMinSlvRt, " + _
              "Max(Case when RmCtg= 'S' then IsNull(" + wIrRmRt + ", 0) else 0 end) as qMaxSlvRt, " + _
              "Min(Case when RmCtg= 'L' or RmCtg = 'M' then IsNull(" + wIrRmRt + ", 0) else 9999 end) as qMinPdRt, " + _
              "Max(Case when RmCtg= 'L' or RmCtg = 'M' then IsNull(" + wIrRmRt + ", 0) else 0 end) as qMaxPdRt, " + _
              "Min(Case when RmCtg= 'X' then IsNull(" + wIrRmRt + ", 0) else 9999 end) as qMinAccRt, " + _
              "Max(Case when RmCtg= 'X' then IsNull(" + wIrRmRt + ", 0) else 0 end) as qMaxAccRt, "
    wSqlStrg3 = " sum(Case when RmCtg= 'G' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qGldWt, " + _
              "0 as qGldLsWt, " + _
              "sum(Case when RmCtg= 'G' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qGldVal, " + _
              "sum(Case when RmCtg= 'P' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qPltWt, " + _
              "0 as qPltLsWt, " + _
              "sum(Case when RmCtg= 'P' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qPltVal, " + _
              "sum(Case when RmCtg= 'S' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qSlvWt, " + _
              "0 as qSlvLsWt, " + _
              "sum(Case when RmCtg= 'S' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qSlvVal, " + _
              "sum(Case when RmCtg= 'L' or RmCtg = 'M' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qPdWt, " + _
              "0 as qPdLsWt, " + _
              "sum(Case when RmCtg= 'L' or RmCtg = 'M' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qPdVal, " + _
              "sum(Case when RmCtg= 'X' then IsNull(" + wIrRmWt + ", 0) else 0 end) as qAccWt, " + _
              "sum(Case when RmCtg= 'X' then IsNull(" + wIrRmVal + ", 0) else 0 end) as qAccVal, "
    wSqlStrg4 = "IsNull((Select Sum(" + wiLVal + ") from OrdLab " + _
              IIF(wb_InBaseCur, " Join InvHd on " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey=OlPrtKey and ", "") + " InCoCd = OlCoCd and InTc = OlTc and InYy = OlYy and InChr = OlChr And InNo = OlNo join InvDsg on InIdNo=IdInIdNo and InCoCd = IdCoCd and InTc = IdTc and InYy = IdYy and InChr = IdChr And InNo = IdNo  and InPrtKey=IdPrtKey ", "") + _
              " join OrdMst on OmCoCd=IdCoCd  and OmTc= IdExpOdTc and OmYy= IdExpOdYy " + _
              " and OmChr=IdExpOdChr and OmNo= IdExpOdNo " + _
              "where " + IIF(gs_Partition = ctCurrPrtn, " OlPrtKey=OdPrtKey and ", "") + " OLCoCd = OdCoCd and " + _
              "OLTc = OdTc And OLYy = OdYy And OLChr =OdChr and OLNo= OdNo and OLSr= OdSr),0) as qLabVal, " + _
              "IsNull((Select Sum(" + wIrSetVal + ") from OrdRm SetRm " + _
              IIF(wb_InBaseCur, " Join InvHd on " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey=SetRm.OrPrtKey and ", "") + " InCoCd = SetRm.OrCoCd and InTc = SetRm.OrTc and InYy = SetRm.OrYy and InChr = SetRm.OrChr And InNo = SetRm.OrNo ", "") + _
              "Where " + IIF(gs_Partition = ctCurrPrtn, " SetRm.OrPrtKey='" + ctCurrPrtn + "' and ", "") + " SetRm.OrTc = OdTc And SetRm.OrYy = OdYy " + _
              "and SetRm.OrChr = OdChr and SetRm.OrNo = OdNo and SetRm.OrSr = OdSr), 0) as qSetVal, " + _
              "max(InPrnDt) as qInvDt, max(" + wIdFOB + ") as IdIFOB, max(InExpNo) as InExpNo, max(DmUom) as DmUom, " + qMainMet + ", " + _
              qMainRmCtgDesc + ", " + qDmCtgDesc + ", " + qDmSalCtgDesc + ", " + qDmColDesc + ", " + qDmKtDesc + ", " + qVaCtgDesc + _
              " From OrdDsg " + _
              "join InvHd on InCoCd = OdCoCd and InTc = OdTc and InYy = OdYy and InChr = OdChr And InNo = OdNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=OdPrtKey ", "") + ws_CurJoin + _
               "left outer join OrdRm Ir on OrOdIdNo=OdIdNo and OrCoCd = OdCoCd and OrTc = OdTc and OrYy = OdYy and " + _
              "OrChr = OdChr And OrNo = OdNo And OrSr = OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OdPrtKey ", "") + _
              "left outer Join RmMst Rm on RmCd = OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + _
              "join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp = OdDmTcTyp and DmCd = OdDmCd and DmSz = '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + " " + _
              "where (Select PValue from Param where Ptyp= 'TC' and PMCd= odTc)= 'FG' " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + wCnd + _
              " group by OdPrtKey, OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr " + _
              " order by qTag, IdTc, IdChr, IdNo, IdSr "
  End If
  '*** Jay 3.3(FG)
  
  '" where IsNull(RmCtg, '') in ('G', 'P', '') "
 
  '*** (Bef 2.13)
  'If wFlagLocalExp = "Y" Then
  '  'Geeta - Emr 2.10.0
  '  If adc("UOINWT") = "Z" Then
  '    'Changed only for Local Exp and Customs
  '    'For Local Exp and MainMetal='Y' InvVal=GoldAsWt*IRt,ZVal=RmZWt*ZRt
  '    'Whereas for Export Invoice InvVal=GoldAsWt*IRt, ZVal=GoldAsWt*(1+GoldLs%)*ZRt
  '
  '    GRP_REP.SetFormula mRep, "wGldRt", "if{ado.qGldWt}<>0 and {ado.qGldType} ='G' then " + _
  '                                     "If Round({ado.qMinGldRt},3)<>Round({ado.qMaxGldRt},3) Then " + _
  '                                     " {ado.qGldVal}/ ({ado.qGldWt}+ ({ado.qGldWt}*{ado.IdGldLs}/100)) " + _
  '                                     " Else Round({ado.qMinGldRt},3) Else 0 "
  '    GRP_REP.SetFormula mRep, "wPltRt", "if{ado.qPltWt}<>0 and {ado.qPltType} ='P' then " + _
  '                                     "If Round({ado.qMinPltRt},3)<>Round({ado.qMaxPltRt},3) Then " + _
  '                                     " {ado.qPltVal}/ ({ado.qPltWt}+ ({ado.qPltWt}*{ado.IdGldLs}/100)) " + _
  '                                     " Else Round({ado.qMinPltRt},3) Else 0 "
  '    'Geeta - Emr 2.10.0
  '    ' ****** Sachin - 2.12 - 05/12/05 - SJM
  '    GRP_REP.SetFormula mRep, "wSlvRt", "if{ado.qSlvWt}<>0 and {ado.qSlvType} ='S' then " + _
  '                                     "If Round({ado.qMinSlvRt},3)<>Round({ado.qMaxSlvRt},3) Then " + _
  '                                     " {ado.qSlvVal}/ ({ado.qSlvWt}+ ({ado.qSlvWt}*{ado.IdGldLs}/100)) " + _
  '                                     " Else Round({ado.qMinSlvRt},3) Else 0 "
  '    ' ****** Sachin - 2.12 - 05/12/05 - SJM
  '  Else
  '    GRP_REP.SetFormula mRep, "wGldRt", "if{ado.qGldWt}<> 0 then {ado.qGldVal}/ ({ado.qGldWt}+ ({ado.qGldWt}*{ado.IdGldLs}/100))"
  '    GRP_REP.SetFormula mRep, "wPltRt", "if{ado.qPltWt}<> 0 then {ado.qPltVal}/ ({ado.qPltWt}+ ({ado.qPltWt}*{ado.IdGldLs}/100))"
  '    ' ****** Sachin - 2.12 - 05/12/05 - SJM
  '    GRP_REP.SetFormula mRep, "wSlvRt", "if{ado.qSlvWt}<> 0 then {ado.qSlvVal}/ ({ado.qSlvWt}+ ({ado.qSlvWt}*{ado.IdGldLs}/100))"
  '    ' ****** Sachin - 2.12 - 05/12/05 - SJM
  '  End If
  '  'GRP_REP.SetFormula mRep, "wGldRt", "if{ado.qGldWt}<>0 then {ado.qGldVal}/ ({ado.qGldWt}+ {@wGldLsWt})"
  '  'GRP_REP.SetFormula mRep, "wPltRt", "if{ado.qPltWt}<>0 then {ado.qPltVal}/ ({ado.qPltWt}+ {@wPltLsWt})"
  '
  '  GRP_REP.SetFormula mRep, "wLabVal", "{ado.qLabVal} + {ado.qSetVal}+ {ado.qIdAdd}"
  'Else
  '  GRP_REP.SetFormula mRep, "wGldRt", "if{ado.qGldWt}<> 0 then {ado.qGldVal}/{ado.qGldWt}"
  '  GRP_REP.SetFormula mRep, "wPltRt", "if{ado.qPltWt}<> 0 then {ado.qPltVal}/{ado.qPltWt}"
  '  ' ****** Sachin - 2.12 - 05/12/05 - SJM
  '  GRP_REP.SetFormula mRep, "wSlvRt", "if{ado.qSlvWt}<> 0 then {ado.qSlvVal}/{ado.qSlvWt}"
  '  ' ****** Sachin - 2.12 - 05/12/05 - SJM
  '  GRP_REP.SetFormula mRep, "wLabVal", "{ado.qLabVal} + {ado.qSetVal}"
  'End If
  '*** (Bef 2.13)
  
  '*** (Jen 2.13) should Local Invoicing have a different formula ???
  If wFlagLocalExp = "Y" Then
'    If adc("UOINWT") = "Z" Then
'      GRP_REP.SetFormula Rep, "wGldRt", "if {ado.qGldWt}<>0 and {ado.qGldType} ='G' then " + _
'                                       "If Round({ado.qMinGldRt},3)<>Round({ado.qMaxGldRt},3) Then " + _
'                                       " {ado.qGldVal}/ ({ado.qGldWt}+ {ado.qGldLsWt}) " + _
'                                       " Else Round({ado.qMinGldRt},3) Else 0 "
'      GRP_REP.SetFormula Rep, "wPltRt", "if {ado.qPltWt}<>0 and {ado.qPltType} ='P' then " + _
'                                       "If Round({ado.qMinPltRt},3)<>Round({ado.qMaxPltRt},3) Then " + _
'                                       " {ado.qPltVal}/ ({ado.qPltWt}+ {ado.qPltLsWt}) " + _
'                                       " Else Round({ado.qMinPltRt},3) Else 0 "
'      GRP_REP.SetFormula Rep, "wSlvRt", "if {ado.qSlvWt}<>0 and {ado.qSlvType} ='S' then " + _
'                                       "If Round({ado.qMinSlvRt},3)<>Round({ado.qMaxSlvRt},3) Then " + _
'                                       " {ado.qSlvVal}/ ({ado.qSlvWt}+ {ado.qSlvLsWt}) " + _
'                                       " Else Round({ado.qMinSlvRt},3) Else 0 "
'    Else
      GRP_REP.SetFormula Rep, "wGldRt", "if {ado.qGldWt}<> 0 then {ado.qGldVal}/ ({ado.qGldWt}+ {ado.qGldLsWt})"
      GRP_REP.SetFormula Rep, "wPltRt", "if {ado.qPltWt}<> 0 then {ado.qPltVal}/ ({ado.qPltWt}+ {ado.qPltLsWt})"
      GRP_REP.SetFormula Rep, "wSlvRt", "if {ado.qSlvWt}<> 0 then {ado.qSlvVal}/ ({ado.qSlvWt}+ {ado.qSlvLsWt})"
'    End If
    'GRP_REP.SetFormula mRep, "wGldRt", "if{ado.qGldWt}<>0 then {ado.qGldVal}/ ({ado.qGldWt}+ {@wGldLsWt})"
    'GRP_REP.SetFormula mRep, "wPltRt", "if{ado.qPltWt}<>0 then {ado.qPltVal}/ ({ado.qPltWt}+ {@wPltLsWt})"
  
    GRP_REP.SetFormula Rep, "wLabVal", "{ado.qLabVal} + {ado.qSetVal}+ {ado.qIdAdd}"
  Else
    GRP_REP.SetFormula Rep, "wGldRt", "if{ado.qGldWt}<> 0 then {ado.qGldVal}/{ado.qGldWt}"
    GRP_REP.SetFormula Rep, "wPltRt", "if{ado.qPltWt}<> 0 then {ado.qPltVal}/{ado.qPltWt}"
    GRP_REP.SetFormula Rep, "wSlvRt", "if{ado.qSlvWt}<> 0 then {ado.qSlvVal}/{ado.qSlvWt}"
    GRP_REP.SetFormula Rep, "wLabVal", "{ado.qLabVal} + {ado.qSetVal}"
  End If
  '*** (Jen 2.13) should Local Invoicing have a different formula ???
  
  '*** Jay 2.14 ***
  GRP_REP.SetFormula Rep, "wAccRt", "if {ado.qAccWt}<> 0 then {ado.qAccVal}/{ado.qAccWt}"
  '*** Jay 2.14 ***
  
  '*** (Jen 2.13) Should Local & Seepz invoice have same formula ???
  'GRP_REP.SetFormula mRep, "wGldRt", "if{ado.qGldWt}<> 0 then {ado.qGldVal}/{ado.qGldWt}"
  'GRP_REP.SetFormula mRep, "wPltRt", "if{ado.qPltWt}<> 0 then {ado.qPltVal}/{ado.qPltWt}"
  'GRP_REP.SetFormula mRep, "wSlvRt", "if{ado.qSlvWt}<> 0 then {ado.qSlvVal}/{ado.qSlvWt}"
  'GRP_REP.SetFormula mRep, "wLabVal", "{ado.qLabVal} + {ado.qSetVal}"
  '*** (Jen 2.13) Should Local & Seepz invoice have same formula ???
  
  ' Zubin 213 (ws_CurJoin added)
  '*** (Jen 2.13)
  If Trim(UCase(moCn.GetFldVal("Select min(InShowGldLs) from InvHd " + ws_CurJoin + "where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' " + wCnd))) = "N" Then
    Rep.hIdGldLs.Suppress = True
    Rep.FldIdGldLs.Suppress = True
  End If
  '*** (Jen 2.13)
  'Hunnar.1 if Show CustDsgcd/Show PoNo not selected then PoNo will move to the top of that section
  'and section height  reduced
  If UCase(adc("UoYN2")) = "N" Or UCase(adc("UoYN3")) = "N" Then
    Rep.FldPoNo.Top = 0
    Rep.CustDmCdSec.Height = 186
  End If
  
  adc.RepRecSource = wSqlStrg1 + wSqlStrg2 + wSqlStrg3 + wSqlStrg4 + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{ado.RmCd}", 7)
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
  '*** DemarkSec is necessary for the lines to extend till the page footer
  
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
'4.1.2
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0       ' Tag = "ShowPic"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  
  Case Is = 1      ' tag = "SortbyCust"
      If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  
  Case Is = 2      ' tag = "ShowCustDsg"
      If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  
  'Hunnar.1 - Show PO No values stored
  Case Is = 3      ' tag = "ShowPONo"
      If .Value = Checked Then adc("UoYN3") = "Y" Else adc("UoYN3") = "N"
      
  End Select
End With
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
    Set moCn = adc.Connection  '*** (Bef speed) bef 24/11/06
    Set adc.FirNKeyCtl = adc("UoInTcFr")
    '*** Jay 3.3(FG)
    gs_InTcTyp = IIF(UCase(adc.MenuCd) <> UCase("RepFgPckLst"), "IN", "FG")
    '*** Jay 3.3(FG)
    Call InitProp(Me)
    '*** Jay 3.3(FG)
    Select Case UCase(adc.MenuCd)
    Case Is = UCase("RepFgPckLst")
      adc("UoInWt") = "I"
      adc("UoInWt").Visible = False: adc("UoInWt").Enabled = False: LblWtAs.Visible = False
      
      adc("UoYN").Visible = False: adc("UoYN").Enabled = False
      ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False
    
      adc("UoYN2").Visible = False: adc("UoYN2").Enabled = False
      ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False
      
      'Hunnar.1 - Show PO No scope is invisible for FG Pck Lst report
      adc("UoYN3").Visible = False: adc("UoYN3").Enabled = False
      ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False
    
    End Select
    '*** Jay 3.3(FG)

    'Hunnar.1 - "Rm Det Level" scope visible for Packing list(Invoice) report only
    If UCase(adc.MenuCd) = UCase("RepPckLstInv") Then
      LblRmGrp.Visible = True
      adc("UoRmGrp").Visible = True
    Else
      LblRmGrp.Visible = False
      adc("UoRmGrp").Visible = False
    End If

    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** Jay 3.3(FG)
  gs_InTcTyp = IIF(UCase(adc.MenuCd) <> UCase("RepFgPckLst"), "IN", "FG")
  '*** Jay 3.3(FG)
  Select Case UCase(IdName)
  'Hunnar.1 checking valid Rm detail level
  Case Is = UCase("UoRmGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                              " PTyp= 'RMGRP' and PMCd ='DTA' and PSCd='" + pv_NewValue + "'")
    ErrMsg = "Invalid Rm Detail Level": Exit Sub
         
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Jay 3.3(FG)
  gs_InTcTyp = IIF(UCase(adc.MenuCd) <> UCase("RepFgPckLst"), "IN", "FG")
  '*** Jay 3.3(FG)
  Select Case UCase(IdName)
  'Hunnar.1 calling help of DTA options
  Case Is = UCase("UoRmGrp")
    gs_RmGrp = "DTA"
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
    adc("UoCoCdFr") = gs_CoCd
' **** Zubin 211 **** '
  If adc("UoInDet") = "" Then adc("UoInDet") = "INVSR"
' **** Zubin 211 **** '
  '*** Jay 3.3(FG)
  If UCase(adc.MenuCd) = UCase("RepFgPckLst") Then adc("UoInWt") = "I"
  '*** Jay 3.3(FG)
  
  '4.1.2
  If adc("UOYN") = "" Then adc("UOYN") = "N"
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UOYN1") = "" Then adc("UOYN1") = "N"
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If adc("UOYN2") = "" Then adc("UOYN2") = "N"
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  'Hunnar.1 - Restore Show PO No scope value
  If adc("UOYN3") = "" Then adc("UOYN3") = "N"
  If adc("UoYN3") = "Y" Then ChkBoxArr(3).Value = Checked Else ChkBoxArr(3).Value = Unchecked
  
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
    'Set mRep = Nothing
    Set Rep = Nothing
    '*** (09/08/05)
End Sub
