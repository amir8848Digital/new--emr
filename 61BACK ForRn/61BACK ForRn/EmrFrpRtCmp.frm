VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpRtCmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raw Material Rate Chart Comparision"
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
      Left            =   10620
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4395
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4785
      TabIndex        =   28
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   29
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
         TabIndex        =   30
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
      Left            =   6000
      TabIndex        =   27
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   24
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
      TabIndex        =   18
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
      TabIndex        =   23
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   20
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpRtCmp.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpRtCmp.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   26
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   22
               Top             =   360
               Width           =   9765
               _ExtentX        =   17224
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   83
            TabIndex        =   25
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraOptSrcCustRt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   4785
               TabIndex        =   45
               ToolTipText     =   "Compare with Source Customer's Cost Rt Or Sales Rate "
               Top             =   2580
               Width           =   3195
               Begin VB.OptionButton OptSrcCustRt 
                  Caption         =   "Sales Rate"
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
                  Left            =   1500
                  TabIndex        =   16
                  ToolTipText     =   "Compare with Source Customer's Cost Rt Or Sales Rate "
                  Top             =   0
                  Width           =   1425
               End
               Begin VB.OptionButton OptSrcCustRt 
                  Caption         =   "Cost Rate"
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
                  TabIndex        =   15
                  ToolTipText     =   "Compare with Source Customer's Cost Rt Or Sales Rate "
                  Top             =   0
                  Width           =   1275
               End
            End
            Begin VB.Frame FraOptRt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   4785
               TabIndex        =   44
               ToolTipText     =   "Compare with Customer's Cost Rt Or Sales Rate "
               Top             =   1140
               Width           =   2985
               Begin VB.OptionButton OptCustRt 
                  Caption         =   "Cost Rate"
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
                  TabIndex        =   3
                  ToolTipText     =   "Compare with Customer's Cost Rt Or Sales Rate "
                  Top             =   0
                  Width           =   1245
               End
               Begin VB.OptionButton OptCustRt 
                  Caption         =   "Sales Rate"
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
                  Left            =   1500
                  TabIndex        =   4
                  ToolTipText     =   "Compare with Customer's Cost Rt Or Sales Rate "
                  Top             =   0
                  Width           =   1425
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   3645
               TabIndex        =   10
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   2010
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "RrCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6120
               TabIndex        =   11
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   2010
               Width           =   8910
               _ExtentX        =   15716
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "RrCd in "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   3645
               TabIndex        =   7
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   1725
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGTO"
               IdName          =   "UORMSCTGTO"
               CmpStr          =   "RrSCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   6120
               TabIndex        =   8
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   1725
               Width           =   8910
               _ExtentX        =   15716
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "RrSCtg in "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1215
               TabIndex        =   5
               ToolTipText     =   "Enter Raw Material Category"
               Top             =   1440
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "RrCtg = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1215
               TabIndex        =   6
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   1725
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
               CmpStr          =   "RrSCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1215
               TabIndex        =   9
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   2010
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "RrCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   1215
               TabIndex        =   13
               ToolTipText     =   "Enter Source Customer Code"
               Top             =   2580
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               ReCalcParent    =   "UOYN"
               IdName          =   "UOCMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1215
               TabIndex        =   1
               ToolTipText     =   "Enter Customer Code"
               Top             =   1140
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               ReCalcParent    =   "UOYN1"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "RrCmCd ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1215
               TabIndex        =   0
               ToolTipText     =   "Enter Customer Category"
               Top             =   840
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOCMCTGFR"
               IdName          =   "UOCMCTGFR"
               CmpStr          =   "RrCmCtg = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1215
               TabIndex        =   17
               ToolTipText     =   "Enter Source YearMonth"
               Top             =   2880
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
               Index           =   0
               Left            =   1215
               TabIndex        =   12
               ToolTipText     =   "Enter Source Customer Category"
               Top             =   2310
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOCMCTGTO"
               IdName          =   "UOCMCTGTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   7905
               TabIndex        =   14
               ToolTipText     =   "Compare with Source Customer's Cost Rt Or Sales Rate ('Y' for Cost / 'N' For Sale)"
               Top             =   2580
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               ReCalcOn        =   "UOCMCDTO"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   7785
               TabIndex        =   2
               ToolTipText     =   "Compare with Customer's Cost Rt Or Sales Rate ('Y' for Cost / 'N' For Sale)"
               Top             =   1140
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               ReCalcOn        =   "UOCMCDFR"
               IdName          =   "UOYN1"
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Compare"
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
               Left            =   3645
               TabIndex        =   43
               Top             =   1140
               Width           =   855
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Compare"
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
               Left            =   3645
               TabIndex        =   42
               Top             =   2580
               Width           =   975
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Src CmCtg"
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
               TabIndex        =   41
               ToolTipText     =   "Location"
               Top             =   2310
               Width           =   1305
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
               Index           =   0
               Left            =   60
               TabIndex        =   40
               Top             =   1140
               Width           =   1305
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
               Index           =   4
               Left            =   60
               TabIndex        =   39
               ToolTipText     =   "Location"
               Top             =   840
               Width           =   1305
            End
            Begin VB.Label LblRhYyyyMm 
               BackStyle       =   0  'Transparent
               Caption         =   "Src YyyyMm"
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
               TabIndex        =   38
               Top             =   2880
               Width           =   1305
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Src Cust"
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
               TabIndex        =   37
               Top             =   2580
               Width           =   1305
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm Ctg"
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
               TabIndex        =   36
               ToolTipText     =   "Location"
               Top             =   1440
               Width           =   1305
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm SubCtg"
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
               Left            =   60
               TabIndex        =   35
               Top             =   1725
               Width           =   1305
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm Code"
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
               Left            =   60
               TabIndex        =   34
               Top             =   2010
               Width           =   1305
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
               Left            =   6120
               TabIndex        =   33
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
               Left            =   3645
               TabIndex        =   32
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
               Left            =   1215
               TabIndex        =   31
               Top             =   480
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpRtCmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim Rep As New EmrRepRmRtCmp
Private Sub SetGroupSort()
  'GRP_REP.Add "Design Ctg", "RrDmCtg", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= RrDmCtg) "
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  'Select RrCmCtg, RrCmCd, RrCmCurCd, RrCtg, RrSCtg, RrCd, RrFrLn, RrToLn, RrSalRt as qNweRt,
  'RrCmCtg as qSrcCmCtg, RrCmCd as qSrcCmCd, RhYyyymm as qRhYyyymm,
  'RrSalRt As qSrcRt
  'from RmRt, RmRtHist where 1=2
  '*** Report Sql ***
  Dim grpflds As String, wSqlStrg As String, wSrcRt As String, wSrcStr As String, wrepcnd As String
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  wrepcnd = adc.RepCond
  wrepcnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")

  If adc("UoYyMmFr") = "0" Then
    wSrcRt = IIF(adc("UoYn") = "Y", "RrCstRt", "RrSalRt")
    wSrcStr = "IsNull((Select " + wSrcRt + " from RmRt Old where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrCmCtg= '" + adc("UoCmCtgTo") + "' and " + _
            "RrCmCd= '" + adc("UoCmCdTo") + "' and " + _
            "Old.RrTcTyp= New.RrTcTyp and Old.RrCmCurCd= New.RrCmCurCd and " + _
            "Old.RrCtg=New.RrCtg and Old.RrSCtg=New.RrSCtg " + _
            "and Old.RrCd=New.RrCd and RrFrLn = New.RrFrLn and RrToLn = New.RrToLn) " + _
            ", 0) as qSrcRt "
  Else
    wSrcRt = IIF(adc("UoYn") = "Y", "RhCstRt", "RhSalRt")
    wSrcStr = "IsNull((Select " + wSrcRt + " from RmRtHist Old where " + IIF(gs_Partition = ctCurrPrtn, " RhPrtKey='" + ctCurrPrtn + "' and ", "") + " RhCmCtg= '" + adc("UoCmCtgTo") + "' and " + _
            "RhCmCd= '" + adc("UoCmCdTo") + "' and RhYyyymm = " + CStr(adc("UoYyMmFr")) + " and " + _
            "Old.RhTcTyp= New.RrTcTyp and Old.RhCmCurCd= New.RrCmCurCd and " + _
            "Old.RhCtg=New.RrCtg and Old.RhSCtg=New.RrSCtg " + _
            "and Old.RhCd=New.RrCd and RhFrLn = New.RrFrLn and RhToLn = New.RrToLn) " + _
            ", 0) as qSrcRt "
  End If
  ' **** Manali 3.03 - 25/06/08 - Cst Rt included for New RmRt
  wSqlStrg = "Select New.RrCmCtg, New.RrCmCd, New.RrCmCurCd, New.RrCtg, New.RrSCtg, New.RrCd, New.RrFrLn, " + _
             "New.RrToLn, " + IIF(adc("UoYn1") = "Y", "New.RrCstRt", "New.RrSalRt") + " as qNweRt, '" + adc("UoCmCtgTo") + "' as qSrcCmCtg, " + _
             "'" + adc("UoCmCdTo") + "' as qSrcCmCd, " + CStr(adc("UoYyMmFr")) + " as qRhYyyymm, " + wSrcStr + _
             "from RmRt New where " + IIF(gs_Partition = ctCurrPrtn, " New.RrPrtKey='" + ctCurrPrtn + "' and ", "") + " New.RrTcTyp= 'RM' " + wrepcnd
             
  ' **** bef  3.03
  'wSqlStrg = "Select New.RrCmCtg, New.RrCmCd, New.RrCmCurCd, New.RrCtg, New.RrSCtg, New.RrCd, New.RrFrLn, " + _
             "New.RrToLn, New.RrSalRt as qNweRt, '" + adc("UoCmCtgTo") + "' as qSrcCmCtg, " + _
             "'" + adc("UoCmCdTo") + "' as qSrcCmCd, " + CStr(adc("UoYyMmFr")) + " as qRhYyyymm, " + wSrcStr + _
             "from RmRt New where " + IIF(gs_Partition = ctCurrPrtn, " New.RrPrtKey='" + ctCurrPrtn + "' and ", "") + " New.RrTcTyp= 'RM' " + wrepcnd
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  If adc("UoYN1") = "Y" Then Rep.TxtHeadNewRt.SetText "New CstRt"
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.RrCd}", 9)
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
Select Case UCase(IdName)
' **** Manali 3.03 - 25/06/08
  Case Is = UCase("UoYn1")
    If (adc("UoCmCdFr") <> "ZSELF" And adc("UoCmCdFr") <> "ZVALUE") Then adc("UoYn1") = "N"
  Case Is = UCase("UoYn")
    If (adc("UoCmCdTo") <> "ZSELF" And adc("UoCmCdTo") <> "ZVALUE") Then adc("UoYn") = "N"
' **** Manali 3.03 - 25/06/08
End Select

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
    gs_PTyp = "RRTCTYP"
    gs_Tbl = "Param"
    gb_ShowYyMmHlp = True     '*** Emr2-06
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    Call SetGroupSort
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  gs_CmCtg = adc("UoCmCtgFr")
  adc("UoYn1") = "N"      ' ******** Manali 3.03 - New Customer Cost Rt Default
  
  ' ***** Manali 3.5.0 - 06/11/08 - Radio buttons added
  If adc("UoYN1") = "Y" Then OptCustRt(0).Value = True
  If adc("UoYN1") = "N" Then OptCustRt(1).Value = True
  If adc("UoYN") = "Y" Then OptSrcCustRt(0).Value = True
  If adc("UoYN") = "N" Then OptSrcCustRt(1).Value = True
  ' ***** Manali 3.5.0 - 06/11/08 - Radio buttons added
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_Tbl = "Param"
  Select Case UCase(IdName)
  Case Is = UCase("UoCmCtgFr")
    gs_CmCtg = pv_NewValue
    
  Case Is = UCase("UoCmCdFr")
    ' ***** Manali 3.5.0 - 06/11/08 - Radio buttons added
    If pv_NewValue <> "ZSELF" And pv_NewValue <> "ZVALUE" Then OptCustRt(1).Value = True: adc("UoYN1") = "N"
    
    Case Is = UCase("UoCmCdTo")
    ' ***** Manali 3.5.0 - 06/11/08 - Radio buttons added
    If pv_NewValue <> "ZSELF" And pv_NewValue <> "ZVALUE" Then OptSrcCustRt(1).Value = True: adc("UoYN") = "N"
  End Select
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_Tbl = "Param"
  Select Case UCase(IdName)
    Case Is = UCase("UoYn1")
      Cancel = (adc("UoCmCdFr") <> "ZSELF" And adc("UoCmCdFr") <> "ZVALUE")
      If Cancel = True Then ErrMsg = "Cannot edit Value When Customer is not 'ZSELF' Or 'ZVALUE'": Exit Sub
  End Select
  SetProp Me, IdName, When
End Sub
' ***** Manali 3.5.0 - 06/11/08 - Radio buttons added
Private Sub OptCustRt_Click(Index As Integer)
Select Case Index
Case Is = 0
  If adc("UoCmCdFr") <> "ZSELF" And adc("UoCmCdFr") <> "ZVALUE" Then
    OptCustRt(0).Value = False: OptCustRt(1).Value = True: adc("UoYN1") = "N"
    DispMsg "Cannot Select Cost Rate When Customer is not 'ZSELF' Or 'ZVALUE'", etError
  Else
    adc("UoYN1") = "Y"
  End If
Case Is = 1
  adc("UoYN1") = "N"
End Select
End Sub
' ***** Manali 3.5.0 - 06/11/08 - Radio buttons added

' ***** Manali 3.5.0 - 06/11/08 - Radio buttons added
Private Sub OptCustRt_GotFocus(Index As Integer)
  DispMsg FraOptRt.ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 06/11/08 - Radio buttons added

' ***** Manali 3.5.0 - 06/11/08 - Radio buttons added
Private Sub OptSrcCustRt_Click(Index As Integer)
Select Case Index
Case Is = 0
  If adc("UoCmCdFr") <> "ZSELF" And adc("UoCmCdFr") <> "ZVALUE" Then
    OptSrcCustRt(0).Value = False: OptSrcCustRt(1).Value = True: adc("UoYN") = "N"
    DispMsg "Cannot Select Cost Rate When Source Customer is not 'ZSELF' Or 'ZVALUE'", etError
  Else
    adc("UoYN") = "Y"
  End If
Case Is = 1
  adc("UoYN") = "N"
End Select
End Sub
' ***** Manali 3.5.0 - 06/11/08 - Radio buttons added

' ***** Manali 3.5.0 - 06/11/08 - Radio buttons added
Private Sub OptSrcCustRt_GotFocus(Index As Integer)
  DispMsg FraOptSrcCustRt.ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 06/11/08 - Radio buttons added

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    adc("UOUSRCD") = adc.UsrCd
    adc("UOMNUCD") = adc.MenuCd
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


