VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpInvBoeLst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice BOE Listing"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11730
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4410
      TabIndex        =   21
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
      Left            =   5835
      TabIndex        =   29
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   30
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
         TabIndex        =   31
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
      Left            =   7050
      TabIndex        =   28
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   25
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
      TabIndex        =   19
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
      Height          =   9615
      Left            =   0
      TabIndex        =   24
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   23
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpInvBoeLst.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpInvBoeLst.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   27
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   20
               Top             =   360
               Width           =   9045
               _ExtentX        =   15954
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   14985
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Box Det                "
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
               Left            =   90
               TabIndex        =   17
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show Box Details"
               Top             =   2595
               Width           =   2025
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1920
               TabIndex        =   13
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   2245
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "M.TzRmCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   2310
               TabIndex        =   16
               ToolTipText     =   "Show Box Detail ? (Yes / No)"
               Top             =   2595
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
               Index           =   20
               Left            =   2700
               TabIndex        =   2
               ToolTipText     =   "Enter Invoice Year"
               Top             =   1080
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "M.TzYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1920
               TabIndex        =   1
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   1080
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "M.TzTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3120
               TabIndex        =   3
               ToolTipText     =   "Enter Invoice Character"
               Top             =   1080
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "M.TzChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1920
               TabIndex        =   7
               ToolTipText     =   "Enter Invoice Number"
               Top             =   1380
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "M.TzNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1920
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   795
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "M.TzCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1920
               TabIndex        =   9
               ToolTipText     =   "Enter From Export Number"
               Top             =   1665
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOFR"
               IdName          =   "UOINEXPNOFR"
               CmpStr          =   "InExpNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   1920
               TabIndex        =   18
               ToolTipText     =   "Show Values In Base Currency (Yes / No)"
               Top             =   2880
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   4680
               TabIndex        =   14
               ToolTipText     =   "EnterTo  Raw Material Category"
               Top             =   2245
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
               CmpStr          =   "M.TzRmCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   7290
               TabIndex        =   15
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   2245
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCTGSEL"
               IdName          =   "UORMCTGSEL"
               CmpStr          =   "M.TzRmCtg In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   5460
               TabIndex        =   5
               ToolTipText     =   "Enter Invoice Year"
               Top             =   1080
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYTO"
               IdName          =   "UOINYYTO"
               CmpStr          =   "M.TzYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   4680
               TabIndex        =   4
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   1080
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCTO"
               IdName          =   "UOINTCTO"
               CmpStr          =   "M.TzTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   5880
               TabIndex        =   6
               ToolTipText     =   "Enter Invoice Character"
               Top             =   1080
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRTO"
               IdName          =   "UOINCHRTO"
               CmpStr          =   "M.TzChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4680
               TabIndex        =   8
               ToolTipText     =   "Enter Invoice Number"
               Top             =   1380
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOTO"
               IdName          =   "UOINNOTO"
               CmpStr          =   "M.TzNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4680
               TabIndex        =   10
               ToolTipText     =   "Enter From Export Number"
               Top             =   1665
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOTO"
               IdName          =   "UOINEXPNOTO"
               CmpStr          =   "InExpNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4680
               TabIndex        =   12
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   1960
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "InDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1920
               TabIndex        =   11
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   1960
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "InDt >="
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch Date"
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
               Index           =   19
               Left            =   120
               TabIndex        =   45
               Top             =   1970
               Width           =   1425
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
               Index           =   8
               Left            =   5820
               TabIndex        =   44
               Top             =   1080
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
               Index           =   7
               Left            =   5400
               TabIndex        =   43
               Top             =   1080
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv No"
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
               Left            =   120
               TabIndex        =   42
               Top             =   1380
               Width           =   1815
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
               Index           =   6
               Left            =   1890
               TabIndex        =   41
               Top             =   330
               Width           =   1335
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
               Left            =   4680
               TabIndex        =   40
               Top             =   330
               Width           =   1275
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
               Left            =   7320
               TabIndex        =   39
               Top             =   330
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "In Base Currency"
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
               Left            =   120
               TabIndex        =   38
               ToolTipText     =   "Location"
               Top             =   2880
               Visible         =   0   'False
               Width           =   1755
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
               Left            =   120
               TabIndex        =   37
               Top             =   1665
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
               Left            =   120
               TabIndex        =   36
               Top             =   795
               Width           =   1335
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
               TabIndex        =   35
               Top             =   1080
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
               TabIndex        =   34
               Top             =   1080
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Tc/Yy/Chr"
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
               Left            =   120
               TabIndex        =   33
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label LblRmCtg 
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
               Left            =   120
               TabIndex        =   32
               ToolTipText     =   "Location"
               Top             =   2270
               Width           =   1425
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpInvBoeLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rep As New EmrRepInvBOELst
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  GRP_REP.Add "VoucherNo", "B.TzTc+B.TzYy+B.TzChr+str(B.TzNo)", "B.TzTc+'/'+B.TzYy+'/'+B.TzChr+'/'+LTrim(Str(B.TzNo))", "", "", "", "", "", ""
  '6.1 BoeNo, Boe date
  GRP_REP.Add "BOENo", "TBOENo", "", "TBOENo", "hTBOENo", "BOE Date", "", "", "(case when TBOEDt= '01/01/80' then '' else convert(varchar, TBOEDt, " + CStr(DtFmtSlahYyyyVal()) + ") end)"
  GRP_REP.Add "BOE Date", "convert(char(8),TBOEDt,112)", "(case when TBOEDt= '01/01/80' then '' else convert(varchar, TBOEDt, " + CStr(DtFmtSlahYyyyVal()) + ") end)", "TBOEDt", "hTBOEDt", "", "", "", ""
  GRP_REP.Add "Rm Code", "RmCtg+RmSCtg+B.TzRmZCd", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ B.TzRmZCd", "QbRmCd", "hTzRmZCd", "", "", "", "(Select RmDesc From RmMst Where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= B.TzRmZCd) "
  GRP_REP.Add "Imp Type", "TImpTyp", "TImpTyp", "TImpTyp", "hTImpTyp", "", "", "", "(Select PDesc From Param Where Ptyp ='IMPTYP' and  PmCd= TImpTyp) "
  GRP_REP.Add "Rm Desc", "B.TzRmCdDesc", "B.TzRmCdDesc", "QbRmCd", "hTzRmZCd", "", "", "", "(Select RmDesc From RmMst Where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= B.TzRmCdDesc) "
'  GRP_REP.Add "Supplier", "TSuppCd", "", "", "", "", "", "", "(Select CmName From CustMst Where CmCtg= 'S' and CmCd= TSuppCd) "
'  GRP_REP.Add "Supp Bill No", "TBillNo", "", "", "", "Supplier", "", "", "(case when TBillDt='01/01/80' then '' else convert(varchar, TBillDt, 103) end)"
'  GRP_REP.Add "Vch Date", "convert(char(8),TDt,112)", "convert(varchar, TDt, 103)", "", "", "", "", "", ""
'  GRP_REP.Add "Purchase Rt", "str(B.TzPurRt,9,2)", "", "TzPurRt", "hTzPurRt", "", "", "", ""
'  GRP_REP.Add "(RmCd) PurRt", "B.TzRmZCd+str(B.TzPurRt,9,2)", "'('+B.TzRmZCd+') '+LTrim(str(B.TzPurRt,9,2))", "TzRmZCd,TzPurRt", "hTzRmZCd,hTzPurRt", "Rm Code,Purchase Rt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
   ' *** Manali 3.03 - 17/06/08 - Join with Param included and RmCtg, Pdesc added
'    Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'    space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'    space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'    TzRmCtg as GRmCtg, PDesc as GRmCtgDesc ,
'    TzTc as qBTc, TzYy as qBYy, TzChr as qBChr, TzNo as qBNo, TzSr as qBSr,
'    TzRmZCd as qBRmCd, TzRmCdDesc as qBRmCdDesc, TzRmQty as qBRmQty, TzRmWt as qBRmWt,
'    TzPurRt as qBPurRt, TzPurAmt as qBPurAmt, TzRmDespWt as qBDespWt, TzTc as qMTc,
'    TzYy as qMYy, TzChr as qMChr, TzNo as qMNo, TzSr as qMSr, TzSrNo as qMSrNo,
'    TzRmQty as qMRmQty, TzRmWt as qMRmWt, TzPurAmt as qMPurAmt, TDt, TSuppCd, TBillNo,
'    TImpTyp, TBillDt, TBOENo, TBOEDt, TCnvRt, TCurrCdm
'    from TxndZ, Txn, RmMst, Param where 1 = 2
' *** Report Sql ***
  
  Dim wrepcnd As String, wCnd As String, i As Integer, wDetPos As Integer, wSqlStrg As String
  Dim grpflds As String
  Dim wExpNo As String
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
 
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  '*** To check if Detail is selected ***
  Dim wGrp(6) As String
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then wDetPos = i: Exit For
  Next i
  
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then
    Rep.BoxHd.Suppress = True: Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True: Rep.DetSec2.Suppress = True
  End If
 
  '*** (Bef 2.13)
  'wSqlStrg = "Set Dateformat DMY Select " + grpflds + ", B.TzTc as qBTc, B.TzYy as qBYy, B.TzChr as qBChr, " + _
  '           "B.TzNo as qBNo, B.TzSr as qBSr, B.TzRmZCd as qBRmCd, " + _
  '           "B.TzRmCdDesc as qBRmCdDesc, B.TzRmQty as qBRmQty, B.TzRmWt as qBRmWt, " + _
  '           "B.TzPurRt as qBPurRt, B.TzPurAmt as qBPurAmt, B.TzRmDespWt as qBDespWt, " + _
  '           "M.TzTc as qMTc, M.TzYy as qMYy, M.TzChr as qMChr, M.TzNo as qMNo, " + _
  '           "M.TzSr as qMSr, M.TzSrNo as qMSrNo, M.TzRmQty as qMRmQty, " + _
  '           "M.TzRmWt as qMRmWt, M.TzPurAmt as qMPurAmt, TDt, TSuppCd, TBillNo, TImpTyp, TBillDt, " + _
  '           "TBOENo, TBOEDt, TCnvRt from TxndZ M " + _
  '           "join InvHd on InCoCd= M.TzCoCd and InTc= M.TzTc and InYy= M.TzYy and InChr= M.TzChr " + _
  '           "and InNo= M.TzNo " + _
  '           "join TxndZ B on B.TzCoCd= M.TzCoCd and B.TzYy= M.TzRefYy and B.TzKey = M.TzRefKey " + _
  '           "join Txn on TCoCd= B.TzCoCd and TTc= B.TzTc and TYy= B.TzYy " + _
  '           "and TChr= B.TzChr and TNo= B.TzNo " + _
  '           "join RmMst on RmCd= B.TzRmZCd where M.TzBM= 'M' " + wCnd
  '*** (Bef 2.13)

  '****** (Bef 2.13) M.TzPurAmt as qMPurAmt,  TCnvRt,  ******
  '*** (Jen 2.13) added TCurrCd
  '****** Sachin 3.02 - Id fields in Joins
  '*** Manali 3.03 - 17/06/08 - Join with Param included and RmCtg, Pdesc added
  '6.1
  wSqlStrg = " Set Dateformat " + DtFmtStr() + " Select " + grpflds + ", M.TzRmCtg as GRmCtg, PDesc as GRmCtgDesc , B.TzTc as qBTc, B.TzYy as qBYy, B.TzChr as qBChr, " + _
             "  B.TzNo as qBNo, B.TzSr as qBSr,  B.TzRmZCd as qBRmCd, " + _
             "  B.TzRmCdDesc as qBRmCdDesc, B.TzRmQty as qBRmQty, B.TzRmWt as qBRmWt, " + _
             "  B.TzPurRt as qBPurRt, B.TzPurAmt as qBPurAmt, B.TzRmDespWt as qBDespWt, " + _
             "  M.TzTc as qMTc, M.TzYy as qMYy, M.TzChr as qMChr, M.TzNo as qMNo, " + _
             "  M.TzSr as qMSr, M.TzSrNo as qMSrNo, M.TzRmQty as qMRmQty, M.TzRmWt as qMRmWt, " + _
             "  (M.TzRmWt* B.TzPurRt) as qMPurAmt, TDt, TSuppCd, TBillNo, TImpTyp, TBillDt, TBOENo, TBOEDt, " + _
             "  (case when TMulDivRs= 'D' and convert (decimal (16, 4), TCnvRtRs)> 0 then 1/ TCnvRtRs else TCnvRtRs end) as TCnvRt, " + _
             " TCurrCd from TxndZ M " + _
             " join InvHd on InIdNo=M.TzInIdNo and InCoCd= M.TzCoCd and InTc= M.TzTc and InYy= M.TzYy and InChr= M.TzChr " + _
             "  and InNo= M.TzNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=M.TzPrtKey ", "") + _
             " join TxndZ B on B.TzCoCd= M.TzCoCd and B.TzYy= M.TzRefYy and B.TzKey = M.TzRefKey " + IIF(gs_Partition = ctCurrPrtn, " And B.TzPrtKey=M.TzPrtKey ", "") + _
             " join Txn on TIdNo=B.TzTIdNo and TCoCd= B.TzCoCd and TTc= B.TzTc and TYy= B.TzYy " + _
             "  and TChr= B.TzChr and TNo= B.TzNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey=B.TzPrtKey ", "") + _
             " join RmMst on RmCd= B.TzRmZCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=B.TzPrtKey ", "") + _
             " join Param on PTyp='RmCtg' and PMCd=M.TzRmCtg " + _
             " Where M.TzBM= 'M' " + IIF(gs_Partition = ctCurrPrtn, " And M.TzPrtKey='" + ctCurrPrtn + "' ", "") + wCnd
             
             ' ***** Bef 3.03 - Manali - not reqd now
             ' IIF(wCnd <> "", " And M.TzRmCtg='" + Trim(ADC("UORMCTGFR")) + "'", " Where M.TzRmCtg='" + Trim(ADC("UORMCTGFR")) + "'")
             
  '*** (Jen 2.13)

  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  '3.11.0
  Dim wFrInNo As String, wToInNo As String, wFrExpNo As String, wToExpNo As String
  wFrInNo = "": wToInNo = ""
  If adc("UoInYyFr") <> "" And adc("UoInChrFr") <> "" And adc("UoInNoFr") <> 0 Then _
    wFrInNo = adc("UoCoCdFr") + "\IN\" + adc("UoInYyFr") + "\" + adc("UoInChrFr") + "\" + CStr(adc("UoInNoFr"))
  
  If adc("UoInYyTo") <> "" And adc("UoInChrTo") <> "" And adc("UoInNoTo") <> 0 Then _
    wToInNo = adc("UoCoCdFr") + "\IN\" + adc("UoInYyTo") + "\" + adc("UoInChrTo") + "\" + CStr(adc("UoInNoTo"))
  
  wFrExpNo = ""
  If Not (adc("UoInTcFr") = "" Or adc("UoInYyFr") = "" Or adc("UoInChrFr") = "" Or adc("UoInNoFr") = 0) Or adc("UoInExpNoFr") <> "" Then _
    wFrExpNo = IIF(adc("UoInExpNoFr") <> "", adc("UoInExpNoFr"), moCn.GetFldVal("Select InExpNo from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and " + _
                            "InTc= '" + adc("UoInTcFr") + "' and InYy= '" + adc("UoInYyFr") + "' and " + _
                            "InChr= '" + adc("UoInChrFr") + "' and InNo= " + CStr(adc("UoInNoFr")) + " "))
                          
  wToExpNo = ""
  If Not (adc("UoInTcTo") = "" Or adc("UoInYyTo") = "" Or adc("UoInChrTo") = "" Or adc("UoInNoTo") <> 0) Or adc("UoInExpNoTo") <> "" Then _
    wToExpNo = IIF(adc("UoInExpNoTo") <> "", adc("UoInExpNoTo"), moCn.GetFldVal("Select InExpNo from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and " + _
                            "InTc= '" + adc("UoInTcTo") + "' and InYy= '" + adc("UoInYyTo") + "' and " + _
                            "InChr= '" + adc("UoInChrTo") + "' and InNo= " + CStr(adc("UoInNoTo")) + " "))
  
  '3.11.0
  Rep.TxtFrInvNo.SetText wFrInNo
  Rep.TxtToInvNo.SetText wToInNo
  Rep.TxtFrExpNo.SetText wFrExpNo
  Rep.TxtToExpNo.SetText wToExpNo
  Rep.TxtFrInvDt.SetText IIF(adc("UoDtFr") <> "01/01/80", adc("UoDtFr"), "")
  Rep.TxtToInvDt.SetText IIF(adc("UoDtTo") <> "01/01/80", adc("UoDtTo"), "")
  'Rep.TxtRmCtg.SetText adc("UoRmCtgFr")      ' **** Manali 3.03 - Not Required
  
  If adc("UoYN") = "N" Then
    Rep.BoxHd.Suppress = True
    Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True
    Rep.DetSec2.Suppress = True
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.qBRmCd}", 8)
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.qBRmCdDesc}", 8)
  
  '*** (Jen 2.13)
  'hTzPurRt Cost Rt||(US$)
  'hExpVal  Exp Value||(US$)
  'h1ExpVal
  'hCnvRt
  'hExpValRs  Exp Value||(Rs.)
  'h1ExpValRs Exp Value||(Rs.)
  '*** (Jen 2.13)
    
  CRV_REP.DisplayGroupTree = False
  
Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
End Sub
' ****** Manali 3.5.0 - 19/11/08 - Check boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case .Tag
  Case Is = "ShowBoxDet"    ' Index = 0
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  End Select
End With
End Sub
' ****** Manali 3.5.0 - 19/11/08 - Check boxes added

' ****** Manali 3.5.0 - 19/11/08 - Check boxes added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 19/11/08 - Check boxes added

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
    Set adc.FirNKeyCtl = adc("UoInTcFr") 'adc("UoCoCdFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gb_RmCtgFor = True
    gs_InTcTyp = "IN"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_InTcTyp = "IN"
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_InTcTyp = "IN"
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
  
  ' ****** Manali 3.5.0 - 19/11/08 - Check boxes added
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  ' ****** Manali 3.5.0 - 19/11/08 - Check boxes added
End Sub
Private Sub DispCoNm()
  'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                            "HCd = '" + ctSelfCmCd + "' ")
  Rep.wCoCd.SetText ws_HName
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  Rep.wCoCdLogo.SetText ws_HName
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

