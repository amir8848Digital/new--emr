VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpRWrk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rework Report"
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
      Left            =   11670
      TabIndex        =   56
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
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5625
      TabIndex        =   65
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   66
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
         TabIndex        =   67
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
      Left            =   6840
      TabIndex        =   64
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   61
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
      TabIndex        =   58
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
      TabIndex        =   60
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   59
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpRwrk.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpRwrk.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   63
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   55
               Top             =   360
               Width           =   9945
               _ExtentX        =   17542
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraPtsVal 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1530
               TabIndex        =   98
               Top             =   6330
               Visible         =   0   'False
               Width           =   5385
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4215
               TabIndex        =   6
               ToolTipText     =   "Enter To Process Location"
               Top             =   1170
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
               CmpStr          =   "Lc.LocPrcs <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   7080
               TabIndex        =   7
               ToolTipText     =   "Enter Prrocess Location Selection"
               Top             =   1170
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
               CmpStr          =   "Lc.LocPrcs In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1440
               TabIndex        =   5
               ToolTipText     =   "Enter From Process Location"
               Top             =   1170
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
               CmpStr          =   "Lc.LocPrcs >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4215
               TabIndex        =   15
               ToolTipText     =   "Enter To Design Category"
               Top             =   2025
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGTO"
               IdName          =   "UODMCTGTO"
               CmpStr          =   "DmCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   7080
               TabIndex        =   16
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2025
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
               CmpStr          =   "DmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   1440
               TabIndex        =   14
               ToolTipText     =   "Enter From Design Category"
               Top             =   2025
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGFR"
               IdName          =   "UODMCTGFR"
               CmpStr          =   "DmCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4215
               TabIndex        =   21
               ToolTipText     =   "Enter To Design Code"
               Top             =   2595
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "BOdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   7080
               TabIndex        =   22
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2595
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "BOdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1440
               TabIndex        =   20
               ToolTipText     =   "Enter From Design Code"
               Top             =   2595
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "BOdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4215
               TabIndex        =   47
               ToolTipText     =   "Enter To Customer"
               Top             =   4590
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "OmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   7080
               TabIndex        =   48
               ToolTipText     =   "Enter Customer Selection"
               Top             =   4590
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1440
               TabIndex        =   46
               ToolTipText     =   "Enter From Customer"
               Top             =   4590
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   4215
               TabIndex        =   50
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   4875
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOTO"
               IdName          =   "UOPONOTO"
               CmpStr          =   "OmPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   7080
               TabIndex        =   51
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   4875
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   1440
               TabIndex        =   49
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   4875
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOFR"
               IdName          =   "UOPONOFR"
               CmpStr          =   "OmPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4215
               TabIndex        =   27
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   3165
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "TDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1440
               TabIndex        =   26
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   3165
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "TDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   4995
               TabIndex        =   42
               ToolTipText     =   "Enter To Order Year"
               Top             =   4020
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "OdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2280
               TabIndex        =   39
               ToolTipText     =   "Enter From Order Year"
               Top             =   4020
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "OdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1440
               TabIndex        =   38
               ToolTipText     =   "Enter From Order Tc"
               Top             =   4020
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "OdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   2760
               TabIndex        =   40
               ToolTipText     =   "Enter From Order Character"
               Top             =   4020
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "OdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   4215
               TabIndex        =   41
               ToolTipText     =   "Enter To Order Tc"
               Top             =   4020
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "OdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   5520
               TabIndex        =   43
               ToolTipText     =   "Enter To Order Character"
               Top             =   4020
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "OdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   1440
               TabIndex        =   44
               ToolTipText     =   "Enter From Order Number"
               Top             =   4305
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "OdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   4215
               TabIndex        =   45
               ToolTipText     =   "Enter To Order Number"
               Top             =   4305
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "OdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   4215
               TabIndex        =   9
               ToolTipText     =   "Enter Rework Reason Code To"
               Top             =   1455
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UORWRKTO"
               IdName          =   "UORWRKTO"
               CmpStr          =   "TrwRwrkCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   7080
               TabIndex        =   10
               ToolTipText     =   "Enter Rework Reason Selection"
               Top             =   1455
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORWRKSEL"
               IdName          =   "UORWRKSEL"
               CmpStr          =   "TrwRwrkCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1440
               TabIndex        =   8
               ToolTipText     =   "Enter From Rework Reason Code From"
               Top             =   1455
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UORWRKFR"
               IdName          =   "UORWRKFR"
               CmpStr          =   "TrwRwrkCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4215
               TabIndex        =   12
               ToolTipText     =   "Enter To By Worker"
               Top             =   1740
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKTO"
               IdName          =   "UOWRKTO"
               CmpStr          =   "TrwByWrk <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   7080
               TabIndex        =   13
               ToolTipText     =   "Enter By Worker Selection"
               Top             =   1740
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOWRKSEL"
               IdName          =   "UOWRKSEL"
               CmpStr          =   "TrwByWrk In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1440
               TabIndex        =   11
               ToolTipText     =   "Enter From By Worker"
               Top             =   1740
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKFR"
               IdName          =   "UOWRKFR"
               Data            =   "ABCDW"
               CmpStr          =   "TrwByWrk >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4215
               TabIndex        =   18
               ToolTipText     =   "Enter To Design Production Category"
               Top             =   2310
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMPRDCTGTO"
               IdName          =   "UODMPRDCTGTO"
               CmpStr          =   "DmPrdCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   7080
               TabIndex        =   19
               ToolTipText     =   "Enter Design Production Category Selection"
               Top             =   2310
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMPRDCTGSEL"
               IdName          =   "UODMPRDCTGSEL"
               CmpStr          =   "DmPrdCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   1440
               TabIndex        =   17
               ToolTipText     =   "Enter From Design Production Category"
               Top             =   2310
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMPRDCTGFR"
               IdName          =   "UODMPRDCTGFR"
               CmpStr          =   "DmPrdCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   4215
               TabIndex        =   24
               ToolTipText     =   "Enter To Karat"
               Top             =   2880
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTTO"
               IdName          =   "UOKTTO"
               CmpStr          =   "OdKt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   7080
               TabIndex        =   25
               ToolTipText     =   "Enter Karat Selection"
               Top             =   2880
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "OdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1440
               TabIndex        =   23
               ToolTipText     =   "Enter From Karat"
               Top             =   2880
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               Data            =   "ABCDW"
               CmpStr          =   "OdKt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   4995
               TabIndex        =   32
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   3450
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
               CmpStr          =   "TrwYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2280
               TabIndex        =   29
               ToolTipText     =   "Enter From Transaction Year"
               Top             =   3450
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
               CmpStr          =   "TrwYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1440
               TabIndex        =   28
               ToolTipText     =   "Enter From Transaction Tc"
               Top             =   3450
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
               CmpStr          =   "TrwTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2760
               TabIndex        =   30
               ToolTipText     =   "Enter From Transaction Character"
               Top             =   3450
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
               CmpStr          =   "TrwChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   4215
               TabIndex        =   31
               ToolTipText     =   "Enter To Transaction Tc"
               Top             =   3450
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
               CmpStr          =   "TrwTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   5450
               TabIndex        =   33
               ToolTipText     =   "Enter To Transaction Character"
               Top             =   3450
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
               CmpStr          =   "TrwChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5160
               TabIndex        =   37
               ToolTipText     =   "Enter To Transaction Serial"
               Top             =   3735
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRTO"
               IdName          =   "UOTDSRTO"
               CmpStr          =   "TrwSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   2450
               TabIndex        =   35
               ToolTipText     =   "Enter From Transaction Serial"
               Top             =   3735
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRFR"
               IdName          =   "UOTDSRFR"
               CmpStr          =   "TrwSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1440
               TabIndex        =   34
               ToolTipText     =   "Enter From Transaction Number"
               Top             =   3735
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOFR"
               IdName          =   "UOTNOFR"
               CmpStr          =   "TrwNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   4215
               TabIndex        =   36
               ToolTipText     =   "Enter To Transaction Number"
               Top             =   3735
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOTO"
               IdName          =   "UOTNOTO"
               CmpStr          =   "TrwNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   76
               Left            =   1515
               TabIndex        =   52
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   6030
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOBAGTYP"
               IdName          =   "UOBAGTYP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   77
               Left            =   4215
               TabIndex        =   53
               ToolTipText     =   "Enter Bag Character"
               Top             =   6030
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRFR"
               IdName          =   "UOBCHRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   78
               Left            =   1515
               TabIndex        =   54
               ToolTipText     =   "Enter ""Y"" for Production Points and ""N"" for Rs. Value"
               Top             =   6075
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   79
               Left            =   4215
               TabIndex        =   0
               ToolTipText     =   "Enter To Company Code"
               Top             =   600
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "TdCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   80
               Left            =   7080
               TabIndex        =   1
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   600
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "TdCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   81
               Left            =   4215
               TabIndex        =   3
               ToolTipText     =   "Enter To Parent Code"
               Top             =   885
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOPRNTCDTO"
               IdName          =   "UOPRNTCDTO"
               CmpStr          =   "LcPrc.LocPrntCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   82
               Left            =   7080
               TabIndex        =   4
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   885
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
               CmpStr          =   "LcPrc.LocPrntCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   83
               Left            =   1440
               TabIndex        =   2
               ToolTipText     =   "Enter From Parent Code"
               Top             =   885
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOPRNTCDFR"
               IdName          =   "UOPRNTCDFR"
               CmpStr          =   "LcPrc.LocPrntCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1440
               TabIndex        =   100
               ToolTipText     =   "Enter From Company Code"
               Top             =   600
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "TdCoCd >="
            End
            Begin VB.Label LblPrntCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Parent Loc"
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
               Left            =   90
               TabIndex        =   99
               Top             =   885
               Width           =   1425
            End
            Begin VB.Label LblBagTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Enter Bag Typ"
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
               Left            =   90
               TabIndex        =   97
               ToolTipText     =   "Location"
               Top             =   6030
               Width           =   1395
            End
            Begin VB.Label LblUoBChrFr 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Chr"
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
               Left            =   3330
               TabIndex        =   96
               Top             =   6030
               Width           =   885
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
               Left            =   90
               TabIndex        =   95
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch Tc/Yy/Chr"
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
               Index           =   11
               Left            =   90
               TabIndex        =   94
               Top             =   3465
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch No/Sr"
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
               Index           =   12
               Left            =   90
               TabIndex        =   93
               Top             =   3750
               Width           =   1605
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
               Index           =   28
               Left            =   5355
               TabIndex        =   92
               Top             =   3465
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
               Index           =   29
               Left            =   2655
               TabIndex        =   91
               Top             =   3465
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
               Index           =   32
               Left            =   4935
               TabIndex        =   90
               Top             =   3465
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
               Index           =   33
               Left            =   2200
               TabIndex        =   89
               Top             =   3465
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
               Index           =   34
               Left            =   5055
               TabIndex        =   88
               Top             =   3750
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
               Index           =   35
               Left            =   2320
               TabIndex        =   87
               Top             =   3750
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Kt"
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
               Left            =   90
               TabIndex        =   86
               Top             =   2895
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Ctg"
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
               Left            =   90
               TabIndex        =   85
               Top             =   2325
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "By Worker"
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
               Left            =   90
               TabIndex        =   84
               Top             =   1755
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rework Code"
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
               Left            =   90
               TabIndex        =   83
               Top             =   1470
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Ord Tc/Yy/Chr"
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
               Index           =   22
               Left            =   90
               TabIndex        =   82
               Top             =   4035
               Width           =   1725
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Ord No"
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
               Index           =   21
               Left            =   90
               TabIndex        =   81
               Top             =   4320
               Width           =   1605
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
               Index           =   24
               Left            =   5355
               TabIndex        =   80
               Top             =   4035
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
               Left            =   2625
               TabIndex        =   79
               Top             =   4035
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
               Index           =   26
               Left            =   4935
               TabIndex        =   78
               Top             =   4035
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
               Left            =   2205
               TabIndex        =   77
               Top             =   4035
               Width           =   105
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
               Left            =   90
               TabIndex        =   76
               Top             =   3180
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "PO No"
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
               Index           =   18
               Left            =   90
               TabIndex        =   75
               Top             =   4890
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
               Left            =   90
               TabIndex        =   74
               Top             =   4605
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Cd"
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
               Index           =   14
               Left            =   90
               TabIndex        =   73
               Top             =   2610
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Ctg"
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
               Left            =   90
               TabIndex        =   72
               Top             =   2040
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Process"
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
               Left            =   90
               TabIndex        =   71
               Top             =   1185
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
               Left            =   7200
               TabIndex        =   70
               Top             =   120
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
               Left            =   4215
               TabIndex        =   69
               Top             =   120
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
               Left            =   1515
               TabIndex        =   68
               Top             =   120
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpRWrk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepRwrk
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Dim moGrph As EmrClsGrph
Dim ws_MultiCoMod As Boolean

Private Sub SetGroupSort()
  GRP_REP.Add "Process", "TdFrBLoc", "", "TdFrBLoc", "hTdFrBLoc", "Company Code,Parent Code", "", "", "IsNull((Select LocDesc From Loc Where LocCoCd = TdCoCd and LocTyp= 'P' and LocCd= TdFrBLoc), '') "
  GRP_REP.Add "Rework Code", "TdFrBLoc+TrwRwrkCd", "'('+ RTrim(TdFrBLoc)+ ') '+ TrwRwrkCd", "TdFrBLoc,TrwRwrkCd", "hTdFrBLoc,hTrwRwrkCd", "Parent Code,Rework", "", "", "IsNull((Select PDesc From Param Where PTyp= 'RWRK' and PMCd= TdFrBLoc and PSCd= TrwRwrkCd), '') "
  GRP_REP.Add "(Process)Worker", "TdFrBLoc+TrwByWrk", "'('+ RTrim(TdFrBLoc)+ ') '+ TrwByWrk", "TdFrBLoc,TrwByWrk", "hTdFrBLoc,hTrwByWrk", "Parent Code,Process", "", "", Array("IsNull((Select vPDesc From vParam Where vPCoCd = TrwCoCd and vPTyp= 'WORK' and vPMCd= TrwByWrk), '') ", "IsNull((Select '(Rs Val:' + LTrim(Str(vPNum* vPNum1, 7, 2)) + ') '+ vPDesc From vParam Where vPCoCd = TrwCoCd and vPTyp= 'WORK' and vPMCd= TrwByWrk), '') ")
  
  GRP_REP.Add "Design Ctg", "DmCtg", "", "DmCtg", "hDmCtg", "", "", "", "IsNull((Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg), '') "
  GRP_REP.Add "Prd Ctg", "DmPrdCtg", "", "DmPrdCtg", "hDmPrdCtg", "", "", "", "IsNull((Select PDesc from Param where PTyp= 'PRDCTG' and PMCd= DmPrdCtg), '') "
  
  GRP_REP.Add "Sal Ctg", "DmSalCtg", "", "DmSalCtg", "", "", "", "", "IsNull((Select PDesc from Param where PTyp= 'SALCTG' And PMCd=DmCtg And PSCd=DmSalCtg), '') "
  
  GRP_REP.Add "Design Code", "BOdDmCd", "", "BOdDmCd", "hBOdDmCd", "Design Ctg,Prd Ctg,(DsgCd)Colour", "", "", "IsNull((Select Substring(DmDesc,1,120) From DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' and ", "") + " DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd and DmSz= ''), '') "
  GRP_REP.Add "Ord Kt", "OdKt", "", "OdKt", "hOdKt", "", "", "", "IsNull((Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt), '') "
  GRP_REP.Add "OrderNo", "BOdTc+BOdYy+BOdChr+str(BOdNo)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "wOrdNo,OmCmCd,OmPoNo", "hOrdNo,hOmCmCd,hOmPoNo", "Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "IsNull((Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd), '') "
  GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))"
'  GRP_REP.Add "Rm Ctg", "TpRmCtg", "", "TpRmCtg", "hTpRmCtg", "", "", "", "IsNull((Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = TpRmCtg), '') "
'  GRP_REP.Add "Set Type", "TpSetTyp", "TpSetTyp", "TpSetTyp", "hSetTyp", "", "", "", "IsNull((Select PDesc From Param Where PTyp= 'LABSCD' and PMCd= 'SET' and PSCd = TpSetTyp), '') "
'  GRP_REP.Add "Rm SubCtg", "TpRmCtg+TpRmSCtg", "'('+ RTrim(TpRmCtg)+ ') '+ TpRmSCtg", "TpRmCtg,TpRmSCtg", "hTpRmCtg,hTpRmSCtg", "Rm Ctg", "", "", "IsNull((Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = TpRmCtg and PSCd= TpRmSCtg), '') "
  GRP_REP.Add "(DsgCd) Colour", "BOdDmCd+OdDmCol", "'('+BOdDmCd+') '+OdDmCol", "BOdDmCd,OdDmCol", "hBOdDmCd,hOdDmCol", "Design Ctg,Prd Ctg,Design Code", "", "", ""
  '6.1
  GRP_REP.Add "Date", "convert(char(10),TDt,112)", "convert(varchar, TDt," + CStr(DtFmtSlahVal()) + ")", "", "", "", "", "", ""
  
  '4.1.0.0 (Sachin) - Addl. Group Sorts based on Date(TDt)
  GRP_REP.Add "Week", "LTrim(Str(DatePart(yyyy,TDt)))+' - '+DateName(WW,TDt)", "", "", "", "Quarter,Month", "", "", ""
  GRP_REP.Add "Month", "Str(DatePart(yyyy,TDt))+Str(DatePart(mm,TDt))", "LTrim(Str(DatePart(yyyy,TDt)))+' '+DateName(MM,TDt)", "", "", "Quarter", "", "", ""
  GRP_REP.Add "Quarter", "LTrim(Str(DatePart(yyyy,TDt)))+' - '+DateName(QQ,TDt)", "", "", "", "", "", "", ""
  GRP_REP.Add "Yearly", "Str(DatePart(yyyy,TDt))", "LTrim(Str(DatePart(yyyy,TDt)))", "", "", "", "", "", ""
  
  '### GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "wOrdNo,OmCmCd,OmPoNo", "hOrdNo,hOmCmCd,hOmPONo", "Customer Code,OrderNo", "", "", "'('+ OmCmCd+ ')'+ BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))"
  If ws_MultiCoMod = True Then
    GRP_REP.Add "Company Code", "TdCoCd", "TdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=TdCoCd) "
    GRP_REP.Add "Parent Location", "Lc.LocPrntCd", "Lc.LocPrntCd", "", "", "", "", "", "(Select Ln.LocDesc from Loc Ln where Ln.LocCd=Lc.LocPrntCd and Ln.LocCoCd='" + ctSelfCoCd + "') "
  End If
  
'  moGrph.AddItem "Prd Quantity", "TpRmQty", "{rdo.TpRmQty}", ""
'  moGrph.AddItem "Prd Points", "TpPts", "{rdo.TpPts}", ""
    
    GRP_REP.Add "Bag Type", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", _
      "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", "", "", "Bag Typ+Chr", "", "", "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')"
    GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
    GRP_REP.Add "Bag Typ+Chr", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+BChr", _
      "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+'/ '+ BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
End Sub
Private Sub adc_setreprecsource()
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String
      
  Dim ws_BagJoin As String
'  Dim ws_TpPts As String
  Dim ws_vParamJoin As String
  Dim ws_ParamJoin As String
  
  Dim wAddlCnd As String   ' ******  New Fields in Where Clause
    
  Set Rep = Nothing
  Set ADC.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' *****  Microway footer
  Call GRP_REP.Gen3LRep
  
  grpflds = GRP_REP.GrpFldLst
    
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  
  ws_BagJoin = "Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                    IIF(ADC("UoBagTyp") <> "", " and PValue= '" + ADC("UoBagTyp") + "' ", " ") + _
                    IIF(ADC("UOBCHRFR") <> "", " and PMCd= '" + ADC("UOBCHRFR") + "' ", " ")
  
  '****** Additional Conditions For New Fields
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(ADC("UOCMCDFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(ADC("UOCMCDTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(ADC("UOCMCDSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(ADC("UODTFR")) <> "01/01/80", " TrwDt >= '" + Trim(ADC("UODTFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(ADC("UODTTO")) <> "01/01/80", " TrwDt <= '" + Trim(ADC("UODTTO")) + "'", "")
  '****** Additional Conditions For New Fields

  ws_vParamJoin = " Join vParam on vPTyp= 'WORK' and vPMCd= TrwByWrk  and vPSCd= '' and TrwCoCd= vPCoCd "
  ws_ParamJoin = "  Join Param Pm on Pm.PTyp= 'RWRK' and Pm.PMCd= LocPrcs and Pm.PSCd= TrwRwrkCd "
  
  Rep.TxtFrVchDt.SetText Format(ADC("UODTFR"), DtShortStrYyyy())  '6.1
  Rep.TxtToVchDt.SetText Format(ADC("UODTTO"), DtShortStrYyyy())
  'sv.36 TrwRwrkCd '-' changed as 'FTP', LcPrc join added
  wSqlStrg = "Select " + grpflds + ", " + _
             "TdCoCd, TdFrBLoc, TrwRwrkCd, TrwByWrk, TrwQty, " + _
             "TrwTc, TrwYy, TrwChr, TrwNo, TrwSr, TrwSrNo, TDt, TdBYy, TdBChr, TdBNo, " + _
             "BOdTc, BOdYy, BOdChr, BOdNo, BOdDmCd, BQty, DmCtg, DmPrdCtg, OdKt, OdDmCol, " + _
             "OmCmCd, OmPoNo, (case TrwRwrkCd when 'FTP' then 'FTP' else 'OTH' end) as qFtpOth, Pm.Pdesc as qRwrkDesc From TxndRwrk " + _
             "Join Txnd on TdIdNo = TrwTdIdNo and TdCoCd= TrwCoCd and TdTc= TrwTc and TdYy= TrwYy and TdChr= TrwChr " + _
             "and TdNo= TrwNo and TrwSr= TdSr and TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=TrwPrtKey ", "") + _
             "Join Txn on TCoCd= TrwCoCd and TTc= TrwTc and TYy= TrwYy and TChr= TrwChr and TNo= TrwNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TrwPrtKey ", "") + _
             "Join Bag on BIdNo=TdBIdNo and BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=TdPrtKey ", "") + ws_BagJoin + _
             "Join Loc Lc on LocCocd= TdCoCd and LocCd=TdFrBLoc " + _
             "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
             "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
             "Join OrdMst on OmIdNo=BOmIdNo and OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy " + _
             "and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=BPrtKey ", "") + _
             "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=DmPrtKey ", "") + _
             ws_vParamJoin + ws_ParamJoin + _
             " Join Loc LcPrc on LcPrc.LocCocd= TdCoCd and LcPrc.LocTyp ='R' and LcPrc.LocCd=Lc.LocPrcs " + _
             wCnd + IIF(gs_Partition = ctCurrPrtn, " and TrwPrtKey='" + ctCurrPrtn + "' ", "") + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + _
             " Order By Grp1, Grp2, Grp3, TdCoCd+TdBYy+TdBChr+Str(TdBNo, 10),TrwByWrk,TrwRwrkCd "
             
  ADC.RepRecSource = wSqlStrg + ctMaxDopOpt

  If ADC("UoBagTyp") = "F" Then
    Rep.TxtBagTypAndChr.SetText "(Flute Bags" & IIF(ADC("UoBChrFr") <> "", ",  Chr = " & ADC("UoBChrFr"), "") & ")"
  ElseIf ADC("UoBagTyp") = "P" Then
    Rep.TxtBagTypAndChr.SetText "(Primary Bags" & IIF(ADC("UoBChrFr") <> "", ",  Chr = " & ADC("UoBChrFr"), "") & ")"
  ElseIf ADC("UoBagTyp") = "C" Then
    Rep.TxtBagTypAndChr.SetText "(Component Bags" & IIF(ADC("UoBChrFr") <> "", ",  Chr = " & ADC("UoBChrFr"), "") & ")"
  End If
  
  CRV_REP.DisplayGroupTree = False

  Dim wsQryStruct As String
  wsQryStruct = "  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3, " + _
    "space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3, " + _
    "space(50) as G1Desc, space(50) as G2Desc, space(50) as G3Desc, " + _
    "TdFrBLoc, TrwRwrkCd, TrwByWrk, TrwQty, TrwTc, TrwYy, TrwChr, TrwNo, " + _
    "TrwSr, TrwSrNo, TDt, TdBYy, TdBChr, TdBNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdDmCd, BQty, DmCtg , " + _
    "DmPrdCtg , OdKt, OdDmCol, OmCmCd, OmPoNo " + _
    "From TxndRwrk, Txnd, Txn, Bag, OrdDsg, OrdMst, DsgMst where 1=2"
'  Call moGrph.SetGraphProp(Rep, wsQryStruct, wSqlStrg + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt

  
  Call DispReport(ADC)
  
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
  
 
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = False
 
  If ws_MultiCoMod = True Then
    Set ADC.FirNKeyCtl = ADC("UoCoCdFr")
  Else
    Set ADC.FirNKeyCtl = ADC("UoIssLocFr")
  End If
  
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_DmTcTyp = "DM"
  gs_TxnTcTyp = "DT"
  gs_OmTcTyp = "SO"
  gs_CmCtg = "C"
  gs_LabMCd = "SET"
  
'  Set moGrph = New EmrClsGrph
'  Call moGrph.SetFormDets(Me, moCn)  '*** (Bef speed) bef 24/11/06
  'Call moGrph.SetFormDets(Me)  '*** (Jen speed) 24/11/06
  
  
  ADC("UOCOCDFR").Enabled = ws_MultiCoMod
  LblPrntCd.Visible = ws_MultiCoMod
  ADC("UOCOCDFR").Enabled = ws_MultiCoMod: ADC("UOCOCDTO").Visible = ws_MultiCoMod
  ADC("UOCOCDSEL").Enabled = ws_MultiCoMod: ADC("UOCOCDSEL").Visible = ws_MultiCoMod
  ADC("UoPrntCdFr").Enabled = ws_MultiCoMod: ADC("UoPrntCdTo").Enabled = ws_MultiCoMod: ADC("UoPrntCdSel").Enabled = ws_MultiCoMod
  ADC("UoPrntCdFr").Visible = ws_MultiCoMod: ADC("UoPrntCdTo").Visible = ws_MultiCoMod: ADC("UoPrntCdSel").Visible = ws_MultiCoMod
  
  Call SetGroupSort
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
'  If UCase$(IdName) = UCase$("UoGrphTyp") Then gs_GrphGrpNo = moGrph.SetGrphGrpNo
  
  Select Case UCase(IdName)
    Case Is = UCase("UoBChrFr")
      If ADC.Mode = xNorm Then
        If ADC("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Bag Type": Exit Sub
        HlpList.PMCd "BCHR", "'" + ADC("UoBagTyp") + "'": Exit Sub
      End If
  End Select
  gs_LocTyp = "'R'"   'sv.36 location type changed
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
    Case Is = UCase("UOBCHRFR")
      ADC("UOBCHRFR") = ""
  End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
'Private Sub OptPtsVal_Click(Index As Integer)
'Select Case Index
'  Case Is = 0
'    adc("UoYN3") = "Y"
'  Case Is = 1
 '   adc("UoYN3") = "N"
'End Select
'End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

Private Sub tab_REP_Click(PreviousTab As Integer)
  fra_tabrep(TAB_REP.Tab).Enabled = True
  fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  ADC("UOUSRCD") = ADC.UsrCd
  ADC("UOMNUCD") = ADC.MenuCd
  ADC("UoBagTyp") = "P"  '********Bhavna
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  If ws_MultiCoMod = False Then _
      ADC("UoCoCdFr") = gs_CoCd: ADC("UoCoCdTo") = gs_CoCd: ADC("UoCoCdSel") = "": ADC("UoCoCdFr").Enabled = False

 
  Me.Caption = GetFrmCaption("", "Rework Report")
'  If adc("UoShowGrph") = "" Then adc("UoShowGrph") = "N"
'  Call moGrph.NonKeyWhen

  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + ctSelfCoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    ADC("UoBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    ADC("UoBChrFr").Visible = False
    ADC("UoBagTyp") = "P": ADC("UoBChrFr") = ""
  End If
  Set wHeadChk = Nothing

End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = False Or UCase(ADC.MenuCd) = UCase("RepWipFgBPcs"), ADC("UOCOCDFR"), ctSelfCoCd) + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
  
  Rep.wCoCd.SetText ws_HName
  
  Rep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    Rep.RHLogo.Suppress = True
  End If
  
  Rep.TxtCoCdFr.SetText ADC("UOCOCDFR")
  If ws_MultiCoMod = False Then
      Rep.TxtCoCdTo.Suppress = True: Rep.TxtCoCdSel.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True: Rep.TxtCoCdSelLbl.Suppress = True
  Else
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdTo.SetText ADC("UOCOCDTO"): Rep.TxtCoCdSel.SetText ADC("UOCOCDSEL")
  End If
  
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  Rep.TxtRepUsrDtTime.SetText Trim(ADC.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  
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

Private Sub xxx()
'    wStr = GRP_REP.GrpOpts
'    If InStr(1, "," + wStr + ",", "," + "Kt" + ",") = 0 Then
'        t3_netwt.Suppress = True: t2_netwt.Suppress = True: t1_netwt.Suppress = True: g_netwt.Suppress = True
'    Else
'        warr = Split(wStr)
'        For i = 0 To 2
'           If warr(i) = "Kt" Then
'              OptPos = i
'              Exit For
'            End If
'        Next i
'        wSupGrand = False
'        For i = 0 To 2
'          If i < OptPos Then
'            Dim xxx As CRAXDRT.Report
'            xxx.Sections("G" + CStr(i) + "F").ReportObjects("T" + CStr(i) + "NetWt").Suppress = True
'            wSupGrand = True
'          End If
'        Next i
'        g_netwt.Suppress = True
'    End If
End Sub
Private Sub GRP_REP_LostFocus()
'  Call moGrph.SetDefaVal
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
'  moGrph.NonKeyValid pr_Cancel, pr_ErrMsg
End Sub
Private Sub TAB_GRPH_Click(PreviousTab As Integer)
'    FraTabGrph(TAB_GRPH.Tab).Enabled = True
'    FraTabGrph(PreviousTab).Enabled = False
End Sub
Private Sub ChkGrph_Click()
'  If ChkGrph.Value = vbChecked Then
'    adc("UoShowGrph") = "Y"
'    FraGrphDet.Enabled = True
'    moGrph.SetDefaVal
'  ElseIf ChkGrph.Value = vbUnchecked Then
'    adc("UoShowGrph") = "N"
'    FraGrphDet.Enabled = False
'  End If
End Sub
Private Sub cboGrphValue_Click()
'  adc("UoGrphValInd") = cboGrphValue.ListIndex
End Sub

