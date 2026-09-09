VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlforspeed8.ocx"
Begin VB.Form EmrFrpVchProfitMrg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profit Margin Report"
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
      Left            =   11430
      TabIndex        =   48
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
      TabIndex        =   47
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
      Left            =   5535
      TabIndex        =   57
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   58
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
         TabIndex        =   59
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
      Left            =   6750
      TabIndex        =   56
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   53
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
      TabIndex        =   49
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
      Height          =   9765
      Left            =   0
      TabIndex        =   52
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   51
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpVchProfitMrg.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpVchProfitMrg.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   55
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   50
               Top             =   360
               Width           =   10155
               _ExtentX        =   17912
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2550
               TabIndex        =   2
               ToolTipText     =   "Enter From Order Year"
               Top             =   1215
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "Od.OdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   2970
               TabIndex        =   3
               ToolTipText     =   "Enter From Order Character"
               Top             =   1215
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "Od.OdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   7830
               TabIndex        =   30
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   3600
               Width           =   7260
               _ExtentX        =   12806
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
               CmpStr          =   "DmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4860
               TabIndex        =   35
               ToolTipText     =   "Enter To Design Code"
               Top             =   4200
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "Od.OdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   7830
               TabIndex        =   36
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   4200
               Width           =   7260
               _ExtentX        =   12806
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "Od.OdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   7830
               TabIndex        =   13
               ToolTipText     =   "Enter Customer Selection"
               Top             =   2070
               Width           =   7260
               _ExtentX        =   12806
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   7830
               TabIndex        =   39
               ToolTipText     =   "Enter Karat Selection"
               Top             =   4500
               Width           =   7260
               _ExtentX        =   12806
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "Od.OdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4860
               TabIndex        =   32
               ToolTipText     =   "Enter To Design Sales Category"
               Top             =   3900
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGTO"
               IdName          =   "UODMSALCTGTO"
               CmpStr          =   "DmSalCtg<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   7830
               TabIndex        =   33
               ToolTipText     =   "Enter Design Sales Category Selection"
               Top             =   3900
               Width           =   7260
               _ExtentX        =   12806
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMSALCTGSEL"
               IdName          =   "UODMSALCTGSEL"
               CmpStr          =   "DmSalCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   7830
               TabIndex        =   22
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   2985
               Width           =   7260
               _ExtentX        =   12806
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPONo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5640
               TabIndex        =   5
               ToolTipText     =   "Enter To Order Year"
               Top             =   1215
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "Od.OdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4860
               TabIndex        =   4
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1215
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "Od.OdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   6060
               TabIndex        =   6
               ToolTipText     =   "Enter To Order Character"
               Top             =   1215
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "Od.OdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   4860
               TabIndex        =   8
               ToolTipText     =   "Enter To Order Number"
               Top             =   1500
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "Od.OdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4860
               TabIndex        =   29
               ToolTipText     =   "Enter To Design Category"
               Top             =   3600
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
               Index           =   40
               Left            =   4860
               TabIndex        =   12
               ToolTipText     =   "Enter To Customer"
               Top             =   2070
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "OmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4860
               TabIndex        =   21
               ToolTipText     =   "Enter To Purchase Order"
               Top             =   2985
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOTO"
               IdName          =   "UOPONOTO"
               CmpStr          =   "OmPONo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   4860
               TabIndex        =   10
               ToolTipText     =   "Enter To Transaction Date"
               Top             =   1785
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "OmDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   7830
               TabIndex        =   42
               ToolTipText     =   "Enter Sales Person Selection"
               Top             =   4785
               Width           =   7260
               _ExtentX        =   12806
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSALPERSEL"
               IdName          =   "UOSALPERSEL"
               CmpStr          =   "OmCmSalPer In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   4860
               TabIndex        =   38
               ToolTipText     =   "Enter To Karat"
               Top             =   4500
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTTO"
               IdName          =   "UOKTTO"
               CmpStr          =   "Od.OdKt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   4860
               TabIndex        =   41
               ToolTipText     =   "Enter To Sales Executive "
               Top             =   4785
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERTO"
               IdName          =   "UOSALPERTO"
               CmpStr          =   "OmCmSalPer<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1770
               TabIndex        =   0
               ToolTipText     =   "Enter Tc Type"
               Top             =   945
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCTYPFR"
               IdName          =   "UOTCTYPFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1770
               TabIndex        =   1
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1230
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "Od.OdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1770
               TabIndex        =   7
               ToolTipText     =   "Enter From Order Number"
               Top             =   1515
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "Od.OdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   1770
               TabIndex        =   28
               ToolTipText     =   "Enter From Design Category"
               Top             =   3600
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
               Index           =   33
               Left            =   1770
               TabIndex        =   34
               ToolTipText     =   "Enter From Design Code"
               Top             =   4200
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "Od.OdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1770
               TabIndex        =   11
               ToolTipText     =   "Enter From Customer"
               Top             =   2085
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   1770
               TabIndex        =   37
               ToolTipText     =   "Enter From Karat"
               Top             =   4515
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               Data            =   "ABCDW"
               CmpStr          =   "Od.OdKt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   1770
               TabIndex        =   31
               ToolTipText     =   "Enter From Design Sales Category"
               Top             =   3900
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGFR"
               IdName          =   "UODMSALCTGFR"
               CmpStr          =   "DmSalCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1770
               TabIndex        =   20
               ToolTipText     =   "Enter From Purchase Order"
               Top             =   2985
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOFR"
               IdName          =   "UOPONOFR"
               CmpStr          =   "OmPONo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   1770
               TabIndex        =   43
               ToolTipText     =   "Specify Whether Base Value Is Cost Or Sale "
               Top             =   5085
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOORDVAL"
               IdName          =   "UOORDVAL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   1770
               TabIndex        =   44
               ToolTipText     =   "Specify Whether Other Value Is Cost Or Sale "
               Top             =   5370
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOORDVAL1"
               IdName          =   "UOORDVAL1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1770
               TabIndex        =   80
               ToolTipText     =   "Enter Company Code"
               Top             =   660
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "Od.OdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1770
               TabIndex        =   9
               ToolTipText     =   "Enter From Transaction Date"
               Top             =   1800
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "OmDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   1770
               TabIndex        =   40
               ToolTipText     =   "Enter From Sales Executive"
               Top             =   4800
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERFR"
               IdName          =   "UOSALPERFR"
               CmpStr          =   "OmCmSalPer >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   1770
               TabIndex        =   45
               ToolTipText     =   "Enter Currency Code Whose Orders You Wish To View Or Leave It Blank If You Wish To View All Orders In Base Currency"
               Top             =   5655
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               IdName          =   "UOCURCDFR"
               CmpStr          =   "OmCmCurCd ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   3060
               TabIndex        =   24
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   3300
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "Od.OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   6140
               TabIndex        =   26
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   3300
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "Od.OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   7830
               TabIndex        =   27
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   3300
               Visible         =   0   'False
               Width           =   7260
               _ExtentX        =   12806
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "Od.OdPoNo in"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   7830
               TabIndex        =   16
               ToolTipText     =   "Enter Customer Group Selection"
               Top             =   2385
               Width           =   7260
               _ExtentX        =   12806
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMGRPSEL"
               IdName          =   "UOCMGRPSEL"
               CmpStr          =   "CmGrp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   4860
               TabIndex        =   15
               ToolTipText     =   "Enter To Customer Group"
               Top             =   2385
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMGRPTO"
               IdName          =   "UOCMGRPTO"
               CmpStr          =   "CmGrp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   1770
               TabIndex        =   14
               ToolTipText     =   "Enter From Customer Group"
               Top             =   2390
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMGRPFR"
               IdName          =   "UOCMGRPFR"
               CmpStr          =   "CmGrp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   7830
               TabIndex        =   19
               ToolTipText     =   "Enter Customer Type Selection"
               Top             =   2685
               Width           =   7260
               _ExtentX        =   12806
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMTYPSEL"
               IdName          =   "UOCMTYPSEL"
               CmpStr          =   "CmTyp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   4860
               TabIndex        =   18
               ToolTipText     =   "Enter To Customer Type"
               Top             =   2685
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMTYPTO"
               IdName          =   "UOCMTYPTO"
               CmpStr          =   "CmTyp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   1770
               TabIndex        =   17
               ToolTipText     =   "Enter From Customer Type"
               Top             =   2690
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMTYPFR"
               IdName          =   "UOCMTYPFR"
               CmpStr          =   "CmTyp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   65
               Left            =   1770
               TabIndex        =   23
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   3300
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTFR"
               IdName          =   "UOODSUBCUSTFR"
               CmpStr          =   "Od.OdSubCust >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   66
               Left            =   4860
               TabIndex        =   25
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   3300
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTTO"
               IdName          =   "UOODSUBCUSTTO"
               CmpStr          =   "Od.OdSubCust <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   118
               Left            =   1770
               TabIndex        =   46
               ToolTipText     =   "Enter Divided By Amount "
               Top             =   5940
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "UOVALDIVBY"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Value Divided By"
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
               Index           =   37
               Left            =   60
               TabIndex        =   88
               Top             =   5940
               Width           =   1665
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
               Index           =   20
               Left            =   7830
               TabIndex        =   87
               Top             =   180
               Width           =   1575
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
               Left            =   6075
               TabIndex        =   86
               Top             =   3300
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
               Index           =   36
               Left            =   3000
               TabIndex        =   85
               Top             =   3300
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Typ"
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
               Left            =   60
               TabIndex        =   84
               Top             =   2700
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Grp"
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
               Left            =   60
               TabIndex        =   83
               Top             =   2400
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "SUB PO No"
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
               Left            =   60
               TabIndex        =   82
               Top             =   3300
               Width           =   1425
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
               TabIndex        =   81
               Top             =   5640
               Width           =   1005
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Executive"
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
               TabIndex        =   79
               Top             =   4785
               Width           =   1785
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order Date"
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
               TabIndex        =   78
               Top             =   1785
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
               Left            =   60
               TabIndex        =   77
               Top             =   645
               Width           =   1335
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Base Value As"
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
               TabIndex        =   76
               Top             =   5070
               Width           =   1845
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Other Value As"
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
               TabIndex        =   75
               Top             =   5355
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
               Index           =   26
               Left            =   5580
               TabIndex        =   74
               Top             =   1215
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
               Index           =   24
               Left            =   6000
               TabIndex        =   73
               Top             =   1215
               Width           =   105
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
               Left            =   4860
               TabIndex        =   72
               Top             =   180
               Width           =   1275
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sal Ctg"
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
               TabIndex        =   71
               Top             =   3900
               Width           =   1365
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
               Left            =   60
               TabIndex        =   70
               Top             =   3600
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
               Left            =   60
               TabIndex        =   69
               Top             =   4215
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
               Left            =   60
               TabIndex        =   68
               Top             =   2070
               Width           =   1425
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
               Left            =   60
               TabIndex        =   67
               Top             =   4500
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
               Index           =   5
               Left            =   60
               TabIndex        =   66
               Top             =   3000
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Type"
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
               TabIndex        =   65
               ToolTipText     =   "Location"
               Top             =   930
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
               Index           =   27
               Left            =   2490
               TabIndex        =   64
               Top             =   1215
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
               Left            =   2910
               TabIndex        =   63
               Top             =   1215
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order No"
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
               Index           =   10
               Left            =   60
               TabIndex        =   62
               Top             =   1500
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order Tc/Yy/Chr"
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
               TabIndex        =   61
               Top             =   1215
               Width           =   1725
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
               Left            =   1770
               TabIndex        =   60
               Top             =   180
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpVchProfitMrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepProfitMrg
Dim ws_hAddJROYN As String
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "", "", "Sales Ctg", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= max(DmCtg)) "
  '***Geeta***Crm**212
  GRP_REP.Add "Sales Ctg", "max(DmSalCtg)", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and Pmcd=max(DmCtg) and PSCd= max(DmSalCtg)) "
  GRP_REP.Add "OrderNo", "OdTc+OdYy+OdChr+str(OdNo)", "OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))", "", "", "Customer,Customer PO,Sales Exec,Customer Grp,Customer Typ", "", "", "max(OmPONo) "
  GRP_REP.Add "Customer PO", "max(OmPONo)", "", "", "", "", "", "", ""
  '*** Jay 3.2.0 [OdPoNo]
  'vk.19 sub customer added
  GRP_REP.Add "Cust SUB PO", "(case when max(OdPoNo)<>'' Or max(OdSubCust)<>'' then '/'+max(OdSubCust)+' : '+max(OdPoNo) else '' end)", "", "", "", "", "", "", "(case when max(OdPoNo)<>'' Or max(OdSubCust)<>'' then '/'+max(OdSubCust)+' : '+max(OdPoNo) else '' end)"
  '*** Jay 3.2.0 [OdPoNo]
  'GRP_REP.Add "Design Code", "max(OdDmCd)", "", "IdDmCd", "hIdDmCd", "Design Ctg,Sales Ctg", "", "", "max(DmDesc) "
  GRP_REP.Add "Design Code", "max(OdDmCd)", "", "", "", "Design Ctg,Sales Ctg", "", "", "max(DmDesc) "
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  
  'Sachin - 4.1.0.0
  GRP_REP.Add "Customer", "max(OmCmCd)", "", "", "", "Sales Exec,Customer Grp,Customer Typ", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmJMDI.TXT_UCD.Text) + "')  <= 5 then  CmName else CmCd end ) From CustMst Where CmCtg= 'C' and CmCd= max(OmCmCd)) "
  
  'Sachin - 4.1.0.0
  GRP_REP.Add "Customer Grp", "max(CmGrp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMGRP' and PmCd = Max(CmGrp))"
  GRP_REP.Add "Customer Typ", "max(CmTyp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMTYP' and PmCd = Max(CmTyp))"
  
  GRP_REP.Add "Karat", "max(OdKt)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = max(OdKt)) "
  GRP_REP.Add "PCS/PRS", "max(DmUom)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMUOM' and PmCd= max(DmUom)) "
  GRP_REP.Add "VaCtg", "max(OdVaCtg)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'VACTG' and PMCd = max(OdVaCtg)) "
  GRP_REP.Add "(DsgCd)Colour", "max(OdDmCd+OdDmCol)", "'('+max(OdDmCd)+')'+max(OdDmCol)", "OdDmCd", "hIdDmCd", "Design Ctg,Sales Ctg,Design Code,Design Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(OdDmCol)) "
  GRP_REP.Add "Design Colour", "max(OdDmCol)", "", "", "", "(DsgCd)Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(OdDmCol)) "
  GRP_REP.Add "Sales Exec", "max(OmCmSalPer)", "", "", "", "", "", "", "(Select PDesc from Param Where PTyp = 'SP' and PMCd= max(OmCmSalPer)) "
  'FldLabVal1
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  ' **** Zubin 100305 (Manoj asked to write in the front end) **** '
'  If UCase(adc("UoOrdVal")) = UCase(adc("UoOrdVal1")) Then
'    pr_ErrMsg = "Base Value and Other Value cannot be same."
'    pr_Cancel = True
'  End If
  ' **** Zubin 100305 **** '
End Sub
Private Sub adc_setreprecsource()
' **** Zubin - Previous Report SQL **** '
  '*** Report Sql ***
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'IdTc, IdYy, IdChr, IdNo, IdSr, IdDmCd, IdQty as IdQty,
  'IdiFOB as qPcVal1, IdiFOB as qPcVal2, IrRmiVal as qMetVal1, IrRmiVal as qMetVal2,
  'IrRmiVal as qDCVal1, IrRmiVal as qDCVal2, IrRmiVal as qXVal1, IrRmiVal as qXVal2,
  'IrSetiVal as qSetVal1, IrSetiVal as qSetVal2, iLiVal as qLabVal1, iLiVal as qLabVal2,
  'IrRmiWt as qPcWt1, IrRmiWt as qPcWt2, IrRmiWt as qMetWt1, IrRmiWt as qMetWt2,
  'IrRmiWt As qDCWt1, IrRmiWt As qDCWt2, IrRmiWt As qXWt1, IrRmiWt As qXWt2
  'From InvDsg, DsgMst, InvHd, InvRm, InvLab, RmMst where 1= 2
  '*** Report Sql ***
'===============================================================================================================

' **** Zubin - New Report SQL **** '
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(45) as DGrp1, space(45) as DGrp2, space(45) as DGrp3,
'  space(45) as G1Desc, space(45) as G2Desc, space(45) as G3Desc,
'  OdTc, OdYy, OdChr, OdNo, OdSr, OdDmCd, OdOrdQty, OdSalPrc as qPcVal1, OdCstPrc as qPcVal2,
'  0 as qMetVal1, 0 as qMetVal2, 0 as qDCVal1, 0 as qDCVal2,  0 as qXVal1, 0 as qXVal2, OrSetSalVal as qSetVal1,
'  OrSetCstVal as qSetVal2, 0 as qLabVal1, 0 as qLabVal2, 0 as qPcWt1, 0 as qPcWt2, 0 as qMetWt1, 0 as qMetWt2,
'  0 as qDCWt1, 0 as qDCWt2, 0 as qXWt1, 0 as qXWt2
'  From OrdDsg, DsgMst, OrdRm, RmMst, OrdMst, CustMst
'  where 1= 2
'===============================================================================================================
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wStr1 As String, wStr2 As String, wSqlStrg As String
  
  Dim wBasePcVal As String, wBaseRmVal As String, wBaseSetVal As String, wBaseLabVal As String
  Dim wOthPcVal As String, wOthRmVal As String, wOthSetVal As String, wOthLabVal As String
  Dim wBase As String, wOth As String, wDiffFirstPart As String, wDiffSecPart As String
  
  Dim wAddlCnd As String '****** Sachin 3.02.0 - New fields in where clause
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
  ws_hAddJROYN = moCn.GetFldVal("Select hAddJROYN from Head where hCoCd='ZZZ'")
  Dim ws_BaseCurCd As String, wb_BaseCur As Boolean
  Dim ws_OmOlJoin As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  If adc("UoCurCdFr") = "" Then
    wb_BaseCur = True
    Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
    ws_OmOlJoin = " join OrdMst on " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey=OlPrtKey and ", "") + " OmCoCd= OlCoCd and OmTc= OlTc and OmYy= OlYy and OmChr= OlChr and OmNo= OlNo "
  Else
    wb_BaseCur = False
    Rep.TxtValIn.SetText "Value In " + adc("UoCurCdFr")
    ws_OmOlJoin = ""
  End If
    
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
   
   
  
    wDiffFirstPart = "1"
    wDiffSecPart = "2"
  
  '*** 2.11 (Jen)
  GRP_REP.SetFormula Rep, "g1_PcVal1", "Sum ({@wVal1}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g2_PcVal1", "Sum ({@wVal1}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g3_PcVal1", "Sum ({@wVal1}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g_PcVal1", "Sum ({@wVal1})"
  
  GRP_REP.SetFormula Rep, "g1_PcVal2", "Sum ({@wVal2}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g2_PcVal2", "Sum ({@wVal2}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g3_PcVal2", "Sum ({@wVal2}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g_PcVal2", "Sum ({@wVal2})"

  GRP_REP.SetFormula Rep, "g1_MetVal1", "Sum ({@wMetVal1}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g2_MetVal1", "Sum ({@wMetVal1}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g3_MetVal1", "Sum ({@wMetVal1}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g_MetVal1", "Sum ({@wMetVal1})"

  GRP_REP.SetFormula Rep, "g1_MetVal2", "Sum ({@wMetVal2}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g2_MetVal2", "Sum ({@wMetVal2}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g3_MetVal2", "Sum ({@wMetVal2}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g_MetVal2", "Sum ({@wMetVal2})"

  GRP_REP.SetFormula Rep, "g1_DCVal1", "Sum ({@wDCVal1}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g2_DCVal1", "Sum ({@wDCVal1}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g3_DCVal1", "Sum ({@wDCVal1}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g_DCVal1", "Sum ({@wDCVal1})"

  GRP_REP.SetFormula Rep, "g1_DCVal2", "Sum ({@wDCVal2}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g2_DCVal2", "Sum ({@wDCVal2}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g3_DCVal2", "Sum ({@wDCVal2}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g_DCVal2", "Sum ({@wDCVal2})"

  GRP_REP.SetFormula Rep, "g1_XVal1", "Sum ({@wXVal1}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g2_XVal1", "Sum ({@wXVal1}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g3_XVal1", "Sum ({@wXVal1}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g_XVal1", "Sum ({@wXVal1})"

  GRP_REP.SetFormula Rep, "g1_XVal2", "Sum ({@wXVal2}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g2_XVal2", "Sum ({@wXVal2}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g3_XVal2", "Sum ({@wXVal2}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g_XVal2", "Sum ({@wXVal2})"

  GRP_REP.SetFormula Rep, "g1_LabVal1", "Sum ({@wLabourVal1}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g2_LabVal1", "Sum ({@wLabourVal1}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g3_LabVal1", "Sum ({@wLabourVal1}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g_LabVal1", "Sum ({@wLabourVal1})"

  GRP_REP.SetFormula Rep, "g1_LabVal2", "Sum ({@wLabourVal2}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g2_LabVal2", "Sum ({@wLabourVal2}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g3_LabVal2", "Sum ({@wLabourVal2}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g_LabVal2", "Sum ({@wLabourVal2})"
  '*** 2.11 (Jen)
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDFR")) <> "", " Od.OdOmCmCd >= '" + Trim(adc("UOCMCDFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDTO")) <> "", " Od.OdOmCmCd <= '" + Trim(adc("UOCMCDTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDSEL")) <> "", " Od.OdOmCmCd In (" + Trim(adc("UOCMCDSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " Od.OdOmDt >= '" + Trim(adc("UODTFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " Od.OdOmDt <= '" + Trim(adc("UODTTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  Dim wValDivBy As Double
  'sv.30 when divied by amount is blank then default value is 1
  wValDivBy = IIF(adc("UoValDivBy") = 0, 1, adc("UoValDivBy"))
  
  Dim wMetCstVal As String, wDCCstVal As String, wXCstVal As String, wSetCstVal As String, wLabCstVal As String, wLabVal As String
  wMetCstVal = "(Select isnull(sum(case when OrRmCtg  in('G', 'P', 'S', 'L','M') then OrSalVal else 0 end),0) from OrdRm where  (OrCoCd= Odm.OdCoCd and " + _
            "OrTc= Odm.OdTc and OrYy= Odm.OdYy and OrChr= Odm.OdChr and OrNo= Odm.OdNo and OrSr= Odm.OdSr) or ('Y'='" + ws_hAddJROYN + "' and OrOdMstIdNo=Odm.OdIdNo) and OrPrtKey=Odm.OdPrtKey) "
          
  wDCCstVal = " (Select isnull(sum(case when OrRmCtg  in('D', 'C') then OrSalVal else 0 end),0) from OrdRm where  (OrCoCd= Odm.OdCoCd and " + _
            "OrTc= Odm.OdTc and OrYy= Odm.OdYy and OrChr= Odm.OdChr and OrNo= Odm.OdNo and OrSr= Odm.OdSr) or ('Y'='" + ws_hAddJROYN + "' and OrOdMstIdNo=Odm.OdIdNo) and OrPrtKey=Odm.OdPrtKey) "
  
  wXCstVal = "(Select isnull(sum(case when OrRmCtg  in('X') then OrSalVal else 0 end),0) from OrdRm where (OrCoCd= Odm.OdCoCd and " + _
            "OrTc= Odm.OdTc and OrYy= Odm.OdYy and OrChr= Odm.OdChr and OrNo= Odm.OdNo and OrSr= Odm.OdSr) or ('Y'='" + ws_hAddJROYN + "' and OrOdMstIdNo=Odm.OdIdNo) and OrPrtKey=Odm.OdPrtKey) "
  
  wSetCstVal = "(Select isnull(sum(OrSetSalVal),0) from OrdRm where  (OrCoCd= Odm.OdCoCd and " + _
            "OrTc= Odm.OdTc and OrYy= Odm.OdYy and OrChr= Odm.OdChr and OrNo= Odm.OdNo and OrSr= Odm.OdSr) or ('Y'='" + ws_hAddJROYN + "' and OrOdMstIdNo=Odm.OdIdNo) and OrPrtKey=Odm.OdPrtKey) "
          
  wLabCstVal = "(Select isnull(sum(OlSalVal),0) from OrdLab where  OlCoCd= Odm.OdCoCd and " + _
            "OlTc= Odm.OdTc and OlYy= Odm.OdYy and OlChr= Odm.OdChr and OlNo= Odm.OdNo and OlSr= Odm.OdSr and OlPrtKey=Odm.OdPrtKey) "
  
  wLabVal = "(Select sum(OlSalVal) from OrdLab where  OlPrtKey=Od.OdPrtKey and  " + _
            "OlCoCd= Od.OdCoCd and OlTc= Od.OdTc and OlYy= Od.OdYy and OlChr= Od.OdChr and OlNo= Od.OdNo and OlSr= Od.OdSr) "

  
  'sv.30 wValDivBy added to qPcVal1, qPcVal2,qMetVal1,qMetVal2,qDCVal1,qDCVal2,qXVal1,qXVal2,qSetVal1,qSetVal2,qLabVal1,qLabVal2
  '578.22 conversion factor added to PCval, MetVal, DiaVal, CSVal, XVal, Setval, LabVal
  wStr1 = "Select " + grpflds + ", " + _
          "OdTc, OdYy, OdChr, OdNo, OdSr, max(OdDmCd) as OdDmCd, max(OdOrdQty) as OdOrdQty, " + _
          "max(qPcVal1/" + CStr(wValDivBy) + ") AS qPcVal1, max(qPcVal2/" + CStr(wValDivBy) + ") AS qPcVal2, sum(qMetVal1/" + CStr(wValDivBy) + ") as qMetVal1,max(qMetVal2/" + CStr(wValDivBy) + ") as qMetVal2, " + _
          "sum(qDCVal1/" + CStr(wValDivBy) + ") as qDCVal1, max(qDCVal2/" + CStr(wValDivBy) + ") as qDCVal2, sum(qXVal1/" + CStr(wValDivBy) + ") as qXVal1, max(qXVal2/" + CStr(wValDivBy) + ") as qXVal2, " + _
          "sum(qSetVal1/" + CStr(wValDivBy) + ") as qSetVal1, max(qSetVal2/" + CStr(wValDivBy) + ") as qSetVal2, max(qLabVal1/" + CStr(wValDivBy) + ") as qLabVal1,max(qLabVal2/" + CStr(wValDivBy) + ") as qLabVal2, " + _
          "0 as qPcWt1, 0 as qPcWt2, 0 as qMetWt1, 0 as qMetWt2, 0 as qDCWt1, 0 as qDCWt2, 0 as qXWt1, 0 as qXWt2 " + _
          ",max(CmGrp) as CmGrp,max(CmTyp) as CmTyp,max(OmCmSalPer) as OmCmSalPer,max(DmCtg) as DmCtg,max(DmSalCtg) as DmSalCtg, " + _
          "max(DmUom) as DmUom,max(OmPONo) as OmPONo,max(OmCmCd) as OmCmCd,max(OdKt) as OdKt,max(OdVaCtg) as OdVaCtg, " + _
          "max(OdDmCol) as OdDmCol,max(DmDesc) as DmDesc,max(OdPoNo) as OdPoNo,max(OdSubCust) as OdSubCust " + _
          "from (Select Od.OdTc, Od.OdYy, Od.OdChr, Od.OdNo, Od.OdSr,(Od.OdDmCd) as OdDmCd,(Od.OdOrdQty) as OdOrdQty," + _
          IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then Od.OdSalPrc/OmCnvFct else Od.OdSalPrc*OmCnvFct end) ", "Od.OdSalPrc") + " as qPcVal1," + _
          IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then isnull(Odm.OdSalPrc,0)/OmCnvFct else isnull(Odm.OdSalPrc,0)*OmCnvFct end) ", "isnull(Odm.OdSalPrc,0)") + " as qPcVal2," + _
          "(case when OrRmCtg in('G', 'P', 'S', 'L','M') then " + IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + " else 0 end) as qMetVal1," + _
          IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then " + wMetCstVal + "/OmCnvFct else " + wMetCstVal + "*OmCnvFct end) ", wMetCstVal) + " as qMetVal2, " + _
          "(case when OrRmCtg in('D', 'C') then " + IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + " else 0 end) as qDCVal1, " + _
          IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then " + wDCCstVal + "/OmCnvFct else " + wDCCstVal + "*OmCnvFct end) ", wDCCstVal) + " as qDCVal2," + _
          "(case when OrRmCtg in('X') then " + IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + " else 0 end) as qXVal1 ," + _
          IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then " + wXCstVal + "/OmCnvFct else " + wXCstVal + "*OmCnvFct end) ", wXCstVal) + " as qXVal2 ," + _
          IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then OrSetSalVal/OmCnvFct else OrSetSalVal*OmCnvFct end) ", "OrSetSalVal") + " as qSetVal1, " + _
          IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then " + wSetCstVal + "/OmCnvFct else " + wSetCstVal + "*OmCnvFct end) ", wSetCstVal) + " as qSetVal2, " + _
          IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then " + wLabVal + "/OmCnvFct else " + wLabVal + "*OmCnvFct end) ", wLabVal) + " as qLabVal1 , " + _
          IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then " + wLabCstVal + "/OmCnvFct else " + wLabCstVal + "*OmCnvFct end) ", wLabCstVal) + "as qLabVal2, "
  
  'U.jEmr.17- Param join added to filter tctype
   wStr2 = "CmGrp,CmTyp,OmCmSalPer,DmCtg,DmSalCtg,DmUom,OmPONo,OmCmCd,(Od.OdKt) as OdKt,(Od.OdVaCtg) as OdVaCtg," + _
          "(Od.OdDmCol) as OdDmCol,DmDesc,(Od.OdPoNo) as OdPoNo,(Od.OdSubCust) as OdSubCust  From OrdDsg Od " + _
          "join DsgMst on DmIdNo=Od.OdDmIdNo and DmTcTyp='" + IIF(adc("UoTcTypFr") <> "QS", "DM", "SM") + "' and DmCd= Od.OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
          "join OrdRm on OrOdIdNo=Od.OdIdNo and OrCoCd= Od.OdCoCd and OrTc= Od.OdTc and OrYy= Od.OdYy and OrChr= Od.OdChr " + _
          "and OrNo= Od.OdNo and OrSr= Od.OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=Od.OdPrtKey ", "") + _
          "join RmMst on RmCd= OrRmCd  " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + _
          "join OrdMst on OmIdNo=Od.OdOmIdNo and OmCoCd= Od.OdCoCd and OmTc= Od.OdTc and OmYy= Od.OdYy and OmChr= Od.OdChr " + _
          "and OmNo= Od.OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=Od.OdPrtKey ", "") + _
          "join CustMst on CmCtg= 'C' and CmCd = OmCmCd " + _
          "join Param on PTyp ='TC' and PMCd =OdTc and PValue ='" + adc("UoTcTypFr") + "' " + _
          "LEFT OUTER JOIN OrdDsg Odm on Odm.OdIdNo=(Select ForMaster.OdIdNo from vwDmStkNumMaster ForMaster where ForMaster.OdDmCd = Od.OdDmCd and ForMaster.OdSfx =Od.OdSfx and ForMaster.OdDmSz =Od.OdDmSz and ForMaster.OdDmStkCoCd =Od.OdDmStkCoCd and ForMaster.OdDmStkYy =Od.OdDmStkYy " + _
          "and ForMaster.OdDmStkChr =Od.OdDmStkChr and ForMaster.OdDmStkNo  =Od.OdDmStkNo and ForMaster.MasterYN='Y') " + wCnd + IIF(gs_Partition = ctCurrPrtn, " and Od.OdPrtKey='" + ctCurrPrtn + "' ", "") + _
          IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + _
          ")tbl Group By tbl.OdTc, tbl.OdYy, tbl.OdChr, tbl.OdNo, tbl.OdSr "


  wSqlStrg = wStr1 + wStr2
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt

  '*** (Ver 2.01)
  GRP_REP.SetFormula Rep, "wPcValPrc", "If {rdo.qPcVal1}>0 then (Round({rdo.qPcVal" + wDiffFirstPart + "}, 2)- Round({rdo.qPcVal" + wDiffSecPart + "}, 2))/{rdo.qPcVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "wMetValPrc", "If {rdo.qMetVal1}>0 then (Round({rdo.qMetVal" + wDiffFirstPart + "}, 2)-Round({rdo.qMetVal" + wDiffSecPart + "}, 2))/{rdo.qMetVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "wDCValPrc", "If {rdo.qDCVal1}>0 then (Round({rdo.qDCVal" + wDiffFirstPart + "}, 2)-Round({rdo.qDCVal" + wDiffSecPart + "}, 2))/{rdo.qDCVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "wXValPrc", "If {rdo.qXVal1}>0 then (Round({rdo.qXVal" + wDiffFirstPart + "}, 2)- Round({rdo.qXVal" + wDiffSecPart + "}, 2))/{rdo.qXVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "wLabValPrc", "If ({rdo.qSetVal1}+ {rdo.qLabVal1})>0 then " + _
                          "(Round(({rdo.qSetVal" + wDiffFirstPart + "} + {rdo.qLabVal" + wDiffFirstPart + "}), 2) - Round(({rdo.qSetVal" + wDiffSecPart + "} + {rdo.qLabVal" + wDiffSecPart + "}), 2))/ " + _
                          "({rdo.qSetVal1} + {rdo.qLabVal1})* 100 else 0 "
  
  GRP_REP.SetFormula Rep, "g1_PcValPrc", "If {@g1_PcVal1}>0 then (Round({@g1_PcVal" + wDiffFirstPart + "}, 2)- Round({@g1_PcVal" + wDiffSecPart + "}, 2))/{@g1_PcVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g2_PcValPrc", "If {@g2_PcVal1}>0 then (Round({@g2_PcVal" + wDiffFirstPart + "}, 2)- Round({@g2_PcVal" + wDiffSecPart + "}, 2))/{@g2_PcVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g3_PcValPrc", "If {@g3_PcVal1}>0 then (Round({@g3_PcVal" + wDiffFirstPart + "}, 2)- Round({@g3_PcVal" + wDiffSecPart + "}, 2))/{@g3_PcVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g_PcValPrc", "If {@g_PcVal1}>0 then (Round({@g_PcVal" + wDiffFirstPart + "}, 2)- Round({@g_PcVal" + wDiffSecPart + "}, 2))/{@g_PcVal1}*100 else 0"
    
  GRP_REP.SetFormula Rep, "g1_MetValPrc", "If {@g1_MetVal1}>0 then (Round({@g1_MetVal" + wDiffFirstPart + "}, 2)- Round({@g1_MetVal" + wDiffSecPart + "}, 2))/{@g1_MetVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g2_MetValPrc", "If {@g2_MetVal1}>0 then (Round({@g2_MetVal" + wDiffFirstPart + "}, 2)- Round({@g2_MetVal" + wDiffSecPart + "}, 2))/{@g2_MetVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g3_MetValPrc", "If {@g3_MetVal1}>0 then (Round({@g3_MetVal" + wDiffFirstPart + "}, 2)- Round({@g3_MetVal" + wDiffSecPart + "}, 2))/{@g3_MetVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g_MetValPrc", "If {@g_MetVal1}>0 then (Round({@g_MetVal" + wDiffFirstPart + "}, 2)- Round({@g_MetVal" + wDiffSecPart + "}, 2))/{@g_MetVal1}*100 else 0"
  
  GRP_REP.SetFormula Rep, "g1_DCValPrc", "If {@g1_DCVal1}>0 then (Round({@g1_DCVal" + wDiffFirstPart + "}, 2)- Round({@g1_DCVal" + wDiffSecPart + "}, 2))/{@g1_DCVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g2_DCValPrc", "If {@g2_DCVal1}>0 then (Round({@g2_DCVal" + wDiffFirstPart + "}, 2)- Round({@g2_DCVal" + wDiffSecPart + "}, 2))/{@g2_DCVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g3_DCValPrc", "If {@g3_DCVal1}>0 then (Round({@g3_DCVal" + wDiffFirstPart + "}, 2)- Round({@g3_DCVal" + wDiffSecPart + "}, 2))/{@g3_DCVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g_DCValPrc", "If {@g_DCVal1}>0 then (Round({@g_DCVal" + wDiffFirstPart + "}, 2)- Round({@g_DCVal" + wDiffSecPart + "}, 2))/{@g_DCVal1}*100 else 0"
  
  GRP_REP.SetFormula Rep, "g1_XValPrc", "If {@g1_XVal1}>0 then (Round({@g1_XVal" + wDiffFirstPart + "}, 2)- Round({@g1_XVal" + wDiffSecPart + "}, 2))/{@g1_XVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g2_XValPrc", "If {@g2_XVal1}>0 then (Round({@g2_XVal" + wDiffFirstPart + "}, 2)- Round({@g2_XVal" + wDiffSecPart + "}, 2))/{@g2_XVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g3_XValPrc", "If {@g3_XVal1}>0 then (Round({@g3_XVal" + wDiffFirstPart + "}, 2)- Round({@g3_XVal" + wDiffSecPart + "}, 2))/{@g3_XVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g_XValPrc", "If {@g_XVal1}>0 then (Round({@g_XVal" + wDiffFirstPart + "}, 2)- Round({@g_XVal" + wDiffSecPart + "}, 2))/{@g_XVal1}*100 else 0"
  
  '*************Geeta **********Emr206 ************Changed Summary Feild
  GRP_REP.SetFormula Rep, "g1_LabValPrc", "If {@g1_LabVal1}>0 then (Round({@g1_LabVal" + wDiffFirstPart + "}, 2)- Round({@g1_LabVal" + wDiffSecPart + "}, 2))/{@g1_LabVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g2_LabValPrc", "If {@g2_LabVal1}>0 then (Round({@g2_LabVal" + wDiffFirstPart + "}, 2)- Round({@g2_LabVal" + wDiffSecPart + "}, 2))/{@g2_LabVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g3_LabValPrc", "If {@g3_LabVal1}>0 then (Round({@g3_LabVal" + wDiffFirstPart + "}, 2)- Round({@g3_LabVal" + wDiffSecPart + "}, 2))/{@g3_LabVal1}*100 else 0"
  GRP_REP.SetFormula Rep, "g_LabValPrc", "If {@g_LabVal1}>0 then (Round({@g_LabVal" + wDiffFirstPart + "}, 2)- Round({@g_LabVal" + wDiffSecPart + "}, 2))/{@g_LabVal1}*100 else 0"
  '****************Geeta**********************************
  
  GRP_REP.SetFormula Rep, "wPcWtPrc", "If {rdo.qPcWt1}>0 then ({rdo.qPcWt" + wDiffFirstPart + "}-{rdo.qPcWt" + wDiffSecPart + "})/{rdo.qPcWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "wMetWtPrc", "If {rdo.qMetWt1}>0 then ({rdo.qMetWt" + wDiffFirstPart + "}-{rdo.qMetWt" + wDiffSecPart + "})/{rdo.qMetWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "wDCWtPrc", "If {rdo.qDCWt1}>0 then ({rdo.qDCWt" + wDiffFirstPart + "}-{rdo.qDCWt" + wDiffSecPart + "})/{rdo.qDCWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "wXWtPrc", "If {rdo.qXWt1}>0 then ({rdo.qXWt" + wDiffFirstPart + "}-{rdo.qXWt" + wDiffSecPart + "})/{rdo.qXWt1}*100 else 0"
  
  GRP_REP.SetFormula Rep, "g1_PcWtPrc", "If {@g1_PcWt1}>0 then ({@g1_PcWt" + wDiffFirstPart + "}-{@g1_PcWt" + wDiffSecPart + "})/{@g1_PcWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "g2_PcWtPrc", "If {@g2_PcWt1}>0 then ({@g2_PcWt" + wDiffFirstPart + "}-{@g2_PcWt" + wDiffSecPart + "})/{@g2_PcWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "g3_PcWtPrc", "If {@g3_PcWt1}>0 then ({@g3_PcWt" + wDiffFirstPart + "}-{@g3_PcWt" + wDiffSecPart + "})/{@g3_PcWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "g_PcWtPrc", "If {@g_PcWt1}>0 then ({@g_PcWt" + wDiffFirstPart + "}-{@g_PcWt" + wDiffSecPart + "})/{@g_PcWt1}*100 else 0"
  
  GRP_REP.SetFormula Rep, "g1_MetWtPrc", "If {@g1_MetWt1}>0 then ({@g1_MetWt" + wDiffFirstPart + "}-{@g1_MetWt" + wDiffSecPart + "})/{@g1_MetWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "g2_MetWtPrc", "If {@g2_MetWt1}>0 then ({@g2_MetWt" + wDiffFirstPart + "}-{@g2_MetWt" + wDiffSecPart + "})/{@g2_MetWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "g3_MetWtPrc", "If {@g3_MetWt1}>0 then ({@g3_MetWt" + wDiffFirstPart + "}-{@g3_MetWt" + wDiffSecPart + "})/{@g3_MetWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "g_MetWtPrc", "If {@g_MetWt1}>0 then ({@g_MetWt" + wDiffFirstPart + "}-{@g_MetWt" + wDiffSecPart + "})/{@g_MetWt1}*100 else 0"
    
  GRP_REP.SetFormula Rep, "g1_DCWtPrc", "If {@g1_DCWt1}>0 then ({@g1_DCWt" + wDiffFirstPart + "}-{@g1_DCWt" + wDiffSecPart + "})/{@g1_DCWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "g2_DCWtPrc", "If {@g2_DCWt1}>0 then ({@g2_DCWt" + wDiffFirstPart + "}-{@g2_DCWt" + wDiffSecPart + "})/{@g2_DCWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "g3_DCWtPrc", "If {@g3_DCWt1}>0 then ({@g3_DCWt" + wDiffFirstPart + "}-{@g3_DCWt" + wDiffSecPart + "})/{@g3_DCWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "g_DCWtPrc", "If {@g_DCWt1}>0 then ({@g_DCWt" + wDiffFirstPart + "}-{@g_DCWt" + wDiffSecPart + "})/{@g_DCWt1}*100 else 0"
  
  GRP_REP.SetFormula Rep, "g1_XWtPrc", "If {@g1_XWt1}>0 then ({@g1_XWt" + wDiffFirstPart + "}-{@g1_XWt" + wDiffSecPart + "})/{@g1_XWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "g2_XWtPrc", "If {@g2_XWt1}>0 then ({@g2_XWt" + wDiffFirstPart + "}-{@g2_XWt" + wDiffSecPart + "})/{@g2_XWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "g3_XWtPrc", "If {@g3_XWt1}>0 then ({@g3_XWt" + wDiffFirstPart + "}-{@g3_XWt" + wDiffSecPart + "})/{@g3_XWt1}*100 else 0"
  GRP_REP.SetFormula Rep, "g_XWtPrc", "If {@g_XWt1}>0 then ({@g_XWt" + wDiffFirstPart + "}-{@g_XWt" + wDiffSecPart + "})/{@g_XWt1}*100 else 0"
  

  Rep.TxtHead.SetText "Profit Margin Report"
  Rep.h1Pc1.SetText "PcVchrVal": Rep.h2Pc1.SetText "PcVchrVal"
  Rep.h1Pc2.SetText "PcCostVal": Rep.h2Pc2.SetText "PcCostVal"
  Rep.h1PcPrc.SetText "Pc %": Rep.h2PcPrc.SetText "Pc %"
  Rep.h1Met1.SetText "MetVchrVal": Rep.h2Met1.SetText "MetVchrVal"
  Rep.h1Met2.SetText "MetCostVal": Rep.h2Met2.SetText "MetCostVal"
  Rep.h1MetPrc.SetText "Met %": Rep.h2MetPrc.SetText "Met %"
  Rep.h1DC1.SetText "DCVchrVal": Rep.h2DC1.SetText "DCVchrVal"
  Rep.h1DC2.SetText "DCCostVal": Rep.h2DC2.SetText "DCCostVal"
  Rep.h1DCPrc.SetText "DC %": Rep.h2DCPrc.SetText "DC %"
  Rep.h1X1.SetText "AccVchrVal": Rep.h2X1.SetText "AccVchrVal"
  Rep.h1X2.SetText "AccCostVal": Rep.h2X2.SetText "AccCostVal"
  Rep.h1AccPrc.SetText "Acc %": Rep.h2AccPrc.SetText "Acc %"
  Rep.h1Lab1.SetText "LabVchrVal": Rep.h2Lab1.SetText "LabVchrVal"
  Rep.h1Lab2.SetText "LabCostVal": Rep.h2Lab2.SetText "LabCostVal"
  Rep.h1LabPrc.SetText "Lab %": Rep.h2LabPrc.SetText "Lab %"
  Rep.FldPcWt1.Suppress = True: Rep.FldPcWt2.Suppress = True: Rep.FldPcWtPrc.Suppress = True
  Rep.FldMetWt1.Suppress = True: Rep.FldMetWt2.Suppress = True: Rep.FldMetWtPrc.Suppress = True
  Rep.FldDCWt1.Suppress = True: Rep.FldDCWt2.Suppress = True: Rep.FldDCWtPrc.Suppress = True
  Rep.FldXWt1.Suppress = True: Rep.FldXWt2.Suppress = True: Rep.FldXWtPrc.Suppress = True
  Rep.G1FWt.Suppress = True: Rep.G2FWt.Suppress = True: Rep.G3FWt.Suppress = True
  Rep.RepFtWt.Suppress = True
  
  'sv.30 passing divide by value to dsr
  GRP_REP.SetFormula Rep, "wDivByVal", adc("UoValDivBy")
  
  'sv.30 when divied by amount>1 then follwing text will display in header
  If adc("UoValDivBy") > 1 Then Rep.TxtDivByVal.SetText "(Values in " + Format(adc("UoValDivBy"), "##,##,##0") + ")"
    
  CRV_REP.DisplayGroupTree = False
  
Call DispReport(adc)
  
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
    gs_EntTyp = "JT"
    Set adc.FirNKeyCtl = adc("UoTcTypFr") 'adc("UoCoCdFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_EntTyp = "JT"
  'U.jEmr.17- validation for tctype
  Select Case UCase(IdName)
  Case Is = UCase("UoTcTypFr")
  If Not moCn.RecSeek("Select 'x' From Param Where PTyp ='TCTYP' and PMCd='" + pv_NewValue + "' and PValue1='" + gs_EntTyp + "'") Then _
          Cancel = True: ErrMsg = "Invalid Tc Type": Exit Sub
  'sv.30 divided by amount cannot be 0
  Case Is = UCase("UoValDivBy")
    If pv_NewValue = 0 Then Cancel = True: ErrMsg = "Divided By Amount Cannot be Zero": Exit Sub
  
  End Select
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = True         '*** (Jen 2.14 Next Patch 1)
  gs_EntTyp = "JT"
  
  Select Case UCase(IdName)
  Case Is = UCase("UoOmTcFr"), UCase("UoOmTcTo")
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_OmTcTyp = adc("UoTcTypFr")
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
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
  'U.jEmr.17- if TcType is blank in form opening then set default value as JT
  If adc("UoTcTypFr") = "" Then adc("UoTcTypFr") = "JT"
  gs_EntTyp = "JT"
  gs_TcTyp = adc("UoTctypFr")     ' ***** Manali 3.5.0 - 08/11/08
  
  adc("UoOrdVal").Enabled = False
  adc("UoOrdVal1").Enabled = False
  
  adc("UoValDivBy") = 1     'sv.30 setting default value
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

' ******** Manali 3.2.2 - User date and time in reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
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
