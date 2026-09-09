VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpLRmTxnLst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Listing (Loose Rm)"
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
      Left            =   10830
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   9630
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4455
      TabIndex        =   43
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
      Left            =   4785
      TabIndex        =   52
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   53
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
         TabIndex        =   54
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
      TabIndex        =   51
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   48
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
      TabIndex        =   44
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
      TabIndex        =   47
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   46
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpLRmTxnLst.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpLRmTxnLst.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   50
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   45
               Top             =   360
               Width           =   8805
               _ExtentX        =   15531
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   49
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4005
               TabIndex        =   12
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   1095
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
               Index           =   48
               Left            =   1515
               TabIndex        =   11
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   1095
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
               Index           =   24
               Left            =   4785
               TabIndex        =   5
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   525
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "B.OdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2295
               TabIndex        =   2
               ToolTipText     =   "Enter From Transaction Year"
               Top             =   525
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "B.OdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1515
               TabIndex        =   1
               ToolTipText     =   "Enter From Transaction Tc"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "B.OdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2715
               TabIndex        =   3
               ToolTipText     =   "Enter From Transaction Character"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "B.OdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   4005
               TabIndex        =   4
               ToolTipText     =   "Enter To Transaction Tc"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "B.OdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   5205
               TabIndex        =   6
               ToolTipText     =   "Enter To Transaction Character"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "B.OdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   4905
               TabIndex        =   10
               ToolTipText     =   "Enter To Transaction Serial"
               Top             =   810
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "B.OdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   2415
               TabIndex        =   8
               ToolTipText     =   "Enter From Transaction Serial"
               Top             =   810
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "B.OdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1515
               TabIndex        =   7
               ToolTipText     =   "Enter From Transaction Number"
               Top             =   810
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "B.OdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   4005
               TabIndex        =   9
               ToolTipText     =   "Enter To Transaction Number"
               Top             =   810
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "B.OdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   4005
               TabIndex        =   17
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   1665
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGTO"
               IdName          =   "UORMSCTGTO"
               CmpStr          =   "RmSCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   6510
               TabIndex        =   18
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   1665
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "RmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1515
               TabIndex        =   13
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   1380
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "RmCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   1515
               TabIndex        =   16
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   1665
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
               CmpStr          =   "RmSCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   4005
               TabIndex        =   14
               ToolTipText     =   "Enter To Raw Material Category"
               Top             =   1380
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
               CmpStr          =   "RmCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   6510
               TabIndex        =   15
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   1380
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCTGSEL"
               IdName          =   "UORMCTGSEL"
               CmpStr          =   "RmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   6510
               TabIndex        =   30
               ToolTipText     =   "Enter Supplier/ Customer Selection"
               Top             =   2880
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   1515
               TabIndex        =   28
               ToolTipText     =   "Enter From Supplier/ Customer"
               Top             =   2865
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OmCmcd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   1515
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   240
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "B.OdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   66
               Left            =   1515
               TabIndex        =   25
               ToolTipText     =   "Enter From Lot No."
               Top             =   2585
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataType        =   4
               DataField       =   "UOLOTNOFR"
               IdName          =   "UOLOTNOFR"
               CmpStr          =   "B.OdJLLotNo>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   67
               Left            =   4005
               TabIndex        =   26
               ToolTipText     =   "Enter To Lot No."
               Top             =   2585
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOLOTNOTO"
               IdName          =   "UOLOTNOTO"
               CmpStr          =   "B.OdJLLotNo<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   68
               Left            =   6510
               TabIndex        =   27
               ToolTipText     =   "Enter Lot No. Selection"
               Top             =   2585
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOLOTNOSEL"
               IdName          =   "UOLOTNOSEL"
               CmpStr          =   "B.OdJLLotNo In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   69
               Left            =   1515
               TabIndex        =   31
               ToolTipText     =   "Enter Currency Code Whose Vouchers You Want To View Or Leave It Blank To View All Vouchers In Base Currency"
               Top             =   3150
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOCURCDFR"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   4005
               TabIndex        =   20
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1970
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOPARENTCDTO"
               IdName          =   "UOPARENTCDTO"
               CmpStr          =   "RmParentCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   72
               Left            =   6510
               TabIndex        =   21
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1970
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
               CmpStr          =   "RmParentCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   73
               Left            =   1515
               TabIndex        =   19
               ToolTipText     =   "Enter From Parent Code"
               Top             =   1970
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOPARENTCDFR"
               IdName          =   "UOPARENTCDFR"
               CmpStr          =   "RmParentCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   4005
               TabIndex        =   23
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   2280
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "RmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6510
               TabIndex        =   24
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   2280
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "RmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1515
               TabIndex        =   22
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   2280
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "RmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   4815
               TabIndex        =   36
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   3435
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
               CmpStr          =   "M.OdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   2325
               TabIndex        =   33
               ToolTipText     =   "Enter From Transaction Year"
               Top             =   3435
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
               CmpStr          =   "M.OdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   1515
               TabIndex        =   32
               ToolTipText     =   "Enter From Transaction Tc"
               Top             =   3435
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
               CmpStr          =   "M.OdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   2775
               TabIndex        =   34
               ToolTipText     =   "Enter From Transaction Character"
               Top             =   3435
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
               CmpStr          =   "M.OdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   4005
               TabIndex        =   35
               ToolTipText     =   "Enter To Transaction Tc"
               Top             =   3435
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
               CmpStr          =   "M.OdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   5265
               TabIndex        =   37
               ToolTipText     =   "Enter To Transaction Character"
               Top             =   3435
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
               CmpStr          =   "M.OdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   4935
               TabIndex        =   41
               ToolTipText     =   "Enter To Transaction Serial"
               Top             =   3720
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRTO"
               IdName          =   "UOTDSRTO"
               CmpStr          =   "M.OdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   2445
               TabIndex        =   39
               ToolTipText     =   "Enter From Transaction Serial"
               Top             =   3720
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRFR"
               IdName          =   "UOTDSRFR"
               CmpStr          =   "M.OdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   1515
               TabIndex        =   38
               ToolTipText     =   "Enter From Transaction Number"
               Top             =   3720
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOFR"
               IdName          =   "UOTNOFR"
               CmpStr          =   "M.OdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   4020
               TabIndex        =   40
               ToolTipText     =   "Enter To Transaction Number"
               Top             =   3720
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOTO"
               IdName          =   "UOTNOTO"
               CmpStr          =   "M.OdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   4020
               TabIndex        =   29
               ToolTipText     =   "Enter From Supplier/ Customer"
               Top             =   2880
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "OmCmcd <="
            End
            Begin VB.Label ALBL 
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
               Index           =   14
               Left            =   2340
               TabIndex        =   82
               Top             =   3720
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Index           =   10
               Left            =   4830
               TabIndex        =   81
               Top             =   3720
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Index           =   9
               Left            =   2220
               TabIndex        =   80
               Top             =   3435
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   4710
               TabIndex        =   79
               Top             =   3435
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   2670
               TabIndex        =   78
               Top             =   3435
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Index           =   6
               Left            =   5160
               TabIndex        =   77
               Top             =   3435
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Base No/Sr"
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
               TabIndex        =   76
               Top             =   3720
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Base Tc/Yy/Chr"
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
               TabIndex        =   75
               Top             =   3435
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Parent Cd"
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
               TabIndex        =   74
               Top             =   1970
               Width           =   1425
            End
            Begin VB.Label LblCurrCd 
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   60
               TabIndex        =   73
               ToolTipText     =   "Location"
               Top             =   3150
               Width           =   1485
            End
            Begin VB.Label LblLotNo 
               BackStyle       =   0  'Transparent
               Caption         =   "Lot No."
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
               TabIndex        =   72
               Top             =   2585
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
               Index           =   0
               Left            =   60
               TabIndex        =   71
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label LblSuppCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Supp / Cust"
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
               TabIndex        =   70
               Top             =   2865
               Width           =   1425
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
               Left            =   60
               TabIndex        =   69
               ToolTipText     =   "Location"
               Top             =   1380
               Width           =   1425
            End
            Begin VB.Label LblRmSCtg 
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
               Left            =   60
               TabIndex        =   68
               Top             =   1665
               Width           =   1425
            End
            Begin VB.Label LblRmCd 
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
               Left            =   60
               TabIndex        =   67
               Top             =   2280
               Width           =   1425
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
               Left            =   60
               TabIndex        =   66
               Top             =   525
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
               Left            =   60
               TabIndex        =   65
               Top             =   810
               Width           =   1605
            End
            Begin VB.Label ALBL 
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
               Left            =   5145
               TabIndex        =   64
               Top             =   525
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               TabIndex        =   63
               Top             =   525
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   4725
               TabIndex        =   62
               Top             =   525
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   2235
               TabIndex        =   61
               Top             =   525
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   4845
               TabIndex        =   60
               Top             =   810
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   2355
               TabIndex        =   59
               Top             =   810
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
               Left            =   60
               TabIndex        =   58
               Top             =   1095
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
               Left            =   6510
               TabIndex        =   57
               Top             =   0
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
               Left            =   4005
               TabIndex        =   56
               Top             =   0
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
               TabIndex        =   55
               Top             =   0
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpLRmTxnLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepLRmTxnLst
Dim mRep As CRAXDRT.Report
Dim moCn As MwfLib.MDOConnection '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean, ws_FrmCaption As String
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
'  GRP_REP.Add "Voucher No", "B.OdTc+B.OdYy+B.OdChr+str(B.OdNo)", "B.OdTc+'/'+B.OdYy+'/'+B.OdChr+'/'+LTrim(Str(B.OdNo))", "", "", "", "", "", "B.OdTc+'/'+B.OdYy+'/'+B.OdChr+'/'+LTrim(Str(B.OdNo))"
  GRP_REP.Add "Voucher No", "B.OdTc+B.OdYy+B.OdChr+str(B.OdNo)", "B.OdTc+'/'+B.OdYy+'/'+B.OdChr+'/'+LTrim(Str(B.OdNo))", "wVchNo", "hVchNo", "", "", "", "B.OdTc+'/'+B.OdYy+'/'+B.OdChr+'/'+LTrim(Str(B.OdNo))"
  GRP_REP.Add "Base Voucher No", "M.OdTc+M.OdYy+M.OdChr+str(M.OdNo)", "M.OdTc+'/'+M.OdYy+'/'+M.OdChr+'/'+LTrim(Str(M.OdNo))", "", "", "", "", "", "M.OdTc+'/'+M.OdYy+'/'+M.OdChr+'/'+LTrim(Str(M.OdNo))"
  '6.1
  GRP_REP.Add "Voucher Date", "convert(varchar(10), OmDt, 112)", "convert(varchar(8),OmDt, " + CStr(DtFmtSlahVal()) + ")", "OmDt", "hVchDt", "Quarter,Month,Week", "", "", ""
  GRP_REP.Add "Cust / Supp", "OmCmCd", "OmCmCd", "OmCmCd", "hOmCmCd", "Customer Grp,Customer Typ", "", "", "(Select ( case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg=OmCmCtg and CmCd= OmCmCd) "
'    GRP_REP.Add "Exp Del Date", "convert(varchar(10), B.OdExpDelDt, 112)", "convert(varchar(8), B.OdExpDelDt, 3)", "", "", "", "", "", ""
    GRP_REP.Add "Rm Ctg", "RmCtg", "", "RmCtg", "hRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = RmCtg) "
    GRP_REP.Add "Rm SubCtg", "RmCtg+RmSCtg", "'('+ RTrim(RmCtg)+ ') '+ RmSCtg", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
    GRP_REP.Add "Rm Code", "RmCtg+RmSCtg+B.OdJLRmCd", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ B.OdJLRmCd", "RmCtg,RmSCtg,OdJLRmCd", "hRmCtg,hRmSCtg,hRmCd", "Rm Ctg,Rm SubCtg,Parent Cd", "", "", "SubString(RmDesc, 1, 30) "
    GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg,Rm SubCtg", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
  
If (moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y' ") And ws_MultiCoMod = False) Or (ws_MultiCoMod = True) Then
    GRP_REP.Add "RM Lot No.", "B.OdJLLotno", "B.OdJLLotNo", "OdJLLotNo", "hLotNo", "", "", "", "B.OdJLLotno"
End If
  
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  '*** Report Sql ***
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wi_Opt As Integer
  Dim wPurStr As String
  
  '*** (Jen 2.12)
  Dim wd_BaseGldRmPurity As Double, wd_BasePltRmPurity As Double, wd_BaseSlvRmPurity As Double
  '*** (Jen 2.12)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  Dim wd_BasePdRmPurity As Double
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  
  ' **** Manali 3.03 - 08/07/08 - RmSzDesc
  Dim wRmSzDesc As String, wRmIdSzJoin As String
  ' **** Manali 3.03 - 08/07/08 - RmSzDesc
  
  Dim wAddlCnd As String  '****** Sachin 3.02.0
  
'*** Jay 2.13(CT) ***
' **** Zubin 211 **** '
'*** Jay 2.13(CT) ***
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Set mRep = Rep
  
  Call DispCoNm
  Call SetMwName(mRep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  '*** This Report should always have the detail section ***
  '*** To check if Grp1, Grp2 is selected and also to check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer, wNonePos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0: wNonePos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
    If wGrp(i) = "(None)" Then
      wNonePos = i: Exit For
    End If
  Next i
  
  '*** (Jen 2.12)
  wd_BaseGldRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst " + _
                       "on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'G'), 1)")
  If wd_BaseGldRmPurity = 0 Then wd_BaseGldRmPurity = 1
  wd_BasePltRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst " + _
                       "on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'P'), 1)")
  If wd_BasePltRmPurity = 0 Then wd_BasePltRmPurity = 1
  wd_BaseSlvRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst " + _
                       "on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'S'), 1)")
  If wd_BaseSlvRmPurity = 0 Then wd_BaseSlvRmPurity = 1
  '*** (Jen 2.12)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L'
  wd_BasePdRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst " + _
                       "on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'L'), 1)")
  If wd_BasePdRmPurity = 0 Then wd_BasePdRmPurity = 1
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L'
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  ' **** Manali 3.03 - 08/07/08 - RmSzDesc
  wRmSzDesc = "(Case when (Select PDesc225 from Param where PTyp='RMSCTG' and PMCd=IsNull(RisRmCtg, '') and PSCd=RmSCtg) = 'Y' Then IsNull(RisRmSzDesc, '') Else '' End) "
  ' **** Manali 3.03 - 08/07/08 - RmSzDesc
  
   Dim ws_BaseCurCd As String, ws_Join As String
        ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                      "HCoCd= '" + adc("UoCoCdFr") + "'")
  
        If adc("UoCurCdFr") <> "" Then
          wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "OmCmCurCd='" + adc("UoCurCdFr") + "'"
          wPurStr = "B.OdJLVchRt, B.OdJLVchVal"
        Else
          wPurStr = " (case when OmMulDiv= 'M' then B.OdJLVchRt/OmCnvFct else B.OdJLVchRt*OmCnvFct end) as OdJLVchRt, (case when OmMulDiv= 'M' then B.OdJLVchVal/OmCnvFct else B.OdJLVchVal*OmCnvFct end) as OdJLVchVal "
        End If
        
        ws_Join = " Join (Select PMCd as qPMCd, (case when " + IIF(adc("UoCurCdFr") = "", "1", "0") + " = 1 then 'Y' " + _
                  "else 'N' end) as qBaseCur from Param where PTyp= 'CURNCY') Cur on qPMCd= OmCmCurCd "
    
   wSqlStrg = "Select " + grpflds + ", OmCoCd, OmTc, OmYy, OmChr, OmNo, OmDt, OmFrDmLoc, OmToDmLoc, OmFrDmDc, OmToDmDc, OmCmCd, IsNull(CmName, '') as CmName, " + _
               "OmCmCtg, OmCnvFct, OmCmCurCd, OmMulDiv, OmCmPayTerms,  " + _
               "B.OdSr, OmFrDmLoc, OmFrDmDc, RmCtg, RmSCtg, RmPurityWt/ " + _
               "(case when RmCtg= 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
               "      when RmCtg= 'P' then " + CStr(wd_BasePltRmPurity) + " " + _
               "      when RmCtg= 'S' then " + CStr(wd_BaseSlvRmPurity) + " " + _
               "      when RmCtg= 'L' then " + CStr(wd_BasePdRmPurity) + " else 1 end) as RmPurityWt, " + _
               "B.OdJLRmCd, B.OdJLLotNo, B.OdJLLn1, B.OdJLLn2, B.OdJLLn3, " + wRmSzDesc + " as qRmSzDesc, 'W' as TdRtByQw,  " + _
               "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
               "B.OdJLQty, B.OdJLWt, B.OdJLVchRt, " + wPurStr + ", RmDesc, IsNull(M.OdTc+'/'+M.OdYy+'/'+M.OdChr+'/'+LTrim(Str(M.OdNo,5))+'/'+LTrim(Str(M.OdSr,3)),'') as qBaseVch " + _
               "From OrdMst " + _
               "Join OrdDsg B on B.OdOmIdNo=OmIdNo and B.OdCoCd= OmCoCd and B.OdTc= OmTc and B.OdYy= OmYy " + _
               " and B.OdChr= OmChr and B.OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and B.OdPrtKey=OmPrtKey ", "") + _
               "Join RmMst on RmCd=B.OdJLRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=B.OdPrtKey ", "") + ws_Join + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=B.OdJLLn1 " + _
               "Left Outer Join CustMst On CmCtg=OmCmCtg and CmCd=OmCmCd " + _
               "Left Outer Join OrdDsg M on B.OdBaseIDKey=M.OdIdNo and (M.OdTc='LIN' Or M.OdCoCd = B.OdCoCd) " + IIF(gs_Partition = ctCurrPrtn, " and M.OdPrtKey=B.OdPrtKey ", "") + _
               "where " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey='" + ctCurrPrtn + "'", "") + _
               wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " order by OmCoCd, OmTc, OmYy, OmChr, OmNo"
     
'    wSqlStrg = "Select " + grpflds + ", " + _
               "TTc, TYy, TChr, TNo, TDt, TFrRmLoc, TToRmLoc, TLsLoc, TLsFrDt, " + _
               "TLsToDt, TGldinLs, TGldRec, TDustWt, TSuppCd, TBillNo,  TBillDt, TBOENo, " + _
               "TBOEDt, TCnvRt," + wTTotAmtStr + _
               "TdSr, TdFrRmLoc, TdFrRmDc, RmCtg, RmSCtg, RmPurityWt/ " + _
               "(case when RmCtg= 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
               "      when RmCtg= 'P' then " + CStr(wd_BasePltRmPurity) + " " + _
               "      when RmCtg= 'S' then " + CStr(wd_BaseSlvRmPurity) + " " + _
               "      when RmCtg= 'L' then " + CStr(wd_BasePdRmPurity) + " else 1 end) as RmPurityWt, " + _
               "TdRmCd, TdLotNo, TdRmSz, " + wRmSzDesc + " as qRmSzDesc, TdRmStkRt, " + _
               "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
               "TdRmQty, TdRmWt, TdToRmLoc, TdRtByQw, " + wPurStr + ", " + _
               "cast((TdRmWt * TdRmStkRt) as decimal (16,2)) as qRmStkAmt, TdDustWt, TdDesc From Txn " + _
               "Join Txnd Td on Td.TdTIdNo=TIdNo and Td.TdCoCd= TCoCd and Td.TdTc= TTc and Td.TdYy= TYy " + _
               "and Td.TdChr= TChr and Td.TdNo= TNo and Td.TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " and Td.TdPrtKey=TPrtKey ", "") + _
               "Join RmMst on RmCd= Td.TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=Td.TdPrtKey ", "") + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
               "where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' and ", "") + " (Select PValue from Param where PTyp= 'TC' and PMCd= TTc)= '" + gs_TxnTcTyp + "' " + _
               wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "")
  
  
             
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  '*********Geeta************Emr207***
    '578.22 showing Purchase Currency was commented now it is removed
    If adc("UoCurCdFr") = "" Then
      mRep.TxtPRCurCd.SetText "Purchase Currency (" + ws_BaseCurCd + ")"
    Else
      mRep.TxtPRCurCd.SetText "Purchase Currency (" + adc("UoCurCdFr") + ")"
    End If
    
  '*********Geeta************Emr207***
 
'  mRep.TxtHead.SetText ws_FrmCaption  '****** Sachin 2.14.0
  '  If wDetPos = 1 Or wNonePos = 2 Then mRep.G1F1.Suppress = True: mRep.G2F1.Suppress = True
  '  If wDetPos = 2 Or wNonePos = 3 Then mRep.G1F1.Suppress = True
  '  If wDetPos = 0 And (gs_TxnTcTyp = "PR" Or gs_TxnTcTyp = "MV") Then
  '      With mRep
        '.hVchNo.Suppress = True:
  '      .DET1.Suppress = True:  .h1TdSr.Suppress = True: .h1TDt.Suppress = True
  '      .h1TSuppCd.Suppress = True: .h1TBoeNo.Suppress = True: .h1TBoeDt.Suppress = True
  '      .h1TCnvRt.Suppress = True: .h1TdRmFrLoc.Suppress = True: .h1TdRmCd.Suppress = True: .h1TdRmDC.Suppress = True: .h1LotNo.Suppress = True
   '     .h1TdRmSz.Suppress = True: .h1TdRmStkRt.Suppress = True: .h1TdRmQW.Suppress = True: .h1TdPurRt.Suppress = True: .h1TdRmToLoc.Suppress = True
  '      Dim wGrpStr As String
  '      Dim wArr() As String
  '     wGrpStr = GRP_REP.GrpOpts
   '    wArr() = Split(wGrpStr, ",")
   '     .h1VchNo.SetText wArr(2)
   '     End With
   ' End If
  ' Zubin 212
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UOCOCDFR") + "' and HLotNoYN= 'Y' ") Then
      Rep.HLotNo.Suppress = True
      Rep.FldDET1LotNo.Suppress = True
  End If
 ' GRP_REP.SetFormula mRep, "wRmSzPtr", "If {rdo.qRmSzDesc} <> '' Then {rdo.qRmSzDesc} " + _
                                      "Else (If {rdo.qPtrYn} = 'Y' Then ToText({rdo.TdRmSz}, 4) Else ToText({rdo.TdRmSz}, 3))"
  
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.OdJLRmCd}", 7)
  
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
  Set moCn = adc.Connection '*** (Bef speed) bef 24/11/06
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = True
  Set adc.FirNKeyCtl = IIF(ws_MultiCoMod = True, adc("UoCoCdFr"), adc("UoOmTcFr"))
      adc("UOCOCDFR").Enabled = ws_MultiCoMod
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******

  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
'  gs_LocTyp = ", 'P'"
'  gb_RmCtgFor = False
'  gb_RmZ = False
'  gs_DmTcTyp = "DM"
  gs_OmTcTyp = "JT"
  gs_CmCtg = ""
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  gb_CmCtgFor = False
  adc("UoCoCdFr") = gs_CoCd
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UOCOCDFR") + "' and HLotNoYN= 'Y' ") Then
    LblLotNo.Visible = False: adc("UoLotNoFr").Visible = False
    adc("UoLotNoTo").Visible = False: adc("UoLotNoSel").Visible = False
    adc("UoLotNoFr") = "": adc("UoLotNoTo") = "": adc("UoLotNoSel") = ""
  End If
  ' Zubin 212

  
  ws_FrmCaption = Me.Caption  '****** Sachin 2.14.0
  Call SetGroupSort
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  adc("UoCoCdFr") = gs_CoCd
  If adc("UoDtFr") = "" Then adc("UoDtFr") = "01/01/1980"
  If adc("UoDtTo") = "" Then adc("UoDtTo") = "01/01/1980"
'  adc("UoLsDtTo") = "01/01/1980"
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(adc("UOCOCDFR"), ws_FrmCaption)
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
 End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  gb_CoCdFor = True
  Select Case UCase(IdName)
    Case Is = UCase("UoCoCdFr")
     If Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'") Then _
          Cancel = True: ErrMsg = "Invalid Company Code": Exit Sub
     Me.Caption = GetFrmCaption(pv_NewValue, ws_FrmCaption)
  End Select
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = True
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
  '*** Jay 3.01 **** [LotNo Err]
  gs_RmLocTyp = "": gs_Loc = ""
  '*** Jay 3.01 **** [LotNo Err]

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
'  Rep1.wCoCd.SetText gs_CoNm
'  Rep2.wCoCd.SetText gs_CoNm
'  Rep3.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + adc("UoCoCdFr") + "' and " + _
                                                    "HCd = '" + ctSelfCmCd + "' ")
  Rep.wCoCd.SetText ws_HName
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  mRep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    mRep.RHMain.UnderlaySection = True: mRep.wCoCd.Suppress = True
  Else
    mRep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  mRep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  
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
  Set moCn = Nothing '*** (Bef speed) bef 24/11/06
  Set mRep = Nothing
  Set Rep = Nothing
 End Sub
