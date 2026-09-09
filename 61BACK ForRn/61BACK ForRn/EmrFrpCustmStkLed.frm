VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpCustmStkLed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Stock Ledger"
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
      Left            =   11550
      TabIndex        =   49
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
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5655
      TabIndex        =   58
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   59
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
         TabIndex        =   60
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
      Left            =   6870
      TabIndex        =   57
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   54
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
      TabIndex        =   50
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
      TabIndex        =   53
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   52
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   56
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   51
               Top             =   360
               Width           =   9375
               _ExtentX        =   16536
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Caption         =   "s"
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   55
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraOptDerVal 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1560
               TabIndex        =   87
               ToolTipText     =   "Select Derived Value or Actual Value"
               Top             =   8040
               Width           =   4875
            End
            Begin VB.Frame FraOptDerRt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1560
               TabIndex        =   86
               ToolTipText     =   "Select Derived Rate or Actual Rate"
               Top             =   7755
               Width           =   4905
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4035
               TabIndex        =   2
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   1470
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1545
               TabIndex        =   1
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   1470
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   4035
               TabIndex        =   20
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   3465
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
               Left            =   6540
               TabIndex        =   21
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   3465
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
               Index           =   4
               Left            =   4035
               TabIndex        =   17
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   3180
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
               Left            =   6540
               TabIndex        =   18
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   3180
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
               Left            =   1545
               TabIndex        =   13
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   2895
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
               Left            =   1545
               TabIndex        =   16
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   3180
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
               Index           =   3
               Left            =   1545
               TabIndex        =   19
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   3465
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
               Index           =   16
               Left            =   4035
               TabIndex        =   14
               ToolTipText     =   "Enter To Raw Material Category"
               Top             =   2895
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
               Left            =   6540
               TabIndex        =   15
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   2895
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
               Index           =   44
               Left            =   4035
               TabIndex        =   9
               ToolTipText     =   "Enter To Supplier"
               Top             =   2325
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "TSuppCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   6540
               TabIndex        =   10
               ToolTipText     =   "Enter Supplier Selection"
               Top             =   2325
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "TSuppCd In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   1545
               TabIndex        =   8
               ToolTipText     =   "Enter From Supplier"
               Top             =   2325
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "TSuppCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   4035
               TabIndex        =   4
               ToolTipText     =   "Enter To Bill Of Entry Number "
               Top             =   1755
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBOETO"
               IdName          =   "UOBOETO"
               CmpStr          =   "TBOENo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   6540
               TabIndex        =   5
               ToolTipText     =   "Enter Bill Of Entry Number Selection"
               Top             =   1755
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOBOESEL"
               IdName          =   "UOBOESEL"
               CmpStr          =   "TBOENo In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   1545
               TabIndex        =   3
               ToolTipText     =   "Enter From Bill Of Entry Number "
               Top             =   1755
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBOEFR"
               IdName          =   "UOBOEFR"
               CmpStr          =   "TBOENo >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   4035
               TabIndex        =   7
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   2040
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOBOEDTTO"
               IdName          =   "UOBOEDTTO"
               CmpStr          =   "TBOEDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1545
               TabIndex        =   6
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   2040
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOBOEDTFR"
               IdName          =   "UOBOEDTFR"
               CmpStr          =   "TBOEDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4035
               TabIndex        =   12
               ToolTipText     =   "Enter To Bill Of Entry Number "
               Top             =   2610
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBILLNOTO"
               IdName          =   "UOBILLNOTO"
               CmpStr          =   "TBillNo <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1545
               TabIndex        =   11
               ToolTipText     =   "Enter From Bill Of Entry Number "
               Top             =   2610
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBILLNOFR"
               IdName          =   "UOBILLNOFR"
               CmpStr          =   "TBillNo >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   4035
               TabIndex        =   25
               ToolTipText     =   "Enter To Purchase Rate"
               Top             =   4035
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOSTKRTTO"
               IdName          =   "UOSTKRTTO"
               CmpStr          =   "TzPurRt <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   1545
               TabIndex        =   24
               ToolTipText     =   "Enter From Purchase Rate"
               Top             =   4035
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOSTKRTFR"
               IdName          =   "UOSTKRTFR"
               CmpStr          =   "TzPurRt >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   4035
               TabIndex        =   27
               ToolTipText     =   "Enter To Average Piece Per Carat"
               Top             =   4320
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOPCPERCTTO"
               IdName          =   "UOPCPERCTTO"
               CmpStr          =   "TzPcPerCt <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   1545
               TabIndex        =   26
               ToolTipText     =   "Enter From Average Piece Per Carat"
               Top             =   4320
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOPCPERCTFR"
               IdName          =   "UOPCPERCTFR"
               CmpStr          =   "TzPcPerCt >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   1545
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   615
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "TzCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1545
               TabIndex        =   28
               ToolTipText     =   "Enter From Import Type Scope"
               Top             =   4605
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
               CmpStr          =   "TImpTyp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4035
               TabIndex        =   29
               ToolTipText     =   "Enter To Import Type Scope"
               Top             =   4605
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
               CmpStr          =   "TImpTyp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   6540
               TabIndex        =   30
               ToolTipText     =   "Enter Import Type Selection"
               Top             =   4605
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPMCDSEL"
               IdName          =   "UOPMCDSEL"
               CmpStr          =   "TImpTyp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1545
               TabIndex        =   22
               ToolTipText     =   "Enter Raw Material Stone Class From Range"
               Top             =   3750
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSFR"
               IdName          =   "UOSTNCLSFR"
               CmpStr          =   "IsNull(PSCd, '') >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4035
               TabIndex        =   23
               ToolTipText     =   "Enter Raw Material Stone Class To Range"
               Top             =   3750
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSTO"
               IdName          =   "UOSTNCLSTO"
               CmpStr          =   "IsNull(PSCd, '') <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   4815
               TabIndex        =   38
               ToolTipText     =   "Enter To Order Year"
               Top             =   5190
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "TzOdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   4035
               TabIndex        =   37
               ToolTipText     =   "Enter To Order Tc"
               Top             =   5190
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "TzOdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   5235
               TabIndex        =   39
               ToolTipText     =   "Enter To Order Character"
               Top             =   5190
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "TzOdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   1545
               TabIndex        =   40
               ToolTipText     =   "Enter From Order Number"
               Top             =   5475
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "TzOdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   4035
               TabIndex        =   41
               ToolTipText     =   "Enter To Order Number"
               Top             =   5475
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "TzOdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   2325
               TabIndex        =   35
               ToolTipText     =   "Enter From Order Year"
               Top             =   5190
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "TzOdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   1545
               TabIndex        =   34
               ToolTipText     =   "Enter From Order Tc"
               Top             =   5190
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "TzOdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   2745
               TabIndex        =   36
               ToolTipText     =   "Enter From Order Character"
               Top             =   5190
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "TzOdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   4035
               TabIndex        =   43
               ToolTipText     =   "Enter To Order Design Code"
               Top             =   5760
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "TzDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   6540
               TabIndex        =   44
               ToolTipText     =   "Enter Order Design Code Selection"
               Top             =   5760
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "TzDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   1545
               TabIndex        =   42
               ToolTipText     =   "Enter From Order Design Code"
               Top             =   5760
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "TzDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   4035
               TabIndex        =   46
               ToolTipText     =   "Enter Order Raw Material Code To Range"
               Top             =   6045
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCD1TO"
               IdName          =   "UORMCD1TO"
               CmpStr          =   "TzRmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   6540
               TabIndex        =   47
               ToolTipText     =   "Enter Order Raw Material Code Selection"
               Top             =   6045
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCD1SEL"
               IdName          =   "UORMCD1SEL"
               CmpStr          =   "TzRmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   1545
               TabIndex        =   45
               ToolTipText     =   "Enter Order Raw Material Code From Range"
               Top             =   6045
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCD1FR"
               IdName          =   "UORMCD1FR"
               CmpStr          =   "TzRmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   64
               Left            =   4035
               TabIndex        =   32
               ToolTipText     =   "Enter To Customer"
               Top             =   4890
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOEXPCMCDTO"
               IdName          =   "UOEXPCMCDTO"
               CmpStr          =   "TCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   65
               Left            =   6540
               TabIndex        =   33
               ToolTipText     =   "Enter Customer Selection"
               Top             =   4890
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOEXPCMCDSEL"
               IdName          =   "UOEXPCMCDSEL"
               CmpStr          =   "TCmCd In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   66
               Left            =   1545
               TabIndex        =   31
               ToolTipText     =   "Enter From Customer"
               Top             =   4890
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOEXPCMCDFR"
               IdName          =   "UOEXPCMCDFR"
               CmpStr          =   "TCmCd >= "
            End
            Begin VB.Label LblCustCd 
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
               Left            =   60
               TabIndex        =   85
               Top             =   4890
               Width           =   1425
            End
            Begin VB.Label Label4 
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
               TabIndex        =   84
               Top             =   6045
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
               TabIndex        =   83
               Top             =   5760
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
               Index           =   26
               Left            =   4755
               TabIndex        =   82
               Top             =   5190
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
               Left            =   5175
               TabIndex        =   81
               Top             =   5190
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               Left            =   60
               TabIndex        =   80
               Top             =   5475
               Width           =   1395
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               Left            =   60
               TabIndex        =   79
               Top             =   5190
               Width           =   1515
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
               Left            =   2685
               TabIndex        =   78
               Top             =   5190
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
               Left            =   2265
               TabIndex        =   77
               Top             =   5190
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Stone Class"
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
               Left            =   60
               TabIndex        =   76
               Top             =   3750
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Import Type"
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
               TabIndex        =   75
               Top             =   4605
               Width           =   1545
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
               Index           =   36
               Left            =   60
               TabIndex        =   74
               Top             =   615
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Avg Pc Per Ct"
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
               TabIndex        =   73
               Top             =   4320
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Purchase Rate"
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
               TabIndex        =   72
               Top             =   4035
               Width           =   1545
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
               Left            =   6540
               TabIndex        =   71
               Top             =   180
               Width           =   1275
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Supp Bill No"
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
               Top             =   2610
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "BOE Date"
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
               TabIndex        =   69
               Top             =   2040
               Width           =   1425
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
               TabIndex        =   68
               Top             =   1470
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
               Top             =   3465
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
               TabIndex        =   66
               Top             =   3180
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
               TabIndex        =   65
               ToolTipText     =   "Location"
               Top             =   2895
               Width           =   1425
            End
            Begin VB.Label LblSuppCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Supplier"
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
               TabIndex        =   64
               Top             =   2325
               Width           =   1425
            End
            Begin VB.Label LblBOENo 
               BackStyle       =   0  'Transparent
               Caption         =   "BOE No"
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
               TabIndex        =   63
               Top             =   1755
               Width           =   1425
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
               Left            =   1545
               TabIndex        =   62
               Top             =   180
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
               Left            =   4035
               TabIndex        =   61
               Top             =   180
               Width           =   1275
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpCustmStkLed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepCustmStkLed
Dim moCn As MwfLib.MDOConnection
Private Sub SetGroupSort()
  GRP_REP.Add "Rm Code", "RmCtg+RmSCtg+TzRmZCd", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ TzRmZCd", "TzRmZCd", "hTzRmZCd", "Rm Ctg,Rm SubCtg", "", "", "SubString(RmDesc, 1, 30) "
  GRP_REP.Add "Rm Ctg", "RmCtg", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = TzRmCtg) "
  GRP_REP.Add "Rm SubCtg", "RmCtg+RmSCtg", "'('+ RTrim(RmCtg)+ ') '+ RmSCtg", "", "", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
  GRP_REP.Add "Purchase Rt", "str(case when qBaseCur= 'Y' then (case when TMulDiv= 'M' then TzPurRt/TCnvRt else TzPurRt*TCnvRt end) else TzPurRt End,9,2)", "", "TzPurRt", "hTzPurRt", "", "", "", ""
  GRP_REP.Add "(RmCd) PurRt", "TzRmZCd+str(case when qBaseCur= 'Y' then (case when TMulDiv= 'M' then TzPurRt/TCnvRt else TzPurRt*TCnvRt end) else TzPurRt End,9,2)", "'('+TzRmZCd+') '+LTrim(str(case when qBaseCur= 'Y' then (case when TMulDiv= 'M' then TzPurRt/TCnvRt else TzPurRt*TCnvRt end) else TzPurRt End,9,2))", "TzRmZCd,TzPurRt", "hTzRmZCd,hTzPurRt", "Rm Code,Purchase Rt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  GRP_REP.Add "Stone Class", "RmCtg+IsNull(PSCd, '')", "'('+ RTrim(RmCtg)+') '+IsNull(PSCd, '')", "", "", "Rm Ctg", "", "", "PDesc"
  
End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  '*** Report Sql ***
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'    space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'    space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'space(1) as qFlag, RmCd, TdLotNo as qLotNo,RmCtg, PSCd As qStnCls, RmSCtg,
'TdCoCd, TdTc, TdYy,TdChr, TdNo, TdSr, TdSrNo,TDt as qDt,TBoeNo,TBoeDt,0.0 as qOpnWt, TdRmWt as qDrWt, TdRmWt as qCrWt,
'TCnvRt,TCnvRtRs,TzPurRt,TzPurAmt,TzPurRt  as qTzPurBaseRt,TzPurAmt as qTzPurBaseAmt,
'        TzPurRt  As qTzPurReRt, TzPurAmt As qTzPurReAmt
'From TxndZ,Txnd, Txn, RmMst,Param  where 1=2
  '*** Report Sql ***
  
  Dim grpflds0 As String, grpflds1 As String, grpflds2 As String, grpflds3 As String, wrepcnd As String
  Dim grpflds As String, wCnd As String, wSqlStrg As String
  Dim wStr As String, wSetTmpTab As String
  Dim wDtCndOpnTd As String, wDtCndDetTd As String
  Dim wDtCndOpnInv As String, wDtCndDetInv As String, wDtCndOpnTz As String, wDtCndDetTz As String
  Dim mwDtFr As String, wDtTo As String
  Dim wFrDt As String, wToDt As String
  
  
  '*** Added Again In 3.2.0
  Dim wFct As String 'wPurStr As String   '*** Jay 3.1.0_1 [Derived Value]
  Dim wBaseFct As String, wReFct As String, wRepRtFlds As String '*** Jay 3.1.1 (All Rates Rpt)
  '*** Added Again In 3.2.0
  
  If (adc("UoStnClsFr") <> "" Or adc("UoStnClsTo") <> "") And ((adc("UoRmCtgFr") <> adc("UoRmCtgTo")) Or (adc("UoRmCtgFr") = "")) Then
      DispMsg "Please select an RmCtg to provide a filter of Stone Class", etError: Exit Sub
  End If
  
  
  Dim wAddlCnd As String '****** Sachin 3.02.0 - New Fields In Where Clause
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Set mRep = Rep
  
  Call DispCoNm
  Call SetMwName(mRep)
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  '*** To check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then wDetPos = i: Exit For
  Next i
  
   grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)
  grpflds2 = GRP_REP.GrpFldLst(2)
  grpflds3 = GRP_REP.GrpFldLst(3)
  
  
wFrDt = CStr(adc("UoDtFr")): wToDt = CStr(adc("UoDtTo"))
wDtCndOpnTd = " and TDt < '" + wFrDt + "'"
wDtCndOpnInv = " and InDt < '" + wFrDt + "'"
wDtCndDetTd = " and TDt >= '" + wFrDt + "' and TDt<='" + wToDt + "'"
wDtCndDetInv = " and InDt >= '" + wFrDt + "' and InDt<='" + wToDt + "'"

wDtCndOpnTz = " and TzClsDt < '" + wFrDt + "'"
wDtCndDetTz = " and TzClsDt >= '" + wFrDt + "' and TzClsDt<='" + wToDt + "'"

wCnd = IIF(wrepcnd <> "", " and " + wrepcnd + " Option (MaxDop 1)", " Option (MaxDop 1)")
  '6.1
wSetTmpTab = "SET DateFormat " + DtFmtStr() + ";Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3, " + _
            " space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3, " + _
            "  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc, " + _
            " space(1) as qFlag, RmCd, TdLotNo as qLotNo,RmCtg, PSCd As qStnCls, RmSCtg, " + _
            " TdCoCd, TdTc, TdYy,TdChr, TdNo, TdSr, TdSrNo,TDt as qDt,TBoeNo,TBoeDt,TdRmWt as qOpnWt, TdRmWt as qDrWt, TdRmWt as qCrWt, " + _
            " TCnvRt,TCnvRtRs,TzPurRt,TzPurAmt,TzPurRt  as qTzPurBaseRt,TzPurAmt as qTzPurBaseAmt, " + _
            " TzPurRt  As qTzPurReRt, TzPurAmt As qTzPurReAmt into #TmpCustmStkLed From TxndZ,Txnd, Txn, RmMst,Param  where 1=2 "
     moCn.CreateTmpTable "#TmpCustmStkLed", wSetTmpTab
  
'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ TzRmSCtg +',%' changed as PSCd = RmStnCls
 wSqlStrg = "Insert Into #TmpCustmStkLed  Select " + grpflds0 + ", " + _
               " 'O' as qFlag,TzRmZCd as RmCd,'' as qLotNo,TzRmCtg as RmCtg,IsNull(PSCd,'') as qStnCls,TzRmSCtg as RmSCtg, " + _
               " TzCoCd as TdCoCd,'' as TdTc,'' as TdYy,'' as TdChr,0 as TdNo, " + _
               " 0 as TdSr,0 as TdSrNo, TDt as qDt,TBoeNo,TBoeDt,TzRmWt as qOpnWt, 0 as qDrWt ,0 as qCrWt, " + _
               " (case when TMulDiv= 'M' and TCnvRt>0 then 1/TCnvRt else TCnvRt end) as TCnvRt , " + _
               " (case when TMulDivRs= 'D' and TCnvRt>0 then 1/TCnvRtRs else TCnvRtRs end) as TCnvRtRs, " + _
               " TzPurRt   as TzPurRt,TzPurAmt  as TzPurAmt, " + _
               " convert (decimal (14, 4),TzPurRt *  (case when TMulDiv= 'M' then 1/TCnvRt else TCnvRt end))  as qTzPurBaseRt, " + _
               " convert (decimal (14, 4),TzPurAmt *  (case when TMulDiv= 'M' then 1/TCnvRt else TCnvRt end))  as qTzPurBaseAmt, " + _
               " convert (decimal (14, 4),TzPurRt *  (case when TMulDivRs= 'D' then 1/TCnvRtRs else TCnvRtRs end))  as qTzPurReRt, " + _
               " convert (decimal (14, 4),TzPurAmt *  (case when TMulDivRs= 'D' then 1/TCnvRtRs else TCnvRtRs end))  as qTzPurReAmt " + _
               " From TxndZ B Join Txn On TzCoCd=TCoCd and TzTc=TTc and TzYy=TYy and TzChr=TChr and TzNo=TNo " + _
               " Join RmMst On RmCd=TzRmZCd " + _
               " Left Outer Join Param on PTyp= 'STNCLS' and PmCd= TzRmCtg and PSCd = RmStnCls " + _
               " where TzBM='B' " + wDtCndOpnTd + wCnd
        moCn.Execute wSqlStrg
'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ TzRmSCtg +',%' changed as PSCd = RmStnCls
   wSqlStrg = "Insert Into #TmpCustmStkLed  Select " + grpflds0 + ", " + _
              "'O' as qFlag,TzRmZCd as RmCd,'' as qLotNo,TzRmCtg as RmCtg,IsNull(PSCd,'') as qStnCls,TzRmSCtg as RmSCtg," + _
              " TzCoCd as TdCoCd,'' as TdTc,'' as TdYy,'' as TdChr,0 as TdNo, " + _
              " 0 as TdSr,0 as TdSrNo, InDt as qDt,InExpNo,InPrnDt,-1*TzRmWt as qOpnWt, 0 as qDrWt ,0 as qCrWt, " + _
              " (case when InMulDiv= 'M' and InCnvFct>0 then 1/InCnvFct else InCnvFct end) as TCnvRt, " + _
              " (case when InRsMulDiv= 'D' and inCnvRt>0 then 1/InCnvRt else InCnvRt end) as TCnvRtRs, " + _
              "  TzPurRt   as TzPurRt, -1*TzPurAmt  as TzPurAmt, " + _
              "  convert (decimal (14, 4),TzPurRt *  (case when InMulDiv= 'M' then 1/InCnvFct else InCnvFct end))  as qTzPurBaseRt, " + _
              "  convert (decimal (14, 4),-1*TzPurAmt *  (case when InMulDiv= 'M'  then 1/InCnvFct else InCnvFct end))  as qTzPurBaseAmt, " + _
              "  convert (decimal (14, 4),TzPurRt *  (case when InRsMulDiv= 'D' and InCnvRt>0 then 1/ InCnvRt else InCnvRt end))  as qTzPurReRt, " + _
              "  convert (decimal (14, 4),-1*TzPurAmt *  (case when InRsMulDiv= 'D' and InCnvRt>0 then 1/ InCnvRt else InCnvRt end))  as qTzPurReAmt " + _
              " From TxndZ  Join InvHd On TzCoCd=InCoCd and TzTc=InTc and TzYy=InYy and TzChr=InChr and TzNo=InNo " + _
              " Join RmMst On RmCd=TzRmZCd " + _
              "  Left Outer Join Param on PTyp= 'STNCLS' and PmCd= TzRmCtg and PSCd = RmStnCls " + _
              "  where TzBM='M' " + wDtCndOpnInv + wCnd
            moCn.Execute wSqlStrg
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ TzRmSCtg +',%' changed as PSCd = RmStnCls
   wSqlStrg = "Insert Into #TmpCustmStkLed  Select " + grpflds0 + ", " + _
              " 'T' as qFlag,TzRmZCd as RmCd,'' as qLotNo,TzRmCtg as RmCtg,IsNull(PSCd,'') as qStnCls,TzRmSCtg as RmSCtg, " + _
              "  TzCoCd as TdCoCd,TzTc as TdTc,TzYy as TdYy,TzChr as TdChr,TzNo as TdNo, " + _
              "  TzSr as TdSr,TzSrNo as TdSrNo, TDt as qDt,TBoeNo,TBoeDt,0 as qOpnWt, TzRmWt as qDrWt ,0 as qCrWt, " + _
              "  (case when TMulDiv= 'M' and TCnvRt>0 then 1/TCnvRt else TCnvRt end) as TCnvRt , " + _
              "  (case when TMulDivRs= 'D' and TCnvRtRs>0 then 1/TCnvRtRs else TCnvRtRs end) as TCnvRtRs, " + _
              "   TzPurRt   as TzPurRt, TzPurAmt  as TzPurAmt, " + _
              "   convert (decimal (14, 4),TzPurRt *  (case when TMulDiv= 'M' and TCnvRt>0 then 1/TCnvRt else TCnvRt end))  as qTzPurBaseRt, " + _
              "   convert (decimal (14, 4),TzPurAmt *  (case when TMulDiv= 'M'  and TCnvRt>0 then 1/TCnvRt else TCnvRt end))  as qTzPurBaseAmt, " + _
              "    convert (decimal (14, 4),TzPurRt *  (case when TMulDivRs= 'D'  and TCnvRtRs>0 then 1/TCnvRtRs else TCnvRtRs end))  as qTzPurReRt, " + _
              "   convert (decimal (14, 4),TzPurAmt *  (case when TMulDivRs= 'D'  and TCnvRtRs>0 then 1/TCnvRtRs else TCnvRtRs end))  as qTzPurReAmt " + _
              "  From TxndZ Join Txn On TzCoCd=TCoCd and TzTc=TTc and TzYy=TYy and TzChr=TChr and TzNo=TNo " + _
              " Join RmMst On RmCd=TzRmZCd " + _
              "  Left Outer Join Param on PTyp= 'STNCLS' and PmCd= TzRmCtg and PSCd = RmStnCls " + _
              "    where TzBM='B' " + wDtCndDetTd + wCnd
            moCn.Execute wSqlStrg
'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ TzRmSCtg +',%' changed as PSCd = RmStnCls
   wSqlStrg = "Insert Into #TmpCustmStkLed  Select " + grpflds0 + ", " + _
              "'T' as qFlag,TzRmZCd as RmCd,'' as qLotNo,TzRmCtg as RmCtg,IsNull(PSCd,'') as qStnCls,TzRmSCtg as RmSCtg, " + _
              " TzCoCd as TdCoCd,TzTc as TdTc,TzYy as TdYy,TzChr as TdChr,TzNo as TdNo, " + _
              " TzSr as TdSr,TzSrNo as TdSrNo, InDt as qDt,InExpNo,InPrnDt,0 as qOpnWt,0 as qDrWt,TzRmWt as qCrWt, " + _
              "  (case when InMulDiv= 'M' and InCnvFct>0 then 1/InCnvFct else InCnvFct end) as TCnvRt, " + _
              " (case when InRsMulDiv= 'D' and InCnvRt>0 then 1/InCnvRt else InCnvRt end) as TCnvRtRs, " + _
              "   TzPurRt   as TzPurRt, -1*TzPurAmt  as TzPurAmt, " + _
              "  convert (decimal (14, 4),TzPurRt *  (case when InMulDiv= 'M' and InCnvFct>0 then 1/InCnvFct else InCnvFct end))  as qTzPurBaseRt, " + _
              "  convert (decimal (14, 4),-1*TzPurAmt *  (case when InMulDiv= 'M' and InCnvFct>0 then 1/InCnvFct else InCnvFct end))  as qTzPurBaseAmt, " + _
              "  convert (decimal (14, 4),TzPurRt *  (case when InRsMulDiv= 'D' and InCnvRt>0  then 1/ InCnvRt else InCnvRt end))  as qTzPurReRt, " + _
              "  convert (decimal (14, 4),-1*TzPurAmt *  (case when InRsMulDiv= 'D' and InCnvRt>0 then 1/ InCnvRt else InCnvRt end))  as qTzPurReAmt " + _
              "  From TxndZ   Join InvHd On TzCoCd=InCoCd and TzTc=InTc and TzYy=InYy and TzChr=InChr and TzNo=InNo " + _
              " Join RmMst On RmCd=TzRmZCd " + _
              "  Left Outer Join Param on PTyp= 'STNCLS' and PmCd= TzRmCtg and PSCd = RmStnCls " + _
              "  where TzBM='M' " + wDtCndDetInv + wCnd
            moCn.Execute wSqlStrg
'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ TzRmSCtg +',%' changed as PSCd = RmStnCls
   wSqlStrg = "Insert Into #TmpCustmStkLed  Select " + grpflds0 + ", " + _
              " 'O' as qFlag,TzRmZCd as RmCd,'' as qLotNo,TzRmCtg as RmCtg,IsNull(PSCd,'') as qStnCls,TzRmSCtg as RmSCtg, " + _
              "  TzCoCd as TdCoCd,'' as TdTc,'' as TdYy,'' as TdChr,0 as TdNo, " + _
              "  0 as TdSr,0 as TdSrNo, TDt as qDt,TBoeNo,TBoeDt,-1*(TzRmWt-TzRmDespWt) as qOpnWt,0 as qDrWt ,0 as qCrWt, " + _
              "  (case when TMulDiv= 'M' and TCnvRt>0 then 1/TCnvRt else TCnvRt end) as TCnvRt ,  " + _
              "  (case when TMulDivRs= 'D' and TCnvRtRs>0 then 1/TCnvRtRs else TCnvRtRs end) as TCnvRtRs, " + _
              "   TzPurRt   as TzPurRt,-1*TzPurAmt  as TzPurAmt, " + _
              "   convert (decimal (14, 4),TzPurRt *  (case when TMulDiv= 'M' and TCnvRt>0 then 1/TCnvRt else TCnvRt end))  as qTzPurBaseRt, " + _
              "   convert (decimal (14, 4),TzPurAmt *  (case when TMulDiv= 'M'  and TCnvRt>0 then 1/TCnvRt else TCnvRt end))  as qTzPurBaseAmt, " + _
              "   convert (decimal (14, 4),TzPurRt *  (case when TMulDivRs= 'D'  and TCnvRtRs>0 then 1/TCnvRtRs else TCnvRtRs end))  as qTzPurReRt, " + _
              "   convert (decimal (14, 4),TzPurAmt *  (case when TMulDivRs= 'D'  and TCnvRtRs>0 then 1/TCnvRtRs else TCnvRtRs end))  as qTzPurReAmt " + _
              "   From TxndZ B Join Txn On TzCoCd=TCoCd and TzTc=TTc and TzYy=TYy and TzChr=TChr and TzNo=TNo " + _
              " Join RmMst On RmCd=TzRmZCd " + _
              "   Left Outer Join Param on PTyp= 'STNCLS' and PmCd= TzRmCtg and PSCd = RmStnCls " + _
              "   where TzBM='B' and TzCls='Y' " + wDtCndOpnTz + wCnd
            moCn.Execute wSqlStrg
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ TzRmSCtg +',%' changed as PSCd = RmStnCls
   wSqlStrg = "Insert Into #TmpCustmStkLed  Select " + grpflds0 + ", " + _
              " 'T' as qFlag,TzRmZCd as RmCd,'' as qLotNo,TzRmCtg as RmCtg,IsNull(PSCd,'') as qStnCls,TzRmSCtg as RmSCtg," + _
              "  TzCoCd as TdCoCd,TzTc as TdTc,TzYy as TdYy,TzChr as TdChr,TzNo as TdNo, " + _
              "  TzSr as TdSr,TzSrNo as TdSrNo, TDt as qDt,TBoeNo,TBoeDt,0 as qOpnWt,0 as qDrWt ,TzRmWt-TzRmDespWt as qCrWt, " + _
              "   (case when TMulDiv= 'M' and TCnvRt>0 then 1/TCnvRt else TCnvRt end) as TCnvRt , " + _
              "   (case when TMulDivRs= 'D' and TCnvRtRs>0 then 1/TCnvRtRs else TCnvRtRs end) as TCnvRtRs, " + _
              "   TzPurRt   as TzPurRt,-1*TzPurAmt  as TzPurAmt, " + _
              "   convert (decimal (14, 4),TzPurRt *  (case when TMulDiv= 'M' and TCnvRt>0 then 1/TCnvRt else TCnvRt end))  as qTzPurBaseRt, " + _
              "   convert (decimal (14, 4),-1*TzPurAmt *  (case when TMulDiv= 'M'  and TCnvRt>0 then 1/TCnvRt else TCnvRt end))  as qTzPurBaseAmt, " + _
              "   convert (decimal (14, 4),TzPurRt *  (case when TMulDivRs= 'D'  and TCnvRtRs>0 then 1/TCnvRtRs else TCnvRtRs end))  as qTzPurReRt, " + _
              "   convert (decimal (14, 4),-1*TzPurAmt *  (case when TMulDivRs= 'D'  and TCnvRtRs>0 then 1/TCnvRtRs else TCnvRtRs end))  as qTzPurReAmt " + _
              "  From TxndZ B Join Txn On TzCoCd=TCoCd and TzTc=TTc and TzYy=TYy and TzChr=TChr and TzNo=TNo " + _
              " Join RmMst On RmCd=TzRmZCd " + _
              "   Left Outer Join Param on PTyp= 'STNCLS' and PmCd= TzRmCtg and PSCd = RmStnCls " + _
              "    where TzBM='B' and TzCls='Y' " + wDtCndDetTz + wCnd
            moCn.Execute wSqlStrg
            
   
  adc.RepRecSource = "Select max(Grp1) as Grp1, max(Grp2) as Grp2, max(Grp3) as Grp3, " + _
      "DGrp1, DGrp2, DGrp3, max(G1Desc) as G1Desc, max(G2Desc) as G2Desc, max(G3Desc) as G3Desc, " + _
      "max(qFlag) as qFlag, max(RmCd) as RmCd, max(qLotNo) as qLotNo, " + _
      " max(RmCtg) as RmCtg, max(qStnCls) as qStnCls,max(RmSCtg) as RmSCtg, " + _
      "TdCoCd, TdTc, TdYy, TdChr, TdNo, max(TdSr) as TdSr,max(TdSrNo) as TdSrNo,Max(qDt) as qDt,Max(TBoeNo) as TBoeNo,Max(TBoeDt) as TBoeDt," + _
      "sum(qOpnWt) as qOpnWt,sum(qDrWt) as qDrWt, sum(qCrWt) as qCrWt, max(TCnvRt) as TCnvRt, max(TCnvRtRs) as TCnvRtRs,max(TzPurRt) as TzPurRt,Sum(TzPurAmt)  as TzPurAmt, " + _
      " max(qTzPurBaseRt) as qTzPurBaseRt,Sum(qTzPurBaseAmt) as qTzPurBaseAmt,Max(qTzPurReRt) as qTzPurReRt, Sum(qTzPurReAmt) as qTzPurReAmt" + _
      " From #TmpCustmStkLed Tmp Group By DGrp1, DGrp2, DGrp3,qFlag,TdCoCd,TdTc,TdYy,TdChr,TdNo,TdSr,TdSrNo " '+ ws_DetGrp
            
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.RmCd}", 7)
  
 ' adc.RepRecSource = wSqlStrg + ctMaxDopOpt  'Sachin 3.01 ctMaxDopOpt added
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
    Set adc.FirNKeyCtl = adc("UoDtFr") 'adc("UoCoCdFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_TxnTcTyp = "TZ"
    ' **** Bef 3.03 - Manali
    'gs_CmCtg = "S"
    
    'URMI 2.04-3 hELP OF RmCd should show only ZRmCd
    gb_RmZ = True
    gb_RmZ1 = False '*** (Jen 3.2.0)
    gb_RmCtgFor = False
    
    gs_PTyp = "ImpTyp"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_PTyp = "ImpTyp"
  
  'URMI 2.04-3 hELP OF RmCd should show only ZRmCd
  gb_RmZ = True
  gb_RmZ1 = False '*** (Jen 3.2.0)
  gb_RmCtgFor = False
  
  ' **** Manali 3.03 - 30/06/08
  Select Case UCase(IdName)
    Case Is = UCase("UoCmCdFr"), UCase("UoCmCdTo"), UCase("UoCmCdSel")
      gs_CmCtg = "S"
    Case Is = UCase("UoExpCmCdFr"), UCase("UoExpCmCdTo"), UCase("UoExpCmCdSel")
      gs_CmCtg = "C"
  End Select
  ' **** Manali 3.03 - 30/06/08
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_PTyp = "ImpTyp"
  
  'URMI 2.04-3 hELP OF RmCd should show only ZRmCd
  gb_RmZ = True
  gb_RmZ1 = False '*** (Jen 3.2.0)
  gb_RmCtgFor = False
  
  ' **** Manali 3.03 - 30/06/08
  Select Case UCase(IdName)
    Case Is = UCase("UoCmCdFr"), UCase("UoCmCdTo"), UCase("UoCmCdSel")
      gs_CmCtg = "S"
    Case Is = UCase("UoExpCmCdFr"), UCase("UoExpCmCdTo"), UCase("UoExpCmCdSel")
      gs_CmCtg = "C"
  End Select
  ' **** Manali 3.03 - 30/06/08
  
  SetProp Me, IdName, When
End Sub

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptBestUse_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UOYN5") = "Y"
Case Is = 1
 adc("UOYN5") = "N"
Case Is = 2
 adc("UOYN5") = ""
End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptDerRt_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UOYN6") = "Y"
Case Is = 1
 adc("UOYN6") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added
Private Sub OptDerRt_GotFocus(Index As Integer)
DispMsg FraOptDerRt.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptDerVal_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UOYN3") = "Y"
Case Is = 1
 adc("UOYN3") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptDerVal_GotFocus(Index As Integer)
DispMsg FraOptDerVal.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

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
 
  
End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                            "HCd = '" + ctSelfCmCd + "' ")
  mRep.wCoCd.SetText ws_HName
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  mRep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    mRep.RHMain.UnderlaySection = True: mRep.wCoCd.Suppress = True
  Else
    mRep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
' ********* Manali 3.2.2 - Before 3.2.2
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  mRep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  mRep.wSrvrTime.SetText ws_DtTm

' ******** Manali 3.2.2 - User date and time in reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  Rep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  'Rep1.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
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
    Set mRep = Nothing
    Set Rep = Nothing
  
    '*** (09/08/05)
End Sub
