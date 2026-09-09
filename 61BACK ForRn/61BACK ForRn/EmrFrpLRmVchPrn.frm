VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpLRmVchPrn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher Printing (Loose Rm)"
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
      Left            =   10890
      TabIndex        =   24
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
      TabIndex        =   25
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
      Left            =   4845
      TabIndex        =   34
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   35
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
         TabIndex        =   36
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
      Left            =   6060
      TabIndex        =   33
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   30
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
      TabIndex        =   26
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
      TabIndex        =   29
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   28
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpLRmVchPrn.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpLRmVchPrn.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   32
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   27
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
            Left            =   60
            TabIndex        =   31
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Rm Desc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   1
               Left            =   30
               TabIndex        =   23
               Tag             =   "ShowRmDesc"
               ToolTipText     =   "Check To Show Rm Description"
               Top             =   3180
               Width           =   1680
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4005
               TabIndex        =   9
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   1395
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
               TabIndex        =   8
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   1395
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
               TabIndex        =   4
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   825
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
               CmpStr          =   "OmYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2295
               TabIndex        =   1
               ToolTipText     =   "Enter From Transaction Year"
               Top             =   825
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
               CmpStr          =   "OmYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1515
               TabIndex        =   0
               ToolTipText     =   "Enter From Transaction Tc"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
               CmpStr          =   "OmTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2715
               TabIndex        =   2
               ToolTipText     =   "Enter From Transaction Character"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
               CmpStr          =   "OmChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   4005
               TabIndex        =   3
               ToolTipText     =   "Enter To Transaction Tc"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
               CmpStr          =   "OmTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   5205
               TabIndex        =   5
               ToolTipText     =   "Enter To Transaction Character"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
               CmpStr          =   "OmChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1515
               TabIndex        =   6
               ToolTipText     =   "Enter From Transaction Number"
               Top             =   1110
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOFR"
               IdName          =   "UOTNOFR"
               CmpStr          =   "OmNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   4005
               TabIndex        =   7
               ToolTipText     =   "Enter To Transaction Number"
               Top             =   1110
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOTO"
               IdName          =   "UOTNOTO"
               CmpStr          =   "OmNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   4005
               TabIndex        =   11
               ToolTipText     =   "Enter To Supplier"
               Top             =   1680
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "OmCmCd<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   6510
               TabIndex        =   12
               ToolTipText     =   "Enter Supplier Selection"
               Top             =   1680
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
               TabIndex        =   10
               ToolTipText     =   "Enter From Supplier"
               Top             =   1680
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OmCmCd>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   4005
               TabIndex        =   17
               ToolTipText     =   "Enter To Loss Location"
               Top             =   2250
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   6510
               TabIndex        =   18
               ToolTipText     =   "Enter Loss Location Selection"
               Top             =   2250
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   1515
               TabIndex        =   16
               ToolTipText     =   "Enter From Loss Location"
               Top             =   2250
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   4005
               TabIndex        =   14
               ToolTipText     =   "Enter To Bill Of Entry Number "
               Top             =   1965
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBOETO"
               IdName          =   "UOBOETO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   6510
               TabIndex        =   15
               ToolTipText     =   "Enter Bill Of Entry Number Selection"
               Top             =   1965
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOBOESEL"
               IdName          =   "UOBOESEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   1515
               TabIndex        =   13
               ToolTipText     =   "Enter From Bill Of Entry Number "
               Top             =   1965
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBOEFR"
               IdName          =   "UOBOEFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   59
               Left            =   4005
               TabIndex        =   20
               ToolTipText     =   "Enter To Loss Date"
               Top             =   2535
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOLSDTTO"
               IdName          =   "UOLSDTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   1515
               TabIndex        =   19
               ToolTipText     =   "Enter From Loss Date"
               Top             =   2535
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOLSDTFR"
               IdName          =   "UOLSDTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   64
               Left            =   1545
               TabIndex        =   52
               ToolTipText     =   "Enter Voucher Currency  US$[Y]/Rs[N]  and ' ' for ALL records in US$ "
               Top             =   3690
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1515
               TabIndex        =   21
               ToolTipText     =   "Enter Currency Code Whose Vouchers You Want To View Or Leave It Blank To View All Vouchers In Base Currency"
               Top             =   2820
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOCURCDFR"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   1965
               TabIndex        =   22
               ToolTipText     =   "Show Pure Weight (Yes / No)"
               Top             =   3120
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
               Index           =   0
               Left            =   1515
               TabIndex        =   55
               ToolTipText     =   "Enter Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "OmCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   4005
               TabIndex        =   56
               ToolTipText     =   "Enter To Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "OmCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   6510
               TabIndex        =   57
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   540
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "OmCoCd In"
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
               Left            =   30
               TabIndex        =   54
               ToolTipText     =   "Location"
               Top             =   2850
               Width           =   1485
            End
            Begin VB.Label LblUSDYN 
               BackStyle       =   0  'Transparent
               Caption         =   "In US$"
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
               TabIndex        =   53
               ToolTipText     =   "Location"
               Top             =   3720
               Visible         =   0   'False
               Width           =   1485
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
               TabIndex        =   51
               Top             =   540
               Width           =   1335
            End
            Begin VB.Label LblLsDt 
               BackStyle       =   0  'Transparent
               Caption         =   "Loss Dt"
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
               TabIndex        =   50
               Top             =   2535
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
               TabIndex        =   49
               Top             =   1965
               Width           =   1425
            End
            Begin VB.Label LblLsLoc 
               BackStyle       =   0  'Transparent
               Caption         =   "Loss Loc"
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
               TabIndex        =   48
               Top             =   2250
               Width           =   1425
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
               TabIndex        =   47
               Top             =   1680
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
               TabIndex        =   46
               Top             =   825
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch No"
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
               TabIndex        =   45
               Top             =   1110
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
               Left            =   5145
               TabIndex        =   44
               Top             =   825
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
               TabIndex        =   43
               Top             =   825
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
               Left            =   4725
               TabIndex        =   42
               Top             =   825
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
               Left            =   2235
               TabIndex        =   41
               Top             =   825
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
               TabIndex        =   40
               Top             =   1395
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
               TabIndex        =   39
               Top             =   60
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
               TabIndex        =   38
               Top             =   60
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
               TabIndex        =   37
               Top             =   60
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpLRmVchPrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepLRmVchPrn
Dim mRep As CRAXDRT.Report
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean, ws_FrmCaption As String
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  adc("UoCoCdFr") = gs_CoCd
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(adc("UOCOCDFR"), ws_FrmCaption)
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
  adc("UoLsDtFr") = "01/01/1980"
  adc("UoLsDtTo") = "01/01/1980"
  
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  '*** SQL for PR, REC, CNV, MV
'Select OmCoCd, OmTc, OmYy, OmChr, OmNo, OmDt, OmFrDmLoc, OmToDmLoc, OmFrDmDc, OmToDmDc, OmCmCd, CmName,
'OmCmCtg, OmCnvFct, OmCmCurCd, OmCmPayTerms,
'OdSr, OmFrDmLoc, OmFrDmDc, RmCtg, RmSCtg, RmPurityWt,
'OdJLRmCd, OdJLLotNo, OdJLLn1, OdJLLn2, OdJLLn3,  '' as qRmSzDesc,  'W' as TdRtByQw,
'''  as qPtrYN, OdJLQty, OdJLWt, OdJLVchRt, OdJLVchVal ,  RmDesc
'From OrdMst, OrdDsg, RmMst, RmIdSz, CustMst where 1=2
  Dim grpflds As String
  Dim wrepcnd As String, wCnd As String, wSqlStrg As String, wi_Opt As Integer
  Dim wPurStr As String, wTTotAmtStr As String
  
  '*** (Jen 2.12)
  Dim wd_BaseGldRmPurity As Double, wd_BasePltRmPurity As Double, wd_BaseSlvRmPurity As Double
  '*** (Jen 2.12)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  Dim wd_BasePdRmPurity As Double
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  
  ' ****** Manali 3.03 - 09/08/07 - RmSzDesc
  Dim wRmSzDesc As String
  ' ****** Manali 3.03 - 09/08/07 - RmSzDesc
  
  Dim wAddlCnd As String  '****** Sachin 3.02.0 - New fields in Where Clause
  
  Set Rep = Nothing
    
    Set adc.RepSource = Rep
    Set mRep = Rep
  
  Call DispCoNm
  Call SetMwName(mRep)   ' ***** Manali 3.6.1 - Microway footer
  'Call GRP_REP.Gen3LRep
'  grpflds = GRP_REP.GrpFldLst
    
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
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  wd_BasePdRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst " + _
                       "on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'L'), 1)")
  If wd_BasePdRmPurity = 0 Then wd_BasePdRmPurity = 1
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
      
  ' ****** Manali 3.03 - 09/08/07 - RmSzDesc
  wRmSzDesc = "(case when (Select PDesc225 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RisRmCtg, '') and PSCd= RmSCtg)= 'Y' then IsNull(RisRmSzDesc, '') Else '' End) "
  ' ****** Manali 3.03 - 09/08/07 - RmSzDesc
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
             
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
'  wAddlCnd = ""
'  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " B.OmDt >= '" + Trim(adc("UODTFR")) + "'", "")
'  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " B.OmDt <= '" + Trim(adc("UODTTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
             
    Dim ws_BaseCurCd As String, ws_Join As String
        ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                      "HCoCd= '" + adc("UoCoCdFr") + "'")
  
        If adc("UoCurCdFr") <> "" Then
          wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "OmCmCurCd='" + adc("UoCurCdFr") + "'"
          wPurStr = "B.OdJLVchRt, B.OdJLVchVal"
 '         wTTotAmtStr = "TNetAmt as qTNetAmt, TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TAddAmt3 as qTAddAmt3, " + _
                         " TAddAmt4 as qTAddAmt4, TTotAmt as qTTotAmt,"
        Else
          wPurStr = " (case when OmMulDiv= 'M' then B.OdJLVchRt/OmCnvFct else B.OdJLVchRt*OmCnvFct end) as OdJLVchRt, (case when OmMulDiv= 'M' then B.OdJLVchVal/OmCnvFct else B.OdJLVchVal*OmCnvFct end) as OdJLVchVal "
  '        wTTotAmtStr = "(case when TMulDiv= 'M' then TNetAmt/TCnvRt Else TNetAmt*TCnvRt End) as qTNetAmt, " + _
                        " (case when TMulDiv= 'M' then TAddAmt1/TCnvRt Else TAddAmt1*TCnvRt End) as qTAddAmt1, " + _
                        " (case when TMulDiv= 'M' then TAddAmt2/TCnvRt Else TAddAmt2*TCnvRt End) as qTAddAmt2, " + _
                        " (case when TMulDiv= 'M' then TAddAmt3/TCnvRt Else TAddAmt3*TCnvRt End) as qTAddAmt3, " + _
                        " (case when TMulDiv= 'M' then TAddAmt4/TCnvRt Else TAddAmt4*TCnvRt End) as qTAddAmt4, " + _
                        " (case when TMulDiv= 'M' then TTotAmt/TCnvRt Else TTotAmt*TCnvRt End) as qTTotAmt, "
        End If
        
  '      ws_Join = " Join (Select PMCd as qPMCd, (case when " + IIF(adc("UoCurCdFr") = "", "1", "0") + " = 1 then 'Y' " + _
                  "else 'N' end) as qBaseCur from Param where PTyp= 'CURNCY') Cur on qPMCd= OmCmCurCd "
   ws_Join = ""
    wSqlStrg = "Select OmCoCd, OmTc, OmYy, OmChr, OmNo, OmDt, OmFrDmLoc, OmToDmLoc, OmFrDmDc, OmToDmDc, OmCmCd, IsNull(CmName, '') as CmName, " + _
               "OmCmCtg, OmCnvFct, OmCmCurCd, OmMulDiv, OmCmPayTerms,  " + _
               "B.OdSr, OmFrDmLoc, OmFrDmDc, RmCtg, RmSCtg, RmPurityWt/ " + _
               "(case when RmCtg= 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
               "      when RmCtg= 'P' then " + CStr(wd_BasePltRmPurity) + " " + _
               "      when RmCtg= 'S' then " + CStr(wd_BaseSlvRmPurity) + " " + _
               "      when RmCtg= 'L' then " + CStr(wd_BasePdRmPurity) + " else 1 end) as RmPurityWt, " + _
               "B.OdJLRmCd, B.OdJLLotNo, B.OdJLLn1, B.OdJLLn2, B.OdJLLn3, " + wRmSzDesc + " as qRmSzDesc, 'W' as TdRtByQw,  " + _
               "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
               "B.OdJLQty, B.OdJLWt, B.OdJLVchRt, " + wPurStr + ", RmDesc, IsNull(M.OdTc+'/'+M.OdYy+'/'+M.OdChr+'/'+LTrim(Str(M.OdNo,5))+'/'+LTrim(str(M.OdSr,3)),'') as qBaseVch " + _
               "From OrdMst " + _
               "Join OrdDsg B on B.OdOmIdNo=OmIdNo and B.OdCoCd= OmCoCd and B.OdTc= OmTc and B.OdYy= OmYy " + _
               " and B.OdChr= OmChr and B.OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and B.OdPrtKey=OmPrtKey ", "") + _
               "Join RmMst on RmCd=B.OdJLRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=B.OdPrtKey ", "") + ws_Join + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=B.OdJLLn1 " + _
               "Left Outer Join CustMst On CmCtg=OmCmCtg and CmCd=OmCmCd " + _
               "Left Outer Join OrdDsg M on B.OdBaseIDKey=M.OdIdNo and (M.OdTc='LIN' Or M.OdCoCd = B.OdCoCd) " + IIF(gs_Partition = ctCurrPrtn, " and M.OdPrtKey=B.OdPrtKey ", "") + _
               "where " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey='" + ctCurrPrtn + "'", "") + _
               wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " order by OmCoCd, OmTc, OmYy, OmChr, OmNo"
 
 
 
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  mRep.TxtHead.SetText ws_FrmCaption '****** Sachin 2.14.0
  
  ' Zubin 212
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UoCoCdFr") + "' and HLotNoYN= 'Y' ") Then
      Rep.h1LotNo.Suppress = True:
      Rep.FldDET1LotNo.Suppress = True:
  End If
  
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc
'  GRP_REP.SetFormula mRep, "wTdRmSz", "If {rdo.qRmSzDesc} <> '' Then {rdo.qRmSzDesc} " + _
                                       "Else (If {rdo.qPtrYn}= 'Y' Then ToText({rdo.OdJLLn1}, 4) Else ToText({rdo.OdJLLn1}, 3))"
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc
  
  If adc("UoYN1") = "N" Then Rep.DETDesc.Suppress = True
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.OdJLRmCd}", 8)
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
  ' Zubin 212
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 1     'Tag = ""
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  End Select
End With
End Sub
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
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
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = True
  Set adc.FirNKeyCtl = IIF(ws_MultiCoMod = True, adc("UoCoCdFr"), adc("UoTTcFr"))
      adc("UOCOCDFR").Enabled = ws_MultiCoMod
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  'gs_LocTyp = "'P'"
  'gb_RmCtgFor = False
  'gb_RmZ = False
  'gs_DmTcTyp = "DM"
  'gs_OmTcTyp = "SO"
  
  gs_CmCtg = ""
  gb_CmCtgFor = False
  
  
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  
    LblBOENo.Visible = False
    adc("UoBoeFr").Visible = False: adc("UoBoeTo").Visible = False: adc("UoBoeSel").Visible = False
      
    LblLsLoc.Visible = False: LblLsDt.Visible = False
    adc("UoIssLocFr").Visible = False: adc("UoIssLocTo").Visible = False: adc("UoIssLocSel").Visible = False
    adc("UoLsDtFr").Visible = False: adc("UoLsDtTo").Visible = False
  
    adc("UoIssLocFr").Enabled = False: adc("UoIssLocTo").Enabled = False: adc("UoIssLocSel").Enabled = False
    adc("UoLsDtFr").Enabled = False: adc("UoLsDtTo").Enabled = False
      
'    adc("UoCmCdFr").CmpStr = "OmCmCd>= ": adc("UoCmCdTo").CmpStr = "TSuppCd<= ": adc("UoCmCdSel").CmpStr = "TSuppCd In "
'    adc("UoBOEFr").CmpStr = "TBOENo>= ": adc("UoBOETo").CmpStr = "TBOENo<= ": adc("UoBOESel").CmpStr = "TBOENo In "
  
  ws_FrmCaption = Me.Caption  '****** Sachin 2.14.0
  
  Call SetGroupSort
  CRV_REP.DisplayGroupTree = False
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
'  If UCase(adc.MenuCd) = UCase("RepVchPrnDT") Then gs_TxnTcTyp = "DT"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnBS") Then gs_TxnTcTyp = "BS"
  ' **** Zubin 211 **** '
'  If UCase(adc.MenuCd) = UCase("RepVchPrnBM") Then gs_TxnTcTyp = "BM"
  ' **** Zubin 211 **** '
'  If UCase(adc.MenuCd) = UCase("RepVchPrnMLT") Then gs_TxnTcTyp = "MLT"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnBV") Then gs_TxnTcTyp = "BV"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnREJ") Then gs_TxnTcTyp = "REJ"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnPR") Then gs_TxnTcTyp = "PR"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnREC") Then gs_TxnTcTyp = "REC"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnCNV") Then gs_TxnTcTyp = "CNV"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnMV") Then gs_TxnTcTyp = "MV"
  gs_TxnTcTyp = "JT"
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
'  If UCase(adc.MenuCd) = UCase("RepVchPrnDT") Then gs_TxnTcTyp = "DT"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnBS") Then gs_TxnTcTyp = "BS"
'  ' **** Zubin 211 **** '
'  If UCase(adc.MenuCd) = UCase("RepVchPrnBM") Then gs_TxnTcTyp = "BM"
'  ' **** Zubin 211 **** '
'  If UCase(adc.MenuCd) = UCase("RepVchPrnMLT") Then gs_TxnTcTyp = "MLT"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnBV") Then gs_TxnTcTyp = "BV"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnREJ") Then gs_TxnTcTyp = "REJ"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnPR") Then gs_TxnTcTyp = "PR"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnREC") Then gs_TxnTcTyp = "REC"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnCNV") Then gs_TxnTcTyp = "CNV"
'  If UCase(adc.MenuCd) = UCase("RepVchPrnMV") Then gs_TxnTcTyp = "MV"
  gs_TxnTcTyp = "JT"
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  gb_CoCdFor = True
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
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
''Rep1.wCoCd.SetText gs_CoNm
''Rep2.wCoCd.SetText gs_CoNm
''Rep3.wCoCd.SetText gs_CoNm
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
  Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
  Set mRep = Nothing
  Set Rep = Nothing
  '*** (09/08/05)
End Sub

