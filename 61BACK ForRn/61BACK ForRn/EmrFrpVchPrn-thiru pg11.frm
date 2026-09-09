VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrpVchPrn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher Printing"
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
      TabIndex        =   25
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
      TabIndex        =   26
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
      TabIndex        =   35
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   36
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
         TabIndex        =   37
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
      TabIndex        =   34
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   31
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
      TabIndex        =   27
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
      TabIndex        =   30
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   29
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpVchPrn.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpVchPrn.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   33
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   28
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
            TabIndex        =   32
            Top             =   360
            Width           =   15105
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1455
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   63
               Top             =   4710
               Width           =   2340
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Stock Rate "
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
               Left            =   4020
               TabIndex        =   62
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Stock Rate"
               Top             =   4710
               Width           =   2100
            End
            Begin VB.Frame FraOptAck 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1500
               TabIndex        =   56
               ToolTipText     =   "Show Open Or Closed Or All Bags"
               Top             =   4200
               Width           =   3465
               Begin VB.OptionButton OptAck 
                  Caption         =   "All"
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
                  Index           =   2
                  Left            =   2700
                  TabIndex        =   59
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptAck 
                  Caption         =   "No"
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
                  Left            =   1260
                  TabIndex        =   58
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   945
               End
               Begin VB.OptionButton OptAck 
                  Caption         =   "Yes"
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
                  TabIndex        =   57
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
            End
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
               TabIndex        =   24
               Tag             =   "ShowRmDesc"
               ToolTipText     =   "Check To Show Rm Description"
               Top             =   3180
               Width           =   1665
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4005
               TabIndex        =   10
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   1395
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
               Left            =   1515
               TabIndex        =   9
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   1395
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
               Index           =   24
               Left            =   4785
               TabIndex        =   5
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   825
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
               CmpStr          =   "TYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2295
               TabIndex        =   2
               ToolTipText     =   "Enter From Transaction Year"
               Top             =   825
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
               CmpStr          =   "TYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1515
               TabIndex        =   1
               ToolTipText     =   "Enter From Transaction Tc"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
               CmpStr          =   "TTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2715
               TabIndex        =   3
               ToolTipText     =   "Enter From Transaction Character"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
               CmpStr          =   "TChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   4005
               TabIndex        =   4
               ToolTipText     =   "Enter To Transaction Tc"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
               CmpStr          =   "TTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   5205
               TabIndex        =   6
               ToolTipText     =   "Enter To Transaction Character"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
               CmpStr          =   "TChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1515
               TabIndex        =   7
               ToolTipText     =   "Enter From Transaction Number"
               Top             =   1110
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOFR"
               IdName          =   "UOTNOFR"
               CmpStr          =   "TNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   4005
               TabIndex        =   8
               ToolTipText     =   "Enter To Transaction Number"
               Top             =   1110
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOTO"
               IdName          =   "UOTNOTO"
               CmpStr          =   "TNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   4005
               TabIndex        =   12
               ToolTipText     =   "Enter To Supplier"
               Top             =   1680
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   6510
               TabIndex        =   13
               ToolTipText     =   "Enter Supplier Selection"
               Top             =   1680
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   1515
               TabIndex        =   11
               ToolTipText     =   "Enter From Supplier"
               Top             =   1680
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   4005
               TabIndex        =   18
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
               TabIndex        =   19
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
               TabIndex        =   17
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
               TabIndex        =   15
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
               TabIndex        =   16
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
               TabIndex        =   14
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
               TabIndex        =   21
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
               TabIndex        =   20
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
               Index           =   35
               Left            =   1515
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "TCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   64
               Left            =   1515
               TabIndex        =   53
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
               TabIndex        =   22
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
               TabIndex        =   23
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
               Index           =   74
               Left            =   5640
               TabIndex        =   61
               ToolTipText     =   "Show Yes Or No Or All Acknowledge"
               Top             =   4200
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
               Height          =   240
               Index           =   89
               Left            =   6465
               TabIndex        =   64
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   4680
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   423
               Enabled         =   0   'False
               DataField       =   "UOYN3"
               IdName          =   "UOYN3"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "SEO Password"
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
               Index           =   51
               Left            =   0
               TabIndex        =   65
               Top             =   4710
               Width           =   1455
            End
            Begin VB.Label LblAck 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Acknowledge"
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
               Height          =   585
               Left            =   30
               TabIndex        =   60
               ToolTipText     =   "Location"
               Top             =   4080
               Visible         =   0   'False
               Width           =   1485
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
               TabIndex        =   55
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
               Left            =   30
               TabIndex        =   54
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
               TabIndex        =   52
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
               TabIndex        =   51
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
               TabIndex        =   50
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
               TabIndex        =   49
               Top             =   2250
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
               TabIndex        =   48
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
               TabIndex        =   47
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
               TabIndex        =   46
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
               TabIndex        =   45
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
               Index           =   32
               Left            =   4725
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
               Index           =   33
               Left            =   2235
               TabIndex        =   42
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
               TabIndex        =   41
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
               Left            =   6480
               TabIndex        =   40
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
               TabIndex        =   39
               Top             =   60
               Width           =   1275
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Fromm"
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
               TabIndex        =   38
               Top             =   60
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpVchPrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep1 As New EmrRepVchPrnOne
Dim Rep2 As New EmrRepVchPrnTwo
Dim Rep3 As New EmrRepVchPrnThree
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
  'MMAck-Ack option button selection changing as per value in UOYN2 textbox
  If adc("UoYN2") = "Y" Then
    OptAck(0).Value = True
  ElseIf adc("UoYN2") = "N" Then
    OptAck(1).Value = True
  ElseIf adc("UoYN2") = "" Then
    OptAck(2).Value = True
  End If
  'StkRt-Set Show StockRate option is No
  adc("UoYN3") = "N"
  
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  ' Zubin 212
  '*** SQL for DT, BS, MLT
  'Select TTc, TYy, TChr, TNo, TDt, TFrRmLoc, TToRmLoc, TFrBLoc, TdSr, TdBYy, TdBChr, TdBNo,
  'DmCd, DmCtg, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, TdBDc, TdBQty, TdPtQty, TdBGrWt,
  'TdFrBLoc, TdToBLoc, TdRjDc, TdRjQty, TdRjPtQty, TdRjCd, TdWrk, TdAck, TdPrdYN,
  'TdByLoc as qBagByLoc, TdByWrk as qBagByWrk, TdDesc as qBagDesc, TdSrNo, TdFrRmLoc,
  'TdFrRmDc, RmCtg, TdRmCd, TdLotNo, TdRmSz, TdRmStkRt, PValue3 as qPtrYN, TdRmQty, TdRmWt,
  'TdToRmLoc, TdByLoc As qRmByLoc, TdByWrk As qRmByWrk, TdDustWt, TdDesc As qRmDesc,
  'LocTyp as qBMLocTyp From Txn, Txnd, RmMst, Bag, DsgMst, Loc, Param Where 1 = 2
  
  '*** SQL for BV, REJ
  'Select TTc, TYy, TChr, TNo, TDt, TdSr, TdBYy, TdBChr, TdBNo, DmCd, DmCtg, BOdTc, BOdYy,
  'BOdChr, BOdNo, BOdSr, TdBDc, TdBQty, TdPtQty, TdBGrWt, TdFrBLoc, TdToBLoc, TdRjCd,
  'TdWrk, TdAck, TdByLoc as qBagByLoc, TdByWrk as qBagByWrk
  'From Txn, Txnd, Bag, DsgMst Where 1 = 2
  
  ' Zubin 212 (TdLotNo added [TdDustWt also added as it was missing])
  ' ***** Manali 3.8.0 - TdRtByQw added
  '*** SQL for PR, REC, CNV, MV
'  Select TTc, Tyy, TChr, Tno, TDt, TFrRmLoc, TToRmLoc, TLsLoc, TLsFrDt, TLsToDt, TGldinLs,
'  TGldRec, TDustWt, TSuppCd, TBillNo,  TBillDt, TBOENo, TBOEDt, TCnvRt, TNetAmt as qTNetAmt,
'  TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TAddAmt3 as qTAddAmt3, TAddAmt4 as qTAddAmt4,
'  TTotAmt as qTTotAmt, TdSr, TdFrRmLoc, TdFrRmDc, RmCtg, RmSCtg, RmPurityWt, TdRmCd , TdLotNo, TdRmSz,
'  TdRmStkRt, TdRtByQw, PValue3 as qPtrYN, TdRmQty, TdRmWt, TdToRmLoc, TdPurRt as qTdPurRt, TdPurAmt as qTdPurAmt, TdDustWt
'  From Txn, Txnd, RmMst, Param Where 1=2
  '*** Report Sql ***
  
  
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
  
  If UCase(adc.MenuCd) = UCase("RepVchPrnDT") Then gs_TxnTcTyp = "DT"
  If UCase(adc.MenuCd) = UCase("RepVchPrnBS") Then gs_TxnTcTyp = "BS"
' **** Zubin 211 **** '
  If UCase(adc.MenuCd) = UCase("RepVchPrnBM") Then gs_TxnTcTyp = "BM"
' **** Zubin 211 **** '
  If UCase(adc.MenuCd) = UCase("RepVchPrnMLT") Then gs_TxnTcTyp = "MLT"
  If UCase(adc.MenuCd) = UCase("RepVchPrnBV") Then gs_TxnTcTyp = "BV"
  If UCase(adc.MenuCd) = UCase("RepVchPrnREJ") Then gs_TxnTcTyp = "REJ"
  If UCase(adc.MenuCd) = UCase("RepVchPrnPR") Then gs_TxnTcTyp = "PR"
  If UCase(adc.MenuCd) = UCase("RepVchPrnREC") Then gs_TxnTcTyp = "REC"
  If UCase(adc.MenuCd) = UCase("RepVchPrnCNV") Then gs_TxnTcTyp = "CNV"
  If UCase(adc.MenuCd) = UCase("RepVchPrnMV") Then gs_TxnTcTyp = "MV"
 
  Set Rep1 = Nothing
  Set Rep2 = Nothing
  Set Rep3 = Nothing
  Select Case UCase(gs_TxnTcTyp)
' **** Zubin 211 ("BM" added) **** '
  Case "DT", "BS", "MLT", "BM"
' **** Zubin 211 **** '
    Set adc.RepSource = Rep1
    Set mRep = Rep1
  Case "BV", "REJ"
    Set adc.RepSource = Rep2
    Set mRep = Rep2
  Case "PR", "REC", "CNV", "MV"
    Set adc.RepSource = Rep3
    Set mRep = Rep3
  End Select
  
  Call DispCoNm
  Call SetMwName(mRep)   ' ***** Manali 3.6.1 - Microway footer
  'Call GRP_REP.Gen3LRep
  'grpflds = GRP_REP.GrpFldLst
    
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
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " Td.TdDt >= '" + Trim(adc("UODTFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " Td.TdDt <= '" + Trim(adc("UODTTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
             
  Select Case UCase(gs_TxnTcTyp)
' **** Zubin 211 ("BM" added) **** '
  Case "DT", "BS", "MLT", "BM"
' **** Zubin 211 **** '
    ' Zubin 212
    ' ****** Manali 3.03 - 09/08/07 - qRmSzDesc added and Join with RmIdSz Added
    ' ****** Manali 3.10.0 - 07/03/12 - TDesc Added
    wSqlStrg = "Select TTc, Tyy, TChr, TNo, TDt, TFrRmLoc, TToRmLoc, TFrBLoc, TDesc, Td.TdSr,Td.TdBYy, Td.TdBChr, Td.TdBNo, DmCd, DmCtg, BOdTc, BOdYy, BOdChr, BOdNo, " + _
               "BOdSr, Td.TdBDc, Td.TdBQty, Td.TdPtQty, Td.TdBGrWt,Td.TdFrBLoc, Td.TdToBLoc, Td.TdRjDc, Td.TdRjQty, Td.TdRjPtQty, " + _
               "Td.TdRjCd, Td.TdWrk, Td.TdAck, Td.TdPrdYN, Td.TdByLoc as qBagByLoc, Td.TdByWrk as qBagByWrk, Td.TdDesc as qBagDesc, IsNull(Rm.TdSrNo, 0) as TdSrNo, IsNull(Rm.TdFrRmLoc, '') as TdFrRmLoc, " + _
               "IsNull(Rm.TdFrRmDc, '') as TdFrRmDc, IsNull(RmCtg, '') as RmCtg, IsNull(Rm.TdRmCd, '') as TdRmCd, " + _
               "IsNull(Rm.TdLotNo, '') as TdLotNo, IsNull(Rm.TdRmSz, 0.0) as TdRmSz, IsNull(Rm.TdRmSz2, 0.0) as TdRmSz2, IsNull(Rm.TdRmSz3, 0.0) as TdRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(Rm.TdRmStkRt, 0.0) as TdRmStkRt, " + _
               "IsNull((Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RmCtg, '') and PSCd= IsNull(RmSCtg, '')), '') as qPtrYN, " + _
               "IsNull(Rm.TdRmQty, 0) as TdRmQty, IsNull(Rm.TdRmWt, 0.0) as TdRmWt,IsNull(Rm.TdToRmLoc, '') as TdToRmLoc, IsNull(Rm.TdByLoc, '') as qRmByLoc, " + _
               "IsNull(Rm.TdByWrk, '') as qRmByWrk, IsNull(Rm.TdDustWt, 0.0) as TdDustWt, " + _
               "IsNull(Rm.TdDesc, '') as qRmDesc,(Select LocTyp from Loc where LocCoCd ='" + adc("UoCoCdFr") + "' and " + _
               "LocCd= IsNull(Rm.TdToRmLoc, '')) as qBMLocTyp, " + _
               "InAmtDesc1, InAmtDesc2, InAmtDesc3, InAmtDesc4, InAmtDesc5, InAmtDesc6, InAmtDesc7 ,InAmtDesc8, InAmtDesc9, InAmtDesc10, " + _
               "InAmtDesc11, InAmtDesc12, InAmt1Per, InAmt2Per, InAmt3Per, InAmt4Per, InAmt5Per, InAmt6Per, InAmt8Per, InAmt9Per, InAmt10Per, InAmt11Per, " + _
               "InAmt1, InAmt2, InAmt3, InAmt4, InAmt5, InAmt6, InAmt7, InAmt8, InAmt9, InAmt10, InAmt11, InAmt12, InGstCd, InBaseAmt  " + _
               "From Txn Join Txnd Td on Td.TdTIdNo=TdTIdNo and Td.TdCoCd= TCoCd and Td.TdTc= TTc and Td.TdYy= TYy " + _
               "and Td.TdChr= TChr and Td.TdNo= TNo and Td.TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " and Td.TdPrtKey=TPrtKey ", "") + _
               "Join Bag on BIdNo=Td.TdBIdNo and BCoCd= Td.TdCoCd and BYy= Td.TdBYy and BChr= Td.TdBChr and BNo= Td.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=Td.TdPrtKey ", "") + _
               "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= 'DM' and DmCd= BOdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=DmPrtKey ", "") + _
               "Join InvHd On InCoCd= TCoCd  and InTc=TTc and InYY=TYy  and InChr=TChr  and InNo=TNo " + _
               "Left Outer Join Txnd Rm on Rm.TdCoCd= Td.TdCoCd and Rm.TdTc= Td.TdTc and " + _
               "Rm.TdYy= Td.TdYy and Rm.TdChr= Td.TdChr and Rm.TdNo= Td.TdNo and " + _
               "Rm.TdSr= Td.TdSr and Rm.TdSrNo<> 0 " + IIF(gs_Partition = ctCurrPrtn, " and Rm.TdPrtKey=Td.TdPrtKey ", "") + " Left Outer Join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=Rm.TdPrtKey and ", "") + " RmCd= Rm.TdRmCd " + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=Rm.TdRmSz " + _
               "where " + IIF(gs_Partition = ctCurrPrtn, " Td.TdPrtKey='" + ctCurrPrtn + "' and ", "") + " (Select PValue from Param where PTyp= 'TC' and PMCd= TTc)= '" + gs_TxnTcTyp + "' " + _
               wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "")
  Case "BV", "REJ"
    '****** Sachin 3.02 - Id fields in Joins
    ' ****** Manali 3.10.0 - 07/03/12 - TDesc Added
    wSqlStrg = "Select TTc, TYy, TChr, TNo, TDt, TdSr, TdBYy, TdBChr, TdBNo, TDesc, DmCd, DmCtg, " + _
               "BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, TdBDc, TdBQty, " + _
               "TdPtQty, TdBGrWt, TdFrBLoc, TdToBLoc, TdRjCd, TdWrk, TdAck, TdByLoc as qBagByLoc,TdByWrk as qBagByWrk, " + _
               "InAmtDesc1, InAmtDesc2, InAmtDesc3, InAmtDesc4, InAmtDesc5, InAmtDesc6, InAmtDesc7 ,InAmtDesc8, InAmtDesc9, InAmtDesc10, " + _
               "InAmtDesc11, InAmtDesc12, InAmt1Per, InAmt2Per, InAmt3Per, InAmt4Per, InAmt5Per, InAmt6Per, InAmt8Per, InAmt9Per, InAmt10Per, InAmt11Per, " + _
               "InAmt1, InAmt2, InAmt3, InAmt4, InAmt5, InAmt6, InAmt7, InAmt8, InAmt9, InAmt10, InAmt11, InAmt12, InGstCd, InBaseAmt " + _
               "From Txn Join Txnd Td on Td.TdTIdNo=TIdNo and Td.TdCoCd= TCoCd and Td.TdTc= TTc and Td.TdYy= TYy " + _
               "and Td.TdChr= TChr and Td.TdNo= TNo and Td.TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " and Td.TdPrtKey=TPrtKey ", "") + _
               "Join Bag on BIdNo=Td.TdBIdNo and BCoCd= Td.TdCoCd and BYy= Td.TdBYy and BChr= Td.TdBChr and " + _
               "BNo= Td.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=Td.TdPrtKey ", "") + " Join DsgMst on  DmIdNo=BDmIdNo and DmTcTyp= 'DM' and DmCd= BOdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
               "Join InvHd On InCoCd= TCoCd  and InTc=TTc and InYY=TYy  and InChr=TChr  and InNo=TNo " + _
               "where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' and ", "") + " (Select PValue from Param where PTyp= 'TC' and PMCd= TTc)= '" + gs_TxnTcTyp + "' " + _
               wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "")
  Case "PR", "REC", "CNV", "MV"
    '************Geeta****************Emr207*******************
    ' ****** Sachin 2.13.0 Correction ******
    Dim ws_BaseCurCd As String, ws_Join As String, wAck As String
        ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                      "HCoCd= '" + adc("UoCoCdFr") + "'")
  
        If adc("UoCurCdFr") <> "" Then
          wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "TCurrCd='" + adc("UoCurCdFr") + "'"
          wPurStr = "TdPurRt as qTdPurRt, TdPurAmt qTdPurAmt"
          wTTotAmtStr = "TNetAmt as qTNetAmt, TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TAddAmt3 as qTAddAmt3, " + _
                         " TAddAmt4 as qTAddAmt4, TTotAmt as qTTotAmt,"
        Else
          wPurStr = " (case when TMulDiv= 'M' then TdPurRt/TCnvRt else TdPurRt*TCnvRt end) as qTdPurRt, (case when TMulDiv= 'M' then TdPurAmt/TCnvRt else TdPurAmt*TCnvRt end) qTdPurAmt "
          wTTotAmtStr = "(case when TMulDiv= 'M' then TNetAmt/TCnvRt Else TNetAmt*TCnvRt End) as qTNetAmt, " + _
                        " (case when TMulDiv= 'M' then TAddAmt1/TCnvRt Else TAddAmt1*TCnvRt End) as qTAddAmt1, " + _
                        " (case when TMulDiv= 'M' then TAddAmt2/TCnvRt Else TAddAmt2*TCnvRt End) as qTAddAmt2, " + _
                        " (case when TMulDiv= 'M' then TAddAmt3/TCnvRt Else TAddAmt3*TCnvRt End) as qTAddAmt3, " + _
                        " (case when TMulDiv= 'M' then TAddAmt4/TCnvRt Else TAddAmt4*TCnvRt End) as qTAddAmt4, " + _
                        " (case when TMulDiv= 'M' then TTotAmt/TCnvRt Else TTotAmt*TCnvRt End) as qTTotAmt, "
        End If
        
        ws_Join = " Join (Select PMCd as qPMCd, (case when " + IIF(adc("UoCurCdFr") = "", "1", "0") + " = 1 then 'Y' " + _
                  "else 'N' end) as qBaseCur from Param where PTyp= 'CURNCY') Cur on qPMCd= TCurrCd "
    
        ' ****** Sachin 2.13.0 Correction ******
        'MMAck-Voucher Ack Y/N added, blank means all voucher
        wAck = IIF(adc("UoYN2") = "Y", " and TdAck= 'Y' ", IIF(adc("UoYN2") = "N", " and TdAck= 'N' ", ""))
    
    
        'If UCase(gs_TxnTcTyp) = "PR" Then
        'If adc("UoYN") <> "" Then
        '   wCnd = wCnd + IIF(wCnd <> "", " and ", "") '+ "TUSDYn='" + adc("UoYN") + "'"
        '   wPurStr = "TdPurRt as qTdPurRt, TdPurAmt qTdPurAmt "
        '   wTTotAmtStr = "TNetAmt as qTNetAmt, TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TAddAmt3 as qTAddAmt3, " + _
                         " TAddAmt4 as qTAddAmt4, TTotAmt as qTTotAmt,"
        ' Else
        '    wPurStr = "TdPurRt as qTdPurRt, TdPurAmt qTdPurAmt "
       'wTTotAmtStr = "TNetAmt as qTNetAmt, TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TAddAmt3 as qTAddAmt3, " + _
                         " TAddAmt4 as qTAddAmt4, TTotAmt as qTTotAmt,"
         
      ' wPurStr = " (case when TUSDYn='N' then TdPurRt/TCnvRt else TdPurRt end ) as qTdPurRt, " + _
                      " (case when TUSDYn='N' then TdPurAmt/TCnvRt else TdPurAmt end ) as qTdPurAmt "
      ' wTTotAmtStr = "(case when TUSDYn='N' then TNetAmt/TCnvRt else TNetAmt end ) as qTNetAmt," + _
                          "(case when TUSDYn='N' then TAddAmt1/TCnvRt else TAddAmt1 end ) as qTAddAmt1," + _
                          "(case when TUSDYn='N' then TAddAmt2/TCnvRt else TAddAmt2 end ) as qTAddAmt2, " + _
                          "(case when TUSDYn='N' then TAddAmt3/TCnvRt else TAddAmt3 end ) as qTAddAmt3, " + _
                          "(case when TUSDYn='N' then TAddAmt4/TCnvRt else TAddAmt4 end ) as qTAddAmt4," + _
                          "(case when TUSDYn='N' then TTotAmt/TCnvRt else TTotAmt end ) as qTTotAmt,"
      'End If
      'Else
      '   wPurStr = "TdPurRt as qTdPurRt, TdPurAmt qTdPurAmt "
      '   wTTotAmtStr = "TNetAmt as qTNetAmt, TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TAddAmt3 as qTAddAmt3, " + _
                       " TAddAmt4 as qTAddAmt4, TTotAmt as qTTotAmt,"
      'End If
      ' ****** Sachin 2.13.0 Correction ******
    '************Geeta****************Emr207*******************
    '*********************Replaced  TdPurRt, TdPurAmt with  wPurStr*********
    '*********************Replaced  TNetAmt, TAddAmt1, TAddAmt2, TAddAmt3, TAddAmt4, TTotAmt, with  wTTotAmtStr*********
    
    '****** Sachin 3.02 - Id fields in Joins
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    ' ***** Manali 3.03 - 09/07/08 - qRmSzDesc added and join with RmIdSz added
    ' ***** Manali 3.8.0 - TdRtByQw added
    ' ****** Manali 3.10.0 - 07/03/12 - TDesc Added
    ' ****** Manali 3.10.0 - 12/03/12 - CmName, TCurrCd, TSuppPayTerms, TdDesc, RmDesc Added
    wSqlStrg = "Select TTc, TYy, TChr, TNo, TDt, TFrRmLoc, TToRmLoc, TDesc, TLsLoc, TLsFrDt, " + _
               "TLsToDt, TGldinLs, TGldRec, TDustWt, TSuppCd, IsNull(CmName, '') as CmName, TBillNo,  TBillDt, TBOENo, " + _
               "TBOEDt, TCnvRt, TCurrCd, TSuppPayTerms, " + wTTotAmtStr + _
               "TdSr, TdFrRmLoc, TdFrRmDc, RmCtg, RmSCtg, RmPurityWt/ " + _
               "(case when RmCtg= 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
               "      when RmCtg= 'P' then " + CStr(wd_BasePltRmPurity) + " " + _
               "      when RmCtg= 'S' then " + CStr(wd_BaseSlvRmPurity) + " " + _
               "      when RmCtg= 'L' then " + CStr(wd_BasePdRmPurity) + " else 1 end) as RmPurityWt, " + _
               "TdRmCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, " + wRmSzDesc + " as qRmSzDesc, TdRmStkRt, TdRtByQw,  " + _
               "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
               "TdRmQty, TdRmWt, TdToRmLoc," + wPurStr + ", TdDustWt, TdDesc, RmDesc,TdAck, " + _
               "InAmtDesc1, InAmtDesc2, InAmtDesc3, InAmtDesc4, InAmtDesc5, InAmtDesc6, InAmtDesc7 ,InAmtDesc8, InAmtDesc9, InAmtDesc10, " + _
               "InAmtDesc11, InAmtDesc12, InAmt1Per, InAmt2Per, InAmt3Per, InAmt4Per, InAmt5Per, InAmt6Per, InAmt8Per, InAmt9Per, InAmt10Per, InAmt11Per, " + _
               "InAmt1, InAmt2, InAmt3, InAmt4, InAmt5, InAmt6, InAmt7, InAmt8, InAmt9, InAmt10, InAmt11, InAmt12, InGstCd, InBaseAmt " + _
               "From Txn " + _
               "Join Txnd Td on Td.TdTIdNo=TIdNo and Td.TdCoCd= TCoCd and Td.TdTc= TTc and Td.TdYy= TYy " + _
               "and Td.TdChr= TChr and Td.TdNo= TNo and Td.TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " and Td.TdPrtKey=TPrtKey ", "") + _
               "Join InvHd On InCoCd= TCoCd  and InTc=TTc and InYY=TYy  and InChr=TChr  and InNo=TNo " + _
               "Join RmMst on RmCd= Td.TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=Td.TdPrtKey ", "") + ws_Join + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
               "Left Outer Join CustMst On CmCtg=(Case When TTc='SA' or TTc='MV' Then 'C' Else 'S' End) and CmCd=TSuppCd " + _
               "where " + IIF(gs_Partition = ctCurrPrtn, " Td.TdPrtKey='" + ctCurrPrtn + "' and ", "") + " (Select PValue from Param where PTyp= 'TC' and PMCd= TTc)= '" + gs_TxnTcTyp + "' " + _
               wCnd + wAck + IIF(wAddlCnd <> "", " And " + wAddlCnd, "")
  End Select
 
 
 
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  '*********Geeta************Emr207***
  If UCase(gs_TxnTcTyp) = "PR" Then
    '****** Sachin 2.13.0 Correction ******
    '     If ADC("UoYN") = "" Or ADC("UoYN") = "Y" Then
    '        mRep.TxtPRCurCd.SetText "In US$"
    '     ElseIf ADC("UoYN") = "N" Then
    '        mRep.TxtPRCurCd.SetText "In RS"
    '     End If
    '****** Sachin 2.13.0 Correction ******
     If adc("UoCurCdFr") = "" Then
      mRep.TxtPRCurCd.SetText "Purchase Currency (" + ws_BaseCurCd + ")"
    Else
      mRep.TxtPRCurCd.SetText "Purchase Currency (" + adc("UoCurCdFr") + ")"
    End If
     
  ElseIf UCase(gs_TxnTcTyp) = "REC" Or UCase(gs_TxnTcTyp) = "CNV" Or UCase(gs_TxnTcTyp) = "MV" Then
   mRep.TxtPRCurCd.SetText ""
  End If
  '*********Geeta************Emr207***
  
  mRep.TxtHead.SetText ws_FrmCaption '****** Sachin 2.14.0
  wi_Opt = 0
  If gs_TxnTcTyp = "DT" Then wi_Opt = 1
  If gs_TxnTcTyp = "BS" Then wi_Opt = 2
' **** Zubin 211 **** '
  If gs_TxnTcTyp = "BM" Then wi_Opt = 2
' **** Zubin 211 **** '
  If gs_TxnTcTyp = "MLT" Then wi_Opt = 3
  
  If gs_TxnTcTyp = "BV" Then wi_Opt = 1
  If gs_TxnTcTyp = "REJ" Then wi_Opt = 2

  If gs_TxnTcTyp = "PR" Then wi_Opt = 1
  If gs_TxnTcTyp = "REC" Then wi_Opt = 2
  If gs_TxnTcTyp = "CNV" Then wi_Opt = 3
  'SRI
  If gs_TxnTcTyp = "MV" Then wi_Opt = 1
  'If gs_TxnTcTyp = "MV" Then wi_Opt = 4
  
  Select Case wi_Opt
  Case Is = 1
    mRep.G1H2.Suppress = True
    mRep.G1H3.Suppress = True
    mRep.G1H4.Suppress = True
    mRep.G2H2.Suppress = True
    mRep.G2H3.Suppress = True
    mRep.G2H4.Suppress = True
    mRep.DET2.Suppress = True
    mRep.DET3.Suppress = True
    mRep.DET4.Suppress = True
    mRep.G1F2.Suppress = True
    mRep.G1F3.Suppress = True
    mRep.G1F4.Suppress = True
  Case Is = 2
    mRep.G1H1.Suppress = True
    mRep.G1H3.Suppress = True
    mRep.G1H4.Suppress = True
    mRep.G2H1.Suppress = True
    mRep.G2H11.Suppress = True
    mRep.G2H3.Suppress = True
    mRep.G2H4.Suppress = True
    mRep.DET1.Suppress = True
    mRep.DET3.Suppress = True
    mRep.DET4.Suppress = True
    mRep.G1F1.Suppress = True
    mRep.G1F3.Suppress = True
    mRep.G1F4.Suppress = True
  Case Is = 3
    mRep.G1H1.Suppress = True
    mRep.G1H2.Suppress = True
    mRep.G1H4.Suppress = True
    mRep.G2H1.Suppress = True
    mRep.G2H11.Suppress = True
    mRep.G2H2.Suppress = True
    mRep.G2H4.Suppress = True
    mRep.DET1.Suppress = True
    mRep.DET2.Suppress = True
    mRep.DET4.Suppress = True
    mRep.G1F1.Suppress = True
    mRep.G1F2.Suppress = True
    mRep.G1F4.Suppress = True
  Case Is = 4
    mRep.G1H1.Suppress = True
    mRep.G1H2.Suppress = True
    mRep.G1H3.Suppress = True
    mRep.G2H1.Suppress = True
    mRep.G2H11.Suppress = True
    mRep.G2H2.Suppress = True
    mRep.G2H3.Suppress = True
    mRep.DET1.Suppress = True
    mRep.DET2.Suppress = True
    mRep.DET3.Suppress = True
    mRep.G1F1.Suppress = True
    mRep.G1F2.Suppress = True
    mRep.G1F3.Suppress = True
  End Select
  ' Zubin 212
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UoCoCdFr") + "' and HLotNoYN= 'Y' ") Then
    Select Case UCase(gs_TxnTcTyp)
    Case "DT", "BS", "MLT", "BM"
      Rep1.h1LotNo.Suppress = True: Rep1.h2LotNo.Suppress = True: Rep1.h3LotNo.Suppress = True
      Rep1.FldDET1LotNo.Suppress = True: Rep1.FldDET2LotNo.Suppress = True: Rep1.FldDET3LotNo.Suppress = True
    Case "PR", "REC", "CNV", "MV"
      Rep3.h1LotNo.Suppress = True: Rep3.h2LotNo.Suppress = True
      Rep3.h3LotNo.Suppress = True: Rep3.h4LotNo.Suppress = True
      Rep3.FldDET1LotNo.Suppress = True: Rep3.FldDET2LotNo.Suppress = True
      Rep3.FldDET3LotNo.Suppress = True: Rep3.FldDET4LotNo.Suppress = True
    End Select
  End If
  
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc
  
'  GRP_REP.SetFormula mRep, "wTdRmSz", "If {rdo.qRmSzDesc} <> '' Then {rdo.qRmSzDesc} " + _
                                       "Else (If {rdo.qPtrYn}= 'Y' Then ToText({rdo.TdRmSz}, 4) Else ToText({rdo.TdRmSz}, 3))"
'  GRP_REP.SetFormula mRep, "wRmSz", "If {rdo.qRmSzDesc} <> '' Then {rdo.qRmSzDesc} "
    GRP_REP.SetFormula mRep, "wRmSz", "If {rdo.qRmSzDesc} <> '' Then {rdo.qRmSzDesc} " + _
                                      "Else CustRmSzDisp ({rdo.TdRmSz},{rdo.TdRmSz2},{rdo.TdRmSz3}) "

                              
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc
  If adc("UoYN1") = "N" Then Rep3.DETDesc.Suppress = True
  
  'StkRt-Getting stock rate display Y\N from head table
  Dim wStkNoDisp As String
  wStkNoDisp = moCn.GetFldVal("Select hStkRtNoDsp from Head where hcocd='" + gs_CoCd + "'")

  'StkRt-when Show Stockrate=yes then stock rate is visible in report
  'if it is NO stock rate is invisible
  If wStkNoDisp = "Y" Then
  If adc("UoYN3") = "Y" Then
    Rep1.hTdRmStkRt.SetText "Stock Rt":   Rep1.hTdRmStkRt2.SetText "Stock Rt":  Rep1.hTdRmStkRt3.SetText "Stock Rt"
    Rep1.FldRmStkRt.Suppress = False:     Rep1.FldRmStkRt2.Suppress = False:    Rep1.FldRmStkRt3.Suppress = False
    Rep3.hTdRmStkRt.SetText "Stock Rt":   Rep3.hTdRmStkRt2.SetText "Stock Rt"
    Rep3.hTdRmStkRt3.SetText "Stock Rt":  Rep3.hTdRmStkRt4.SetText "Stock Rt"
    Rep3.FldRmStkRt.Suppress = False:     Rep3.FldRmStkRt2.Suppress = False
    Rep3.FldRmStkRt3.Suppress = False:    Rep3.FldRmStkRt4.Suppress = False
  Else
    Rep1.hTdRmStkRt.SetText "":       Rep1.hTdRmStkRt2.SetText "":      Rep1.hTdRmStkRt3.SetText ""
    Rep1.FldRmStkRt.Suppress = True:  Rep1.FldRmStkRt2.Suppress = True: Rep1.FldRmStkRt3.Suppress = True
    Rep3.hTdRmStkRt.SetText "":       Rep3.hTdRmStkRt2.SetText ""
    Rep3.hTdRmStkRt3.SetText "":      Rep3.hTdRmStkRt4.SetText ""
    Rep3.FldRmStkRt.Suppress = True:  Rep3.FldRmStkRt2.Suppress = True
    Rep3.FldRmStkRt3.Suppress = True: Rep3.FldRmStkRt4.Suppress = True
  End If
  End If

  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
  ' Zubin 212
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 1     'Tag = ""
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  'StkRt-To select this option(Show Stock rate) either Seo user can give their password or
  ' can enter SEOPWD password otherwise this option can't not select
  Case Is = 2
    If .Value = Checked Then
       If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
            "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
            "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
            "Isnull ((Select PValue1 from Param where " + _
            "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ), ''))") Then _
            .Value = Unchecked
    End If
    If .Value = Checked Then adc("UoYN3") = "Y" Else adc("UoYN3") = "N"
 
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
  gs_CmCtg = "S"
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  
  
  '************Geeta****************Emr207*******************
   ' LblUSDYN.Visible = False: ADC("UoYN").Visible = False
  '************Geeta****************Emr207*******************
  Select Case UCase(adc.MenuCd)
  Case Is = UCase("RepVchPrnDT"), UCase("RepVchPrnBS"), UCase("RepVchPrnMLT"), UCase("RepVchPrnBM")
    '*** General Properties ***
    If UCase(adc.MenuCd) = UCase("RepVchPrnDT") Then gs_TxnTcTyp = "DT"
    If UCase(adc.MenuCd) = UCase("RepVchPrnBS") Then gs_TxnTcTyp = "BS"
    ' **** Zubin 211 **** '
    If UCase(adc.MenuCd) = UCase("RepVchPrnBM") Then gs_TxnTcTyp = "BM"
    ' **** Zubin 211 **** '
    If UCase(adc.MenuCd) = UCase("RepVchPrnMLT") Then gs_TxnTcTyp = "MLT"
    
    LblSuppCd.Visible = False: LblBOENo.Visible = False
    LblLsLoc.Visible = False: LblLsDt.Visible = False
    adc("UoCmCdFr").Visible = False: adc("UoCmCdTo").Visible = False: adc("UoCmCdSel").Visible = False
    adc("UoBOEFr").Visible = False: adc("UoBOETo").Visible = False: adc("UoBOESel").Visible = False
    adc("UoIssLocFr").Visible = False: adc("UoIssLocTo").Visible = False: adc("UoIssLocSel").Visible = False
    adc("UoLsDtFr").Visible = False: adc("UoLsDtTo").Visible = False
  
    adc("UoCmCdFr").Enabled = False: adc("UoCmCdTo").Enabled = False: adc("UoCmCdSel").Enabled = False
    adc("UoBOEFr").Enabled = False: adc("UoBOETo").Enabled = False: adc("UoBOESel").Enabled = False
    adc("UoIssLocFr").Enabled = False: adc("UoIssLocTo").Enabled = False: adc("UoIssLocSel").Enabled = False
    adc("UoLsDtFr").Enabled = False: adc("UoLsDtTo").Enabled = False
    LblCurrCd.Visible = False: adc("UOCURCDFR").Visible = False
    '*** General Properties ***
    ChkBoxArr(1).Visible = False: ChkBoxArr(1).Enabled = False      ' ***** Manali 3.10.0 - TdDesc in Loose Rm Vch
    FraOptAck.Visible = False: FraOptAck.Enabled = False  'MMAck-Disable visible property of Ack option buttons
  Case Is = UCase("RepVchPrnBV"), UCase("RepVchPrnREJ")
    '*** General Properties ***
    LblSuppCd.Visible = False: LblBOENo.Visible = False
    LblLsLoc.Visible = False: LblLsDt.Visible = False
    adc("UoCmCdFr").Visible = False: adc("UoCmCdTo").Visible = False: adc("UoCmCdSel").Visible = False
    adc("UoBOEFr").Visible = False: adc("UoBOETo").Visible = False: adc("UoBOESel").Visible = False
    adc("UoIssLocFr").Visible = False: adc("UoIssLocTo").Visible = False: adc("UoIssLocSel").Visible = False
    adc("UoLsDtFr").Visible = False: adc("UoLsDtTo").Visible = False
    
    adc("UoCmCdFr").Enabled = False: adc("UoCmCdTo").Enabled = False: adc("UoCmCdSel").Enabled = False
    adc("UoBOEFr").Enabled = False: adc("UoBOETo").Enabled = False: adc("UoBOESel").Enabled = False
    adc("UoIssLocFr").Enabled = False: adc("UoIssLocTo").Enabled = False: adc("UoIssLocSel").Enabled = False
    adc("UoLsDtFr").Enabled = False: adc("UoLsDtTo").Enabled = False
    LblCurrCd.Visible = False: adc("UOCURCDFR").Visible = False
    '*** General Properties ***
    If UCase(adc.MenuCd) = UCase("RepVchPrnBV") Then gs_TxnTcTyp = "BV"
    If UCase(adc.MenuCd) = UCase("RepVchPrnREJ") Then gs_TxnTcTyp = "REJ"
    
    ChkBoxArr(1).Visible = False: ChkBoxArr(1).Enabled = False      ' ***** Manali 3.10.0 - TdDesc in Loose Rm Vch
    FraOptAck.Visible = False: FraOptAck.Enabled = False  'MMAck-Disable visible property of Ack option buttons
  Case Is = UCase("RepVchPrnPR"), UCase("RepVchPrnREC"), UCase("RepVchPrnCNV"), UCase("RepVchPrnMV")
    '*** General Properties ***
    LblSuppCd.Visible = False: LblBOENo.Visible = False
    LblLsLoc.Visible = False: LblLsDt.Visible = False
    adc("UoCmCdFr").Visible = False: adc("UoCmCdTo").Visible = False: adc("UoCmCdSel").Visible = False
    adc("UoBOEFr").Visible = False: adc("UoBOETo").Visible = False: adc("UoBOESel").Visible = False
    adc("UoIssLocFr").Visible = False: adc("UoIssLocTo").Visible = False: adc("UoIssLocSel").Visible = False
    adc("UoLsDtFr").Visible = False: adc("UoLsDtTo").Visible = False
  
    adc("UoCmCdFr").Enabled = False: adc("UoCmCdTo").Enabled = False: adc("UoCmCdSel").Enabled = False
    adc("UoBOEFr").Enabled = False: adc("UoBOETo").Enabled = False: adc("UoBOESel").Enabled = False
    adc("UoIssLocFr").Enabled = False: adc("UoIssLocTo").Enabled = False: adc("UoIssLocSel").Enabled = False
    adc("UoLsDtFr").Enabled = False: adc("UoLsDtTo").Enabled = False
    LblCurrCd.Visible = False: adc("UOCURCDFR").Visible = False
    '*** General Properties ***
    If UCase(adc.MenuCd) <> UCase("RepTxnREC") Then _
    FraOptAck.Visible = True: FraOptAck.Enabled = True: LblAck.Visible = True  'MMAck-Enable visible property of Ack label & option buttons
 
    '*** Only For Purchase ***
    If UCase(adc.MenuCd) = UCase("RepVchPrnPR") Then
      gs_TxnTcTyp = "PR"
      adc("UoCmCdFr").CmpStr = "TSuppCd>= ": adc("UoCmCdTo").CmpStr = "TSuppCd<= ": adc("UoCmCdSel").CmpStr = "TSuppCd In "
      adc("UoBOEFr").CmpStr = "TBOENo>= ": adc("UoBOETo").CmpStr = "TBOENo<= ": adc("UoBOESel").CmpStr = "TBOENo In "
      LblSuppCd.Visible = True: LblBOENo.Visible = True
      adc("UoCmCdFr").Visible = True: adc("UoCmCdTo").Visible = True: adc("UoCmCdSel").Visible = True
      adc("UoBOEFr").Visible = True: adc("UoBOETo").Visible = True: adc("UoBOESel").Visible = True
      adc("UoCmCdFr").Enabled = True: adc("UoCmCdTo").Enabled = True: adc("UoCmCdSel").Enabled = True
      adc("UoBOEFr").Enabled = True: adc("UoBOETo").Enabled = True: adc("UoBOESel").Enabled = True
      LblCurrCd.Visible = True: adc("UOCURCDFR").Visible = True
      '************Geeta****************Emr207*******************
      'LblUSDYN.Visible = True: ADC("UoYN").Visible = True
      '************Geeta****************Emr207*******************
    End If
    '*** Only For Purchase ***
    '*** Only For Recovery ***
    If UCase(adc.MenuCd) = UCase("RepVchPrnREC") Then
      gs_TxnTcTyp = "REC"
      adc("UoIssLocFr").CmpStr = "TLsLoc>= ": adc("UoIssLocTo").CmpStr = "TLsLoc<= ": adc("UoIssLocSel").CmpStr = "TLsLoc In "
      adc("UoLsDtFr").CmpStr = "TLsFrDt>= ": adc("UoLsDtTo").CmpStr = "TLsToDt<= "
      LblLsLoc.Visible = True: LblLsDt.Visible = True
      adc("UoIssLocFr").Visible = True: adc("UoIssLocTo").Visible = True: adc("UoIssLocSel").Visible = True
      adc("UoLsDtFr").Visible = True: adc("UoLsDtTo").Visible = True
      adc("UoIssLocFr").Enabled = True: adc("UoIssLocTo").Enabled = True: adc("UoIssLocSel").Enabled = True
      adc("UoLsDtFr").Enabled = True: adc("UoLsDtTo").Enabled = True
    End If
    '*** Only For Recovery ***
    If UCase(adc.MenuCd) = UCase("RepVchPrnCNV") Then gs_TxnTcTyp = "CNV"
    If UCase(adc.MenuCd) = UCase("RepVchPrnMV") Then gs_TxnTcTyp = "MV"
    ChkBoxArr(1).Visible = True: ChkBoxArr(1).Enabled = True         ' ***** Manali 3.10.0 - TdDesc in Loose Rm Vch
  End Select
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  
  ws_FrmCaption = Me.Caption  '****** Sachin 2.14.0
  
  Call SetGroupSort
  CRV_REP.DisplayGroupTree = False
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  If UCase(adc.MenuCd) = UCase("RepVchPrnDT") Then gs_TxnTcTyp = "DT"
  If UCase(adc.MenuCd) = UCase("RepVchPrnBS") Then gs_TxnTcTyp = "BS"
  ' **** Zubin 211 **** '
  If UCase(adc.MenuCd) = UCase("RepVchPrnBM") Then gs_TxnTcTyp = "BM"
  ' **** Zubin 211 **** '
  If UCase(adc.MenuCd) = UCase("RepVchPrnMLT") Then gs_TxnTcTyp = "MLT"
  If UCase(adc.MenuCd) = UCase("RepVchPrnBV") Then gs_TxnTcTyp = "BV"
  If UCase(adc.MenuCd) = UCase("RepVchPrnREJ") Then gs_TxnTcTyp = "REJ"
  If UCase(adc.MenuCd) = UCase("RepVchPrnPR") Then gs_TxnTcTyp = "PR"
  If UCase(adc.MenuCd) = UCase("RepVchPrnREC") Then gs_TxnTcTyp = "REC"
  If UCase(adc.MenuCd) = UCase("RepVchPrnCNV") Then gs_TxnTcTyp = "CNV"
  If UCase(adc.MenuCd) = UCase("RepVchPrnMV") Then gs_TxnTcTyp = "MV"
  
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
  If UCase(adc.MenuCd) = UCase("RepVchPrnDT") Then gs_TxnTcTyp = "DT"
  If UCase(adc.MenuCd) = UCase("RepVchPrnBS") Then gs_TxnTcTyp = "BS"
  ' **** Zubin 211 **** '
  If UCase(adc.MenuCd) = UCase("RepVchPrnBM") Then gs_TxnTcTyp = "BM"
  ' **** Zubin 211 **** '
  If UCase(adc.MenuCd) = UCase("RepVchPrnMLT") Then gs_TxnTcTyp = "MLT"
  If UCase(adc.MenuCd) = UCase("RepVchPrnBV") Then gs_TxnTcTyp = "BV"
  If UCase(adc.MenuCd) = UCase("RepVchPrnREJ") Then gs_TxnTcTyp = "REJ"
  If UCase(adc.MenuCd) = UCase("RepVchPrnPR") Then gs_TxnTcTyp = "PR"
  If UCase(adc.MenuCd) = UCase("RepVchPrnREC") Then gs_TxnTcTyp = "REC"
  If UCase(adc.MenuCd) = UCase("RepVchPrnCNV") Then gs_TxnTcTyp = "CNV"
  If UCase(adc.MenuCd) = UCase("RepVchPrnMV") Then gs_TxnTcTyp = "MV"
  
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
  Rep1.wCoCd.SetText ws_HName
  Rep2.wCoCd.SetText ws_HName
  Rep3.wCoCd.SetText ws_HName
  
  
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
    ws_Dt = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
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
  Set Rep1 = Nothing
  Set Rep2 = Nothing
  Set Rep3 = Nothing
  '*** (09/08/05)
End Sub

Private Sub OptAck_Click(Index As Integer)
Select Case Index
Case Is = 0
    adc("UoYN2") = "Y"
Case Is = 1
    adc("UoYN2") = "N"
Case Is = 2
    adc("UoYN2") = ""
End Select
End Sub

Private Sub OptAck_GotFocus(Index As Integer)
  DispMsg FraOptAck.ToolTipText, etInfo
End Sub

'StkRt-if password is change to blank, Show StockRate check box will turn into blank
Private Sub TxtPwd_Change()
    If Trim(TxtPwd.text) = "" Then
        ChkBoxArr(2).Value = Unchecked
    End If
End Sub
'StkRt-if user is Seo user then validate with their Seo password else checks with SEOPWD password
Private Sub TxtPwd_Validate(Cancel As Boolean)
  Dim ws_VldPwd As String
  ws_VldPwd = moCn.GetFldVal("select isnull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + adc.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ")
  If Trim(TxtPwd.text) <> "" And Trim(TxtPwd.text) <> ws_VldPwd Then
      Cancel = True:  DispMsg "Invalid Password / User Does Not Have Seo Rights ", etError
  End If
End Sub


