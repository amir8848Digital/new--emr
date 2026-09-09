VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrpWipBPcs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WIP Bag Pieces Report"
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
      TabIndex        =   76
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
      TabIndex        =   75
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
      TabIndex        =   85
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   86
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
         TabIndex        =   87
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
      TabIndex        =   84
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   81
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
      TabIndex        =   77
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
      Height          =   9645
      Left            =   0
      TabIndex        =   80
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   79
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpWipBPcs.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpWipBPcs.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraLoc 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1440
               TabIndex        =   139
               Top             =   680
               Width           =   4305
               Begin VB.OptionButton OptLoc 
                  Caption         =   "Process Loc"
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
                  Left            =   2800
                  TabIndex        =   141
                  ToolTipText     =   "Select Production or Process"
                  Top             =   0
                  Width           =   1545
               End
               Begin VB.OptionButton OptLoc 
                  Caption         =   "Production Loc"
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
                  Left            =   100
                  TabIndex        =   140
                  ToolTipText     =   "Select Production or Process"
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   2085
               End
            End
            Begin VB.Frame FraOptBagIn 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1545
               TabIndex        =   133
               ToolTipText     =   "Show Bags In FG / INV / Both"
               Top             =   7665
               Width           =   3825
               Begin VB.OptionButton OptBagIn 
                  Caption         =   "Both"
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
                  TabIndex        =   136
                  ToolTipText     =   "Show Bags In FG / INV / Both"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptBagIn 
                  Caption         =   "INV"
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
                  TabIndex        =   135
                  ToolTipText     =   "Show Bags In FG / INV / Both"
                  Top             =   0
                  Width           =   945
               End
               Begin VB.OptionButton OptBagIn 
                  Caption         =   "FG"
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
                  TabIndex        =   134
                  ToolTipText     =   "Show Bags In FG / INV / Both"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.Frame FraOptBagPcs 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1545
               TabIndex        =   123
               ToolTipText     =   "Show Normal Or Rejected Bag Pieces"
               Top             =   7950
               Width           =   2595
               Begin VB.OptionButton OptBagPcs 
                  Caption         =   "Normal"
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
                  TabIndex        =   69
                  ToolTipText     =   "Show Normal Or Rejected Bag Pieces"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptBagPcs 
                  Caption         =   "Rejected"
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
                  TabIndex        =   70
                  ToolTipText     =   "Show Normal Or Rejected Bag Pieces"
                  Top             =   0
                  Width           =   1185
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Picture "
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
               Left            =   30
               TabIndex        =   71
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   8235
               Width           =   1725
            End
            Begin VB.Frame FraOptBag 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1545
               TabIndex        =   122
               ToolTipText     =   "Show Open Or Closed Or All Bags"
               Top             =   7380
               Width           =   3825
               Begin VB.OptionButton OptBag 
                  Caption         =   "Open"
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
                  TabIndex        =   66
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptBag 
                  Caption         =   "Closed"
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
                  TabIndex        =   67
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   945
               End
               Begin VB.OptionButton OptBag 
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
                  TabIndex        =   68
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1545
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   61
               Top             =   6165
               Width           =   2340
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4245
               TabIndex        =   7
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   1305
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6930
               TabIndex        =   8
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1305
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4245
               TabIndex        =   19
               ToolTipText     =   "Enter To Design Category"
               Top             =   2445
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
               Left            =   6930
               TabIndex        =   20
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2445
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
               Index           =   31
               Left            =   4245
               TabIndex        =   22
               ToolTipText     =   "Enter To Design Code"
               Top             =   2730
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
               Left            =   6930
               TabIndex        =   23
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2730
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
               Index           =   41
               Left            =   6930
               TabIndex        =   39
               ToolTipText     =   "Enter Customer Selection"
               Top             =   3885
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
               Index           =   5
               Left            =   5025
               TabIndex        =   31
               ToolTipText     =   "Enter To Order Year"
               Top             =   3300
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
               Left            =   2325
               TabIndex        =   28
               ToolTipText     =   "Enter From Order Year"
               Top             =   3315
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
               Index           =   13
               Left            =   2745
               TabIndex        =   29
               ToolTipText     =   "Enter From Order Character"
               Top             =   3315
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
               Left            =   4245
               TabIndex        =   30
               ToolTipText     =   "Enter To Order Tc"
               Top             =   3300
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
               Left            =   5445
               TabIndex        =   32
               ToolTipText     =   "Enter To Order Character"
               Top             =   3300
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
               Index           =   54
               Left            =   4245
               TabIndex        =   35
               ToolTipText     =   "Enter To Order Number"
               Top             =   3585
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
               Index           =   15
               Left            =   4245
               TabIndex        =   25
               ToolTipText     =   "Enter To Karat"
               Top             =   3015
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
               Left            =   6930
               TabIndex        =   26
               ToolTipText     =   "Enter Karat Selection"
               Top             =   3015
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
               Index           =   43
               Left            =   4245
               TabIndex        =   41
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   4155
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
               Left            =   6930
               TabIndex        =   42
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   4170
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
               Index           =   21
               Left            =   5130
               TabIndex        =   36
               ToolTipText     =   "Enter To Order Serial"
               Top             =   3585
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "OdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   2430
               TabIndex        =   34
               ToolTipText     =   "Enter From Order Serial"
               Top             =   3600
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "OdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   1545
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   365
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "BCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   6930
               TabIndex        =   11
               ToolTipText     =   "Enter FgSubLoc Location Selection"
               Top             =   1590
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOFGSUBLOCSEL"
               IdName          =   "UOFGSUBLOCSEL"
               CmpStr          =   "BFgSubLoc In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   4245
               TabIndex        =   60
               ToolTipText     =   "Show Bags which are Dalayed By Less  than or equal to Entered Days: (Enter zero to see all Bags)"
               Top             =   5865
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "##0"
               MaxLength       =   3
               DataType        =   1
               DataField       =   "UOBDLDAYSTO"
               IdName          =   "UOBDLDAYSTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2085
               TabIndex        =   74
               ToolTipText     =   "Show Picture? ( Enter 'Y'/'N')"
               Top             =   8235
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   1545
               TabIndex        =   64
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   7020
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOBAGTYP"
               ReCalcParent    =   "UOBCHRFR"
               IdName          =   "UOBAGTYP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   3015
               TabIndex        =   65
               ToolTipText     =   "Enter Bag Character"
               Top             =   7095
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRFR"
               ReCalcOn        =   "UOYN2"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "BChr = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1545
               TabIndex        =   6
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   1305
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   1545
               TabIndex        =   18
               ToolTipText     =   "Enter From Design Category"
               Top             =   2460
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
               Left            =   1545
               TabIndex        =   21
               ToolTipText     =   "Enter From Design Code"
               Top             =   2745
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
               Index           =   10
               Left            =   1545
               TabIndex        =   27
               ToolTipText     =   "Enter From Order Tc"
               Top             =   3315
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
               Index           =   53
               Left            =   1545
               TabIndex        =   33
               ToolTipText     =   "Enter From Order Number"
               Top             =   3600
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
               Index           =   3
               Left            =   5955
               TabIndex        =   73
               ToolTipText     =   "Show Normal Or Rejected Pieces ?"
               Top             =   7950
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOBAGPCS"
               IdName          =   "UOBAGPCS"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1545
               TabIndex        =   24
               ToolTipText     =   "Enter From Karat"
               Top             =   3030
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
               Index           =   45
               Left            =   1545
               TabIndex        =   40
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   4170
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
               Index           =   6
               Left            =   5955
               TabIndex        =   72
               ToolTipText     =   "Show Open Or Closed Or All Bags ? (Open - 'Y', Closed - 'N', All - Blank)"
               Top             =   7380
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1545
               TabIndex        =   9
               ToolTipText     =   "Enter From FgSubLoc Location"
               Top             =   1590
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOFGSUBLOCFR"
               IdName          =   "UOFGSUBLOCFR"
               CmpStr          =   "BFgSubLoc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   1965
               TabIndex        =   54
               ToolTipText     =   "Enter From Packing List Character"
               Top             =   5595
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFDPLCHRFR"
               IdName          =   "UOFDPLCHRFR"
               CmpStr          =   "FdPlChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   2745
               TabIndex        =   55
               ToolTipText     =   "Enter From Packing List Number"
               Top             =   5595
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFDPLNOFR"
               IdName          =   "UOFDPLNOFR"
               CmpStr          =   "FdPlNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   4245
               TabIndex        =   56
               ToolTipText     =   "Enter To Packing List Year"
               Top             =   5580
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOFDPLYYTO"
               IdName          =   "UOFDPLYYTO"
               CmpStr          =   "FdPlYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   4665
               TabIndex        =   57
               ToolTipText     =   "Enter To Packing List Character"
               Top             =   5580
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFDPLCHRTO"
               IdName          =   "UOFDPLCHRTO"
               CmpStr          =   "FdPlChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   5445
               TabIndex        =   58
               ToolTipText     =   "Enter To Packing List Number"
               Top             =   5580
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFDPLNOTO"
               IdName          =   "UOFDPLNOTO"
               CmpStr          =   "FdPlNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   1545
               TabIndex        =   62
               ToolTipText     =   "Specify If The Value To Be Shown Should Be Stock Value, Cost Value Or Sale Value Or No Value"
               Top             =   6450
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   3000
               TabIndex        =   63
               ToolTipText     =   "Enter Currency Code Whose Orders You Wish To View Or Leave It Blank If You Wish To View All Orders In Base Currency"
               Top             =   6480
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               ReCalcOn        =   "UOINVAL"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   4245
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   365
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "BCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   6930
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   360
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "BCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   4245
               TabIndex        =   4
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1005
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDTO"
               IdName          =   "UOPRNTCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   6930
               TabIndex        =   5
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1020
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   1545
               TabIndex        =   3
               ToolTipText     =   "Enter From Parent Code"
               Top             =   1005
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDFR"
               IdName          =   "UOPRNTCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   1545
               TabIndex        =   43
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   4455
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   4245
               TabIndex        =   44
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   4440
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   6930
               TabIndex        =   45
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   4455
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "OdPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   4245
               TabIndex        =   50
               ToolTipText     =   "Enter To Order Date"
               Top             =   5010
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
               Index           =   58
               Left            =   1545
               TabIndex        =   49
               ToolTipText     =   "Enter From Order Date"
               Top             =   5025
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
               Index           =   59
               Left            =   4245
               TabIndex        =   52
               ToolTipText     =   "Enter To Production Delivery Date"
               Top             =   5295
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1TO"
               IdName          =   "UODELDT1TO"
               CmpStr          =   "OdDelDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   1545
               TabIndex        =   51
               ToolTipText     =   "Enter From Production Delivery Date"
               Top             =   5310
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1FR"
               IdName          =   "UODELDT1FR"
               CmpStr          =   "OdDelDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   4245
               TabIndex        =   47
               ToolTipText     =   "Enter To Priority Code"
               Top             =   4725
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOPRTCDTO"
               IdName          =   "UOPRTCDTO"
               CmpStr          =   "OdPrtCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   6930
               TabIndex        =   48
               ToolTipText     =   "Enter Priority Code Selection"
               Top             =   4740
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRTCDSEL"
               IdName          =   "UOPRTCDSEL"
               CmpStr          =   "OdPrtCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   64
               Left            =   1545
               TabIndex        =   46
               ToolTipText     =   "Enter From Priority Code"
               Top             =   4740
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOPRTCDFR"
               IdName          =   "UOPRTCDFR"
               CmpStr          =   "OdPrtCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   66
               Left            =   6930
               TabIndex        =   14
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1875
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORECLOCSEL"
               IdName          =   "UORECLOCSEL"
               CmpStr          =   "IsNull(TdFrBLoc,BLoc) In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   4245
               TabIndex        =   10
               ToolTipText     =   "Enter To FgSubLoc Location"
               Top             =   1590
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOFGSUBLOCTO"
               IdName          =   "UOFGSUBLOCTO"
               CmpStr          =   "BFgSubLoc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   65
               Left            =   4245
               TabIndex        =   13
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   1875
               Visible         =   0   'False
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORECLOCTO"
               IdName          =   "UORECLOCTO"
               CmpStr          =   "IsNull(TdFrBLoc,BLoc) <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4245
               TabIndex        =   16
               ToolTipText     =   "Enter To By Worker"
               Top             =   2160
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKTO"
               IdName          =   "UOWRKTO"
               CmpStr          =   "BWrk <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   6930
               TabIndex        =   17
               ToolTipText     =   "Enter By Worker Selection"
               Top             =   2160
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOWRKSEL"
               IdName          =   "UOWRKSEL"
               CmpStr          =   "BWrk In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1545
               TabIndex        =   15
               ToolTipText     =   "Enter From By Worker"
               Top             =   2160
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKFR"
               IdName          =   "UOWRKFR"
               CmpStr          =   "BWrk >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   67
               Left            =   1545
               TabIndex        =   12
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   1890
               Visible         =   0   'False
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORECLOCFR"
               IdName          =   "UORECLOCFR"
               CmpStr          =   "IsNull(TdFrBLoc,BLoc) >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4245
               TabIndex        =   38
               ToolTipText     =   "Enter To Customer"
               Top             =   3885
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
               Index           =   42
               Left            =   1545
               TabIndex        =   37
               ToolTipText     =   "Enter From Customer"
               Top             =   3885
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
               Index           =   22
               Left            =   1545
               TabIndex        =   59
               ToolTipText     =   "Show Bags which are Dalayed By More than or equal to Entered Days: (Enter zero to see all Bags)"
               Top             =   5880
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "##0"
               MaxLength       =   3
               DataType        =   1
               DataField       =   "UOBDLDAYS"
               IdName          =   "UOBDLDAYS"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1545
               TabIndex        =   53
               ToolTipText     =   "Enter From Packing List Year"
               Top             =   5595
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOFDPLYYFR"
               IdName          =   "UOFDPLYYFR"
               CmpStr          =   "FdPlYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   68
               Left            =   1545
               TabIndex        =   131
               ToolTipText     =   "Enter From Source Sketch / Design Code"
               Top             =   6765
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOBOEFR"
               IdName          =   "UOBOEFR"
               CmpStr          =   "DmSrcDsgCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   69
               Left            =   4245
               TabIndex        =   132
               ToolTipText     =   "Enter To Source Sketch / Design Code"
               Top             =   6765
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOBOETO"
               IdName          =   "UOBOETO"
               CmpStr          =   "DmSrcDsgCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   5955
               TabIndex        =   137
               ToolTipText     =   "Show Bags In FG / INV / Both"
               Top             =   7665
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   6960
               TabIndex        =   142
               ToolTipText     =   "Show Production or Process Location"
               Top             =   720
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOLOCTYPFR"
               IdName          =   "UOLOCTYPFR"
            End
            Begin VB.Label LblBagIn 
               BackStyle       =   0  'Transparent
               Caption         =   "Bags In"
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
               TabIndex        =   138
               Top             =   7665
               Width           =   1185
            End
            Begin VB.Label LblSrcDsgCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Src Sk/Dsg Cd"
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
               TabIndex        =   130
               ToolTipText     =   "Location"
               Top             =   6735
               Width           =   1485
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Delay Days"
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
               TabIndex        =   129
               Top             =   5880
               Width           =   1485
            End
            Begin VB.Label LblPckLst 
               BackStyle       =   0  'Transparent
               Caption         =   "Packing List"
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
               TabIndex        =   128
               Top             =   5595
               Width           =   1485
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "From Process"
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
               TabIndex        =   127
               Top             =   1890
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Priority"
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
               TabIndex        =   126
               Top             =   4740
               Width           =   1425
            End
            Begin VB.Label lblPrdDelDt 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Del Dt"
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
               TabIndex        =   125
               Top             =   5310
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order Dt"
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
               TabIndex        =   124
               Top             =   5025
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sub PO No"
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
               TabIndex        =   121
               Top             =   4455
               Width           =   1425
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
               Left            =   60
               TabIndex        =   120
               Top             =   1005
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
               Height          =   285
               Left            =   2040
               TabIndex        =   119
               Top             =   6480
               Width           =   1005
            End
            Begin VB.Label lblvaltyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Value Type"
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
               TabIndex        =   118
               ToolTipText     =   "Location"
               Top             =   6450
               Width           =   1485
            End
            Begin VB.Label lblseopwd 
               BackStyle       =   0  'Transparent
               Caption         =   "Enter SEO Pwd"
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
               TabIndex        =   117
               Top             =   6165
               Width           =   1575
            End
            Begin VB.Label LblPkLstDeLimit4 
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
               Left            =   5385
               TabIndex        =   116
               Top             =   5580
               Width           =   105
            End
            Begin VB.Label LblPkLstDeLimit3 
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
               Left            =   4605
               TabIndex        =   115
               Top             =   5580
               Width           =   105
            End
            Begin VB.Label LblPkLstDeLimit2 
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
               Left            =   2685
               TabIndex        =   114
               Top             =   5595
               Width           =   105
            End
            Begin VB.Label LblPkLstDeLimit1 
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
               Left            =   1905
               TabIndex        =   113
               Top             =   5595
               Width           =   105
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
               Left            =   2100
               TabIndex        =   112
               Top             =   7095
               Width           =   885
            End
            Begin VB.Label LblBagTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Type"
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
               TabIndex        =   111
               ToolTipText     =   "Location"
               Top             =   7020
               Width           =   1035
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Picture"
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   16
               Left            =   120
               TabIndex        =   110
               ToolTipText     =   "Location"
               Top             =   8280
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label lblFgSubLoc 
               BackStyle       =   0  'Transparent
               Caption         =   "Fg Sub Loc"
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
               TabIndex        =   109
               Top             =   1590
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
               TabIndex        =   108
               Top             =   365
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Open Bags"
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
               TabIndex        =   107
               Top             =   7380
               Width           =   1185
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
               Index           =   31
               Left            =   2370
               TabIndex        =   106
               Top             =   3600
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
               Index           =   30
               Left            =   5070
               TabIndex        =   105
               Top             =   3585
               Width           =   105
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
               Left            =   60
               TabIndex        =   104
               Top             =   4170
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
               TabIndex        =   103
               Top             =   3030
               Width           =   1425
            End
            Begin VB.Label LblBagPcs 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Pcs"
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
               TabIndex        =   102
               Top             =   7950
               Width           =   1275
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Worker"
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
               TabIndex        =   101
               Top             =   2175
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
               Left            =   60
               TabIndex        =   100
               Top             =   3315
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Ord No/ Sr"
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
               TabIndex        =   99
               Top             =   3600
               Width           =   1515
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
               Index           =   24
               Left            =   5385
               TabIndex        =   98
               Top             =   3300
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
               Index           =   25
               Left            =   2685
               TabIndex        =   97
               Top             =   3315
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
               Index           =   26
               Left            =   4965
               TabIndex        =   96
               Top             =   3300
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
               Index           =   27
               Left            =   2265
               TabIndex        =   95
               Top             =   3315
               Width           =   105
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
               TabIndex        =   94
               Top             =   3885
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
               TabIndex        =   93
               Top             =   2745
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
               Left            =   60
               TabIndex        =   92
               Top             =   2460
               Width           =   1425
            End
            Begin VB.Label LblProcess 
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
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
               TabIndex        =   91
               Top             =   1305
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
               Left            =   6900
               TabIndex        =   90
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
               TabIndex        =   89
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
               TabIndex        =   88
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   83
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   78
               Top             =   360
               Width           =   9045
               _ExtentX        =   15954
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpWipBPcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepWipBPcs
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  'Repair - repair bag fields included
  If UCase(adc.MenuCd) = UCase("RprWIP") Then
    GRP_REP.Add "Location", "BRepLoc", "", "BLoc", "hBLoc", "Company Code,Parent Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'P' and LocCd= BRepLoc) "
    GRP_REP.Add "Process", "LocPrcs", "", "BLoc", "hBLoc", "Company Code,Parent Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'R' and LocCd= Lc.LocPrcs) "
    GRP_REP.Add "(Location)Worker", "BRepLoc+BWrk", "'('+ RTrim(BRepLoc)+ ') '+ BWrk", "BLoc,BWrk", "hBLoc,hBWrk", "Parent Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'WORK' and vPMCd= BWrk) "
    GRP_REP.Add "(Process)Worker", "LocPrcs+BWrk", "'('+ RTrim(LocPrcs)+ ') '+ BWrk", "BLoc,BWrk", "hBLoc,hBWrk", "Parent Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'WORK' and vPMCd= BWrk) "
    
    GRP_REP.Add "OrderNo", "BRepOdTc+BRepOdYy+BRepOdChr+str(BRepOdNo)", "BRepOdTc+'/'+BRepOdYy+'/'+BRepOdChr+'/'+LTrim(Str(BRepOdNo))", "OmCmCd,OmPoNo", "hOmCmCd,hOmPoNo", "Company Code,Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
    GRP_REP.Add "OrderSr", "BRepOdTc+BRepOdYy+BRepOdChr+str(BRepOdNo)+str(BRepOdSr)", "BRepOdTc+'/'+BRepOdYy+'/'+BRepOdChr+'/'+LTrim(Str(BRepOdNo))+'/'+LTrim(Str(BRepOdSr))", "OmCmCd,OmPoNo", "hOmCmCd,hOmPoNo", "Company Code,OrderNo,Ord Kt,Design Cd,Design Ctg,(DsgCd)Colour,Customer Code,Customer PO,Prd Del Date,Priority,(OrdKt)Colour", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
    GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "Company Code", "", "", "BRepOdTc+'/'+BRepOdYy+'/'+BRepOdChr+'/'+LTrim(Str(BRepOdNo))"
    GRP_REP.Add "Cust Sub PO", "OmCmCd+OmPONo+OdPoNo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)+(case when OdPONo<> '' then '/'+ LTrim(OdPONo) else '' end )", "OmCmCd,OmPoNo,OdPoNo", "hOmCmCd,hOmPONo,hOdPONo", "Customer Code,Customer PO", "", "", "BRepOdTc+'/'+BRepOdYy+'/'+BRepOdChr+'/'+LTrim(Str(BRepOdNo))+'/'+LTrim(Str(BRepOdSr))+(case when OdPONo<> '' then '/'+ LTrim(OdPONo) else '' end )"
  Else
    GRP_REP.Add "Location", "BLoc", "", "BLoc", "hBLoc", "Company Code,Parent Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'P' and LocCd= BLoc) "
    GRP_REP.Add "Process", "LocPrcs", "", "BLoc", "hBLoc", "Company Code,Parent Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'R' and LocCd= Lc.LocPrcs) "
    GRP_REP.Add "(Location)Worker", "BLoc+BWrk", "'('+ RTrim(BLoc)+ ') '+ BWrk", "BLoc,BWrk", "hBLoc,hBWrk", "Parent Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'WORK' and vPMCd= BWrk) "
    GRP_REP.Add "(Process)Worker", "LocPrcs+BWrk", "'('+ RTrim(LocPrcs)+ ') '+ BWrk", "BLoc,BWrk", "hBLoc,hBWrk", "Parent Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'WORK' and vPMCd= BWrk) "
    
    GRP_REP.Add "OrderNo", "BOdTc+BOdYy+BOdChr+str(BOdNo)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "OmCmCd,OmPoNo", "hOmCmCd,hOmPoNo", "Company Code,Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
    GRP_REP.Add "OrderSr", "BOdTc+BOdYy+BOdChr+str(BOdNo)+str(BOdSr)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))", "OmCmCd,OmPoNo", "hOmCmCd,hOmPoNo", "Company Code,OrderNo,Ord Kt,Design Cd,Design Ctg,(DsgCd)Colour,Customer Code,Customer PO,Prd Del Date,Priority,(OrdKt)Colour", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
    GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "Company Code", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))"
    GRP_REP.Add "Cust Sub PO", "OmCmCd+OmPONo+OdPoNo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)+(case when OdPONo<> '' then '/'+ LTrim(OdPONo) else '' end )", "OmCmCd,OmPoNo,OdPoNo", "hOmCmCd,hOmPONo,hOdPONo", "Customer Code,Customer PO", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))+(case when OdPONo<> '' then '/'+ LTrim(OdPONo) else '' end )"
    
    GRP_REP.Add "FgSubLoc", "BFgSubLoc", "", "", "", "Company Code", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'FGSUBLOC' and vPMCd= BFgSubLoc) "
  End If
  
  GRP_REP.Add "Design Ctg", "DmCtg", "", "DmCtg", "hDmCtg", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
  GRP_REP.Add "Design Code", "BOdDmCd", "", "BOdDmCd", "hBOdDmCd", "Design Ctg,(DsgCd)Colour", "", "", "Substring(DmDesc,1,120) "
  GRP_REP.Add "Ord Kt", "OdKt", "", "OdKt", "hOdKt", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt) "
  'GRP_REP.Add "OrderNo", "BOdTc+BOdYy+BOdChr+str(BOdNo)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "wOrdNo,OmCmCd,OmPoNo", "hOrdNo,hOmCmCd,hOmPoNo", "Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "

'?????? check
  ' Zubin 212 (Correction 'Order No' changed to 'OrderNo', 'Priority' and (OrdKt)Colour shld be hidden)
  
   '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  ' ***** Manali 3.8.0 - EmrFrmMDI.TXT_UCD.text Replaced by gs_UsrCd
  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
  'GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
  
  '*** Jay 3.2.0 *** [OdPoNo]
  
  '*** Jay 3.2.0 *** [OdPoNo]
  GRP_REP.Add "(DsgCd)Colour", "BOdDmCd+OdDmCol", "'('+BOdDmCd+') '+OdDmCol", "BOdDmCd,OdDmCol", "hBOdDmCd,hOdDmCol", "Design Ctg,Design Code", "", "", "IsNull(Substring(DmDesc,1,120), '')+ ' ('+(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) +')'"
  GRP_REP.Add "Prd Del Date", "convert(varchar(8), OdDelDt, 112)", "convert(varchar, OdDelDt, 103)", "", "", "", "", "", ""
  
  GRP_REP.Add "(OrdKt)Colour", "OdKt+OdDmCol", "'('+OdKt+') '+OdDmCol", "OdKt,OdDmCol", "hOdKt,hOdDmCol", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt)+ ' ('+(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) +')'"
  GRP_REP.Add "Priority", "OdPrtCd", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'PRTCD' and PMCd= OdPrtCd) "
  
  
'  GRP_REP.Add "Fr Process", "TdFrBLoc", "", "", "", "", "", "", "IsNull(TdFrBLoc,BLoc)"

  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ,Process,OrderNo,OrderSr,Customer PO,FgSubLoc ******
  If ws_MultiCoMod = True Then
    GRP_REP.Add "Company Code", "BCoCd", "BCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=OdCoCd) "
    GRP_REP.Add "Parent Production", "LocPrntCd", "LocPrntCd", "", "", "", "", "", "(Select Ln.LocDesc from Loc Ln where Ln.LocCd=Lc.LocPrntCd and Ln.LocCoCd='" + ctSelfCoCd + "') "
     'GRP_REP.Add "Parent Process", "LocPrntCd", "LocPrntCd", "", "", "", "", "", "(Select Ln.LocDesc from Loc Ln where Ln.LocCd=Lc.LocPrntCd and Ln.LocCoCd='" + ctSelfCoCd + "') "
     GRP_REP.Add "Parent Process", "(Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs)", "(Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs)", "", "", "", "", "", "(Select Ln.LocDesc from Loc Ln where Ln.LocCd=(Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs) and Ln.LocCoCd='" + ctSelfCoCd + "') "
                                                                              '(Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs)
 
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******
 
' **** Zubin 211 **** '
  If UCase(adc.MenuCd) = UCase("RepWipBPcs") Or UCase(adc.MenuCd) = UCase("RprWIP") Then
    '*** Jay 2.13(CT) ***
    'Uma *** new 'D' btyp
    GRP_REP.Add "Bag Type", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", _
      "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", "", "", "Bag Typ+Chr", "", "", "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')"
    GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
    GRP_REP.Add "Bag Typ+Chr", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+BChr", _
      "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+'/ '+ BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
    'GRP_REP.Add "Bag Type", "PValue", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end )", "", "", "Bag Typ+Chr", "", "", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)"
    'GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
    'GRP_REP.Add "Bag Typ+Chr", "PValue+BChr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+ '/'+ BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
    ' **** Zubin 211 **** '
    '*** Jay 2.13(CT) ***
  '*** 2.11 (Jen)
    
  ElseIf UCase(adc.MenuCd) = UCase("RepWipFgBPcs") Then
    GRP_REP.Add "Pck List", "FdPlYy+FdPlChr+str(FdPlNo)", "(case when FdPlNo> 0 then FdPlYy+'/'+FdPlChr+'/'+LTrim(Str(FdPlNo)) else '' end)", "", "", "", "", "", ""
    GRP_REP.Add "FG / INV", " (Case when exists (Select 'x' From InvFgd where IfBIdNo = BIdNo and IfFdBYy = BYy and IfFdBChr = BChr and IfFdBNo = BNo and IfCoCd = BCoCd) then 'INV' else 'FG' end)", " (Case when exists (Select 'x' From InvFgd where IfBIdNo = BIdNo and IfFdBYy = BYy and IfFdBChr = BChr and IfFdBNo = BNo and IfCoCd = BCoCd) then 'INV' else 'FG' end)", "", "", "", "", "", ""
  '*** 2.11 (Jen)
  End If

End Sub
Private Sub adc_setreprecsource()
  '/** urmi BDlDays 2.06 */
  '*** Report Sql ***
'*** Jay 3.2.0 [PicNm]
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(45) as DGrp1, space(45) as DGrp2, space(45) as DGrp3,
'  space(45) as G1Desc, space(45) as G2Desc, space(45) as G3Desc,
'  BCoCd, BYy, BChr, BNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr,DmTcTyp, BOdDmCd, OdDmSz, DmCtg,
'  OdKt, OdDmCol, OdPicNm, BLoc, LocPrntCd, BWrk, BGrWt,
'  9999 as BDlDays,
'  BQty As qQty, BPtQty As qPtQty, OmCmCd, OmPoNo
'  From Bag, DsgMst, OrdDsg, OrdMst where 1=2
  '*** Report Sql ***
  On Error GoTo RepErr
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wQty As String
  Dim wPtQty As String, wSelCnd As String, wBagClsCnd As String
  
  '*** 2.11 (Jen)
  Dim ws_FgJoin As String
  '*** 2.11 (Jen)
  
  ' ****** sachin 2.12 - 24/11/05 ****** '
  Dim wFrVal As String
  ' ****** sachin 2.12 - 24/11/05 ****** '
  
  ' urmi BDlDays 2.06
  Dim wBagDelayCnd As String

' **** Zubin 211 **** '
  Dim ws_BagJoin As String
' **** Zubin 211 **** '
  Dim wAddlCnd As String  '****** Sachin 3.02.0 - New Fields In Where Clause

  ' Zubin 213
  Dim wb_InBaseCur As Boolean
  Dim ws_BaseCurCd As String
  wb_InBaseCur = IIF(UCase(adc.MenuCd) = UCase("RepWipFgBPcs") And UCase(adc("UoCurCdFr")) = "", True, False)
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  ' Zubin 213
      
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
    grpflds = GRP_REP.GrpFldLst
   
   '*** To check if Grp1, Grp2 is selected and also to check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer
  Dim wi_PckLstPos As Integer   '*** 2.11 (Jen)
  Dim ws_FGINVChk As String
  ws_FGINVChk = ""
    
    wGrp(1) = GRP_REP.Value(gltGroup1)
    wGrp(2) = GRP_REP.Value(gltGroup2)
    wGrp(3) = GRP_REP.Value(gltGroup3)
    wDetPos = 0: wi_PckLstPos = 0
  
  For i = 1 To 6
    If UCase(wGrp(i)) = UCase("Pck List") Then
      wi_PckLstPos = i
    End If
    
    '*** 2.11 (Jen)
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
    '*** 2.11 (Jen)
  Next i
   ' If wDetPos = 0 Then DispMsg "Detail Option Has To Be Selected For This Report", etError
  If wDetPos = 0 Then
     Rep.PicSec2.Suppress = True
  End If

  ' ****** sachin 2.12 - 24/11/05 ****** '
    If UCase(adc("UoPMCdFr")) = "C" Then
      ' Zubin 213 (Bef 213)
      'wFrVal = "OdCstPrc*BQty as qSalVal "
      'Rep.TxtValTyp.SetText "(Value = Cost)"
      ' Zubin 213
      wFrVal = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdCstPrc/OmCnvFct else OdCstPrc*OmCnvFct end) ", "OdCstPrc") + "*BQty as qSalVal "
      Rep.TxtValTyp.SetText "(Value = Cost [" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + "])"
    ElseIf UCase(adc("UoPMCdFr")) = "S" Then
      ' Zubin 213 (Bef 213)
      'wFrVal = "OdSalPrc*BQty as qSalVal "
      'Rep.TxtValTyp.SetText "(Value = Sale)"
      ' Zubin 213
      wFrVal = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc") + "*BQty as qSalVal "
      Rep.TxtValTyp.SetText "(Value = Sale [" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + "])"
    Else
      wFrVal = " 0.00 as qSalVal "
      Rep.TxtValTyp.SetText ""
      Rep.hSalVal.Suppress = True
      Rep.h1SalVal.Suppress = True
    End If
  ' ****** sachin 2.12 - 24/11/05 ****** '

  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    'Uma Location / Process selection
    If OptLoc(0).Value = True Then
      If UCase(adc.MenuCd) = UCase("RprWIP") Then
        If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and BRepLoc >= '" + adc("UoIssLocFr") + "'"
        If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and BRepLoc <= '" + adc("UoIssLocTo") + "'"
        If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and BRepLoc In (" + adc("UoIssLocSel") + ")"
      Else
        If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and BLoc >= '" + adc("UoIssLocFr") + "'"
        If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and BLoc <= '" + adc("UoIssLocTo") + "'"
        If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and BLoc In (" + adc("UoIssLocSel") + ")"
      End If
      If adc("UOPRNTCDFR") <> "" Then wCnd = wCnd + " and LocPrntCd >= '" + adc("UOPRNTCDFR") + "'"
      If adc("UOPRNTCDTO") <> "" Then wCnd = wCnd + " and LocPrntCd <= '" + adc("UOPRNTCDTO") + "'"
      If adc("UOPRNTCDSEL") <> "" Then wCnd = wCnd + " and LocPrntCd In (" + adc("UOPRNTCDSEL") + ")"
      
    ElseIf OptLoc(1).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and LocPrcs >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and LocPrcs <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and LocPrcs In (" + adc("UoIssLocSel") + ")"
            
      If adc("UOPRNTCDFR") <> "" Then wCnd = wCnd + " and (Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs) >= '" + adc("UOPRNTCDFR") + "'"
      If adc("UOPRNTCDTO") <> "" Then wCnd = wCnd + " and (Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs) <= '" + adc("UOPRNTCDTO") + "'"
      If adc("UOPRNTCDSEL") <> "" Then wCnd = wCnd + " and (Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs) In (" + adc("UOPRNTCDSEL") + ")"
            
    End If

  
  ' Zubin 213
  wCnd = wCnd + IIF(UCase(adc.MenuCd) = UCase("RepWipFgBPcs") And wb_InBaseCur = False, _
                                " and OmCmCurCd= '" + adc("UoCurCdFr") + "'", "")
  
  'wSelCnd = IIf(adc("UoBagPcs") = "N", " where (BQty<> 0 or BPtQty<> 0 or (BQty= 0 and BPtQty= 0 and BRjQty= 0 and BRjPtQty= 0 and Round(BGrWt, 3)<> 0)) ", " where (BRjQty<> 0 or BRjPtQty<> 0) ")
  '### wSelCnd = IIf(ADC("UoBagPcs") = "N", " where (BQty<> 0 or BPtQty<> 0) ", " where (BRjQty<> 0 or BRjPtQty<> 0) ")    '(This condition should have given the same result as above but it does not give)
  '### wSelCnd = IIf(ADC("UoBagPcs") = "N", " where (BRjQty= 0 and BRjPtQty= 0) ", " where (BRjQty<> 0 or BRjPtQty<> 0) ")
  
  wBagClsCnd = IIF(adc("UoYN") = "Y", " and BCls= 'N' ", IIF(adc("UoYN") = "N", " and BCls= 'Y' ", ""))
  
  ' urmi BDlDays 2.06
  wBagDelayCnd = IIF(adc("UoBDlDays") = 0, "", "and datediff(dd, BRecvDt, '" + CStr(moCn.SrvrDate) + "') >= " + CStr(adc("UoBDlDays")))
  '*****Geeta*****Emr206********
  wBagDelayCnd = wBagDelayCnd + IIF(adc("UoBDlDaysTo") = 0, "", " and  datediff(dd, BRecvDt, '" + CStr(moCn.SrvrDate) + "') <= " + CStr(adc("UoBDlDaysTo")))
  '*****Geeta*****Emr206********
  Dim ws_TdYyKeyStr As String
  If UCase(adc.MenuCd) = UCase("RepWipBPcs") Then
    wQty = IIF(adc("UoBagPcs") = "N", " BQty as qQty", " BRjQty as qQty")
    wPtQty = IIF(adc("UoBagPcs") = "N", " BPtQty as qPtQty", " BRjPtQty as qPtQty")
    wSelCnd = IIF(adc("UoBagPcs") = "N", " where (BQty<> 0 or BPtQty<> 0 or (BQty= 0 and BPtQty= 0 and BRjQty= 0 and BRjPtQty= 0 and Round(BGrWt, 3)<> 0)) ", " where (BRjQty<> 0 or BRjPtQty<> 0) ")
    wSelCnd = wSelCnd + " and BLoc<> 'PFG' "
    If adc("UoBagPcs") = "N" Then Rep.TxtHead.SetText "WIP Bag Pieces Report (Normal)"
    If adc("UoBagPcs") = "R" Then Rep.TxtHead.SetText "WIP Bag Pieces Report (Rejected)"
    Rep.TxtLocFr.SetText adc("UoIssLocFr")
    Rep.TxtLocTo.SetText adc("UoIssLocTo")
    Rep.TxtLocSel.SetText adc("UoIssLocSel")
    
    '*** Jay 2.13(CT) ***
    'Uma *** blank <> 'D'
    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                    IIF(adc("UoBagTyp") <> "", " and PValue= '" + adc("UoBagTyp") + "' ", " and PValue <> 'D'") + _
                    IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
      ' **** Zubin 211 **** '
'    If UCase(adc("UoBagTyp")) = "P" Then
'      ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'P' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'    ElseIf UCase(adc("UoBagTyp")) = "F" Then
'      ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'F' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'    ElseIf UCase(adc("UoBagTyp")) = "C" Then
'      ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'C' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'    Else
'      ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'    End If
    '*** Jay 2.13(CT) ***
    ws_FgJoin = " "   '*** 2.11 (Jen)
    ws_TdYyKeyStr = " (Select Max(TdYyKey) From Txnd where TdCoCd=BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo and TdSrNo=0 and TdTc='BV') "
    
  ' **** Zubin 211 **** '
  ElseIf UCase(adc.MenuCd) = UCase("RepWipFgBPcs") Then
    wQty = " BQty as qQty"
    wPtQty = " BPtQty as qPtQty"
    '/*** Where had disappeared from 2.01 so and has replaces by 'where' again ***/
    'wSelCnd = " and BLoc= 'PFG' "
    wSelCnd = " where BLoc= 'PFG' "
    Rep.TxtHead.SetText "Fg Bag Pieces Report"
    Rep.TxthLocFr.SetText "Loc :": Rep.TxthLocTo.SetText "Bags In": Rep.TxthLocSel.SetText " "
    Rep.TxtLocFr.SetText "PFG"
    Rep.TxtLocTo.SetText IIF(adc("UoYn2") = "Y", " FG ", IIF(adc("UoYn2") = "N", " INV ", "FG & INV"))
    
    '*** 2.11 (Jen) (changed Jen 2.14 Next Patch 1)
    ws_FgJoin = " "
    ws_TdYyKeyStr = " BLstYyKey "
    '****** Sachin 3.02 s- Id fields in Joins
    If wi_PckLstPos > 0 Or adc("UoFdPlYyFr") <> "" Or adc("UoFdPlChrFr") <> "" Or adc("UoFdPlNoFr") <> 0 Or _
       adc("UoFdPlYyTo") <> "" Or adc("UoFdPlChrTo") <> "" Or adc("UoFdPlNoTo") <> 0 Then
      ws_FgJoin = " Join Fgd On FdBIdNo=BIdNo and FdCoCd= BCoCd and FdBYy= BYy and FdBChr= BChr and FdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=BPrtKey ", "") + _
                  "     and exists (select 'x' from Param where PTyp= 'TC' and PMCd= FdTc and PValue= 'FB') " + _
                  "     and (FdDespQty= 0 or " + _
                  "          exists (Select 'x' From InvFgd where IfFdIdNo=FdIdNo and IfCoCd= FdCoCd and IfFdTc= FdTc and IfFdYy= FdYy " + _
                  "                  and IfFdChr= FdChr and IfFdNo= FdNo and IfFdSr= FdSr and (Select PValue from Param Where PTyp='TC' and PMCd=IfTc)='" + ctInTcTyp + "' " + IIF(gs_Partition = ctCurrPrtn, " and IfPrtKey=FdPrtKey ", "") + ")) "
    End If
    
    '*** 2.11 (Jen) (changed Jen 2.14 Next Patch 1)
    
    If adc("UoYn2") <> "" Then
        ws_FGINVChk = IIF(adc("UoYn2") = "Y", " Not ", "") + _
            "exists (Select 'x' From InvFgd where IfBIdNo = BIdNo and " + _
            "IfFdBYy = BYy and IfFdBChr = BChr and IfFdBNo = BNo and IfCoCd = BCoCd)"
    End If
  'for Repair
  ElseIf UCase(adc.MenuCd) = UCase("RprWIP") Then
    wQty = " BRepQty as qQty"
    'wQty = " OdOrdQty as qQty"
    wPtQty = " 0 as qPtQty"
    wSelCnd = " where BLoc= 'PFG' "
    If adc("UoBagPcs") = "N" Then Rep.TxtHead.SetText "WIP Bag Pieces Report (Normal)"
    If adc("UoBagPcs") = "R" Then Rep.TxtHead.SetText "WIP Bag Pieces Report (Rejected)"
    Rep.TxtLocFr.SetText adc("UoIssLocFr")
    Rep.TxtLocTo.SetText adc("UoIssLocTo")
    Rep.TxtLocSel.SetText adc("UoIssLocSel")
        
    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                    IIF(adc("UoBagTyp") <> "", " and PValue= '" + adc("UoBagTyp") + "' ", " and PValue <> 'D'") + _
                    IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
      
    ws_FgJoin = " "
    ws_TdYyKeyStr = " (Select Max(TdYyKey) From Txnd where TdCoCd=BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo and TdSrNo=0 and TdTc='BV') "
    
  End If
  
  

  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(adc("UOCMCDFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(adc("UOCMCDTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(adc("UOCMCDSEL")) + ")", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields

' **** Zubin 211 ( ws_BagJoin added) **** '
  '/** urmi BDlDays 2.06 */
  '*** Jay 3.2.0 [PicNm]
'             "datediff(dd, BRecvDt, '" + CStr(moCn.SrvrDate) + "') as BDlDays, " + _
             wQty + "," + wFrVal + "," + wPtQty + ", OmCmCd, " + _
             "OmPoNo, IsNull((Case when TdTc='BV' and BLoc='PFG' Then TdToBLoc Else TdFrBLoc End),BLoc) as TdFrBLoc From Bag " + ws_BagJoin + ws_FgJoin + _
             "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= 'DM' and DmCd= BOdDmCd and DmSz= ''  " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
             "Join Loc Lc on LocCocd= BCoCd and BLoc= LocCd " + _
  '*********Bhavna added DmSrcDsgCd scope and fld
  
  'Query for repair
  If UCase(adc.MenuCd) = UCase("RprWIP") Then
  wSqlStrg = "Select " + grpflds + ", " + _
             "BCoCd, BYy, BChr, BNo, BRepOdTc as BOdTc,BRepOdYy as BOdYy ,BRepOdChr as BOdChr,BRepOdNo as BOdNo ,BRepOdSr as BOdSr," + _
             "DmTcTyp, DmSrcDsgCd, BOdDmCd, OdDmSz, DmCtg, OdKt, OdDmCol, OdPicNm, BRepLoc as BLoc, " + _
             IIF(OptLoc(0).Value = True, "LocPrntCd", "(Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs)") + " as LocPrntCd, BWrk, BRepGrWt as BGrWt, " + _
             "datediff(dd, BRecvDt, '" + CStr(moCn.SrvrDate) + "') as BDlDays, " + _
             wQty + "," + wFrVal + "," + wPtQty + ", OmCmCd, " + _
             "OmPoNo, LocPrcs From Bag " + ws_BagJoin + ws_FgJoin + _
             "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= 'DM' and DmCd= BOdDmCd and DmSz= ''  " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
             "Join Loc Lc on LocCocd= BCoCd and BRepLoc= LocCd " + _
             "Join OrdDsg on OdCoCd= BCoCd and OdTc= BRepOdTc and OdYy= BRepOdYy " + _
             "and OdChr= BRepOdChr and OdNo= BRepOdNo and OdSr= BRepOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
             "Join OrdMst on OmCoCd= BCoCd and OmTc= BRepOdTc and OmYy= BRepOdYy " + _
             "and OmChr= BRepOdChr and OmNo= BRepOdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=BPrtKey ", "") + _
             wSelCnd + IIF(adc("UoYN") = "Y", " and BRepYN= 'Y' ", IIF(adc("UoYN") = "N", " and BRepYN= 'N' ", " and BRepOdTc <>'' ")) + _
             wBagDelayCnd + _
             wCnd + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", "") + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
             IIF(ws_FGINVChk <> "", " And " + ws_FGINVChk, "")
  Else
  wSqlStrg = "Select " + grpflds + ", " + _
             "BCoCd, BYy, BChr, BNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, DmTcTyp, DmSrcDsgCd, BOdDmCd, OdDmSz, DmCtg, " + _
             "OdKt, OdDmCol, OdPicNm," + IIF(OptLoc(0).Value = True, "BLoc", "LocPrcs") + " as BLoc, " + _
             IIF(OptLoc(0).Value = True, "LocPrntCd", "(Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs)") + " as LocPrntCd, BWrk, BGrWt, " + _
             "datediff(dd, BRecvDt, '" + CStr(moCn.SrvrDate) + "') as BDlDays, " + _
             wQty + "," + wFrVal + "," + wPtQty + ", OmCmCd, " + _
             "OmPoNo, LocPrcs From Bag " + ws_BagJoin + ws_FgJoin + _
             "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= 'DM' and DmCd= BOdDmCd and DmSz= ''  " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
             "Join Loc Lc on LocCocd= BCoCd and BLoc= LocCd " + _
             "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
             "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
             "Join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy " + _
             "and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=BPrtKey ", "") + _
             wSelCnd + wBagClsCnd + wBagDelayCnd + _
             wCnd + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", "") + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
             IIF(ws_FGINVChk <> "", " And " + ws_FGINVChk, "")
  End If
  
'             "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
             "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
             "Join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy " + _
             "and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=BPrtKey ", "") + _
             " Left Outer Join Txnd  On TdCoCd=BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo and TdSrNo=0 and  " + _
             " TdYyKey = " + ws_TdYyKeyStr + _
             wSelCnd + wBagClsCnd + wBagDelayCnd + _
             wCnd + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", "") + IIF(wAddlCnd <> "", " And " + wAddlCnd, "")


'/************************/

  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  '***Geeta**Emr208 **Picture
  If adc("UoYN1") = "N" Or wDetPos = 0 Then
    Rep.PICSEC.Suppress = True
  Else
    Rep.PICSEC.Suppress = False
  End If
  
  '***(Jen 2.11)
  '*** Jay 2.13(CT) ***
  'Uma *** new 'D' btyp
  If adc("UoBagTyp") = "F" Then
    Rep.TxtBagTypAndChr.SetText "(Flute Bags" & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
  ElseIf adc("UoBagTyp") = "P" Then
    Rep.TxtBagTypAndChr.SetText "(Primary Bags" & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
  ElseIf adc("UoBagTyp") = "C" Then
    Rep.TxtBagTypAndChr.SetText "(Component Bags" & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
  ElseIf adc("UoBagTyp") = "D" Then
    Rep.TxtBagTypAndChr.SetText "(PD Bags" & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
  End If
  'If adc("UoBagTyp") <> "" Then Rep.TxtBagTypAndChr.SetText IIF(adc("UoBagTyp") = "F", "(Flute Bags", "(Primary Bags") & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
  '*** Jay 2.13(CT) ***
  '***(Jen 2.11)

  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 1     'Tag = "ShowPic"
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  End Select
End With
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added



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
  gb_CoCdFor = False
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******


  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    If ws_MultiCoMod = True Then
      Set adc.FirNKeyCtl = adc("UoCoCdFr")
      LblPrntCd.Visible = True
      adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
      adc("UoPrntCdFr").Visible = True: adc("UoPrntCdTo").Visible = True: adc("UoPrntCdSel").Visible = True
    Else
      LblPrntCd.Visible = False
      Set adc.FirNKeyCtl = adc("UoIssLocFr")
      adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
      adc("UoPrntCdFr").Visible = False: adc("UoPrntCdTo").Visible = False: adc("UoPrntCdSel").Visible = False
    End If

  '  Set adc.FirNKeyCtl = adc("UoCoCdFr")
  If UCase(adc.MenuCd) = UCase("RepWipBPcs") Or UCase(adc.MenuCd) = UCase("RprWIP") Then
    
    'adc("UoIssLocFr").CmpStr = "BLoc >= "
    'adc("UoIssLocTo").CmpStr = "BLoc <= "
    'adc("UoIssLocSel").CmpStr = "BLoc In "
    LblProcess.Visible = True
    
    FraLoc.Visible = True: FraLoc.Enabled = True
    
    adc("UoIssLocFr").Visible = True: adc("UoIssLocFr").Enabled = True
    adc("UoIssLocTo").Visible = True: adc("UoIssLocTo").Enabled = True
    adc("UoIssLocSel").Visible = True: adc("UoIssLocSel").Enabled = True
    adc("UoFgSubLocFr").CmpStr = ""
    adc("UoFgSubLocTo").CmpStr = ""
    adc("UoFgSubLocSel").CmpStr = ""
    lblFgSubLoc.Visible = False
    adc("UoFgSubLocFr").Visible = False: adc("UoFgSubLocFr").Enabled = False
    adc("UoFgSubLocTo").Visible = False: adc("UoFgSubLocTo").Enabled = False
    adc("UoFgSubLocSel").Visible = False: adc("UoFgSubLocSel").Enabled = False
    '**********Bhavna src dsg cd for Fg bag pcs
    adc("UoBOEFr").CmpStr = ""
    adc("UoBOETo").CmpStr = ""
    LblSrcDsgCd.Visible = False
    adc("UoBOEFr").Visible = False: adc("UoBOEFr").Enabled = False
    adc("UoBOETo").Visible = False: adc("UoBOETo").Enabled = False
    '**********Bhavna src dsg cd for Fg bag pcs
    
    '*** (2.11) Jen
    LblPckLst.Visible = False
    adc("UoFdPlYyFr").Visible = False: adc("UoFdPlChrFr").Visible = False: adc("UoFdPlNoFr").Visible = False
    adc("UoFdPlYyTo").Visible = False: adc("UoFdPlChrTo").Visible = False: adc("UoFdPlNoTo").Visible = False
    LblPkLstDeLimit1.Visible = False: LblPkLstDeLimit2.Visible = False: LblPkLstDeLimit3.Visible = False: LblPkLstDeLimit4.Visible = False
    adc("UoFdPlYyFr").CmpStr = "": adc("UoFdPlChrFr").CmpStr = "": adc("UoFdPlNoFr").CmpStr = ""
    adc("UoFdPlYyTo").CmpStr = "": adc("UoFdPlChrTo").CmpStr = "": adc("UoFdPlNoTo").CmpStr = ""
    '*** (2.11) Jen

    ' **** Zubin 211 **** '
    '    adc("UoFdPlYyFr").CmpStr = "": adc("UoFdPlYyTo").CmpStr = ""
    '    adc("UoFdPlChrFr").CmpStr = "": adc("UoFdPlChrTo").CmpStr = ""
    '    adc("UoFdPlNoFr").CmpStr = "": adc("UoFdPlNoTo").CmpStr = ""
    '    LblPL.Visible = False
    '    adc("UoFdPlYyFr").Visible = False: adc("UoFdPlYyTo").Visible = False
    '    adc("UoFdPlChrFr").Visible = False: adc("UoFdPlChrTo").Visible = False
    '    adc("UoFdPlNoFr").Visible = False: adc("UoFdPlNoTo").Visible = False
    '    LblPlFrSlsh1.Visible = False: LblPlFrSlsh2.Visible = False
    '    LblPlFrSlsh3.Visible = False: LblPlFrSlsh4.Visible = False
    ' **** Zubin 211 **** '
    
    '    adc("UoFdPlYyFr").CmpStr = ""
    '    adc("UoFdPlYyTo").CmpStr = ""
    '    adc("UoFdPlChrFr").CmpStr = ""
    '    adc("UoFdPlChrTo").CmpStr = ""
    '    adc("UoFdPlNoFr").CmpStr = ""
    '    adc("UoFdPlNoTo").CmpStr = ""
    '    LblPckLst.Visible = False: LblPlSlash1.Visible = False: LblPlSlash2.Visible = False
    '    LblPlSlash3.Visible = False: LblPlSlash4.Visible = False
    '    adc("UoFdPlYyFr").Visible = False: adc("UoFdPlYyFr").Enabled = False
    '    adc("UoFdPlYyTo").Visible = False: adc("UoFdPlYyTo").Enabled = False
    '    adc("UoFdPlChrFr").Visible = False: adc("UoFdPlChrFr").Enabled = False
    '    adc("UoFdPlChrTo").Visible = False: adc("UoFdPlChrTo").Enabled = False
    '    adc("UoFdPlNoFr").Visible = False: adc("UoFdPlNoFr").Enabled = False
    '    adc("UoFdPlNoTo").Visible = False: adc("UoFdPlNoTo").Enabled = False
    
'    LblBagPcs.Visible = True
'    adc("UoBagPcs").Visible = True: adc("UoBagPcs").Enabled = True
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
    FraOptBagPcs.Enabled = True: FraOptBagPcs.Visible = True
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
    FraOptBagIn.Enabled = False: FraOptBagIn.Visible = False
    OptBagIn(0).Value = False: OptBagIn(1).Value = False: OptBagIn(2).Value = False
    LblBagIn.Visible = False
    ' ****** sachin 2.12 - 24/11/05 ****** '
      adc("UoPmCdFr").Visible = False
      TxtPwd.Visible = False
      lblseopwd.Visible = False
      LblValTyp.Visible = False
    ' ****** sachin 2.12 - 24/11/05 ****** '
    ' Zubin 213
    adc("UoCurCdFr").Visible = False
    LblCurrCd.Visible = False
    ' Zubin 213
    
  ElseIf UCase(adc.MenuCd) = UCase("RepWipFgBPcs") Then
    ' **** Zubin 211 **** '
    '''''    Set adc.FirNKeyCtl = adc("UoWrkFr")
    'Set adc.FirNKeyCtl = adc("UoFgSubLocFr")
    
    ' **** Zubin 211 **** '
    adc("UoIssLocFr").CmpStr = ""
    adc("UoIssLocTo").CmpStr = ""
    adc("UoIssLocSel").CmpStr = ""
    LblProcess.Visible = False
    
    FraLoc.Visible = False: FraLoc.Enabled = False
    adc("UoIssLocFr").Visible = False: adc("UoIssLocFr").Enabled = False
    adc("UoIssLocTo").Visible = False: adc("UoIssLocTo").Enabled = False
    adc("UoIssLocSel").Visible = False: adc("UoIssLocSel").Enabled = False
    lblFgSubLoc.Visible = True
    adc("UoFgSubLocFr").Visible = True: adc("UoFgSubLocFr").Enabled = True
    adc("UoFgSubLocTo").Visible = True: adc("UoFgSubLocTo").Enabled = True
    adc("UoFgSubLocSel").Visible = True: adc("UoFgSubLocSel").Enabled = True

    '*** (2.11) Jen
    LblPckLst.Visible = True
    adc("UoFdPlYyFr").Visible = True: adc("UoFdPlChrFr").Visible = True: adc("UoFdPlNoFr").Visible = True
    adc("UoFdPlYyTo").Visible = True: adc("UoFdPlChrTo").Visible = True: adc("UoFdPlNoTo").Visible = True
    LblPkLstDeLimit1.Visible = True: LblPkLstDeLimit2.Visible = True: LblPkLstDeLimit3.Visible = True: LblPkLstDeLimit4.Visible = True
    
    adc("UoFdPlYyFr").CmpStr = "FdPlYy >=": adc("UoFdPlChrFr").CmpStr = "FdPlChr >=": adc("UoFdPlNoFr").CmpStr = "FdPlNo >="
    adc("UoFdPlYyTo").CmpStr = "FdPlYy <=": adc("UoFdPlChrTo").CmpStr = "FdPlChr <=": adc("UoFdPlNoTo").CmpStr = "FdPlNo <="
    adc("UoPrntCdFr").CmpStr = "LocPrntCd <=": adc("UoPrntCdTo").CmpStr = "LocPrntCd <=": adc("UoPrntCdSel").CmpStr = "LocPrntCd In "
    '*** (2.11) Jen
    
    '**********Bhavna src dsg cd for Fg bag pcs
    adc("UoBOEFr").CmpStr = "DmSrcDsgCd >= "
    adc("UoBOETo").CmpStr = "DmSrcDsgCd <= "
    LblSrcDsgCd.Visible = True
    adc("UoBOEFr").Visible = True: adc("UoBOEFr").Enabled = True
    adc("UoBOETo").Visible = True: adc("UoBOETo").Enabled = True
    '**********Bhavna src dsg cd for Fg bag pcs
    

' **** Zubin 211 **** '
 '   adc("UoFdPlYyFr").CmpStr = "FdPlYy >=": adc("UoFdPlYyTo").CmpStr = "FdPlYy <="
 '   adc("UoFdPlChrFr").CmpStr = "FdPlChr >=": adc("UoFdPlChrTo").CmpStr = "FdPlChr <="
 '   adc("UoFdPlNoFr").CmpStr = "FdPlNo >=": adc("UoFdPlNoTo").CmpStr = "FdPlNo <="
' **** Zubin 211 **** '

'    adc("UoFdPlYyFr").CmpStr = "FdPlYy >="
'    adc("UoFdPlYyTo").CmpStr = "FdPlYy <="
'    adc("UoFdPlChrFr").CmpStr = "FdPlChr >="
'    adc("UoFdPlChrTo").CmpStr = "FdPlChr <="
'    adc("UoFdPlNoFr").CmpStr = "FdPlNo >="
'    adc("UoFdPlNoTo").CmpStr = "FdPlNo <="
'    LblPckLst.Visible = True: LblPlSlash1.Visible = True: LblPlSlash2.Visible = True
'    LblPlSlash3.Visible = True: LblPlSlash4.Visible = True
'    adc("UoFdPlYyFr").Visible = True: adc("UoFdPlYyFr").Enabled = True
'    adc("UoFdPlYyTo").Visible = True: adc("UoFdPlYyTo").Enabled = True
'    adc("UoFdPlChrFr").Visible = True: adc("UoFdPlChrFr").Enabled = True
'    adc("UoFdPlChrTo").Visible = True: adc("UoFdPlChrTo").Enabled = True
'    adc("UoFdPlNoFr").Visible = True: adc("UoFdPlNoFr").Enabled = True
'    adc("UoFdPlNoTo").Visible = True: adc("UoFdPlNoTo").Enabled = True
    LblBagPcs.Visible = False
    adc("UoBagPcs").Visible = False: adc("UoBagPcs").Enabled = False
    adc("UoBagPcs") = ""
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
    FraOptBagPcs.Enabled = False: FraOptBagPcs.Visible = False
    OptBagPcs(0).Value = False: OptBagPcs(1).Value = False
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
    
    FraOptBagIn.Enabled = True: FraOptBagIn.Visible = True
    LblBagIn.Visible = True
    
    LblBagTyp.Visible = False
    adc("UoBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    adc("UoBChrFr").Visible = False
    
    '****** Sachin 2.14.0 - [Multi-Company]
    LblPrntCd.Visible = False
    'adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
    adc("UoPrntCdFr").Visible = False: adc("UoPrntCdTo").Visible = False: adc("UoPrntCdSel").Visible = False
    adc("UoPrntCdFr").CmpStr = "": adc("UoPrntCdTo").CmpStr = "": adc("UoPrntCdSel").CmpStr = ""
    '****** Sachin 2.14.0 - [Multi-Company]
    
  End If
  
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_LocTyp = "'P'"
  'gb_RmCtgFor = True
  'gb_RmZ = False
  gs_DmTcTyp = "DM"
  'gs_TxnTcTyp = "BV"
  gs_OmTcTyp = "SO"
  gs_CmCtg = "C"
  
  ' ****** sachin 2.12 - 24/11/05 ****** '
    gs_PTyp = "ORDVAL"
    gs_Tbl = "Param"
  ' ****** sachin 2.12 - 24/11/05 ****** '
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  Call SetGroupSort
End Sub

Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
    
 ' ****** sachin 2.12 - 24/11/05 ****** '
    gs_PTyp = "ORDVAL"
    gs_Tbl = "Param"
    gb_CoCdFor = False '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
    
    Select Case UCase(IdName)
    Case Is = UCase("UoPmCdFr")
    '*** Jay 2.14 *** (SEO Change)
      If pv_NewValue <> "" Then
          If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
              "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
              "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
              "(Select PValue1 from Param where " + _
              "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ))") Then _
              Cancel = True: ErrMsg = "Enter SEO Password to View Values ": Exit Sub
      End If
      
    '*** Jay 2.14 *** (SEO Change)

    ' Bef 2.14
    '  If pv_NewValue <> "" And Trim(TxtPwd.text) <> moCn.GetFldVal("Select PValue1 from Param where " + _
                                   "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
    '  Cancel = True: DispMsg "Enter SEO Password to View Values", etError
    'End If
    ' Bef 2.14
    '*** Jay 2.14 *** (SEO Change)
    
                
        End Select
  ' ****** sachin 2.12 - 24/11/05 ****** '
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
 
    ' ****** sachin 2.12 - 24/11/05 ****** '
    gs_PTyp = "ORDVAL"
    gs_Tbl = "Param"
    gb_CoCdFor = False
  ' ****** sachin 2.12 - 24/11/05 ****** '
  
  '*** (Jen 2.11)
  Select Case UCase(IdName)
    Case Is = UCase("UoBChrFr")
      If adc.Mode = xNorm Then
        If adc("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Bag Type.": Exit Sub
        HlpList.PMCd "BCHR", "'" + adc("UoBagTyp") + "'": Exit Sub
      End If
            
    Case Is = UCase("UOISSLOCFR"), UCase("UOISSLOCTO"), UCase("UOISSLOCSEL"), UCase("UOPRNTCDFR"), UCase("UOPRNTCDTO"), UCase("UOPRNTCDSEL")
        If adc("UOLOCTYPFR") = "R" Then gs_LocTyp = "'R'"
        If adc("UOLOCTYPFR") = "P" Then gs_LocTyp = "'P'"

    Case Is = UCase("UOOMTCFR")
          If UCase(adc.MenuCd) = UCase("RprWIP") Then HlpList.PMCd "TC", "'JT'":  Exit Sub
    Case Is = UCase("UOOMTCTO")
          If UCase(adc.MenuCd) = UCase("RprWIP") Then HlpList.PMCd "TC", "'JT'":  Exit Sub
        
  End Select
  '*** (Jen 2.11)
  
 SetProp Me, IdName, When
   
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
' **** Zubin 211 **** '
  Select Case UCase(IdName)
    Case Is = UCase("UOBCHRFR")
      'If ADC("UoBagTyp") = "N" Then
      adc("UOBCHRFR") = ""
  End Select
' **** Zubin 211 **** '
End Sub

Private Sub GRP_REP_GotFocus()
'uuuuu
'If OptLoc(0).Value = True Then
'  GRP_REP.Add "Location", "BLoc", "", "BLoc", "hBLoc", "Company Code,Parent Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'P' and LocCd= BLoc) "
''  GRP_REP.Add "(Location)Worker", "BLoc+BWrk", "'('+ RTrim(BLoc)+ ') '+ BWrk", "BLoc,BWrk", "hBLoc,hBWrk", "Parent Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'WORK' and vPMCd= BWrk) "
'ElseIf OptLoc(1).Value = True Then
'  GRP_REP.Add "Process", "LocPrcs", "", "LocPrcs", "hBLoc", "Company Code,Parent Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and  LocCd= BLoc) "
'  GRP_REP.Add "(Process)Worker", "BLoc+BWrk", "'('+ RTrim(BLoc)+ ') '+ BWrk", "BLoc,BWrk", "hBLoc,hBWrk", "Parent Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'WORK' and vPMCd= BWrk) "
'End If
  
'GRP_REP.Add "Process", "(Select LocPrcs From Loc Where LocCoCd = BCoCd and LocTyp= 'R' and LocCd= BLoc)", "", "LocPrcs", "hBLoc", "Company Code,Parent Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'R' and LocCd= BLoc) "
End Sub

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptBag_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN") = "Y"
Case Is = 1
  adc("UoYN") = "N"
Case Is = 2
  adc("UoYN") = ""
End Select
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptBagIn_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN2") = "Y"
Case Is = 1
  adc("UoYN2") = "N"
Case Is = 2
  adc("UoYN2") = ""
End Select
End Sub

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptBag_GotFocus(Index As Integer)
  DispMsg FraOptBag.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
Private Sub OptBagPcs_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoBagPcs") = "N"
Case Is = 1
  adc("UoBagPcs") = "R"
End Select
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
Private Sub OptBagPcs_GotFocus(Index As Integer)
DispMsg FraOptBagPcs.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added

Private Sub OptLocation_Click(Index As Integer)

End Sub

Private Sub OptLoc_GotFocus(Index As Integer)
  DispMsg FraLoc.ToolTipText, etInfo
Select Case Index
Case Is = 0
  adc("UOISSLOCFR") = "":   adc("UOISSLOCTO") = "":   adc("UOISSLOCsel") = ""
  adc("UOPRNTCDFR") = "":   adc("UOPRNTCDTO") = "":   adc("UOPRNTCDSEL") = ""
  adc("UoLocTypFr") = "P"
Case Is = 1
  adc("UOISSLOCFR") = "":   adc("UOISSLOCTO") = "":   adc("UOISSLOCsel") = ""
  adc("UOPRNTCDFR") = "":   adc("UOPRNTCDTO") = "":   adc("UOPRNTCDSEL") = ""
  adc("UoLocTypFr") = "R"
End Select

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
  
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = False Then _
      adc("UoCoCdFr") = gs_CoCd: adc("UoCoCdTo") = gs_CoCd: adc("UoCoCdSel") = "": adc("UoCoCdFr").Enabled = False
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
  '**Geeta*Emr208*Default Show picture is always no
  adc("UoYN1") = "N"
  ' ****** sachin 2.12 ****** '
  adc("UoPmCdFr") = ""
  ' ****** sachin 2.12 ****** '

  '*** Jay 2.13(CT) ***
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + gs_CoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    adc("UoBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    adc("UoBChrFr").Visible = False
    adc("UoBagTyp") = "P": adc("UoBChrFr") = ""
  End If
  Set wHeadChk = Nothing
  '*** Jay 2.13(CT) ***
'*** Jay 2.13(CT) ***
'  '*** (Jen 2.11)
'  If moCn.GetFldVal("Select hFluteBagYN from Head where HCoCd= '" + adc("UoCoCdFr") + "' and " + _
'                                                    "HCd= '" + ctSelfCmCd + "' ") = "N" Then
'    LblBagTyp.Visible = False
'    adc("UoBagTyp").Visible = False
'    LblUoBChrFr.Visible = False
'    adc("UoBChrFr").Visible = False
'    adc("UoBagTyp") = "P": adc("UoBChrFr") = ""
'  End If
'  '*** (Jen 2.11)
'*** Jay 2.13(CT) ***

  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  If adc("UoYN") = "Y" Then
    OptBag(0).Value = True
  ElseIf adc("UoYN") = "N" Then
    OptBag(1).Value = True
  ElseIf adc("UoYN") = "" Then
    OptBag(2).Value = True
  End If
  
  If adc("UoYN2") = "Y" Then
    OptBagIn(0).Value = True
  ElseIf adc("UoYN2") = "N" Then
    OptBagIn(1).Value = True
  ElseIf adc("UoYN2") = "" Then
    OptBagIn(2).Value = True
  End If
  
  If UCase(adc.MenuCd) = UCase("RepWipBPcs") Then
    If adc("UoBagPcs") = "N" Then OptBagPcs(0).Value = True
    If adc("UoBagPcs") = "R" Then OptBagPcs(1).Value = True
  End If
  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  
  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  If adc("UOLOCTYPFR") = "P" Then
    OptLoc(0).Value = True
  ElseIf adc("UOLOCTYPFR") = "R" Then
    OptLoc(1).Value = True
  End If

End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = False Or UCase(adc.MenuCd) = UCase("RepWipFgBPcs"), adc("UOCOCDFR"), gs_CoCd) + "' and " + _
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
  
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  Rep.TxtCoCdFr.SetText adc("UOCOCDFR")
  If ws_MultiCoMod = False Or UCase(adc.MenuCd) = UCase("RepWipFgBPcs") Then
      Rep.TxtCoCdTo.Suppress = True: Rep.TxtCoCdSel.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True: Rep.TxtCoCdSelLbl.Suppress = True
  Else
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdTo.SetText adc("UOCOCDTO"): Rep.TxtCoCdSel.SetText adc("UOCOCDSEL")
  End If
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
' ********* Manali 3.2.2 - Before 3.2.2
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  Rep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  Rep.wSrvrTime.SetText ws_DtTm

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

Private Sub TxtPwd_Validate(Cancel As Boolean)
'*** Jay 2.14 *** (SEO Change)
  Dim ws_VldPwd As String
  ws_VldPwd = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + adc.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ")
  If Trim(TxtPwd.text) = "" And ws_VldPwd <> "" Then adc("UoPMCdFr") = "": Exit Sub
  If Trim(TxtPwd.text) <> "" And Trim(TxtPwd.text) <> ws_VldPwd Then
      Cancel = True:  DispMsg "Invalid Password / User Does Not Have Seo Rights ", etError
  End If
' Bef 2.14
' ****** sachin 2.12 ****** '
'    If Trim(TxtPwd.text) = "" Then adc("UoPmCDFr") = ""
'    If Trim(TxtPwd.text) = "" Or Trim(TxtPwd.text) = moCn.GetFldVal("Select PValue1 from Param where " + _
'                           "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'
'      Exit Sub
'    Else
'      Cancel = True
'      DispMsg "Invalid Password", etError
'    End If
' ****** sachin 2.12 ****** '
' Bef 2.14
'*** Jay 2.14 *** (SEO Change)
End Sub
