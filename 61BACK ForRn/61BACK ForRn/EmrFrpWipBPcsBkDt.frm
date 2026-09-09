VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrpWipBPcsBkDt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WIP Bag Pieces Report (As Of Date)"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   60
      TabIndex        =   72
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11190
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   9780
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4455
      TabIndex        =   70
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
      Left            =   5145
      TabIndex        =   80
      Top             =   9690
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   81
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
         TabIndex        =   82
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
      Left            =   6360
      TabIndex        =   79
      Top             =   9705
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   76
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
      TabIndex        =   75
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   74
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpWipBPcsBkDt.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpWipBPcsBkDt.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   78
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   73
               Top             =   360
               Width           =   9045
               _ExtentX        =   15954
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Caption         =   "Customer"
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   77
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraLoc 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1500
               TabIndex        =   121
               Top             =   1000
               Width           =   4305
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
                  Left            =   0
                  TabIndex        =   123
                  ToolTipText     =   "Select Production or Process"
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   2085
               End
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
                  Left            =   2940
                  TabIndex        =   122
                  ToolTipText     =   "Select Production or Process"
                  Top             =   0
                  Width           =   1545
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
               TabIndex        =   60
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   5850
               Width           =   1680
            End
            Begin VB.Frame FraOptBagPcs 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1515
               TabIndex        =   116
               ToolTipText     =   "Show Normal Or Rejected Bag Pieces"
               Top             =   6420
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
                  TabIndex        =   66
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
                  TabIndex        =   67
                  ToolTipText     =   "Show Normal Or Rejected Bag Pieces"
                  Top             =   0
                  Width           =   1185
               End
            End
            Begin VB.Frame FraOptBag 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1515
               TabIndex        =   115
               ToolTipText     =   "Show Open Or Closed Or All Bags"
               Top             =   6135
               Width           =   3705
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
                  TabIndex        =   62
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
                  TabIndex        =   63
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
                  TabIndex        =   64
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4455
               TabIndex        =   7
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   1290
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
               Left            =   7380
               TabIndex        =   8
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1290
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1515
               TabIndex        =   6
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   1290
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4455
               TabIndex        =   16
               ToolTipText     =   "Enter To Design Category"
               Top             =   2145
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
               Left            =   7380
               TabIndex        =   17
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2145
               Width           =   7590
               _ExtentX        =   13388
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
               Left            =   1515
               TabIndex        =   15
               ToolTipText     =   "Enter From Design Category"
               Top             =   2145
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
               Left            =   4455
               TabIndex        =   19
               ToolTipText     =   "Enter To Design Code"
               Top             =   2430
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
               Left            =   7380
               TabIndex        =   20
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2430
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "BOdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1515
               TabIndex        =   18
               ToolTipText     =   "Enter From Design Code"
               Top             =   2430
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
               Left            =   4455
               TabIndex        =   35
               ToolTipText     =   "Enter To Customer"
               Top             =   3570
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
               Index           =   41
               Left            =   7380
               TabIndex        =   36
               ToolTipText     =   "Enter Customer Selection"
               Top             =   3570
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1515
               TabIndex        =   34
               ToolTipText     =   "Enter From Customer"
               Top             =   3570
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
               Index           =   5
               Left            =   5235
               TabIndex        =   28
               ToolTipText     =   "Enter To Order Year"
               Top             =   3000
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
               Index           =   17
               Left            =   4455
               TabIndex        =   27
               ToolTipText     =   "Enter To Order Tc"
               Top             =   3000
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
               Left            =   5655
               TabIndex        =   29
               ToolTipText     =   "Enter To Order Character"
               Top             =   3000
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
               Left            =   1515
               TabIndex        =   30
               ToolTipText     =   "Enter From Order Number"
               Top             =   3285
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
               Left            =   4455
               TabIndex        =   32
               ToolTipText     =   "Enter To Order Number"
               Top             =   3285
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
               Index           =   7
               Left            =   4455
               TabIndex        =   13
               ToolTipText     =   "Enter To By Worker"
               Top             =   1860
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKTO"
               IdName          =   "UOWRKTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   7380
               TabIndex        =   14
               ToolTipText     =   "Enter By Worker Selection"
               Top             =   1860
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOWRKSEL"
               IdName          =   "UOWRKSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1515
               TabIndex        =   12
               ToolTipText     =   "Enter From By Worker"
               Top             =   1860
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKFR"
               IdName          =   "UOWRKFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   4455
               TabIndex        =   22
               ToolTipText     =   "Enter To Karat"
               Top             =   2715
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
               Left            =   7380
               TabIndex        =   23
               ToolTipText     =   "Enter Karat Selection"
               Top             =   2715
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "OdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1515
               TabIndex        =   21
               ToolTipText     =   "Enter From Karat"
               Top             =   2715
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
               Index           =   43
               Left            =   4455
               TabIndex        =   38
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   3855
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
               Left            =   7380
               TabIndex        =   39
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   3855
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   1515
               TabIndex        =   37
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   3855
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
               Index           =   21
               Left            =   5340
               TabIndex        =   33
               ToolTipText     =   "Enter To Order Serial"
               Top             =   3285
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
               Left            =   2400
               TabIndex        =   31
               ToolTipText     =   "Enter From Order Serial"
               Top             =   3285
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
               Left            =   1515
               TabIndex        =   58
               ToolTipText     =   "Enter Pending As Of Date"
               Top             =   5565
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOASOFDT"
               IdName          =   "UOASOFDT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   5400
               TabIndex        =   65
               ToolTipText     =   "Show Normal Or Rejected Pieces ?"
               Top             =   6420
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
               Index           =   6
               Left            =   5400
               TabIndex        =   61
               ToolTipText     =   "Show Open Or Closed Or All Bags ? (Open - 'Y', Closed - 'N', All - Blank)"
               Top             =   6135
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
               Index           =   12
               Left            =   1515
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   400
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
               Index           =   20
               Left            =   2295
               TabIndex        =   25
               ToolTipText     =   "Enter From Order Year"
               Top             =   3000
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
               Left            =   1515
               TabIndex        =   24
               ToolTipText     =   "Enter From Order Tc"
               Top             =   3000
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
               Left            =   2715
               TabIndex        =   26
               ToolTipText     =   "Enter From Order Character"
               Top             =   3000
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
               Index           =   27
               Left            =   2055
               TabIndex        =   59
               ToolTipText     =   "Show Picture? ( Enter 'Y'/'N')"
               Top             =   5850
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
               Left            =   1515
               TabIndex        =   68
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   6705
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
               Left            =   3045
               TabIndex        =   69
               ToolTipText     =   "Enter Bag Character"
               Top             =   6705
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRFR"
               ReCalcOn        =   "UOBAGTYP"
               IdName          =   "UOBCHRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1515
               TabIndex        =   52
               ToolTipText     =   "Enter From Packing List Year"
               Top             =   5280
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
               Index           =   37
               Left            =   1935
               TabIndex        =   53
               ToolTipText     =   "Enter From Packing List Character"
               Top             =   5280
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
               Left            =   2715
               TabIndex        =   54
               ToolTipText     =   "Enter From Packing List Number"
               Top             =   5280
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
               Left            =   4455
               TabIndex        =   55
               ToolTipText     =   "Enter To Packing List Year"
               Top             =   5280
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
               Left            =   4875
               TabIndex        =   56
               ToolTipText     =   "Enter To Packing List Character"
               Top             =   5280
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
               Left            =   5655
               TabIndex        =   57
               ToolTipText     =   "Enter To Packing List Number"
               Top             =   5280
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
               Index           =   47
               Left            =   4455
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   405
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
               Left            =   7380
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   405
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "BCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   4455
               TabIndex        =   4
               ToolTipText     =   "Enter To Parent Code"
               Top             =   705
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDTO"
               IdName          =   "UOPRNTCDTO"
               CmpStr          =   "LocPrntCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   7380
               TabIndex        =   5
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   705
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
               CmpStr          =   "LocPrntCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   1515
               TabIndex        =   3
               ToolTipText     =   "Enter From Parent Code"
               Top             =   705
               Width           =   1720
               _ExtentX        =   3043
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDFR"
               IdName          =   "UOPRNTCDFR"
               CmpStr          =   "LocPrntCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2810
               TabIndex        =   41
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   4140
               Width           =   1600
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   5750
               TabIndex        =   43
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   4140
               Width           =   1600
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   7380
               TabIndex        =   44
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   4140
               Visible         =   0   'False
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "OdPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   4455
               TabIndex        =   10
               ToolTipText     =   "Enter To Fg Sub Location"
               Top             =   1575
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOFGSUBLOCTO"
               IdName          =   "UOFGSUBLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   7380
               TabIndex        =   11
               ToolTipText     =   "Enter Fg Sub Location Selection"
               Top             =   1575
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOFGSUBLOCSEL"
               IdName          =   "UOFGSUBLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   1515
               TabIndex        =   9
               ToolTipText     =   "Enter From Fg Sub Location"
               Top             =   1575
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOFGSUBLOCFR"
               IdName          =   "UOFGSUBLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   4455
               TabIndex        =   49
               ToolTipText     =   "Enter To Order Date"
               Top             =   4710
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
               Left            =   1515
               TabIndex        =   48
               ToolTipText     =   "Enter From Order Date"
               Top             =   4710
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
               Left            =   4455
               TabIndex        =   51
               ToolTipText     =   "Enter To Production Delivery Date"
               Top             =   4995
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
               Left            =   1515
               TabIndex        =   50
               ToolTipText     =   "Enter From Production Delivery Date"
               Top             =   4995
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
               Left            =   4455
               TabIndex        =   46
               ToolTipText     =   "Enter To Priority Code"
               Top             =   4425
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
               Left            =   7380
               TabIndex        =   47
               ToolTipText     =   "Enter Priority Code Selection"
               Top             =   4425
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRTCDSEL"
               IdName          =   "UOPRTCDSEL"
               CmpStr          =   "OdPrtCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   64
               Left            =   1515
               TabIndex        =   45
               ToolTipText     =   "Enter From Priority Code"
               Top             =   4425
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
               Index           =   55
               Left            =   6960
               TabIndex        =   124
               ToolTipText     =   "Show Production or Process Location"
               Top             =   960
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOLOCTYPFR"
               IdName          =   "UOLOCTYPFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   1515
               TabIndex        =   40
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   4140
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTFR"
               IdName          =   "UOODSUBCUSTFR"
               CmpStr          =   "OdSubCust >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   76
               Left            =   4455
               TabIndex        =   42
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   4140
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTTO"
               IdName          =   "UOODSUBCUSTTO"
               CmpStr          =   "OdSubCust <="
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
               Left            =   5680
               TabIndex        =   126
               Top             =   4140
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
               Index           =   9
               Left            =   2740
               TabIndex        =   125
               Top             =   4140
               Width           =   105
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
               TabIndex        =   120
               Top             =   4440
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
               Left            =   30
               TabIndex        =   119
               Top             =   4710
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
               TabIndex        =   118
               Top             =   4995
               Width           =   1425
            End
            Begin VB.Label LblFgSubLoc 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "FgSubLoc"
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
               Top             =   1575
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
               TabIndex        =   114
               Top             =   4140
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
               TabIndex        =   113
               Top             =   700
               Width           =   1425
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
               Left            =   1875
               TabIndex        =   112
               Top             =   5280
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
               Left            =   2655
               TabIndex        =   111
               Top             =   5280
               Width           =   105
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
               TabIndex        =   110
               Top             =   5280
               Width           =   1485
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
               Left            =   4815
               TabIndex        =   109
               Top             =   5280
               Width           =   105
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
               Left            =   5595
               TabIndex        =   108
               Top             =   5280
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
               Left            =   2190
               TabIndex        =   107
               Top             =   6705
               Width           =   795
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
               Left            =   60
               TabIndex        =   106
               ToolTipText     =   "Location"
               Top             =   6705
               Width           =   1395
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
               Left            =   60
               TabIndex        =   105
               ToolTipText     =   "Location"
               Top             =   5850
               Visible         =   0   'False
               Width           =   1395
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
               Left            =   2235
               TabIndex        =   104
               Top             =   3000
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
               Left            =   2655
               TabIndex        =   103
               Top             =   3000
               Width           =   105
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
               TabIndex        =   102
               Top             =   3000
               Width           =   1725
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
               TabIndex        =   101
               Top             =   400
               Width           =   1335
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
               TabIndex        =   100
               Top             =   6420
               Width           =   1425
            End
            Begin VB.Label lblOpnBag 
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
               Left            =   60
               TabIndex        =   99
               Top             =   6135
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "As Of Date"
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
               TabIndex        =   98
               Top             =   5565
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
               Index           =   31
               Left            =   2340
               TabIndex        =   97
               Top             =   3285
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
               Index           =   30
               Left            =   5280
               TabIndex        =   96
               Top             =   3285
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   95
               Top             =   3855
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   94
               Top             =   2715
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   93
               Top             =   1860
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   92
               Top             =   3285
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
               Left            =   5595
               TabIndex        =   91
               Top             =   3000
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
               Left            =   5175
               TabIndex        =   90
               Top             =   3000
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   89
               Top             =   3570
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   88
               Top             =   2430
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   87
               Top             =   2145
               Width           =   1425
            End
            Begin VB.Label LblProcess 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   86
               Top             =   1290
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               Left            =   7380
               TabIndex        =   85
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               Left            =   4455
               TabIndex        =   84
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   83
               Top             =   120
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpWipBPcsBkDt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepWipBPcs
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_SqlStr As String
Dim ms_AsOfDt As String   ' ***** Manali 3.4.1 - 02/12/08

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  GRP_REP.Add "Location", Array("ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN T1.TdCoCd+T1.TdFrBLoc ELSE T1.TdCoCd+T1.TdToBLoc END, BCoCd+BOpnLoc)", "FgCoCd+FgToBLoc"), Array("ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN '['+T1.TdCoCd+'] '+T1.TdFrBLoc ELSE '['+T1.TdCoCd+'] '+T1.TdToBLoc END, '['+BCoCd+'] '+BOpnLoc)", "'['+FgCoCd+'] '+FgToBLoc"), "BLoc", "hBLoc", "Company Code,Parent Code", "", "", Array("(Select LocDesc From Loc Where LocTyp= 'P' And LocCoCd = BCoCd And LocCd= ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN T1.TdFrBLoc ELSE T1.TdToBLoc END, BOpnLoc)) ", "(Select LocDesc From Loc Where LocTyp= 'P' And LocCoCd = BCoCd And LocCd= FgToBLoc) ")
  If UCase(ADC.MenuCd) <> UCase("RepWipFgBPcsBkDt") Then
          GRP_REP.Add "Process", Array("T1.TdCoCd+LocPrcs", "FgCoCd+FgToBLoc"), Array("'['+T1.TdCoCd+']'+LocPrcs", "'['+FgCoCd+'] '+FgToBLoc"), "BLoc", "hBLoc", "Company Code,Parent Code", "", "", Array("(Select LocDesc From Loc Where LocTyp= 'R' And LocCoCd = BCoCd And LocCd= Lc.LocPrcs)", "(Select LocDesc From Loc Where LocTyp= 'R' And LocCoCd = BCoCd And LocCd= Lc.LocPrcs)")
  End If


  
  GRP_REP.Add "(Location)Worker", Array("ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN T1.TdFrBLoc ELSE T1.TdToBLoc END, BOpnLoc)+ ISNULL(T1.TdWrk, BWrk)", "FgToBLoc + TdWrk"), _
                Array("'('+ RTrim(ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN T1.TdFrBLoc ELSE T1.TdToBLoc END, BOpnLoc))+ ') '+ ISNULL(T1.TdWrk, BWrk)", "'(' + RTrim(FgToBLoc) + ') ' + TdWrk"), "BLoc,BWrk", "hBLoc,hBWrk", "Parent Code,Process", "", "", Array("(Select vPDesc From vParam Where vPCoCd = BCoCd And vPTyp= 'WORK' and vPMCd= ISNULL(T1.TdWrk, BWrk)) ", "(Select vPDesc From vParam Where vPCoCd = BCoCd And vPTyp= 'WORK' and vPMCd = TdWrk)")
  GRP_REP.Add "(Process)Worker", Array("LocPrcs+ ISNULL(T1.TdWrk, BWrk)", "FgToBLoc + TdWrk"), _
                Array("'('+ RTrim(LocPrcs)+ ') '+ ISNULL(T1.TdWrk, BWrk)", "'(' + RTrim(FgToBLoc) + ') ' + TdWrk"), "BLoc,BWrk", "hBLoc,hBWrk", "Parent Code,Process", "", "", Array("(Select vPDesc From vParam Where vPCoCd = BCoCd And vPTyp= 'WORK' and vPMCd= ISNULL(T1.TdWrk, BWrk)) ", "(Select vPDesc From vParam Where vPCoCd = BCoCd And vPTyp= 'WORK' and vPMCd = TdWrk)")

  GRP_REP.Add "Design Ctg", "DmCtg", "", "DmCtg", "hDmCtg", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
  GRP_REP.Add "Design Code", "BOdDmCd", "", "BOdDmCd", "hBOdDmCd", "Design Ctg,(DsgCd)Colour", "", "", "Substring(DmDesc,1,120) "
  GRP_REP.Add "Ord Kt", "OdKt", "", "OdKt", "hOdKt", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt) "
  GRP_REP.Add "OrderNo", "BCoCd+BOdTc+BOdYy+BOdChr+str(BOdNo)", "BCoCd+'/'+BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "OmCmCd,OmPoNo", "hOmCmCd,hOmPoNo", "Company Code,Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
  GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))"
  '*** Jay 3.2.0 *** [OdPoNo]
  'vk.19 sub customer added
  GRP_REP.Add "Cust SUB PO", "OmCmCd+OmPONo+OdSubCust+OdPoNo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)+(case when OdPONo<> '' Or OdSubCust<>'' then '/'+LTrim(OdSubCust)+' : '+LTrim(OdPONo) else '' end )", "OmCmCd,OmPoNo,OdPoNo", "hOmCmCd,hOmPONo,hOdPONo", "Customer Code,Customer PO", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))+(case when OdPONo<> '' Or OdSubCust<>'' then '/'+LTrim(OdSubCust)+' : '+LTrim(OdPONo) else '' end )"
  '*** Jay 3.2.0 *** [OdPoNo]
  GRP_REP.Add "(DsgCd)Colour", "BOdDmCd+OdDmCol", "'('+BOdDmCd+') '+OdDmCol", "BOdDmCd,OdDmCol", "hBOdDmCd,hOdDmCol", "Design Ctg,Design Code", "", "", "IsNull(Substring(DmDesc,1,120), '')+ ' ('+(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) +')'"
  GRP_REP.Add "Prd Del Date", "convert(varchar(8), OdDelDt, 112)", "convert(varchar, OdDelDt, 103)", "", "", "", "", "", ""
  GRP_REP.Add "(OrdKt)Colour", "OdKt+OdDmCol", "'('+OdKt+') '+OdDmCol", "OdKt,OdDmCol", "hOdKt,hOdDmCol", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt)+ ' ('+(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) +')'"
  GRP_REP.Add "Priority", "OdPrtCd", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'PRTCD' and PMCd= OdPrtCd) "
  
  '***** Manali 3.5.0 - 20/12/08 - following Code was commented earlier, now again included in query
  If UCase(ADC.MenuCd) = UCase("RepWipFgBPcsBkDt") Then
    GRP_REP.Add "FgSubLoc", "FgToBLoc+ISNULL(FmdToFgSubLoc, FdSubLoc)", "'('+FgToBLoc+') '+ISNULL(FmdToFgSubLoc, FdSubLoc)", "BLoc", "hBLoc", "Process", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd And vPTyp= 'FGSUBLOC' and vPMCd = ISNULL(FmdToFgSubLoc, FdSubLoc))"
  End If
  '***** Manali 3.5.0 - 20/12/08 - above Code was commented earlier, now again included in query

' **** Zubin 211 **** '
  If UCase(ADC.MenuCd) = UCase("RepWipBPcsBkDt") Then
'*** Jay 2.13(CT) ***
    GRP_REP.Add "Bag Type", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", _
      "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", "", "", "Bag Typ+Chr", "", "", "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')"
    GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
    GRP_REP.Add "Bag Typ+Chr", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+BChr", _
      "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+'/ '+ BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
  '    GRP_REP.Add "Bag Type", "(Select (case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", "(Select case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", "", "", "Bag Typ+Chr", "", "", "(Select case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')"
  '    GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
  '    GRP_REP.Add "Bag Typ+Chr", "(Select (case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+BChr", "(Select (case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+'/ '+ BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
  '*** Jay 2.13(CT) ***
  '*** 2.11 (Jen)
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ,Process,OrderNo,OrderSr,Customer PO,FgSubLoc ******
  If ws_MultiCoMod = True Then _
    GRP_REP.Add "Company Code", "BCoCd", "BCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=OdCoCd) "
    GRP_REP.Add "Parent Location", "LocPrntCd", "LocPrntCd", "", "", "", "", "", "(Select Ln.LocDesc from Loc Ln where Ln.LocCd=Lc.LocPrntCd and Ln.LocCoCd='" + ctSelfCoCd + "') "
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******
  
  ElseIf UCase(ADC.MenuCd) = UCase("RepWipFgBPcsBkDt") Then
    'GRP_REP.Add "Pck List", "FdPlYy+FdPlChr+str(FdPlNo)", "FdPlYy+'/'+FdPlChr+'/'+LTrim(Str(FdPlNo))", "", "", "", "", "", ""
    GRP_REP.Add "Pck List", "FdPlYy+FdPlChr+str(FdPlNo)", "(case when FdPlNo> 0 then FdPlYy+'/'+FdPlChr+'/'+LTrim(Str(FdPlNo)) else '' end)", "", "", "", "", "", ""
  '*** 2.11 (Jen)
  End If
' **** Zubin 211 **** '
  
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  '*** Jay 3.2.0 [PicNm]
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  BCoCd, BYy, BChr, BNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr,DmTcTyp, BOdDmCd, OdDmSz, DmCtg,
'  OdKt, OdDmCol, OdPicNm, BLoc, LocPrntCd, BWrk, BGrWt,
'  9999 as BDlDays,
'  BQty As qQty, BPtQty As qPtQty, OmCmCd, OmPoNo
'  From Bag, DsgMst, OrdDsg, OrdMst where 1=2
  '*** Report Sql ***
  
  
  Set Rep = Nothing
  Set ADC.RepSource = Rep
  
  ' ***** Manali 3.4.1 - "01/01/80" excluded from check
  Dim ws_PurgeDt As String
  'sv.9 getting max(PurgeDate) from vParam for given company code
  ws_PurgeDt = moCn.GetFldVal("Select max(convert(date,vPDesc225,3)) from vParam " + _
                          " where vPCoCd >='" + ADC("UoCoCdFr") + "' and " + _
                          " vPCoCd <='" + IIF(ADC("UoCoCdTo") <> "", ADC("UoCoCdTo"), "ZZZ") + "'" + _
                          IIF(ADC("UoCoCdSel") <> "", " and vPCoCd In(" + ADC("UoCoCdSel") + ")", "") + _
                          " and vPTyp= 'PURGEDT' and vPMCd= 'PURGEDT' and vPSCd =''")
  
  If ws_PurgeDt = "" Then ws_PurgeDt = "01/01/80"
  
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And _
      ((UCase(ADC.MenuCd) = UCase("RepWipBPcsBkDt") Or UCase(ADC.MenuCd) = UCase("RepWipFgBPcsBkDt")) And _
      CDate(ADC("UOASOFDT")) <= CDate(ws_PurgeDt) And CDate(ADC("UoAsOfDt")) <> MWLib.EmptyDate) Then
      DispMsg "As Of Date Can Be Either [01/01/80] Or Greater Than Purge Date [" + CStr(CDate(ws_PurgeDt)) + "]", etError: Exit Sub
      'DispMsg "As Of Date Cannot Be Less Than Purge Date [" + CStr(gs_PurgeDt) + "]", etError: Exit Sub
  End If
  
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  
   '*** To check if Grp1, Grp2 is selected and also to check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer
  
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then
      wDetPos = i
    End If
  Next i
  ' If wDetPos = 0 Then DispMsg "Detail Option Has To Be Selected For This Report", etError
 If wDetPos = 0 Then
    Rep.PicSec2.Suppress = True
 End If
 
  Rep.TxtHead.SetText "WIP Report Back Dated"
  Rep.TxtAsOfDt.SetText "As Of Date:  " + CStr(ADC("UoAsOfDt"))
  Rep.TxtLocFr.SetText ADC("UoIssLocFr")
  Rep.TxtLocTo.SetText ADC("UoIssLocTo")
  'Rep.TxtLocSel.SetText adc("UoIssLocSel") 22 Loc sel removed from header
  
  Rep.Text6.Suppress = True
  Rep.Text11.Suppress = True
  
  '****** Sachin 2.12 02/12/05 - Value in WIP Bag Pieces
  Rep.TxtValTyp.SetText ""
  Rep.hSalVal.Suppress = True
  Rep.h1SalVal.Suppress = True
  '****** Sachin 2.12 02/12/05 - Value in WIP Bag Pieces
  
  ' ****** Manali 3.4.1 - "01/01/80" converted to purge dt
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And _
      ((UCase(ADC.MenuCd) = UCase("RepWipBPcsBkDt") Or UCase(ADC.MenuCd) = UCase("RepWipFgBPcsBkDt")) And _
      CDate(ADC("UoAsOfDt")) = MWLib.EmptyDate) Then
      ms_AsOfDt = CDate(ws_PurgeDt)
  Else
      ms_AsOfDt = ADC("UoAsOfDt")
  End If

  Call setSqlStr
  'Debug.Print ms_SqlStr
  ADC.RepRecSource = ms_SqlStr + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
 
  '***Geeta**Emr209 **Picture
  If ADC("UoYN1") = "N" Or wDetPos = 0 Then
    Rep.PICSEC.Suppress = True
  Else
    Rep.PICSEC.Suppress = False
  End If
   
  '***(Jen 2.11)
  '*** Jay 2.13(CT) ***
  If ADC("UoBagTyp") = "F" Then
    Rep.TxtBagTypAndChr.SetText "(Flute Bags" & IIF(ADC("UoBChrFr") <> "", ",  Chr = " & ADC("UoBChrFr"), "") & ")"
  ElseIf ADC("UoBagTyp") = "P" Then
    Rep.TxtBagTypAndChr.SetText "(Primary Bags" & IIF(ADC("UoBChrFr") <> "", ",  Chr = " & ADC("UoBChrFr"), "") & ")"
  ElseIf ADC("UoBagTyp") = "C" Then
    Rep.TxtBagTypAndChr.SetText "(Component Bags" & IIF(ADC("UoBChrFr") <> "", ",  Chr = " & ADC("UoBChrFr"), "") & ")"
  End If
  'If adc("UoBagTyp") <> "" Then Rep.TxtBagTypAndChr.SetText IIF(adc("UoYn2") = "Y", "(Flute Bags", "(Primary Bags") & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
  '*** Jay 2.13(CT) ***
  '***(Jen 2.11)
  'Rn.158 suppress the fields and section that are not related to weights
  Rep.DetWithWt.Suppress = True
  Rep.PHDetWithWt.Suppress = True
  Rep.g1MetWt1.Suppress = True: Rep.g1PureWt1.Suppress = True: Rep.g1DCWt1.Suppress = True:  Rep.g1XWt1.Suppress = True
  Rep.g2MetWt1.Suppress = True: Rep.g2PureWt1.Suppress = True: Rep.g2DCWt1.Suppress = True:  Rep.g2XWt1.Suppress = True
  Rep.g3MetWt1.Suppress = True: Rep.g3PureWt1.Suppress = True: Rep.g3DCWt1.Suppress = True:  Rep.g3XWt1.Suppress = True
  Rep.gMetWt1.Suppress = True: Rep.gPureWt1.Suppress = True: Rep.gDCWt1.Suppress = True:  Rep.gXWt1.Suppress = True
  Rep.h1MetWt.Suppress = True:  Rep.h1PureWt.Suppress = True:  Rep.h1DCWt.Suppress = True:   Rep.h1XWt.Suppress = True
   
   
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******
  
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 1     'Tag = "ShowPic"
    If .Value = Checked Then ADC("UoYN1") = "Y" Else ADC("UoYN1") = "N"
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
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = False
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
'  Set adc.FirNKeyCtl = adc("UoCoCdFr")
  If UCase(ADC.MenuCd) = UCase("RepWipBPcsBkDt") Then
    
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    If ws_MultiCoMod = True Then
      Set ADC.FirNKeyCtl = ADC("UoCoCdFr")
      LblPrntCd.Visible = True
      ADC("UOCOCDFR").Enabled = True: ADC("UOCOCDTO").Visible = True: ADC("UOCOCDSEL").Visible = True
      ADC("UoPrntCdFr").Visible = True: ADC("UoPrntCdTo").Visible = True: ADC("UoPrntCdSel").Visible = True
    Else
      LblPrntCd.Visible = False
      Set ADC.FirNKeyCtl = ADC("UoIssLocFr")
      ADC("UOCOCDFR").Enabled = False: ADC("UOCOCDTO").Visible = False: ADC("UOCOCDSEL").Visible = False
      ADC("UoPrntCdFr").Visible = False: ADC("UoPrntCdTo").Visible = False: ADC("UoPrntCdSel").Visible = False
    End If
    If ADC("UoBagPcs") = "" Then ADC("UoBagPcs") = "N"
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    'Set adc.FirNKeyCtl = adc("UoIssLocFr")
    LblProcess.Visible = True
    
    FraLoc.Visible = True: FraLoc.Enabled = True
    
    ADC("UoIssLocFr").Visible = True: ADC("UoIssLocFr").Enabled = True
    ADC("UoIssLocTo").Visible = True: ADC("UoIssLocTo").Enabled = True
    ADC("UoIssLocSel").Visible = True: ADC("UoIssLocSel").Enabled = True
    
    '***** Manali 3.5.0 - 20/12/08 - following Code was commented earlier, now again included in query
    lblFgSubLoc.Visible = False
    ADC("UoFgSubLocFr").Visible = False: ADC("UoFgSubLocFr").Enabled = False
    ADC("UoFgSubLocTo").Visible = False: ADC("UoFgSubLocTo").Enabled = False
    ADC("UoFgSubLocSel").Visible = False: ADC("UoFgSubLocSel").Enabled = False
    '***** Manali 3.5.0 - 20/12/08 - above Code was commented earlier, now again included in query
    
    lblOpnBag.Visible = False
     ADC("UoYn").Visible = False: ADC("UoYn").Enabled = False
    ADC("UoYn") = ""
'    LblBagPcs.Visible = True
'    adc("UoBagPcs").Visible = True: adc("UoBagPcs").Enabled = True
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
    FraOptBagPcs.Enabled = True: FraOptBagPcs.Visible = True
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
    
    ' ****** Manali 3.5.0 - 06/11/08 - Radio buttons Added
    FraOptBag.Enabled = False: FraOptBag.Visible = False: OptBag(2).Value = True
    ' ****** Manali 3.5.0 - 06/11/08 - Radio buttons Added
    
    '*** (2.11) Jen
    LblPckLst.Visible = False
    ADC("UoFdPlYyFr").Visible = False: ADC("UoFdPlChrFr").Visible = False: ADC("UoFdPlNoFr").Visible = False
    ADC("UoFdPlYyTo").Visible = False: ADC("UoFdPlChrTo").Visible = False: ADC("UoFdPlNoTo").Visible = False
    LblPkLstDeLimit1.Visible = False: LblPkLstDeLimit2.Visible = False: LblPkLstDeLimit3.Visible = False: LblPkLstDeLimit4.Visible = False
    ADC("UoFdPlYyFr").CmpStr = "": ADC("UoFdPlChrFr").CmpStr = "": ADC("UoFdPlNoFr").CmpStr = ""
    ADC("UoFdPlYyTo").CmpStr = "": ADC("UoFdPlChrTo").CmpStr = "": ADC("UoFdPlNoTo").CmpStr = ""
    ADC("UoPrntCdFr").CmpStr = "LocPrntCd>=": ADC("UoPrntCdTo").CmpStr = "LocPrntCd<=": ADC("UoPrntCdSel").CmpStr = "LocPrntCd In "
    '*** (2.11) Jen
  ElseIf UCase(ADC.MenuCd) = UCase("RepWipFgBPcsBkDt") Then
    ' Emr211
    'Set ADC.FirNKeyCtl = ADC("UoIssLocFr")
     'Set adc.FirNKeyCtl = adc("UOWRKFR")
     Set ADC.FirNKeyCtl = ADC("UoFgSubLocFr")     ' ****** Manali 3.5.0 - FGSubLoc included in Fg Pcs bckDtd report
    
    ' Emr211
    LblProcess.Visible = False
    
    FraLoc.Visible = False: FraLoc.Enabled = False
    ADC("UoIssLocFr").Visible = False: ADC("UoIssLocFr").Enabled = False
    ADC("UoIssLocTo").Visible = False: ADC("UoIssLocTo").Enabled = False
    ADC("UoIssLocSel").Visible = False: ADC("UoIssLocSel").Enabled = False
    
    '***** Manali 3.5.0 - 20/12/08 - following Code was commented earlier, now again included in query
    lblFgSubLoc.Visible = True
    ADC("UoFgSubLocFr").Visible = True: ADC("UoFgSubLocFr").Enabled = True
    ADC("UoFgSubLocTo").Visible = True: ADC("UoFgSubLocTo").Enabled = True
    ADC("UoFgSubLocSel").Visible = True: ADC("UoFgSubLocSel").Enabled = True
    '***** Manali 3.5.0 - 20/12/08 - above Code was commented earlier, now again included in query

    lblOpnBag.Visible = True
    ' adc("UoYn").Visible = True: adc("UoYn").Enabled = True  ' ****** Manali 3.5.0 - 06/11/08 - Radio buttons Added
    LblBagPcs.Visible = False
    ADC("UoBagPcs").Visible = False: ADC("UoBagPcs").Enabled = False
    ADC("UoBagPcs") = ""
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
    FraOptBagPcs.Enabled = False: FraOptBagPcs.Visible = False
    OptBagPcs(0).Value = False: OptBagPcs(1).Value = False
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
      
    ' ****** Manali 3.5.0 - 06/11/08 - Radio buttons Added
    FraOptBag.Enabled = True: FraOptBag.Visible = True
    ' ****** Manali 3.5.0 - 06/11/08 - Radio buttons Added

    ADC("UoBagTyp").Visible = False
    LblBagTyp.Visible = False
    ADC("UoBChrFr").Visible = False
    LblUoBChrFr.Visible = False
    
    '*** (2.11) Jen
    LblPckLst.Visible = True
    ADC("UoFdPlYyFr").Visible = True: ADC("UoFdPlChrFr").Visible = True: ADC("UoFdPlNoFr").Visible = True
    ADC("UoFdPlYyTo").Visible = True: ADC("UoFdPlChrTo").Visible = True: ADC("UoFdPlNoTo").Visible = True
    LblPkLstDeLimit1.Visible = True: LblPkLstDeLimit2.Visible = True: LblPkLstDeLimit3.Visible = True: LblPkLstDeLimit4.Visible = True
    
    ADC("UoFdPlYyFr").CmpStr = "FdPlYy >=": ADC("UoFdPlChrFr").CmpStr = "FdPlChr >=": ADC("UoFdPlNoFr").CmpStr = "FdPlNo >="
    ADC("UoFdPlYyTo").CmpStr = "FdPlYy <=": ADC("UoFdPlChrTo").CmpStr = "FdPlChr <=": ADC("UoFdPlNoTo").CmpStr = "FdPlNo <="
    '*** (2.11) Jen
    
    '****** Sachin 2.14.0 - [Multi-Company]
    LblPrntCd.Visible = False
    ADC("UOCOCDFR").Enabled = False: ADC("UOCOCDTO").Visible = False: ADC("UOCOCDSEL").Visible = False
    ADC("UoPrntCdFr").Visible = False: ADC("UoPrntCdTo").Visible = False: ADC("UoPrntCdSel").Visible = False
    ADC("UoPrntCdFr").CmpStr = "": ADC("UoPrntCdTo").CmpStr = "": ADC("UoPrntCdSel").CmpStr = ""
    '****** Sachin 2.14.0 - [Multi-Company]
    
  End If
  
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_LocTyp = "'P'"
  gs_DmTcTyp = "DM"
  gs_OmTcTyp = "SO"
  gs_CmCtg = "C"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  Call SetGroupSort

End Sub

Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gb_CoCdFor = False '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
 gs_LocTyp = "'P'"
   gb_CoCdFor = False '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

  '*** (Jen 2.11)
  Select Case UCase(IdName)
    '*** Jay 2.13(CT) ***
    Case Is = UCase("UoBChrFr")
      If ADC.Mode = xNorm Then
        If ADC("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Bag Type": Exit Sub
        HlpList.PMCd "BCHR", "'" + ADC("UoBagTyp") + "'": Exit Sub
      End If
    '*** Jay 2.13(CT) ***
    Case Is = UCase("UOISSLOCFR"), UCase("UOISSLOCTO"), UCase("UOISSLOCSEL")
        If ADC("UOLOCTYPFR") = "R" Then gs_LocTyp = "'R'"
        If ADC("UOLOCTYPFR") = "P" Then gs_LocTyp = "'P'"
 
  End Select
  '*** (Jen 2.11)

  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
' **** Zubin 211 **** '
  Select Case UCase(IdName)
    Case Is = UCase("UOBCHRFR")
      'If ADC("UOYN2") = "N" Then
      ADC("UOBCHRFR") = ""
  End Select
' **** Zubin 211 **** '
End Sub

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptBag_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN") = "Y"
Case Is = 1
  ADC("UoYN") = "N"
Case Is = 2
  ADC("UoYN") = ""
End Select
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptBag_GotFocus(Index As Integer)
  DispMsg FraOptBag.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
Private Sub OptBagPcs_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoBagPcs") = "N"
Case Is = 1
  ADC("UoBagPcs") = "R"
End Select
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
Private Sub OptBagPcs_GotFocus(Index As Integer)
DispMsg FraOptBagPcs.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added

Private Sub OptLoc_GotFocus(Index As Integer)
  DispMsg FraLoc.ToolTipText, etInfo
Select Case Index
Case Is = 0
  ADC("UOISSLOCFR") = "":   ADC("UOISSLOCTO") = "":   ADC("UOISSLOCsel") = ""
  ADC("UoLocTypFr") = "P"
Case Is = 1
  ADC("UOISSLOCFR") = "":   ADC("UOISSLOCTO") = "":   ADC("UOISSLOCsel") = ""
  ADC("UoLocTypFr") = "R"
End Select

End Sub

Private Sub tab_REP_Click(PreviousTab As Integer)
  fra_tabrep(TAB_REP.Tab).Enabled = True
  fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  ADC("UOUSRCD") = ADC.UsrCd
  ADC("UOMNUCD") = ADC.MenuCd
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)

  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = False Or UCase(ADC.MenuCd) = UCase("RepWipFgBPcsBkDt") Then _
      ADC("UoCoCdFr") = gs_CoCd: ADC("UoCoCdTo") = gs_CoCd: ADC("UoCoCdSel") = "": ADC("UoCoCdFr").Enabled = False
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
  ADC("UoAsOfDt") = moCn.SrvrDate
  'Manoj
  ADC("UOYN1") = "N"
  '*** Jay 2.13(CT) ***
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + gs_CoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    ADC("UoBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    ADC("UoBChrFr").Visible = False
    ADC("UoBagTyp") = "P": ADC("UoBChrFr") = ""
  End If
  Set wHeadChk = Nothing
  '*** Jay 2.13(CT) ***
  
'*** Jay 2.13(CT) ***
'  '*** (Jen 2.11)
'  If moCn.GetFldVal("Select hFluteBagYN from Head where HCoCd= '" + adc("UoCoCdFr") + "' and " + _
'                                                    "HCd= '" + ctSelfCmCd + "' ") = "N" Then
'    LblBagTyp.Visible = False
'    adc("UoYN2").Visible = False
'    LblUoBChrFr.Visible = False
'    adc("UoBChrFr").Visible = False
'    adc("UOYN2") = "N": adc("UoBChrFr") = ""
'  End If
'  '*** (Jen 2.11)
'*** Jay 2.13(CT) ***

  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  If ADC("UoYN") = "Y" Then
    OptBag(0).Value = True
  ElseIf ADC("UoYN") = "N" Then
    OptBag(1).Value = True
  ElseIf ADC("UoYN") = "" Then
    OptBag(2).Value = True
  End If
  If ADC("UoBagPcs") = "" Then ADC("UoBagPcs") = "N"
  If ADC("UoBagPcs") = "N" Then OptBagPcs(0).Value = True
  If ADC("UoBagPcs") = "R" Then OptBagPcs(1).Value = True
  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  If ADC("UoYN1") = "" Then ADC("UoYN1") = "N"
  
  If ADC("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  If ADC("UOLOCTYPFR") = "P" Then
    OptLoc(0).Value = True
  ElseIf ADC("UOLOCTYPFR") = "R" Then
    OptLoc(1).Value = True
  End If

End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = False Or UCase(ADC.MenuCd) = UCase("RepWipFgBPcsBkDt"), ADC("UOCOCDFR"), gs_CoCd) + "' and " + _
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
  Rep.TxtCoCdFr.SetText ADC("UOCOCDFR")
  
  If ws_MultiCoMod = False Or UCase(ADC.MenuCd) = UCase("RepWipFgBPcsBkDt") Then
      Rep.TxtCoCdTo.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True
  Else
      Rep.TxtCoCdToLbl.Suppress = False
      Rep.TxtCoCdTo.SetText ADC("UOCOCDTO")
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
Private Sub Form_Resize()
    ADC.Top = IIF(Me.ScaleHeight > 450, Me.ScaleHeight - 450, Me.ScaleHeight)
    ADC.Width = IIF(Me.ScaleWidth > 50, Me.ScaleWidth - 50, Me.ScaleWidth)
    DST_REP.Top = ADC.Top
    FRA_KEYS.Top = ADC.Top
    FRA_REPOBJ.Height = IIF(Me.ScaleHeight > 450, Me.ScaleHeight - 450, Me.ScaleHeight)
    FRA_REPOBJ.Width = IIF(Me.ScaleWidth > 25, Me.ScaleWidth - 25, Me.ScaleWidth)
    CRV_REP.Height = FRA_REPOBJ.Height
    CRV_REP.Width = FRA_REPOBJ.Width
End Sub
Public Sub setSqlStr()
  Dim wsCmnFlds1 As String, wsCmnFlds2 As String, wsCmnJoin As String
  Dim wsTxnQtyFlds As String, wsTxnQtyWhr As String, wsBQty As String, wsBPtQty As String, wsRjQty As String, wsRjPtQty As String
  Dim wsTxnFlds As String, wsTxnJoin As String, wsTxnWhere As String, wsTxnCnd As String
  Dim wsFgFlds As String, wsFgJoin As String, wsFgWhere As String, wsFgCnd As String, wsFgTxnjoin As String
  Dim wsTxnTbl As String, wsTxnTblJoin As String, wsFgTbl As String, wsFgTblJoin As String, wsPrmTbl As String
  
  Dim wrepcnd As String, wCnd As String, wsDtCnd As String
  Dim grpflds0 As String, grpflds1 As String
  
  Dim wAddlCnd As String  '****** Sachin 3.02.0
  
'*** Jay 2.13(CT) ***
  'Dim ws_BagTypJoin As String ' **** Zubin 211 **** '
  Dim ws_BagJoin As String
  Dim ws_BagJoinTd As String    ' Manali 3.8.1
'*** Jay 2.13(CT) ***

  
  wsTxnTbl = " Txn "
  wsTxnTblJoin = IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' and ", "") + " TIdNo=TdTIdNo and TCoCd = TdCoCd ANd TTc = TdTc And TYy = TdYy And TChr = TdChr And TNo = TdNo"
  wsFgTbl = " Fg "
  wsFgTblJoin = IIF(gs_Partition = ctCurrPrtn, " FgPrtKey='" + ctCurrPrtn + "' and ", "") + " FgIdNo=FdFgIdNo and FgCoCd = FdCoCd And FgTc = FdTc And FgYy = FdYy And FgChr = FdChr And FgNo = FdNo"
  wsPrmTbl = " (SELECT PMCD, PValue FROM Param WHERE PTyp='TC')"
' Manoj
'  wsCmnFlds1 = " BYy, BChr, BNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BOdDmCd, OdDmSz, DmCtg, OdKt, OdDmCol"
  '*** Jay 3.2.0 [PicNm]
  '11.b OdSfx added
  wsCmnFlds1 = " BCoCd, BYy, BChr, BNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, DmTcTyp, BOdDmCd,OdSfx, OdDmSz, DmCtg, OdKt, OdDmCol, OdPicNm"
  '*** Jay 3.2.0 [PicNm]
  wsCmnFlds2 = " OmCmCd, OmPoNo"
  
  wsCmnJoin = " JOIN DsgMst ON DmIdNo=BDmIdNo and DmTcTyp= BOdDmTcTyp AND DmCd= BOdDmCd AND DmSz= ''" + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
      " JOIN (OrdMst JOIN OrdDsg ON OmIdNo=OdOmIdNo and OmCoCd= OdCoCd AND OmTc= OdTc AND OmYy= OdYy AND OmChr= OdChr AND OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + ")" + _
      " ON " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey=BPrtKey and ", "") + " OmIdNo=BOmIdNo and OmCoCd= BCoCd AND OmTc= BOdTc AND OmYy= BOdYy AND OmChr= BOdChr AND OmNo= BOdNo AND OdSr= BOdSr"
  wsFgTxnjoin = " JOIN Txnd ON TdSrNo = 0 AND TdCoCd = FdCoCd And TdYyKey = FdBYyKey" + _
      " And TdBYy = FdBYy And TdBChr = FdBChr And TdBNo = FdBNo " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=FdPrtKey ", "")

  ' ***** Manali 3.4.1 - 02/12/08 - [adc("UoAsOfDt") replaced by ms_AsOfDt
  wsDtCnd = " <= '" + Format$(ms_AsOfDt, "dd/mm/yy") + "'"
  'wsDtCnd = " <= '" + Format$(adc("UoAsOfDt"), "dd/mm/yy") + "'"
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(ADC("UOCMCDFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(ADC("UOCMCDTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(ADC("UOCMCDSEL")) + ")", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  If UCase(ADC.MenuCd) = UCase("RepWipBPcsBkDt") Then
' **** Zubin 211 **** '
'    wsBQty = " ISNULL(CASE WHEN PValue = 'BS' THEN (BOpnQty + TB.TdBQty) ELSE T1.TdBQty END, BOpnQty)"
'    wsBPtQty = " ISNULL(CASE WHEN PValue = 'BS' THEN (BOpnPtQty + TB.TdBPtQty) ELSE T1.TdPtQty END, BOpnPtQty)"
'    wsRjQty = " ISNULL(CASE WHEN PValue = 'BS' THEN TB.TdRjQty ELSE T1.TdRjQty END, BRjQty)"
'    wsRjPtQty = " ISNULL(CASE WHEN PValue = 'BS' THEN TB.TdRjPtQty ELSE T1.TdRjPtQty END, BRjPtQty)"
    
    wsBQty = " ISNULL(CASE WHEN TC.PValue = 'BS' THEN (BOpnQty + TB.TdBQty) " + _
                "          WHEN TC.PValue = 'BM' AND FB.PValue = 'F' THEN (BOpnQty + TB.TdBQty) " + _
                "          ELSE T1.TdBQty END, BOpnQty)"
                
    wsBPtQty = " ISNULL(CASE WHEN TC.PValue = 'BS' THEN (BOpnPtQty + TB.TdBPtQty) " + _
                "            WHEN TC.PValue = 'BM' AND FB.PValue = 'F' THEN (BOpnPtQty + TB.TdBPtQty) " + _
                "            ELSE T1.TdPtQty END, BOpnPtQty)"
                
    wsRjQty = " ISNULL(CASE WHEN TC.PValue = 'BS' THEN TB.TdRjQty " + _
                "           WHEN TC.PValue = 'BM' AND FB.PValue = 'F' THEN TB.TdRjQty " + _
                "           ELSE T1.TdRjQty END, BRjQty)"
                
    wsRjPtQty = " ISNULL(CASE WHEN TC.PValue = 'BS' THEN TB.TdRjPtQty " + _
                "             WHEN TC.PValue = 'BM' AND FB.PValue = 'F' THEN TB.TdRjPtQty " + _
                "             ELSE T1.TdRjPtQty END, BRjPtQty)"
' **** Zubin 211 **** '

    If UCase$(ADC("UOBAGPCS")) = UCase$("N") Then
      wsTxnQtyFlds = wsBQty + " AS qQty,0.00 As qSalVal," + wsBPtQty + " AS qPtQty"
      wsTxnQtyWhr = " (" + wsBQty + " <> 0 OR" + wsBPtQty + " <> 0 OR"
    ElseIf UCase$(ADC("UOBAGPCS")) = UCase$("R") Then
      '*** Jay 2.14Next [0 As qSalVal, added as not mention earlier - Report Error Rectified]
      wsTxnQtyFlds = wsRjQty + " AS qQty,0.00 As qSalVal," + wsRjPtQty + " AS qPtQty"
      '*** Jay 2.14Next [0 As qSalVal, added as not mention earlier - Report Error Rectified]
      wsTxnQtyWhr = " (" + wsRjQty + " <> 0 OR" + wsRjPtQty + " <> 0 OR"
    End If
    
    'BLoc BWrk BGrWt BDlDays qQty qPtQty
'inuu
    wsTxnFlds = IIF(OptLoc(0).Value = True, "ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN T1.TdFrBLoc ELSE T1.TdToBLoc END, BLoc)", "LocPrcs") + " AS BLoc, LocPrntCd, " + _
        " ISNULL(T1.TdWrk, BWrk) AS BWrk, CAST(ISNULL(T1.TdBGrWt, 0) AS DECIMAL(10,3)) BGrWt, 0 AS BDlDays" + _
        IIF(wsTxnQtyFlds <> "", " , ", "") + wsTxnQtyFlds
    'inuu
    wsFgFlds = IIF(OptLoc(0).Value = True, "FgToBLoc", "LocPrcs") + " as FgToBLoc, LocPrntCd, TdWrk AS FdWrk, CAST(FdGrWt AS DECIMAL(10,3)) FdBGrWt, 0 AS BDlDays, FdQty, 0.00 as qSalVal, 0 AS FdPtQty"
    
' **** Zubin 211 **** '
'''''    wsTxnJoin = " (Select * From Bag Where BOpnDt" + wsDtCnd + " AND " + _
        " (case when BClsDt= '" + Format$(MWLib.EmptyDate, "dd/mm/yy") + "' then '" + Format$(CDate("31/12/2028"), "dd/mm/yy") + _
        "' else BClsDt end) > '" + Format$(adc("UoAsOfDt"), "dd/mm/yy") + "' ) B LEFT OUTER JOIN Txnd T1 " + " JOIN " + wsPrmTbl + _
        " P ON PMCd= T1.TdTc ON " + " BCoCd= T1.TdCoCd AND BYy= T1.TdBYy AND BChr= T1.TdBChr AND BNo= T1.TdBNo" + _
        " AND T1.TdSrNo = 0" + " AND T1.TdYyKey= (SELECT MAX(TdYyKey) FROM Txnd TdB Join " + wsTxnTbl + " TB ON" + _
        " TB.TCoCd = TdB.TdCoCd ANd TB.TTc = TdB.TdTc And TB.TYy = TdB.TdYy And TB.TChr = TdB.TdChr And TB.TNo = TdB.TdNo" + _
        " WHERE TdB.TdCoCd= BCoCd AND TdB.TdBYy= BYy AND TdB.TdBChr= BChr AND TdB.TdBNo= BNo AND TB.TDt" + wsDtCnd + ")" + _
        " LEFT OUTER JOIN (SELECT TdCoCd, TdBYy, TdBChr, TdBNo," + _
        " SUM(TdBQty * CASE WHEN TdBDC = 'C' THEN -1 WHEN TdBDC = 'D' THEN 1 END) TdBQty," + _
        " SUM(TdPtQty * CASE WHEN TdBDC = 'C' THEN -1 WHEN TdBDC = 'D' THEN 1 END) TdBPtQty," + _
        " SUM(TdRjQty * CASE WHEN TdBDC = 'C' THEN -1 WHEN TdBDC = 'D' THEN 1 END) TdRjQty," + _
        " SUM(TdRjPtQty * CASE WHEN TdBDC = 'C' THEN -1 WHEN TdBDC = 'D' THEN 1 END) TdRjPtQty" + _
        " FROM Txnd Join " + wsTxnTbl + " T ON" + wsTxnTblJoin + " WHERE TdSrNo = 0 AND TdBDC <> '' And TDt" + wsDtCnd + _
        " Group By TdCoCd, TdBYy, TdBChr, TdBNo) TB ON" + _
        " BYy = TB.TdBYy AND BChr = TB.TdBChr  AND BNo = TB.TdBNo AND BCoCd = TB.TdCoCd"

'*** Jay 2.13(CT) ***
    ' ***** Manali 3.8.0 -
    ws_BagJoin = " Join Param FB On FB.PTyp= 'BCHR' and FB.PMCd= BChr and FB.PSCd= '' " + _
                    IIF(ADC("UoBagTyp") <> "", " and FB.PValue= '" + ADC("UoBagTyp") + "' ", " ") + _
                    IIF(ADC("UOBCHRFR") <> "", " and FB.PMCd= '" + ADC("UOBCHRFR") + "' ", " ")

    ws_BagJoinTd = " Join Param FB On FB.PTyp= 'BCHR' and FB.PMCd= T1.TdBChr and FB.PSCd= '' " + _
                    IIF(ADC("UoBagTyp") <> "", " and FB.PValue= '" + ADC("UoBagTyp") + "' ", " ") + _
                    IIF(ADC("UOBCHRFR") <> "", " and FB.PMCd= '" + ADC("UOBCHRFR") + "' ", " ")

' Bef  3.8.0
'    ws_BagJoin = " Join Param FB On FB.PTyp= 'BCHR' and FB.PMCd= BChr and FB.PSCd= '' " + _
'                    IIF(adc("UoBagTyp") <> "", " and FB.PValue= '" + adc("UoBagTyp") + "' ", " ") + _
'                    IIF(adc("UOBCHRFR") <> "", " and FB.PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'    If UCase(adc("UoBagTyp")) = "N" Then
'      ws_BagTypJoin = " and FB.PValue = 'N' " + IIF(adc("UOBCHRFR") <> "", " and FB.PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'    ElseIf UCase(adc("UoBagTyp")) = "Y" Then
'      ws_BagTypJoin = " and FB.PValue = 'Y' " + IIF(adc("UOBCHRFR") <> "", " and FB.PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'    Else
'      ws_BagTypJoin = " " + IIF(adc("UOBCHRFR") <> "", " and FB.PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'    End If
'    ws_BagJoin = " Join Param FB On FB.PTyp= 'BCHR' and FB.PMCd= BChr and FB.PSCd= '' " + ws_BagTypJoin
'*** Jay 2.13(CT) ***
        
    '****** Sachin 2.14.0 - [Multi-Company] - Loc Join added below ******
    ' ***** Manali 3.4.1 - 02/12/08 - [adc("UoAsOfDt") replaced by ms_AsOfDt]
    ' ***** Manali 3.8.0 - ws_BagJoinTd added in inner table (TB), Alias For Txnd inside TB given as T1
    wsTxnJoin = " (Select * From Bag Where BOpnDt" + wsDtCnd + " AND " + _
        " (case when BClsDt= '" + Format$(MWLib.EmptyDate, "dd/mm/yy") + "' then '" + Format$(CDate("31/12/2028"), "dd/mm/yy") + _
        "' else BClsDt end) > '" + Format$(ms_AsOfDt, "dd/mm/yy") + "'  /* '" + Format$(ADC("UoAsOfDt"), "dd/mm/yy") + "' */ ) B " + _
        " LEFT OUTER JOIN Txnd T1 " + " JOIN " + wsPrmTbl + _
        " TC ON PMCd= T1.TdTc ON " + " BIdNo=T1.TdBIdNo and BCoCd= T1.TdCoCd AND BYy= T1.TdBYy AND BChr= T1.TdBChr AND BNo= T1.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=T1.TdPrtKey ", "") + _
        " AND T1.TdSrNo = 0" + " AND T1.TdYyKey= (SELECT MAX(TdYyKey) FROM Txnd TdB Join " + wsTxnTbl + " TB ON" + _
        " TB.TIdNo=TdB.TdTIdNo and TB.TCoCd = TdB.TdCoCd ANd TB.TTc = TdB.TdTc And TB.TYy = TdB.TdYy And TB.TChr = TdB.TdChr And TB.TNo = TdB.TdNo" + IIF(gs_Partition = ctCurrPrtn, " and TB.TPrtKey=TdB.TdPrtKey ", "") + _
        " WHERE TdB.TdBIdNo=BIdNo and TdB.TdCoCd= BCoCd AND TdB.TdBYy= BYy AND TdB.TdBChr= BChr AND TdB.TdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " and TdB.TdPrtKey=BPrtKey ", "") + " AND TB.TDt" + wsDtCnd + ")" + _
        " JOIN LOC Lc ON LocCoCd=IsNull(T1.TdCoCd,BCoCd) And LocCd=ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN T1.TdFrBLoc ELSE T1.TdToBLoc END, BLoc) " + _
        " LEFT OUTER JOIN (SELECT TdPrtKey, TdBIdNo, TdCoCd, TdBYy, TdBChr, TdBNo," + _
        " SUM(TdBQty * CASE WHEN TdBDC = 'C' THEN -1 WHEN TdBDC = 'D' THEN 1 END) TdBQty," + _
        " SUM(TdPtQty * CASE WHEN TdBDC = 'C' THEN -1 WHEN TdBDC = 'D' THEN 1 END) TdBPtQty," + _
        " SUM(TdRjQty * CASE WHEN TdBDC = 'C' THEN -1 WHEN TdBDC = 'D' THEN 1 END) TdRjQty," + _
        " SUM(TdRjPtQty * CASE WHEN TdBDC = 'C' THEN -1 WHEN TdBDC = 'D' THEN 1 END) TdRjPtQty" + _
        " FROM Txnd T1 Join " + wsTxnTbl + " T ON" + wsTxnTblJoin + " " + ws_BagJoinTd + " WHERE " + IIF(gs_Partition = ctCurrPrtn, " TdPrtKey='" + ctCurrPrtn + "' and ", "") + " TdSrNo = 0 AND (TdTc<>'BM' OR PValue='F') AND TdBDC <> '' And TDt" + wsDtCnd + _
        " Group By TdPrtKey, TdBIdNo, TdCoCd, TdBYy, TdBChr, TdBNo) TB ON" + _
        " BIdNo=TB.TdBIdNo and BYy = TB.TdBYy AND BChr = TB.TdBChr  AND BNo = TB.TdBNo AND BCoCd = TB.TdCoCd " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=TB.TdPrtKey ", "")
    
'  ***** Bef  3.8.0
'    wsTxnJoin = " (Select * From Bag Where BOpnDt" + wsDtCnd + " AND " + _
'        " (case when BClsDt= '" + Format$(MWLib.EmptyDate, "dd/mm/yy") + "' then '" + Format$(CDate("31/12/2028"), "dd/mm/yy") + _
'        "' else BClsDt end) > '" + Format$(ms_AsOfDt, "dd/mm/yy") + "'  /* '" + Format$(adc("UoAsOfDt"), "dd/mm/yy") + "' */ ) B " + _
'        " LEFT OUTER JOIN Txnd T1 " + " JOIN " + wsPrmTbl + _
'        " TC ON PMCd= T1.TdTc ON " + " BIdNo=T1.TdBIdNo and BCoCd= T1.TdCoCd AND BYy= T1.TdBYy AND BChr= T1.TdBChr AND BNo= T1.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=T1.TdPrtKey ", "") + _
'        " AND T1.TdSrNo = 0" + " AND T1.TdYyKey= (SELECT MAX(TdYyKey) FROM Txnd TdB Join " + wsTxnTbl + " TB ON" + _
'        " TB.TIdNo=TdB.TdTIdNo and TB.TCoCd = TdB.TdCoCd ANd TB.TTc = TdB.TdTc And TB.TYy = TdB.TdYy And TB.TChr = TdB.TdChr And TB.TNo = TdB.TdNo" + IIF(gs_Partition = ctCurrPrtn, " and TB.TPrtKey=TdB.TdPrtKey ", "") + _
'        " WHERE TdB.TdBIdNo=BIdNo and TdB.TdCoCd= BCoCd AND TdB.TdBYy= BYy AND TdB.TdBChr= BChr AND TdB.TdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " and TdB.TdPrtKey=BPrtKey ", "") + " AND TB.TDt" + wsDtCnd + ")" + _
'        " JOIN LOC Lc ON LocCoCd=IsNull(T1.TdCoCd,BCoCd) And LocCd=ISNULL(CASE WHEN T1.TdToBLoc = '' THEN T1.TdFrBLoc ELSE T1.TdToBLoc END, BLoc) " + _
'        " LEFT OUTER JOIN (SELECT TdPrtKey, TdBIdNo, TdCoCd, TdBYy, TdBChr, TdBNo," + _
'        " SUM(TdBQty * CASE WHEN TdBDC = 'C' THEN -1 WHEN TdBDC = 'D' THEN 1 END) TdBQty," + _
'        " SUM(TdPtQty * CASE WHEN TdBDC = 'C' THEN -1 WHEN TdBDC = 'D' THEN 1 END) TdBPtQty," + _
'        " SUM(TdRjQty * CASE WHEN TdBDC = 'C' THEN -1 WHEN TdBDC = 'D' THEN 1 END) TdRjQty," + _
'        " SUM(TdRjPtQty * CASE WHEN TdBDC = 'C' THEN -1 WHEN TdBDC = 'D' THEN 1 END) TdRjPtQty" + _
'        " FROM Txnd Join " + wsTxnTbl + " T ON" + wsTxnTblJoin + " WHERE " + IIF(gs_Partition = ctCurrPrtn, " TdPrtKey='" + ctCurrPrtn + "' and ", "") + " TdSrNo = 0 AND TdBDC <> '' And TDt" + wsDtCnd + _
'        " Group By TdPrtKey, TdBIdNo, TdCoCd, TdBYy, TdBChr, TdBNo) TB ON" + _
'        " BIdNo=TB.TdBIdNo and BYy = TB.TdBYy AND BChr = TB.TdBChr  AND BNo = TB.TdBNo AND BCoCd = TB.TdCoCd " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=TB.TdPrtKey ", "")
' **** Zubin 211 **** '

    '****** Sachin 3.02 - Id fields in Join
    '578-34 STR(FdKey, 7) changed to STR(FdKey, 8)
    wsFgJoin = " Fgd JOIN " + wsFgTbl + " F ON" + wsFgTblJoin + _
        " Join Loc Lc On LocCoCd=FgCoCd And LocCd=FgToBLoc " + _
        " JOIN " + wsPrmTbl + " P ON PMCD = FdTc" + _
        " JOIN Bag ON BIdNo=FdBIdNo and BCoCd= FdCoCd AND BYy= FdBYy AND BChr= FdBChr AND BNo= FdBNo" + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=FdPrtKey ", "") + _
        " AND (FdYy + STR(FdKey, 8)) = (SELECT MAX(FdYy + STR(FdKey, 8)) FROM Fgd" + _
        " JOIN " + wsFgTbl + " F ON" + wsFgTblJoin + _
        " WHERE FdBIdNo=BIdNo and FdCoCd= BCoCd AND FdBYy= BYy AND FdBChr= BChr AND FdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=BPrtKey ", "") + " And FgDt" + wsDtCnd + ")" + _
        " AND FdBYyKey >= (SELECT MAX(TdYyKey) FROM Txnd" + _
        " Join " + wsTxnTbl + " T ON " + wsTxnTblJoin + _
        " WHERE TdBIdNo=BIdNo and TdCoCd= BCoCd AND TdBYy= BYy AND TdBChr= BChr AND TdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=BPrtKey ", "") + " And TDt" + wsDtCnd + ")" + wsFgTxnjoin
    
' **** Zubin 211 (TC. added with PValue) **** '
    wsTxnWhere = IIF(gs_Partition = ctCurrPrtn, " B.BPrtKey='" + ctCurrPrtn + "' and ", "") + " (" + wsTxnQtyWhr + " (" + wsBQty + " = 0 AND" + wsBPtQty + " = 0 AND" + wsRjQty + " = 0 AND" + wsRjPtQty + " = 0 AND" + _
        " CAST(ISNULL(T1.TdBGrWt, 0) AS DECIMAL(10,3)) <> 0))" + _
        " AND ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN  T1.TdFrBLoc ELSE T1.TdToBLoc END, BOpnLoc) <> 'PFG'" + _
        " AND ISNULL(T1.TdYyKey,'') NOT IN (SELECT FdBYyKey FROM fgd JOIN " + wsFgTbl + " F ON" + wsFgTblJoin + _
        " WHERE FdBIdNo=BIdNo and FdCoCd = BCoCd AND FdByy = BYy AND FdBChr = BChr AND FdBNo = BNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=BPrtKey ", "") + " And FgDt" + wsDtCnd + ") AND ISNULL(TC.PValue , '') " + _
        IIF(UCase$(ADC("UOBAGPCS")) = UCase$("N"), " <> 'MLT' AND Not (ISNULL(TC.PValue , '') = 'REJ' AND TdBDC = 'C') ", _
          IIF(UCase$(ADC("UOBAGPCS")) = UCase$("R"), " = 'REJ' AND TdBDC = 'C'", ""))
' **** Zubin 211 (TC. added with PValue) **** '
    wsFgWhere = " PValue IN ( 'FWB') AND BOpnDt " + wsDtCnd
    
    '***** Manali 3.5.0 - 20/12/08 - following Code was commented earlier, now again included in query
    ADC("UOFGSUBLOCFR").CmpStr = ""
    ADC("UOFGSUBLOCTO").CmpStr = ""
    ADC("UOFGSUBLOCSEL").CmpStr = ""
    '***** Manali 3.5.0 - 20/12/08 - above Code was commented earlier, now again included in query
    
    'adc("UoIssLocFr").CmpStr = " ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN T1.TdFrBLoc ELSE T1.TdToBLoc END, BOpnLoc) >= "
    'adc("UoIssLocTo").CmpStr = " ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN T1.TdFrBLoc ELSE T1.TdToBLoc END, BOpnLoc) <= "
    'adc("UoIssLocSel").CmpStr = " ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN T1.TdFrBLoc ELSE T1.TdToBLoc END, BOpnLoc) In "
    If OptLoc(0).Value = True Then
      ADC("UoIssLocFr").CmpStr = " ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN T1.TdFrBLoc ELSE T1.TdToBLoc END, BOpnLoc) >= "
      ADC("UoIssLocTo").CmpStr = " ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN T1.TdFrBLoc ELSE T1.TdToBLoc END, BOpnLoc) <= "
      ADC("UoIssLocSel").CmpStr = " ISNULL(CASE WHEN T1.TdToBLoc = '' Or T1.TdAck='N' THEN T1.TdFrBLoc ELSE T1.TdToBLoc END, BOpnLoc) In "
    ElseIf OptLoc(1).Value = True Then
      ADC("UoIssLocFr").CmpStr = " LocPrcs >= "
      ADC("UoIssLocTo").CmpStr = " LocPrcs <= "
      ADC("UoIssLocSel").CmpStr = " LocPrcs In "
    End If
    
    ADC("UOWRKFR").CmpStr = " ISNULL(T1.TdWrk, BWrk) >= "
    ADC("UOWRKTO").CmpStr = " ISNULL(T1.TdWrk, BWrk)  <= "
    ADC("UOWRKSEL").CmpStr = " ISNULL(T1.TdWrk, BWrk)  In "
    '****** Sachin 2.14.0 - [Multi-Company] *****
    ADC("UoPrntCdFr").CmpStr = "LocPrntCd>=": ADC("UoPrntCdTo").CmpStr = "LocPrntCd<=": ADC("UoPrntCdSel").CmpStr = "LocPrntCd In "
    wrepcnd = ADC.RepCond
    wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    wsTxnCnd = wCnd
    
    'adc("UoIssLocFr").CmpStr = " FgToBLoc >= "
    'adc("UoIssLocTo").CmpStr = " FgToBLoc <= "
    'adc("UoIssLocSel").CmpStr = " FgToBLoc In "
    If OptLoc(0).Value = True Then
      ADC("UoIssLocFr").CmpStr = " FgToBLoc >= "
      ADC("UoIssLocTo").CmpStr = " FgToBLoc <= "
      ADC("UoIssLocSel").CmpStr = " FgToBLoc In "
    ElseIf OptLoc(1).Value = True Then
      ADC("UoIssLocFr").CmpStr = " LocPrcs >= "
      ADC("UoIssLocTo").CmpStr = " LocPrcs <= "
      ADC("UoIssLocSel").CmpStr = " LocPrcs In "
    End If
    
    ADC("UOWRKFR").CmpStr = ""
    ADC("UOWRKTO").CmpStr = ""
    ADC("UOWRKSEL").CmpStr = ""
    wrepcnd = ADC.RepCond
    wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    wsFgCnd = wCnd
    
    
    grpflds0 = GRP_REP.GrpFldLst(0)
    grpflds1 = GRP_REP.GrpFldLst(1)

' **** Zubin 211 (ws_BagJoin added) **** '
    ms_SqlStr = "Set DateFormat DMY" + " SELECT " + grpflds0 + ", " + wsCmnFlds1 + IIF(wsTxnFlds <> "", ", ", "") + wsTxnFlds + ", " + wsCmnFlds2 + _
      " FROM" + wsTxnJoin + ws_BagJoin + wsCmnJoin + _
      " WHERE" + wsTxnWhere + wsTxnCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
      IIF(UCase$(ADC("UOBAGPCS")) = UCase$("N"), ") UNION ALL" + _
      " SELECT " + grpflds1 + ", " + wsCmnFlds1 + ", " + wsFgFlds + ", " + wsCmnFlds2 + _
      " FROM" + wsFgJoin + wsCmnJoin + _
      " WHERE" + wsFgWhere + wsFgCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, ""), ")")
' **** Zubin 211 (ws_BagJoin added) **** '

    If ADC("UoBagPcs") = "N" Then Rep.TxtHead.SetText "WIP Bag Pieces Report Back Dated (Normal)"
    If ADC("UoBagPcs") = "R" Then Rep.TxtHead.SetText "WIP Bag Pieces Report Back Dated (Rejected)"
    Rep.TxtLocFr.SetText ADC("UoIssLocFr")
    Rep.TxtLocTo.SetText ADC("UoIssLocTo")
    'Rep.TxtLocSel.SetText adc("UoIssLocSel") 22 Loc sel removed from header

  ElseIf UCase(ADC.MenuCd) = UCase("RepWipFgBPcsBkDt") Then
    
    ADC("UoIssLocFr").CmpStr = " FgToBLoc >= "
    ADC("UoIssLocTo").CmpStr = " FgToBLoc <= "
    ADC("UoIssLocSel").CmpStr = " FgToBLoc In "
    
    ADC("UoPrntCdFr").CmpStr = "": ADC("UoPrntCdTo").CmpStr = "": ADC("UoPrntCdSel").CmpStr = ""
    '***** Manali 3.5.0 - 20/12/08 - following Code was commented earlier, now again included in query
    ADC("UOFGSUBLOCFR").CmpStr = " ISNULL(FmdToFgSubLoc, FdSubLoc) >= "
    ADC("UOFGSUBLOCTO").CmpStr = " ISNULL(FmdToFgSubLoc, FdSubLoc) <= "
    ADC("UOFGSUBLOCSEL").CmpStr = " ISNULL(FmdToFgSubLoc, FdSubLoc) In "
    '***** Manali 3.5.0 - 20/12/08 - above Code was commented earlier, now again included in query
    ADC("UOWRKFR").CmpStr = ""
    ADC("UOWRKTO").CmpStr = ""
    ADC("UOWRKSEL").CmpStr = ""
    
    wrepcnd = ADC.RepCond
    wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    wsTxnCnd = wCnd
    
    grpflds1 = GRP_REP.GrpFldLst(1)
    '****** Sachin 3.02 - Id fields in Joins
    ' ***** Manali 3.5.0 - 20/12/08 - LEFT OUTER JOIN with Fmd was removed ealier from query, now its again included
    ' ***** Manali 3.10.0 - 14/05/12 - Jt Conversion bags exluded in open bags option
    ' ***** Manali 3.10.0 - 14/05/12 - Error rectified - Join with Fmd changed
    '578-34 STR(FdKey, 7) changed to STR(FdKey, 8) and STR(FmdKey, 7) changed to STR(FmdKey, 8)
    ms_SqlStr = "Set DateFormat DMY" + " SELECT " + grpflds1 + ", " + wsCmnFlds1 + ", FgToBLoc AS BLoc, TdWrk AS BWrk," + _
      " CAST(FdGrWt AS DECIMAL(10,3)) BGrWt, 0 AS BDlDays, FdQty AS qQty, 0.00 As qSalVal, 0 AS qPtQty," + wsCmnFlds2 + _
      " FROM Fgd JOIN " + wsFgTbl + " F ON" + wsFgTblJoin + " JOIN " + wsPrmTbl + " P ON PMCD = FdTc" + _
      " JOIN Bag ON " + IIF(gs_Partition = ctCurrPrtn, " FdPrtKey=BPrtKey and ", "") + " BCoCd= FdCoCd AND BYy= FdBYy AND BChr= FdBChr AND BNo= FdBNo" + _
      " AND (FdYy + STR(FdKey, 8)) = (SELECT MAX(FdYy + STR(FdKey, 8)) FROM Fgd JOIN " + wsFgTbl + " F ON" + wsFgTblJoin + _
      " WHERE " + IIF(gs_Partition = ctCurrPrtn, " FdPrtKey=BPrtKey and ", "") + " FdCoCd= BCoCd AND FdBYy= BYy AND FdBChr= BChr AND FdBNo= BNo And FgDt" + wsDtCnd + ")" + _
      wsFgTxnjoin + wsCmnJoin + _
      " LEFT OUTER JOIN Fmd ON" + _
      " FdYy = FmdFdYy And FdKey = FmdFdKey And (FmdYy + STR(FmdKey, 8)) = (SELECT MAX(FmdYy + STR(FmdKey, 8))" + _
      " FROM Fmd JOIN Fm ON FmCoCd = FmdCoCd And FmTc = FmdTc And FmYy = FmdYy And FmChr = FmdChr And FmNo = FmdNo And FmdFmIdNo=FmIdNo " + IIF(gs_Partition = ctCurrPrtn, " And FmdPrtKey=FmPrtKey ", "") + _
      " WHERE FmdCoCd= FdCoCd AND FmdFdYy= FdYy AND FmdFdKey= FdKey " + IIF(gs_Partition = ctCurrPrtn, " And FmdPrtKey=FdPrtKey ", "") + _
      " And FmDt" + wsDtCnd + ")" + _
      " WHERE " + IIF(gs_Partition = ctCurrPrtn, " FgPrtKey='" + ctCurrPrtn + "' and ", "") + " PValue = 'FB' " + IIF(ADC("UOYN") <> "", "AND " + IIF(UCase$(ADC("UOYN")) = UCase$("Y"), " Not ", "") + " ( Exists (Select 1 " + _
      " From InvFgd JOIN InvHd On InIdNo=IfInIdNo and InCoCd = IfCoCd And InTc = IfTc And InYy = IfYy And InChr = IfChr And InNo =IfNo and IfSr<>0 " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IfPrtKey ", "") + _
      " Where (Select PValue from Param Where PTyp='TC' and PMCd=IfTc)='" + ctInTcTyp + "' and InCoCd = FdCoCd AND IfFdBYy = FdBYy And IfFdBChr = FdBChr And IfFdBNo = FdBNo " + IIF(gs_Partition = ctCurrPrtn, " and IfPrtKey=FdPrtKey ", "") + " And InDt" + wsDtCnd + ") or (BCnvJtYn='Y' and BClsDt<= '14/05/12'))", "") + _
      " And FgDt" + wsDtCnd + " And BOpnDt" + wsDtCnd + wCnd + _
      IIF(wAddlCnd <> "", " And " + wAddlCnd, "")
    

    Rep.TxtHead.SetText "Fg Bag Pieces Report Back Dated"
    Rep.TxthLocFr.SetText "Loc :": Rep.TxthLocTo.SetText " ": 'Rep.TxthLocSel.SetText " " 22 Loc sel removed from header
    Rep.TxtLocFr.SetText "PFG"
  
  End If
  '22 Customer From & To added to header
  Rep.TxtCustFr.SetText ADC("UoCmCdFr"): Rep.TxtCustTo.SetText ADC("UoCmCdTo")
End Sub
