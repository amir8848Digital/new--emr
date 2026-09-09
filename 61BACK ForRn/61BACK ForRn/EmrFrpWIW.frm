VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrpWIW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "What Is Where"
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
      Left            =   10710
      TabIndex        =   73
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
      TabIndex        =   74
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
      Left            =   4665
      TabIndex        =   83
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   84
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
         TabIndex        =   85
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
      Left            =   5880
      TabIndex        =   82
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   79
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
      TabIndex        =   75
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
      Height          =   9705
      Left            =   0
      TabIndex        =   78
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   77
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpWIW.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group&Sort"
         TabPicture(1)   =   "EmrFrpWIW.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9030
            Index           =   1
            Left            =   -74940
            TabIndex        =   81
            Top             =   360
            Width           =   15090
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   76
               Top             =   360
               Width           =   9435
               _ExtentX        =   16642
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9045
            Index           =   0
            Left            =   120
            TabIndex        =   80
            Top             =   360
            Width           =   15075
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
               Left            =   4720
               TabIndex        =   129
               ToolTipText     =   "Select Production or Process"
               Top             =   1200
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
               Left            =   1680
               TabIndex        =   128
               ToolTipText     =   "Select Production or Process"
               Top             =   1200
               Value           =   -1  'True
               Width           =   2085
            End
            Begin VB.Frame FraOptQty 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1725
               TabIndex        =   123
               ToolTipText     =   "Select Bag Qty/ No. of Bags"
               Top             =   7200
               Width           =   4455
               Begin VB.OptionButton OptQty 
                  Caption         =   "Bag Qty"
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
                  TabIndex        =   125
                  ToolTipText     =   "Select Bag Qty"
                  Top             =   0
                  Width           =   1695
               End
               Begin VB.OptionButton OptQty 
                  Caption         =   "No. of Bags"
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
                  Left            =   2950
                  TabIndex        =   124
                  ToolTipText     =   "Select Number of Bags"
                  Top             =   0
                  Width           =   1725
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Diamond Status "
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
               Index           =   6
               Left            =   30
               TabIndex        =   120
               Tag             =   "AppCol"
               ToolTipText     =   "Check To See Diamond Status (Report will be slow if checked)"
               Top             =   8640
               Width           =   1890
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Apply Colors       "
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
               Index           =   5
               Left            =   30
               TabIndex        =   72
               Tag             =   "AppCol"
               ToolTipText     =   "Check To Apply Colors"
               Top             =   8355
               Width           =   1890
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Remark     "
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
               Index           =   4
               Left            =   30
               TabIndex        =   70
               Tag             =   "ShowRem"
               ToolTipText     =   "Check To Show Special Remark"
               Top             =   8070
               Width           =   1890
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show 2nd Line   "
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
               Left            =   30
               TabIndex        =   66
               Tag             =   "ShowLine"
               ToolTipText     =   "Check To Show 2nd Line"
               Top             =   7515
               Width           =   1890
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Picture      "
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
               TabIndex        =   68
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   7800
               Width           =   1890
            End
            Begin VB.Frame FraOptDelDt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1725
               TabIndex        =   117
               ToolTipText     =   "Select Export Delivery Date Or Production Delivery Date"
               Top             =   6960
               Width           =   4455
               Begin VB.OptionButton OptDelDt 
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
                  Height          =   285
                  Index           =   1
                  Left            =   2950
                  TabIndex        =   64
                  ToolTipText     =   "Select Export Delivery Date Or Production Delivery Date"
                  Top             =   0
                  Width           =   1725
               End
               Begin VB.OptionButton OptDelDt 
                  Caption         =   "Exp Del Dt"
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
                  TabIndex        =   63
                  ToolTipText     =   "Select Export Delivery Date Or Production Delivery Date"
                  Top             =   0
                  Width           =   1695
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1725
               TabIndex        =   5
               ToolTipText     =   "Enter Location Sequence"
               Top             =   1485
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSEQCDFR"
               IdName          =   "UOSEQCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4720
               TabIndex        =   50
               ToolTipText     =   "Enter To Design Category"
               Top             =   5535
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
               Left            =   7590
               TabIndex        =   51
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   5550
               Width           =   7420
               _ExtentX        =   13097
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
               Left            =   1725
               TabIndex        =   49
               ToolTipText     =   "Enter From Design Category"
               Top             =   5535
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
               Left            =   4720
               TabIndex        =   18
               ToolTipText     =   "Enter To Design Code"
               Top             =   2370
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "OdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   7590
               TabIndex        =   19
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2370
               Width           =   7420
               _ExtentX        =   13097
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "OdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1725
               TabIndex        =   17
               ToolTipText     =   "Enter From Design Code"
               Top             =   2370
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "OdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4720
               TabIndex        =   30
               ToolTipText     =   "Enter To Customer"
               Top             =   3795
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
               Left            =   7590
               TabIndex        =   31
               ToolTipText     =   "Enter Customer Selection"
               Top             =   3810
               Width           =   7420
               _ExtentX        =   13097
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1725
               TabIndex        =   29
               ToolTipText     =   "Enter From Customer"
               Top             =   3795
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
               Left            =   4720
               TabIndex        =   39
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   4680
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
               Left            =   7590
               TabIndex        =   40
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   4695
               Width           =   7420
               _ExtentX        =   13097
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   1725
               TabIndex        =   38
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   4680
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
               Left            =   4720
               TabIndex        =   24
               ToolTipText     =   "Enter To Order Date"
               Top             =   2940
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
               Left            =   1725
               TabIndex        =   23
               ToolTipText     =   "Enter From Order Date"
               Top             =   2940
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
               Index           =   5
               Left            =   5500
               TabIndex        =   10
               ToolTipText     =   "Enter To Order Year"
               Top             =   1800
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
               Left            =   2475
               TabIndex        =   7
               ToolTipText     =   "Enter From Order Year"
               Top             =   1800
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
               Left            =   1725
               TabIndex        =   6
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1800
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
               Left            =   2895
               TabIndex        =   8
               ToolTipText     =   "Enter From Order Character"
               Top             =   1800
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
               Left            =   4720
               TabIndex        =   9
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1800
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
               Left            =   5920
               TabIndex        =   11
               ToolTipText     =   "Enter To Order Character"
               Top             =   1800
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
               Left            =   1725
               TabIndex        =   12
               ToolTipText     =   "Enter From Order Number"
               Top             =   2085
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
               Left            =   4720
               TabIndex        =   14
               ToolTipText     =   "Enter To Order Number"
               Top             =   2085
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
               Left            =   4720
               TabIndex        =   47
               ToolTipText     =   "Enter To Karat"
               Top             =   5250
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
               Left            =   7590
               TabIndex        =   48
               ToolTipText     =   "Enter Karat Selection"
               Top             =   5265
               Width           =   7420
               _ExtentX        =   13097
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "OdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1725
               TabIndex        =   46
               ToolTipText     =   "Enter From Karat"
               Top             =   5250
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               CmpStr          =   "OdKt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   5635
               TabIndex        =   15
               ToolTipText     =   "Enter To Order Serial"
               Top             =   2085
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
               Index           =   34
               Left            =   2640
               TabIndex        =   13
               ToolTipText     =   "Enter From Order Serial"
               Top             =   2085
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
               Index           =   1
               Left            =   4720
               TabIndex        =   21
               ToolTipText     =   "Enter To Priority Code"
               Top             =   2655
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
               Index           =   2
               Left            =   7590
               TabIndex        =   22
               ToolTipText     =   "Enter Priority Code Selection"
               Top             =   2655
               Width           =   7420
               _ExtentX        =   13097
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRTCDSEL"
               IdName          =   "UOPRTCDSEL"
               CmpStr          =   "OdPrtCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1725
               TabIndex        =   20
               ToolTipText     =   "Enter From Priority Code"
               Top             =   2655
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
               Index           =   4
               Left            =   4720
               TabIndex        =   26
               ToolTipText     =   "Enter To Prd Delivery Date"
               Top             =   3225
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTTO"
               IdName          =   "UODELDTTO"
               CmpStr          =   "OdDelDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1725
               TabIndex        =   25
               ToolTipText     =   "Enter From Prd Delivery Date"
               Top             =   3225
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTFR"
               IdName          =   "UODELDTFR"
               CmpStr          =   "OdDelDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   4720
               TabIndex        =   59
               ToolTipText     =   "Enter To Balance Quantity"
               Top             =   6390
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOBALPRDTO"
               IdName          =   "UOBALPRDTO"
               CmpStr          =   "(OdPrdQty-  OdFgQty) <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   1725
               TabIndex        =   58
               ToolTipText     =   "Enter From Balance Quantity"
               Top             =   6390
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOBALPRDFR"
               IdName          =   "UOBALPRDFR"
               CmpStr          =   "(OdOrdQty- OdPrdQty)>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   2355
               TabIndex        =   65
               ToolTipText     =   "Show Second Line Of Detail? ('N', 'Y', Blank)"
               Top             =   7515
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               ReCalcParent    =   "UOYN6"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1725
               TabIndex        =   0
               ToolTipText     =   "Enter From Company Code"
               Top             =   270
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "OdCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4720
               TabIndex        =   53
               ToolTipText     =   "Enter To Sales Executive"
               Top             =   5820
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERTO"
               IdName          =   "UOSALPERTO"
               CmpStr          =   "OmCmSalPer <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1725
               TabIndex        =   52
               ToolTipText     =   "Enter From Sales Executive"
               Top             =   5820
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
               Index           =   14
               Left            =   6660
               TabIndex        =   62
               ToolTipText     =   "Show Export/ Production Delivery Date (Enter ""Y"" for Export and ""N"" for Production)"
               Top             =   6960
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
               Index           =   47
               Left            =   4720
               TabIndex        =   56
               ToolTipText     =   "Enter To Production Sequence"
               Top             =   6105
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   1725
               TabIndex        =   55
               ToolTipText     =   "Enter From Productiopn Sequence"
               Top             =   6105
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   7590
               TabIndex        =   54
               ToolTipText     =   "Enter Sales Executive Selection"
               Top             =   5835
               Width           =   7420
               _ExtentX        =   13097
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSALPERSEL"
               IdName          =   "UOSALPERSEL"
               CmpStr          =   "OmCmSalPer In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   7590
               TabIndex        =   57
               ToolTipText     =   "Enter Production Sequence Selection"
               Top             =   6120
               Width           =   7420
               _ExtentX        =   13097
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPMCDSEL"
               IdName          =   "UOPMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   4720
               TabIndex        =   28
               ToolTipText     =   "Enter To Export Delivery Date"
               Top             =   3510
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1TO"
               IdName          =   "UODELDT1TO"
               CmpStr          =   "OdExpDelDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1725
               TabIndex        =   27
               ToolTipText     =   "Enter From Export Delivery Date"
               Top             =   3510
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1FR"
               IdName          =   "UODELDT1FR"
               CmpStr          =   "OdExpDelDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2355
               TabIndex        =   67
               ToolTipText     =   "Show Picture? ( Enter 'Y'/'N')"
               Top             =   7800
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
               Index           =   35
               Left            =   1725
               TabIndex        =   60
               ToolTipText     =   "Specify Bag Type"
               Top             =   6675
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
               Index           =   36
               Left            =   4720
               TabIndex        =   61
               ToolTipText     =   "Enter Bag Character"
               Top             =   6675
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
               Index           =   37
               Left            =   4720
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   270
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "OdCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   1725
               TabIndex        =   3
               ToolTipText     =   "Enter Parent Location Sequence"
               Top             =   555
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   1725
               TabIndex        =   4
               ToolTipText     =   "Enter Location Sequence"
               Top             =   840
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   4
               DataField       =   "UOPRNTSEQCDFR"
               IdName          =   "UOPRNTSEQCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   7590
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   300
               Width           =   7420
               _ExtentX        =   13097
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "OdCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   3000
               TabIndex        =   42
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   4965
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   6000
               TabIndex        =   44
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   4965
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   7590
               TabIndex        =   45
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   4980
               Visible         =   0   'False
               Width           =   7420
               _ExtentX        =   13097
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "OdPoNo in"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   2355
               TabIndex        =   69
               ToolTipText     =   "Show Special Remarks? ( Enter 'Y'/'N')"
               Top             =   8070
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN4"
               IdName          =   "UOYN4"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   2355
               TabIndex        =   71
               ToolTipText     =   "Apply Colours? ( Enter 'Y'/'N')"
               Top             =   8355
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN5"
               IdName          =   "UOYN5"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   7590
               TabIndex        =   34
               ToolTipText     =   "Enter Customer Group Selection"
               Top             =   4095
               Width           =   7420
               _ExtentX        =   13097
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMGRPSEL"
               IdName          =   "UOCMGRPSEL"
               CmpStr          =   "CmGrp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   59
               Left            =   4720
               TabIndex        =   33
               ToolTipText     =   "Enter To Customer Group"
               Top             =   4095
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
               Index           =   60
               Left            =   1725
               TabIndex        =   32
               ToolTipText     =   "Enter From Customer Group"
               Top             =   4095
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
               Index           =   61
               Left            =   7590
               TabIndex        =   37
               ToolTipText     =   "Enter Customer Type Selection"
               Top             =   4395
               Width           =   7420
               _ExtentX        =   13097
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMTYPSEL"
               IdName          =   "UOCMTYPSEL"
               CmpStr          =   "CmTyp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   4720
               TabIndex        =   36
               ToolTipText     =   "Enter To Customer Type"
               Top             =   4395
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
               Index           =   63
               Left            =   1725
               TabIndex        =   35
               ToolTipText     =   "Enter From Customer Type"
               Top             =   4395
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
               Index           =   64
               Left            =   2355
               TabIndex        =   121
               ToolTipText     =   "Diamond Status Required? ( Enter 'Y'/'N') (if ""y"", Report will be slow)"
               Top             =   8640
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN6"
               ReCalcOn        =   "UOYN"
               IdName          =   "UOYN6"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   65
               Left            =   6660
               TabIndex        =   126
               ToolTipText     =   "Show Bag Qty/ No. of Bags (Enter ""Y"" for Bag Qty and ""N"" for No.of Bags)"
               Top             =   7200
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN7"
               IdName          =   "UOYN7"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   90
               Left            =   7440
               TabIndex        =   130
               ToolTipText     =   "Show Production or Process Location"
               Top             =   1200
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
               Index           =   66
               Left            =   1725
               TabIndex        =   41
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   4965
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
               Index           =   67
               Left            =   4720
               TabIndex        =   43
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   4965
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTTO"
               IdName          =   "UOODSUBCUSTTO"
               CmpStr          =   "OdSubCust <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   82
               Left            =   7590
               TabIndex        =   16
               ToolTipText     =   "Enter Order Number Selection"
               Top             =   2085
               Width           =   7420
               _ExtentX        =   13097
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOOMNOSEL"
               IdName          =   "UOOMNOSEL"
               CmpStr          =   "OdNo In"
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
               Index           =   23
               Left            =   2940
               TabIndex        =   132
               Top             =   4965
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
               Index           =   21
               Left            =   5940
               TabIndex        =   131
               Top             =   4965
               Width           =   105
            End
            Begin VB.Label LblShowQty 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Qty"
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
               Top             =   7200
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Note: Selecting the option as ‘Y’es , may slow down the processing of report.(Diamond Status On 2nd Line Of Detail Row)"
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
               Height          =   315
               Index           =   16
               Left            =   2880
               TabIndex        =   122
               Top             =   8640
               Width           =   11775
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
               TabIndex        =   119
               Top             =   4365
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
               Index           =   6
               Left            =   60
               TabIndex        =   118
               Top             =   4080
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
               TabIndex        =   116
               Top             =   4965
               Width           =   1425
            End
            Begin VB.Label LblLocSeq 
               BackStyle       =   0  'Transparent
               Caption         =   "Prnt Loc Seq"
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
               Height          =   255
               Left            =   60
               TabIndex        =   115
               Top             =   855
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Loc Seq"
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
               Height          =   255
               Index           =   20
               Left            =   60
               TabIndex        =   114
               Top             =   1500
               Width           =   1425
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
               TabIndex        =   113
               ToolTipText     =   "Location"
               Top             =   6675
               Width           =   1035
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
               Left            =   3780
               TabIndex        =   112
               ToolTipText     =   "Location"
               Top             =   6675
               Width           =   915
            End
            Begin VB.Label LblExpDelDt 
               BackStyle       =   0  'Transparent
               Caption         =   "Exp Del Date"
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
               Top             =   3510
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Sequence"
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
               TabIndex        =   110
               Top             =   6105
               Width           =   1485
            End
            Begin VB.Label LblPRdExpDt 
               BackStyle       =   0  'Transparent
               Caption         =   "Delivery Date"
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
               Height          =   225
               Left            =   60
               TabIndex        =   109
               ToolTipText     =   "Location"
               Top             =   6990
               Width           =   1845
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
               Index           =   8
               Left            =   60
               TabIndex        =   108
               Top             =   5820
               Width           =   1600
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
               TabIndex        =   107
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bal Qty"
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
               TabIndex        =   106
               Top             =   6390
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Del Date"
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
               TabIndex        =   105
               Top             =   3225
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
               TabIndex        =   104
               Top             =   2655
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
               Left            =   2580
               TabIndex        =   103
               Top             =   2085
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
               Left            =   5575
               TabIndex        =   102
               Top             =   2085
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order No/Sr"
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
               TabIndex        =   101
               Top             =   2085
               Width           =   1605
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
               TabIndex        =   100
               Top             =   5250
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
               TabIndex        =   99
               Top             =   1800
               Width           =   1725
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
               Left            =   5860
               TabIndex        =   98
               Top             =   1800
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
               Left            =   2835
               TabIndex        =   97
               Top             =   1800
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
               Left            =   5440
               TabIndex        =   96
               Top             =   1800
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
               Left            =   2415
               TabIndex        =   95
               Top             =   1800
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Ord Date"
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
               TabIndex        =   94
               Top             =   2940
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
               Left            =   60
               TabIndex        =   93
               Top             =   4680
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
               TabIndex        =   92
               Top             =   3795
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
               TabIndex        =   91
               Top             =   2370
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
               TabIndex        =   90
               Top             =   5535
               Width           =   1425
            End
            Begin VB.Label LblPrntYN 
               BackStyle       =   0  'Transparent
               Caption         =   "Prnt(Y/N) "
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
               Height          =   255
               Left            =   60
               TabIndex        =   89
               Top             =   540
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
               Left            =   7590
               TabIndex        =   88
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
               Left            =   4720
               TabIndex        =   87
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
               Left            =   1725
               TabIndex        =   86
               Top             =   0
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpWIW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepWIW
Dim Rep1 As New EmrRepWIWLong
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
Dim wi_ScopeCos As Integer, ws_ScoCnd As String
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  GRP_REP.Add "OrderNo", "OdCoCd+OdTc+OdYy+OdChr+str(OdNo)", "OdCoCd+'/'+OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))", "", "", "Company Code,Customer Code,Customer PO,Sales Exec,Customer Grp,Customer Typ", "", "", "'('+max(OmCmCd)+ '/'+ max(OmPONo)+ ')' "
  'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp= max(OdDmTcTyp) in DsgMst where condition
  GRP_REP.Add "Design Code", "max(OdDmCd)", "", "", "", "Design Ctg", "", "", "(Select Substring(DmDesc,1,120) From DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' and ", "") + " DmTcTyp= max(OdDmTcTyp) and DmCd= max(OdDmCd) and DmSz= '') "
  GRP_REP.Add "Priority", "max(OdPrtCd)", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'PRTCD' and PMCd= max(OdPrtCd)) "
  GRP_REP.Add "Ord Date", "convert(varchar(8), max(OmDt), 112)", "convert(varchar, max(OmDt), 103)", "", "", "", "", "", ""
  GRP_REP.Add "Prd Del Date", "convert(varchar(8), max(OdDelDt), 112)", "convert(varchar, max(OdDelDt), 103)", "", "", "", "", "", ""
  'IG.164 -Export delievery date should not visible for Production Short & Long reports
  If UCase(ADC.MenuCd) = UCase("RepWIWExp") Or UCase(ADC.MenuCd) = UCase("RepWIWLongExp") Then _
    GRP_REP.Add "Exp Del Date", "convert(varchar(8), max(OdExpDelDt), 112)", "convert(varchar, max(OdExpDelDt), 103)", "", "", "", "", "", ""
    
   '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
    'Sachin 4.1.0.0
  GRP_REP.Add "Customer Code", "max(OmCmCd)", "", "", "", "Sales Exec,Customer Grp,Customer Typ", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= max(OmCmCd)) "
  
  'Sachin - 4.1.0.0
  GRP_REP.Add "Customer Grp", "max(CmGrp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMGRP' and PmCd = Max(CmGrp))"
  GRP_REP.Add "Customer Typ", "max(CmTyp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMTYP' and PmCd = Max(CmTyp))"
  
  GRP_REP.Add "Customer PO", "max(OmCmCd)+max(OmPONo)", "'('+ max(OmCmCd)+ ') '+ LTrim(max(OmPONo))", "", "", "Company Code", "", "", "OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))"
  '*** Jay 3.2.0 [OdPoNo]
  'vk.19 Sub Customer added
  GRP_REP.Add "Cust SUB PO", "max(OmCmCd)+max(OmPONo)+max(OdSubCust)+max(OdPONo)", "'('+ max(OmCmCd)+ ') '+ LTrim(max(OmPONo))+(case when Max(OdPONo)<> '' Or Max(OdSubCust)<> '' then '/'+ LTrim(max(OdSubCust))+' : '+LTrim(Max(OdPONo)) else '' end )", "", "", "Customer Code", "", "", "OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))+'/'+LTrim(Str(OdSr))+(case when max(OdPONo)<> '' Or Max(OdSubCust)<> '' then '/'+LTrim(max(OdSubCust))+' : '+ LTrim(max(OdPONo)) else '' end )"
  '*** Jay 3.2.0 [OdPoNo]
  GRP_REP.Add "Ord Kt", "max(OdKt)", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(OdKt)) "
  GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) "
  'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp= max(OdDmTcTyp) in DsgMst where condition
  GRP_REP.Add "(DsgCd)Colour", "max(OdDmCd)+max(OdDmCol)", "'('+max(OdDmCd)+') '+max(OdDmCol)", "", "", "Design Ctg", "", "", "(Select Substring(DmDesc,1,120) From DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' and ", "") + " DmTcTyp= max(OdDmTcTyp) and DmCd= max(OdDmCd) and DmSz= '') "
  
  'Sachin 3.11.0
  GRP_REP.Add "Design Colour", "max(OdDmCol)", "max(OdDmCol)", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol))"
  'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp= max(OdDmTcTyp) in DsgMst where condition
  GRP_REP.Add "(DsgCd)Size", "max(OdDmCd)+max(OdDmSz)", "'('+max(OdDmCd)+') '+max(OdDmSz)", "", "", "Design Ctg", "", "", "(Select Substring(DmDesc,1,120) From DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' and ", "") + " DmTcTyp= max(OdDmTcTyp) and DmCd= max(OdDmCd) and DmSz= '') "
  
  GRP_REP.Add "Sales Exec", "max(OmCmSalPer)", "", "", "", "", "", "", "(Select PDesc from Param Where PTyp = 'SP' and PMCd= max(OmCmSalPer)) "
  '/** urmi 2.06 PrdSeq*//
  GRP_REP.Add "Ord PrdSeq", "max(OdPrdSeq)", "", "", "", "", "", "", "(Select vPDesc from vParam where vPTyp= 'PrdSeq' and vPMCd= max(OdPrdSeq) and vPCoCd= max(OdCoCd)) "
  GRP_REP.Add "Ord PrcsSeq", "max(vPValue1)", "", "", "", "", "", "", "(Select vPDesc from vParam where vPTyp= 'PrcsSeq' and vPMCd= max(vp.vPValue1) and vPCoCd= max(OdCoCd)) "
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******
  If ws_MultiCoMod = True Then _
    GRP_REP.Add "Company Code", "OdCoCd", "", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=OdCoCd) "
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******
  
'  GRP_REP.Add "OrderNo", "OdTc+OdYy+OdChr+str(OdNo)", "OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))", "wOrdNo,OmCmCd,OmPoNo", "hOrdNo,hOmCmCd,hOmPoNo", "Customer Code,Customer PO", "", "", "'('+max(OmCmCd)+ '/'+ max(OmPONo)+ ')' "
'  GRP_REP.Add "Design Code", "max(OdDmCd)", "", "OdDmCd", "hOdDmCd", "Design Ctg", "", "", "(Select Substring(DmDesc,1,120) From DsgMst Where DmTcTyp= 'DM' and DmCd= max(OdDmCd) and DmSz= '') "
'  GRP_REP.Add "Priority", "max(OdPrtCd)", "", "OdPrtCd", "hOdPrtCd", "", "", "", "(Select PDesc from Param where PTyp= 'PRTCD' and PMCd= max(OdPrtCd)) "
'  GRP_REP.Add "Ord Date", "max(OmDt)", "convert(varchar, max(OmDt), 103)", "OmDt", "hOmDt", "", "", "", ""
'  GRP_REP.Add "Del Date", "max(OdDelDt)", "convert(varchar, max(OdDelDt), 103)", "OdDelDt", "hOdDelDt", "", "", "", ""
'  GRP_REP.Add "Customer Code", "max(OmCmCd)", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= max(OmCmCd)) "
'  GRP_REP.Add "Customer PO", "max(OmCmCd)+max(OmPONo)", "'('+ max(OmCmCd)+ ') '+ LTrim(max(OmPONo))", "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "", "", "", "OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))"
'  GRP_REP.Add "Ord Kt", "max(OdKt)", "", "OdKt", "hOdKt", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(OdKt)) "
'  GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "DmCtg", "hDmCtg", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) "

' **** Zubin 211 **** '
  'GRP_REP.Add "Bag Character", "max(BChr)", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'BCHR' and PMCd= max(BChr) and PSCd= '') "
' **** Zubin 211 **** '
End Sub

Private Sub adc_setreprecsource()
'*** (Jen 2.12)
On Error GoTo RepErr
'*** (Jen 2.12)
  '*** Jay 3.2.0 [PicNm]
  ' ***** Manali 3.5.0 - 10/11/08 - Spl Rem section added [OdSalRem, qDetpos, qDispalRemYN], Colors Applied [qColNm, qColCrCd, qColAppYN]
  '*** Report Sql For WIW (Prd & Exp) ***               '****Geeta***New Fld DmTcTyp**Emr208
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr,DmTcTyp, OdDmCd, OmDt, OmCmCd, OdPrtCd,
  'OdKt, OdDelDt, OdSfx, OdDmSz, DmCtg, OdDmCol, OdPicNm, space(30) as qUniqRmCd,
  'OdPrdSeq,OdOrdQty, OdOrdQty as qBalQty, BQty as qFlrQty, 0.0 as qQty1,
  '0.0 as qQty2, 0.0 as qQty3, 0.0 as qQty4, 0.0 as qQty5, 0.0 as qQty6,
  '0.0 as qQty7, 0.0 as qQty8, 0.0 as qQty9, 0.0 as qQty10, 0.0 as qQty11,
  'OdSalRem, qDetpos, qDispalRemYN, qColNm, qColCrCd, qColAppYN
  'from OrdDsg, OrdMst, DsgMst, Bag where 1=2
  '*** Report Sql For WIW (Prd & Exp) ***
  
  '*** Report Sql For WIWLong (Prd & Exp) ***
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr,DmTcTyp,OdDmCd, OmDt, OmCmCd, OdPrtCd,
  'OdKt, OdDelDt, OdSfx, OdDmSz, DmCtg, OdDmCol, space(30) as qUniqRmCd,
  'OdPrdSeq,OdOrdQty, OdOrdQty as qBalQty, BQty as qFlrQty, 0.0 as qQty1,
  '0.0 as qQty2, 0.0 as qQty3, 0.0 as qQty4, 0.0 as qQty5, 0.0 as qQty6,
  '0.0 as qQty7, 0.0 as qQty8, 0.0 as qQty9, 0.0 as qQty10, 0.0 as qQty11,
  '0.0 as qQty12, 0.0 as qQty13, 0.0 as qQty14, 0.0 as qQty15,
  'OdSalRem, qDetpos, qDispalRemYN, qColNm, qColCrCd, qColAppYN
  'from OrdDsg, OrdMst, DsgMst, Bag where 1=2
  '*** Report Sql For WIWLong (Prd & Exp) ***
  
  '*** Sub Rep Sql For WIWLong ***
  'Select OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCd, 0 as qSrt
  'from OrdRm where 1= 2
  '*** Sub Rep Sql For WIWLong ***
  
  Dim wOrdDtStr As String
  Dim grpflds As String, wrepcnd As String, wSel As String, wCnd As String, wStr As String
  Dim wSqlStrg As String, wFldLst As String, wWsQty As String, wHsQty As String
  
  Dim ws_FGStr As String    '*** (Jen 2.14 Next 11/05/07)
  
' **** Zubin 211 **** '
  Dim ws_BagJoin As String
  Dim ws_BChrCnd As String
' **** Zubin 211 **** '
  
  Dim wAddlCnd As String  '****** Sachin 3.02.0
  
  Dim wDelDt As String    ' **** Manali 3.5.0 - 12/11/08 - Del Dt to be set for applying colors
  Dim wNonePos As Integer, wOrdNoPos As Integer, wSec As String   ' ***** Manali 3.6.0 - 12/10/09 - Apply Colors for OrdNo Group Sort
  
  Set Rep = Nothing
  Set Rep1 = Nothing
  'Set ADC.RepSource = Rep
  
  If UCase(ADC.MenuCd) = UCase("RepWIWPrd") Or UCase(ADC.MenuCd) = UCase("RepWIWExp") Then
    Set ADC.RepSource = Rep
    Set mRep = Rep
  ElseIf UCase(ADC.MenuCd) = UCase("RepWIWLongPrd") Or UCase(ADC.MenuCd) = UCase("RepWIWLongExp") Then
    Set ADC.RepSource = Rep1
    Set mRep = Rep1
  End If
  Call DispCoNm
  Call SetMwName(mRep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
  wrepcnd = ADC.RepCond
  
    '*** To check if Detail is selected ***
  Dim wGrp(3) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 3
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
  Next i
  
  ' ***** Manali 3.6.0 - 12/10/09 - Apply Colors for OrdNo Group Sort
  If wDetPos = 0 Then
    For i = 1 To 3
      If wGrp(i) = "OrderNo" Then
        wOrdNoPos = i: Exit For
      End If
    Next i
    
    If wOrdNoPos <> 0 Then
      For i = 1 To 3
        If wGrp(i) = "(None)" Then
          wNonePos = i: Exit For
        End If
      Next i
    End If
  End If
  
  If wDetPos = 0 Then
    If wOrdNoPos = 1 And wNonePos = 2 Then
      wSec = "G3F"
    ElseIf wOrdNoPos = 1 And wNonePos = 3 Then
      wSec = "G2F"
    ElseIf wOrdNoPos = 1 And wNonePos = 0 Then
      wSec = "G1F"
    ElseIf wOrdNoPos = 2 And wNonePos = 3 Then
      wSec = "G3F"
    ElseIf wOrdNoPos = 2 And wNonePos = 0 Then
      wSec = "G2F"
    ElseIf wOrdNoPos = 3 Then
      wSec = "G3F"
    End If
  End If
  ' ***** Manali 3.6.0 - 12/10/09 - Apply Colors for OrdNo Group Sort
  
  ' **** CRM 2.12 - 23-02-06 **** '
  Dim wCrmCnd As String
  wCrmCnd = " And OmSubmitYN <> 'N'"
  ' **** CRM 2.12 - 23-02-06 **** '
  
  If UCase(ADC.MenuCd) = UCase("RepWIWPrd") Then
    wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    wSel = " max(OdPrdQty) as OdOrdQty, max(OdPrdQty- OdFgQty) as qBalQty, "
    'uma for No of bag
    If ADC("UOYN7") = "N" Then wSel = " 0 as OdOrdQty, 0 as qBalQty, "
'    wStr = " where OdPrdQty> OdFgQty and IsNull(BLoc,'')<> '" + ctFgLoc + "' "
    
    'wStr = " and IsNull(BLoc,'')<> '" + ctFgLoc + "' where OdPrdQty> OdFgQty "   '*** (Bef 11/05/07)
    '*** (Jen 2.14 Next 11/05/07)
    ws_FGStr = " and IsNull(BLoc,'')<> '" + ctFgLoc + "' "
    wStr = " where OdPrdQty> OdFgQty "
    '*** (Jen 2.14 Next 11/05/07)
    'Loc to Prcs-Report tilte & Report field list based on Option button selection
    If OptLoc(0).Value = True Then
      mRep.TxtHead.SetText "What Is Where (Production)"
      wFldLst = FldLst("WIWSEQ", 11)
    Else
      mRep.TxtHead.SetText "What Is Where (Process)"
      wFldLst = FldLst("RWIWSEQ", 11)
    End If
    
    wDelDt = "OdDelDt"      ' ***** Manali 3.5.0 - 12/11/08 - Colors applied
  ElseIf UCase(ADC.MenuCd) = UCase("RepWIWLongPrd") Then
    wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    wSel = " max(OdPrdQty) as OdOrdQty, max(OdPrdQty- OdFgQty) as qBalQty, "
    'uma for No of bag
    If ADC("UOYN7") = "N" Then wSel = " 0 as OdOrdQty, 0 as qBalQty, "
'    wStr = " where OdPrdQty> OdFgQty and IsNull(BLoc,'')<> '" + ctFgLoc + "' "
    
    'wStr = " and IsNull(BLoc,'')<> '" + ctFgLoc + "' where OdPrdQty> OdFgQty "    '*** (Bef 11/05/07)
    '*** (Jen 2.14 Next 11/05/07)
    ws_FGStr = " and IsNull(BLoc,'')<> '" + ctFgLoc + "' "
    wStr = " Where OdPrdQty> OdFgQty "
    '*** (Jen 2.14 Next 11/05/07)
    'Loc to Prcs-Report tilte & Report field list based on Option button selection
    If OptLoc(0).Value = True Then
      mRep.TxtHead.SetText "What Is Where Long (Production)"
      wFldLst = FldLst("WIWLONG", 15)
    Else
      mRep.TxtHead.SetText "What Is Where Long (Process)"
      wFldLst = FldLst("RWIWLONG", 15)
    End If
        
    wDelDt = "OdDelDt"      ' ***** Manali 3.5.0 - 12/11/08 - Colors applied
  ElseIf UCase(ADC.MenuCd) = UCase("RepWIWExp") Then
    wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    wSel = " max(OdOrdQty) as OdOrdQty, max(OdOrdQty- OdExpQty) as qBalQty, "
    ws_FGStr = " "    '*** (Jen 2.14 Next 11/05/07)
    wStr = " where OdOrdQty> OdExpQty "
    'Loc to Prcs-Report tilte & Report field list based on Option button selection
    If OptLoc(0).Value = True Then
      mRep.TxtHead.SetText "What Is Where (Export)"
      wFldLst = FldLst("WIWSEQ", 11)
    Else
      mRep.TxtHead.SetText "What Is Where (Export)"
      wFldLst = FldLst("RWIWSEQ", 11)
    End If
        
    wDelDt = "OdExpDelDt"      ' ***** Manali 3.5.0 - 12/11/08 - Colors applied
  ElseIf UCase(ADC.MenuCd) = UCase("RepWIWLongExp") Then
    wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    wSel = " max(OdOrdQty) as OdOrdQty, max(OdOrdQty- OdExpQty) as qBalQty, "
    ws_FGStr = " "    '*** (Jen 2.14 Next 11/05/07)
    wStr = " where OdOrdQty> OdExpQty "
    'Loc to Prcs-Report tilte & Report field list based on Option button selection
    If OptLoc(0).Value = True Then
      mRep.TxtHead.SetText "What Is Where Long (Export)"
      wFldLst = FldLst("WIWLONG", 15)
    Else
      mRep.TxtHead.SetText "What Is Where Long (Export)"
      wFldLst = FldLst("RWIWLONG", 15)
    End If
       
    wDelDt = "OdExpDelDt"      ' ***** Manali 3.5.0 - 12/11/08 - Colors applied
  End If
  
  'Loc to Prcs-If user choose location then selecting OdPrdSeq,
  'if user choose Procss getting Prdseq by Prcs-Loc seq. link in vparam
  If OptLoc(0).Value = True Then
    If ADC("UOPMCDFR") <> "" Then wCnd = wCnd + " and OdPrdSeq >= '" + ADC("UOPMCDFR") + "'"
    If ADC("UOPMCDTO") <> "" Then wCnd = wCnd + " and OdPrdSeq <= '" + ADC("UOPMCDTO") + "'"
    If ADC("UOPMCDSEL") <> "" Then wCnd = wCnd + " and OdPrdSeq In( '" + ADC("UOPMCDSEL") + "')"
  Else
    If ADC("UOPMCDFR") <> "" Then wCnd = wCnd + " and vPValue1 >= '" + ADC("UOPMCDFR") + "'"
    If ADC("UOPMCDTO") <> "" Then wCnd = wCnd + " and vPValue1 <= '" + ADC("UOPMCDTO") + "'"
    If ADC("UOPMCDSEL") <> "" Then wCnd = wCnd + " and vPValue1 In ('" + ADC("UOPMCDSEL") + "')"
  End If
  
  If ADC("UOYN3") = "Y" Then
    wOrdDtStr = "OdExpDelDt"
    mRep.TxtOdDelDt.SetText "Exp Del Dt"
  ElseIf ADC("UOYN3") = "N" Then
    wOrdDtStr = "OdDelDt"
    mRep.TxtOdDelDt.SetText "Prd Del Dt"
  End If

  ' ***** Manali 3.5.0 - 25/11/08 - CustDsgCd added
  Dim ws_CustDmCdStr As String
  ws_CustDmCdStr = "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                  "Where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = Max(OmCmCd) and " + _
                  "CdOurDmCd = Max(OdDmCd) and CdOurSfx = Max(OdSfx) and " + _
                  "CdOurDmSz = Max(OdDmSz)), " + _
                  "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                      "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = Max(OmCmCd) and " + _
                      "CdOurDmCd = Max(OdDmCd) and CdOurSfx = Max(OdSfx) and " + _
                      "CdOurDmSz = '' ), " + _
                    "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                          "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = Max(OmCmCd) and " + _
                          "CdOurDmCd = Max(OdDmCd) and CdOurSfx = '' and " + _
                          "CdOurDmSz = Max(OdDmSz) ), " + _
                        "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                              "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = Max(OmCmCd) and " + _
                              "CdOurDmCd = Max(OdDmCd) and CdOurSfx = '' and " + _
                              "CdOurDmSz= ''), '')))) as qCustDmCd, "


' **** Zubin 211 **** '
'''''  wSqlStrg = "Select " + grpflds + ", " + _
             "OdTc, OdYy, OdChr, OdNo, OdSr, max(DmTcTyp) as DmTcTyp, max(OdDmCd) as OdDmCd, max(OmDt) as OmDt, " + _
             "max(OmCmCd) as OmCmCd, " + _
             "max(OdPrtCd) as OdPrtCd, max(OdKt) as OdKt, max(" + wOrdDtStr + ") as OdDelDt, max(OdSfx) as OdSfx, " + _
             "max(OdDmSz) as OdDmSz, max(DmCtg) as DmCtg, " + _
             "max(OdDmCol) as OdDmCol, space(30) as qUniqRmCd, max(OdPrdSeq) as OdPrdSeq, " + wSel + _
             " Sum(IsNull(BQty, 0)) as qFlrQty " + _
             wFldLst + _
             " from OrdDsg join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy " + _
             "and OmChr= OdChr and OmNo= OdNo " + _
             "Join CustMst on CmCtg= 'C' and CmCd = OmCmCd " + _
             "join DsgMst on DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + _
             "left outer join Bag on BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy " + _
             "and BOdChr= OdChr and BOdNo= OdNo and BOdSr= OdSr and BCls= 'N' " + _
             wStr + wCnd + _
             " Group By OdTc, OdYy, OdChr, OdNo, OdSr "
             
  If ADC("UOBCHRFR") <> "" Then
    ws_BChrCnd = " having MAX(BChr)= '" + ADC("UOBCHRFR") + "'"
  Else
    ws_BChrCnd = " "
  End If
  
  '*** Jay 2.13(CT) ***
  'uma
  'ws_BagJoin = IIF(UCase(adc("UoBagTyp")) = "", " (Select PValue from Param where PTyp= 'BCHR' and " + _
                    "PMCd= IsNull(BChr, '') and PSCd= '')<> 'D' ", " and (Select PValue from Param where " + _
                    "PTyp= 'BCHR' and PMCd= IsNull(BChr, '') and PSCd= '')= '" + UCase(adc("UoBagTyp")) + "' And BChr= '" + adc("UOBCHRFR") + "' ")
  
'  <> 'D' ", " and PValue = '" + UCase(adc("UoBagTyp")) + "' And BChr= '" + adc("UOBCHRFR") + "'")
  
  
  ws_BagJoin = IIF(UCase(ADC("UoBagTyp")) = "P", " and (Select PValue from Param where PTyp= 'BCHR' and " + _
                    "PMCd= IsNull(BChr, '') and PSCd= '')= 'P' ", " and (Select PValue from Param where " + _
                    "PTyp= 'BCHR' and PMCd= IsNull(BChr, '') and PSCd= '')= '" + UCase(ADC("UoBagTyp")) + "' And BChr= '" + ADC("UOBCHRFR") + "' ")
                    
  '*** Jay 2.13(CT) ***

  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(ADC("UOCMCDFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(ADC("UOCMCDTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(ADC("UOCMCDSEL")) + ")", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields

' ****** Bhavna emr 5.1.0 Wax set and hand set diamonds available for ..If diamond status Y

If ADC("UoYN6") = "Y" Then
wHsQty = ", IsNull((select sum(BQty) from Bag " + _
"Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd='' where (PValue= 'P' or PValue = 'F') " + _
"and BOdIdNo=OdIdNo and BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr and " + _
"BOdNo= OdNo and BOdSr= OdSr and " + _
"(BCls= 'Y' or BLoc= 'PFG'  or " + _
"Exists (select Top 1 TdRmWt from Txnd where TdPrtKey=BPrtKey and  TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and " + _
"TdBChr= BChr and TdBNo= BNo and TdRmCd=" + _
"(select Top 1 OrRmCd from ordrm where  bOdTc= OrTc and bOdYy= OrYy and bOdChr= OrChr and bOdNo= OrNo and" + _
"  bOdSr= OrSr and bOdIdNo = OrOdIdNo and orHsQty<>0) ))),0) as qHsQty, 'HDia' as qhHs "

wWsQty = ", IsNull((select sum(BQty) from Bag " + _
"Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd='' where (PValue= 'P' or PValue = 'F') " + _
"and BOdIdNo=OdIdNo and BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr and " + _
"BOdNo= OdNo and BOdSr= OdSr and " + _
"(BCls= 'Y' or BLoc= 'PFG'  or " + _
"Exists (select Top 1 TdRmWt from Txnd where TdPrtKey=BPrtKey and  TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and " + _
"TdBChr= BChr and TdBNo= BNo and TdRmCd=" + _
"(select Top 1 OrRmCd from ordrm where  bOdTc= OrTc and bOdYy= OrYy and bOdChr= OrChr and bOdNo= OrNo and" + _
"  bOdSr= OrSr and bOdIdNo = OrOdIdNo and orWsQty<>0) ))),0) as qWsQty, 'WDia' as qhWs "

'"join OrdRm on OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and OdChr= OrChr and OdNo= OrNo " + _
"and OdSr= OrSr and OdIdNo = OrOdIdNo and orhsqty<>0 " + _
"where BOdIdNo=OdIdNo and BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr and " + _
"BOdNo= OdNo and BOdSr= OdSr and " + _
"(BCls= 'Y' or BLoc= 'PFG'  or " + _
"Exists (select Top 1 'x' from Txnd where TdPrtKey=BPrtKey and  TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and " + _
"TdBChr= BChr and TdBNo= BNo and TdRmCd= OrRmCd and tdRmWt>0) ) ),0)   as qHsQty, 'HDia' as qhHs "

'select sum(BQty) from Bag
'Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd=''
'where (PValue= 'P' or PValue = 'F') and
'Bno=44974 and byy=14 and (BCls= 'Y' or BLoc= 'PFG'  or
'exists(select top 1 Tdrmwt from Txnd where TdPrtKey=BPrtKey and
'TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo and TdRmCd=
'(select Top 1 OrRmCd from ordrm where  bOdTc= OrTc and bOdYy= OrYy and bOdChr= OrChr and bOdNo= OrNo and
'bOdSr= OrSr and bOdIdNo = OrOdIdNo and orHsQty<>0)  ))
    
    
'wWsQty = ", isNull((select sum(BQty) from Bag " + _
"Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd='' and (PValue= 'P' or PValue = 'F') " + _
"join OrdRm on OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and OdChr= OrChr and OdNo= OrNo " + _
"and OdSr= OrSr and OdIdNo = OrOdIdNo and orWsQty<>0 " + _
"where BOdIdNo=OdIdNo and BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr and " + _
"BOdNo= OdNo and BOdSr= OdSr and " + _
"(BCls= 'Y' or BLoc= 'PFG'  or " + _
"Exists (select Top 1 'x' from Txnd where TdPrtKey=BPrtKey and  TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and " + _
"TdBChr= BChr and TdBNo= BNo and TdRmCd= OrRmCd and TdRmWt>0) ) ),0)   as qWsQty, 'WDia' as qhWs "

'wHsQty = ", IsNull((select sum(BQty) from Bag " + _
"Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd='' and (PValue= 'P' or PValue = 'F') " + _
"join OrdRm on OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and OdChr= OrChr and OdNo= OrNo " + _
"and OdSr= OrSr and OdIdNo = OrOdIdNo and orhsqty<>0 " + _
"where BOdIdNo=OdIdNo and BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr and " + _
"BOdNo= OdNo and BOdSr= OdSr and " + _
"(BCls= 'Y' or BLoc= 'PFG'  or (select sum(tdrmwt* " + _
"( Case when TdFrRmLoc='W' and TdFrRmDc='D' then 1 " + _
"when TdFrRmLoc='W' and TdFrRmDc='C' then -1 " + _
"when TdToRmLoc='W' and TdToRmDc='D' then 1 " + _
"when TdToRmLoc='W' and TdToRmDc='C' then -1 Else 0 end) " + _
") from Txnd where TdPrtKey=BPrtKey and  TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and " + _
"TdBChr= BChr and TdBNo= BNo and TdRmCd= OrRmCd) > 0 ) ),0)   as qHsQty, 'HDia' as qhHs "
    
'wWsQty = ", isNull((select sum(BQty) from Bag " + _
"Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd='' and (PValue= 'P' or PValue = 'F') " + _
"join OrdRm on OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and OdChr= OrChr and OdNo= OrNo " + _
"and OdSr= OrSr and OdIdNo = OrOdIdNo and orWsQty<>0 " + _
"where BOdIdNo=OdIdNo and BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr and " + _
"BOdNo= OdNo and BOdSr= OdSr and " + _
"(BCls= 'Y' or BLoc= 'PFG'  or (select sum(tdrmwt* " + _
"( Case when TdFrRmLoc='W' and TdFrRmDc='D' then 1 " + _
"when TdFrRmLoc='W' and TdFrRmDc='C' then -1 " + _
"when TdToRmLoc='W' and TdToRmDc='D' then 1 " + _
"when TdToRmLoc='W' and TdToRmDc='C' then -1 Else 0 end) " + _
") from Txnd where TdPrtKey=BPrtKey and  TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and " + _
"TdBChr= BChr and TdBNo= BNo and TdRmCd= OrRmCd) > 0 ) ),0)   as qWsQty, 'WDia' as qhWs "

Else
    wHsQty = ", 0 as qHsQty, '    ' as qhHs "
    wWsQty = ", 0 as qWsQty, '    ' as qhWs "
    Rep.qHsQty1.Suppress = True
    Rep.qWsQty1.Suppress = True
    Rep1.qHsQty1.Suppress = True
    Rep1.qWsQty1.Suppress = True
End If

' ****** Bhavna emr 5.1.0 Wax set and hand set diamonds available for ..If diamond status Y
  
  '****** (Jen 2.14 Next 11/05/07) added ws_FGStr ******
  '*** jay 2.14 *** (Check for 'SO' added)
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - Loc Join And CoCd added in Group By Clause ******
  '****** Sachin 3.02.0 - Id fields in Joins
  '*** Jay 3.2.0 [PicNm]
  ' ***** Manali 3.5.0 - 12/11/08 - OSalRem fields, Color fields and Join with param table added
  ' ***** Manali 3.5.0 - 25/11/08 - CustDsgCd added
  ' ***** Manali 3.6.0 - 12/10/09 - Apply Colors for OrdNo Group Sort - [qOMColCrCd, qSec added]
  'indigo.164 PValue= 'SO' changed as PValue In('SO','OS')
  'DmTcTyp= 'DM' changed as DmTcTyp=OdDmTcTyp in DsgMst join
  wSqlStrg = "Set DateFormat DMY " + _
             "Select " + grpflds + ", " + _
             "OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, max(DmTcTyp) as DmTcTyp, max(OdDmCd) as OdDmCd, " + ws_CustDmCdStr + " max(OmDt) as OmDt, " + _
             "max(OmCmCd) as OmCmCd, " + _
             "max(OdPrtCd) as OdPrtCd, max(OdKt) as OdKt, max(" + wOrdDtStr + ") as OdDelDt, max(OdSfx) as OdSfx, " + _
             "max(OdDmSz) as OdDmSz, max(DmCtg) as DmCtg, " + _
             "max(OdDmCol) as OdDmCol, max(OdPicNm) as OdPicNm, space(30) as qUniqRmCd, max(OdPrdSeq) as OdPrdSeq, " + wSel + _
             " Sum(case when '" + ADC("UOYN7") + "'='Y' Then IsNull(BQty, 0) when '" + ADC("UOYN7") + "'='N' Then IsNull(1, 0) Else 0 End) as qFlrQty " + wHsQty + wWsQty + _
             wFldLst + ", Replace(max(OdSalRem),char(13)+ char(10),' ') as OdSalRem, '" + CStr(wDetPos) + "' as qDetPos, '" + ADC("UoYN4") + "' as qDispSalRemYN,  " + _
             " IsNull(max(WC.PMCd), '') as qColNm, IsNull(max(RC.PValue), '') as qColCrCd, '" + ADC("UoYn5") + "' as qAppColYN, " + _
            "IsNull((Select RC1.PValue from Param RC1 Join Param WC1 On WC1.PTYP='WIWCOL' and RC1.PMCd=WC1.PMCd " + _
            "Where RC1.PTyp='REPCOL'  and DateDiff(dd, '" + CStr(moCn.SrvrDate) + "', max(" + wDelDt + ")) >= WC1.PNum and DateDiff(dd, '" + CStr(moCn.SrvrDate) + "', max(" + wDelDt + "))<= WC1.PNUm1), '') as qOmColCrCd, " + _
             "'" + wSec + "' as qSec, max(vp.vPValue1) as vPValue1  " + _
             " From OrdDsg Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy " + _
             " and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
             "Join CustMst on CmCtg= 'C' and CmCd = OmCmCd " + _
             "Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp=OdDmTcTyp and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
             "Left Outer join Bag on BOdIdNo=OdIdNo and BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy " + _
             "  And BOdChr= OdChr and BOdNo= OdNo and BOdSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=OdPrtKey ", "") + " and BCls= 'N' " + _
              ws_BagJoin + ws_FGStr + " Left Outer Join Loc On LocCd= BLoc And LocCoCd='" + ctSelfCoCd + "' " + _
             " left outer Join Param WC on WC.PTYP='WIWCOL' and DateDiff(dd, '" + CStr(moCn.SrvrDate) + "', " + wDelDt + ") >= WC.PNum and DateDiff(dd, '" + CStr(moCn.SrvrDate) + "', " + wDelDt + ")<= WC.PNUm1" + _
             " Left outer Join Param RC on  RC.PTyp='REPCOL' and RC.PMCd=WC.PMCd " + _
             " Left outer Join vParam vp on vPTyp ='PRDSEQ' and vPCoCd=OdCoCd and vPMCd=OdPrdSeq " + _
             wStr + wCnd + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", "") + wAddlCnd, "") + wCrmCnd + " and exists (select 'x' from param where PTyp= 'TC' and PMCd= OdTc and PValue In('SO','OS')) " + _
             " Group By OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OdIdNo " + ws_BChrCnd
             
'Loc to Prcs-Left outer Join vParam added
'Sum(IsNull(BQty, 0)) as qFlrQty  'uma flrqty changed for no.of bag

'*** jay 2.14 ***
' **** Zubin 211 **** '
  
'  wSqlStrg = "Select " + grpflds + ", " + _
'             "OdTc, OdYy, OdChr, OdNo, OdSr, max(OdDmCd), max(OmDt), max(OmCmCd), " + _
'             "max(OdPrtCd), max(OdKt), max(OdDelDt), max(OdSfx), max(OdDmSz), max(DmCtg), " + _
'             "max(OdDmCol), space(30) as qUniqRmCd, max(OdPrdSeq), " + wSel + _
'             " Sum(IsNull(BQty, 0)) as qFlrQty " + _
'             wFldLst + _
'             " from OrdDsg join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy " + _
'             "and OmChr= OdChr and OmNo= OdNo " + _
'             "Join CustMst on CmCtg= 'C' and CmCd = OmCmCd " + _
'             "join DsgMst on DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + _
'             "left outer join Bag on BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy " + _
'             "and BOdChr= OdChr and BOdNo= OdNo and BOdSr= OdSr and BCls= 'N' " + _
'             wStr + wCnd + _
'             " Group By OdTc, OdYy, OdChr, OdNo, OdSr "

  ADC.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - CoCd added in Group By Clause ******
  If ADC("UoYN") = "Y" And (UCase(ADC.MenuCd) = UCase("RepWIWLongPrd") Or UCase(ADC.MenuCd) = UCase("RepWIWLongExp")) Then
    Dim wSubRepSql As String
    '****** Sachin 3.02.0 - Id fields in Joins
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
'    wSubRepSql = " Select OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCd, " + _
'                 " (case max(OrRmCtg) when 'D' then 1 when 'C' then 2 when 'G' then 3 " + _
'                 " when 'P' then 4 When 'S' then 5 When 'L' then 6 when 'X' then 7 when 'A' then 8 " + _
'                 " when 'M' then 9 else 0 end) as qSrt " + _
'                 " from OrdMst join OrdDsg on OdOmIdNo=OmIdNo and OdCoCd= OmcoCd and OdTc= OmTc and OdYy= OmYy " + _
'                 " and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
'                 " Join CustMst on CmCtg= 'C' and CmCd = OmCmCd " + _
'                 " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd " + _
'                 " and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + " join OrdRm on OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy " + _
'                 " and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OdPrtKey ", "") + _
'                 wStr + wCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
'                 " and Exists (select 'x' from param where PTyp= 'TC' and PMCd= OdTc and PValue= 'SO') " + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " ") + wAddlCnd, "") + wCrmCnd + _
'                 " Group by OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCd " + _
'                 " Order By OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, qSrt, OrRmCd "
    'indigo.164 PValue= 'SO' changed as PValue In('SO','OS')
    'DmTcTyp= 'DM' changed as DmTcTyp= OdDmTcTyp in DsgMst join
    wSubRepSql = " Select OdTc as OrTc, OdYy as OrYy, OdChr as OrChr, OdNo as OrNo, OdSr as OrSr, '' as OrRmCd, 0 as qSrt, dbo.MwFn_GetColStnStr(OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, '','N') as qRmCdStr " + _
                 " From OrdMst " + _
                 "      Join OrdDsg on OdOmIdNo=OmIdNo and OdCoCd= OmcoCd and OdTc= OmTc and OdYy= OmYy  and OdChr= OmChr and OdNo= OmNo  and OdPrtKey=OmPrtKey " + _
                 "      Join CustMst on CmCtg= 'C' and CmCd = OmCmCd " + _
                 "      Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp and DmCd= OdDmCd  and DmSz= '' and DmPrtKey=OdPrtKey " + _
                 wStr + wCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
                 " and Exists (select 'x' from param where PTyp= 'TC' and PMCd= OdTc and PValue In('SO','OS')) " + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " ") + wAddlCnd, "") + wCrmCnd + _
                 " Order By OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr"
                 
    Call moCn.RepRes(Rep1.Subreport1.OpenSubreport, wSubRepSql + ctMaxDopOpt)   '****** Sachin 3.01 added ctMaxDopOpt
  End If
  
  
  If ADC("UoYn5") = "Y" Then
    Dim wColNm As String, wColDesc As String, wRs As MDORowSet
    wColDesc = "Select P1.PmCd as PMCd, P1.PDesc as PDesc, P2.PValue as qColCd from param P1 Join Param P2 On p2.PtYp='REPCOL' and P2.PMCd=P1.PMCd where P1.PTyp='WIWCOL' Order By P1.PNum"
    Call moCn.RepRes(mRep.SubRepColDesc.OpenSubreport, wColDesc + ctMaxDopOpt)
  Else
    mRep.SecColDesc.Suppress = True
  End If

  
'  '*** To check if Detail is selected ***
'  Dim wGrp(3) As String, i As Integer, wDetPos As Integer
'  wGrp(1) = GRP_REP.Value(gltGroup1)
'  wGrp(2) = GRP_REP.Value(gltGroup2)
'  wGrp(3) = GRP_REP.Value(gltGroup3)
'  wDetPos = 0
'  For i = 1 To 3
'    If wGrp(i) = "(Detail)" Then
'      wDetPos = i: Exit For
'    End If
'  Next i
  '*** If Detail not selected then suppress Section 'DetSec2' Of Detail ***
  If wDetPos = 0 Then mRep.DetSec2.Suppress = True: mRep.DetLineForSec1.Suppress = True: mRep.PICSEC.Suppress = True
  
  
  If ADC("UoYN") = "N" Then
    mRep.DetSec2.Suppress = True
    mRep.TxtOdDelDt.Suppress = True
    mRep.TxtOdSfx.Suppress = True
    mRep.TxtOdDmSz.Suppress = True
    mRep.TxtDmCtg.Suppress = True
    mRep.TxtOdPrtCd.Suppress = True
    mRep.TxtUniqRmSeq.Suppress = True
'    mRep.TxtWsQty.Suppress = True
'    mRep.TxtHsQty.Suppress = True
'  ElseIf ADC("UoYN") = "Y" Then
'    mRep.DetLineForSec1.Suppress = True
  End If
  
  '***Geeta**Emr208 **Picture
  If ADC("UoYN1") = "N" Or wDetPos = 0 Then
    mRep.PICSEC.Suppress = True
    '*** (Bef 2.12)
    'If UCase(adc.MenuCd) = UCase("RepWIWPrd") Or UCase(adc.MenuCd) = UCase("RepWIWExp") Then _
    'mRep.DetSec2.UnderlaySection = False
    '*** (Bef 2.12)
  Else
    mRep.PICSEC.Suppress = IIF(wDetPos = 0, True, False)
    '*** (Bef 2.12) (This part of code has been commented because the line overlaps the second section if no picture is present)
    'If UCase(adc.MenuCd) = UCase("RepWIWPrd") Or UCase(adc.MenuCd) = UCase("RepWIWExp") Then _
    'mRep.DetSec2.UnderlaySection = True
    '*** (Bef 2.12) (This part of code has been commented because the line overlaps the second section if no picture is present)
  End If
  '***

'*** Jay 2.13(CT) ***
  ' ***** Manali 3.8.0 - Rep Changed to mRep
  If ADC("UoBagTyp") = "F" Then
    mRep.TxtBagTypAndChr.SetText "(Flute Bags" & IIF(ADC("UoBChrFr") <> "", ",  Chr = " & ADC("UoBChrFr"), "") & ")"
  ElseIf ADC("UoBagTyp") = "P" Then
    mRep.TxtBagTypAndChr.SetText "(Primary Bags" & IIF(ADC("UoBChrFr") <> "", ",  Chr = " & ADC("UoBChrFr"), "") & ")"
  ElseIf ADC("UoBagTyp") = "C" Then
    mRep.TxtBagTypAndChr.SetText "(Component Bags" & IIF(ADC("UoBChrFr") <> "", ",  Chr = " & ADC("UoBChrFr"), "") & ")"
  End If
'  '***(Jen 2.11)
'  If adc("UoBagTyp") <> "" Then mRep.TxtBagTypAndChr.SetText IIF(adc("UoBagTyp") = "Y", "(Flute Bags", "(Primary Bags") & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
'  '***(Jen 2.11)
 '*** Jay 2.13(CT) ***
 
 ' ***** Manali 3.5.0 - 10/11/08 - Spl Rem section added
' If wDetPos = 0 Then mRep.DetSalRem.Suppress = True
' If adc("UoYN4") = "Y" Then mRep.DetSalRem.Suppress = False Else mRep.DetSalRem.Suppress = True
 ' ***** Manali 3.5.0 - 10/11/08 - Spl Rem section added
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******

'*** (Jen 2.12)
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
'*** (Jen 2.12)
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowLine"
    If .Value = Checked Then ADC("UoYN") = "Y" Else ADC("UoYN") = "N"
  Case Is = 1     'Tag = "ShowPic"
    If .Value = Checked Then ADC("UoYN1") = "Y" Else ADC("UoYN1") = "N"
  Case Is = 4     'Tag = "ShowRem"
    If .Value = Checked Then ADC("UoYN4") = "Y" Else ADC("UoYN4") = "N"
  Case Is = 5     'Tag = "AppCol"
    If .Value = Checked Then ADC("UoYN5") = "Y" Else ADC("UoYN5") = "N"
' ******* Bhavna 5.0.2 - 14/07/2015 - show diamond status wax set and hand set
  Case Is = 6     'Tag = "DiaStatus"
    If .Value = Checked Then ADC("UoYN6") = "Y" Else ADC("UoYN6") = "N"
' ******* Bhavna 5.0.2 - 14/07/2015
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
  
  If ws_MultiCoMod = True Then
    Set ADC.FirNKeyCtl = ADC("UoCoCdFr")
  Else
    Set ADC.FirNKeyCtl = ADC("UoSeqCdFr")
  End If
  
  gb_CoCdFor = False
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  'Manoj
  If UCase(ADC.MenuCd) = UCase("RepWIWPrd") Or UCase(ADC.MenuCd) = UCase("RepWIWLongPrd") Then
    LblPRdExpDt.Visible = False
    ADC("UOYN3").Visible = False
    LblExpDelDt.Visible = False
    ADC("UoDelDt1Fr").Visible = False
    ADC("UoDelDt1To").Visible = False
    ' ****** Manali 3.5.0 - 12/11/08 - Radio Butttons added
    FraOptDelDt.Enabled = False: FraOptDelDt.Visible = False
    ' ****** Manali 3.5.0 - 12/11/08 - Radio Butttons added
  End If
  'Manoj
  If UCase(ADC.MenuCd) = UCase("RepWIWExp") Or UCase(ADC.MenuCd) = UCase("RepWIWLongExp") Then
    LblShowQty.Visible = False
    FraOptQty.Enabled = False: FraOptQty.Visible = False
  End If
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = False Then
    LblLocSeq.Visible = False: LblPrntYN.Visible = False
    ADC("UOYN2") = "N": ADC("UOYN2").Visible = False
    ADC("UOPRNTSEQCDFR") = "": ADC("UOPRNTSEQCDFR").Visible = False
    ADC("UOCOCDFR").Enabled = False: ADC("UOCOCDTO").Visible = False: ADC("UOCOCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
      
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  'gs_LocTyp="'P'"
  'gb_RmCtgFor = True
  'gb_RmZ = False
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"
  'gs_TxnTcTyp = "BV"
  
  gs_CmCtg = "C"
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  Call SetGroupSort
End Sub

Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)

  If UCase(ADC.MenuCd) = UCase("RepWIWPrd") Or UCase(ADC.MenuCd) = UCase("RepWIWExp") Then
    gs_PTyp = "WIWSEQ"
  ElseIf UCase(ADC.MenuCd) = UCase("RepWIWLongPrd") Or UCase(ADC.MenuCd) = UCase("RepWIWLongExp") Then
    gs_PTyp = "WIWLONG"
  End If
  
  '/* urmi 2.06 */
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr"), UCase("UoPMCdTo"), UCase("UoPMCdSel")
    gs_Tbl = "vParam"
    gs_PTyp = "PRDSEQ"
    
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Case Is = UCase("UoYN2")
    If pv_NewValue = "Y" Then
        ADC("UOSEQCDFR") = ""
    Else
        ADC("UOPRNTSEQCDFR") = ""
    End If
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Case Is = UCase("UoYN")
    If pv_NewValue = "Y" Then
        ChkBoxArr(6).Enabled = True
    Else
        ChkBoxArr(6).Enabled = False
        ADC("UoYn6") = "N"
    End If
  '578-28 validating order no. selection should be a number
  Case Is = UCase("UoOmNoSel")
  Dim wOmNoSel() As String, i As Integer
    wOmNoSel = Split(pv_NewValue, ",")
    If IsArray(wOmNoSel) Then
      For i = 0 To UBound(wOmNoSel)
        Cancel = Not IsNumeric(wOmNoSel(i))
        If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is Not Numeric. Enter Order Numbers like 1,2,3 ": Exit Sub
      Next i
    End If
    
  End Select
   
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  
  gb_CoCdFor = False '****** Sachin 2.14.0
  'Loc to Prcs-Report field list(WIW Seq) depends on option button
  If UCase(ADC.MenuCd) = UCase("RepWIWPrd") Or UCase(ADC.MenuCd) = UCase("RepWIWExp") Then
        gs_PTyp = IIF(OptLoc(0).Value = True, "WIWSEQ", "RWIWSEQ")
    
  ElseIf UCase(ADC.MenuCd) = UCase("RepWIWLongPrd") Or UCase(ADC.MenuCd) = UCase("RepWIWLongExp") Then
        gs_PTyp = IIF(OptLoc(0).Value = True, "WIWLONG", "RWIWLONG")
        
  End If
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"
  '/* urmi 2.06 */
 Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr"), UCase("UoPMCdTo"), UCase("UoPMCdSel")
    gs_Tbl = "vParam"
    'Loc to Prcs-Prdseq/Prcsseq depends on option button
    gs_PTyp = IIF(OptLoc(0).Value = True, "PRDSEQ", "PRCSSEQ")
    If ADC.Mode = xNorm Then
      If IdName = UCase("UoPMCdSel") Then HlpList.MultiSelect = True
      HlpList.vPMCd ADC("UoCoCdFr"), gs_PTyp: Exit Sub
    End If
    
  Case Is = UCase("UoSeqCdFr")
      '****** Sachin 2.14.0 - Original - Seeking No of Companies in the Scope ******
      ws_ScoCnd = IIF(ADC("UoCoCdFr") <> "", " And hCoCd>='" + ADC("UoCoCdFr") + "'", " ")
      ws_ScoCnd = ws_ScoCnd + IIF(ADC("UoCoCdTo") <> "", " And hCoCd<='" + ADC("UoCoCdTo") + "'", " ")
      ws_ScoCnd = ws_ScoCnd + IIF(ADC("UoCoCdSel") <> "", " And hCoCd In (" + ADC("UoCoCdSel") + ")", " ")
      wi_ScopeCos = moCn.GetFldVal("Select Count('x') From Head Where HCd='" + ctSelfCmCd + "' " + ws_ScoCnd)
      If ws_MultiCoMod = True And wi_ScopeCos = 1 Then
          ws_ScoCnd = moCn.GetFldVal("Select HCoCd From Head Where HCd='" + ctSelfCmCd + "' " + ws_ScoCnd)
      Else
          ws_ScoCnd = ADC("UoCoCdFr")
      End If
      

        
      '****** Sachin 2.14.0 - Original - Seeking No of Companies in the Scope ******
      If ADC.Mode = xNorm Then
           If (ADC("UOYN2") = "Y" Or ADC("UOYN2") = "") Then Cancel = True: ErrMsg = "Location Sequence Can Be Specified If Parent(Y/N) Is Set to 'N'": Exit Sub
           HlpList.vPMCd IIF(wi_ScopeCos = 1, ws_ScoCnd, ctSelfCoCd), gs_PTyp, , , , "N": Exit Sub
      End If
      
    ' **** Zubin 211 **** '
    Case Is = UCase("UOBCHRFR")
        If ADC.Mode = xNorm Then
          If ADC("UoBagTyp") = "P" Or ADC("UoBagTyp") = "" Then Cancel = True: ErrMsg = "This Option Is Valid Only For Flute/ Component Bag ": Exit Sub
          HlpList.PMCd "BCHR", "'" + ADC("UoBagTyp") + "'": Exit Sub
        End If
    ' **** Zubin 211 **** '
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
      
  Case Is = UCase("UOPRNTSEQCDFR")
      If ADC.Mode = xNorm Then
        If (ADC("UOYN2") = "N" Or ADC("UOYN2") = "") Then Cancel = True: ErrMsg = "Parent Location Sequence Can Be Specified If Parent(Y/N) Is Set to 'Y'": Exit Sub
        HlpList.vPMCd ctSelfCoCd, gs_PTyp, , , , "Y": Exit Sub
      End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******

 End Select
  '/* urmi 2.06 */

  SetProp Me, IdName, When

End Sub
Private Function FldLst(ByVal mSeq As String, ByVal mCtr As Integer)
  Dim wRsLocDesc As MDORowSet
  Dim ws_Seq As String, ws_str As String, wi_cPos As Integer, wi_Ctr As Integer, wi_Pos As Integer
  Dim ws_fldlst As String, wReqStr As String, ws_LocDesc As Variant
  
  '****** Sachin 2.14.0 - Original - Seeking No of Companies in the Scope ******
  ws_ScoCnd = IIF(ADC("UoCoCdFr") <> "", " And hCoCd>='" + ADC("UoCoCdFr") + "'", " ")
  ws_ScoCnd = ws_ScoCnd + IIF(ADC("UoCoCdTo") <> "", " And hCoCd<='" + ADC("UoCoCdTo") + "'", " ")
  ws_ScoCnd = ws_ScoCnd + IIF(ADC("UoCoCdSel") <> "", " And hCoCd In (" + ADC("UoCoCdSel") + ")", " ")
  wi_ScopeCos = moCn.GetFldVal("Select Count('x') From Head Where HCd='" + ctSelfCmCd + "' " + ws_ScoCnd)
  If ws_MultiCoMod = True And wi_ScopeCos = 1 Then
      ws_ScoCnd = moCn.GetFldVal("Select HCoCd From Head Where HCd='" + ctSelfCmCd + "' " + ws_ScoCnd)
  Else
      ws_ScoCnd = ADC("UoCoCdFr")
  End If
  '****** Sachin 2.14.0 - Original - Seeking No of Companies in the Scope ******
  'Loc to Prcs-selecting P /R Loc description
  Set wRsLocDesc = moCn.OpenRes("Select LocCd, LocDesc from Loc where LocCoCd ='" + IIF(ADC("UOYN2") = "Y", ctSelfCoCd, IIF(wi_ScopeCos = 1, ws_ScoCnd, ctSelfCoCd)) + "' and LocTyp='" + IIF(OptLoc(0).Value = True, "P", "R") + "'")
  ws_Seq = moCn.GetFldVal("select vPDesc225 from vParam where vPCoCd= '" + IIF(ADC("UOYN2") = "Y", ctSelfCoCd, IIF(wi_ScopeCos = 1, ws_ScoCnd, ctSelfCoCd)) + "' and vPTyp= '" + mSeq + "' and vPMCd= '" + IIF(ADC("UOYN2") = "N", ADC("UoSeqCdFr"), ADC("UoPrntSeqCdFr")) + "'")
  'ws_Seq = "PMKT,PWAX,PDIA,PWST,PCST,PFIL,PHST,PQC,PREP,PDLV"
  ws_str = "," + ws_Seq + ","
  wi_cPos = 1
  wi_Ctr = 1
  Do While wi_Ctr <= mCtr
    wi_Pos = InStr(wi_cPos + 1, ws_str, ",")
    If wi_Pos = 0 Then
      ws_fldlst = ws_fldlst + ",0 AS qQty" + CStr(wi_Ctr)
      GRP_REP.SetFormula mRep, "hQty" + CStr(wi_Ctr), "''"
      SetRepText mRep, "TxtLocCd" + CStr(wi_Ctr), " "
      SetRepText mRep, "TxtLocDesc" + CStr(wi_Ctr), " "
    Else
      wReqStr = Mid(ws_str, wi_cPos + 1, wi_Pos - wi_cPos - 1)
      wi_cPos = wi_Pos
      'ws_fldlst = ws_fldlst + ",Sum(Case When IsNull(" + IIF(adc("UOYN2") = "N", "BLoc", "LocPrntCd") + ", '')='" + wReqStr + "' Then " + _
                "IsNull(BQty, 0) Else 0 End) AS qQty" + CStr(wi_Ctr)
      'uma to validate bag qty / no. of bag
      'ws_fldlst = ws_fldlst + ",Sum(Case When IsNull(" + IIF(adc("UOYN2") = "N", "BLoc", "LocPrntCd") + ", '')='" + wReqStr + "' and '" + adc("UOYN7") + "'='Y' Then " + _
                "IsNull(BQty, 0) When IsNull(" + IIF(adc("UOYN2") = "N", "BLoc", "LocPrntCd") + ", '')='" + wReqStr + "' and '" + adc("UOYN7") + "'='N' Then " + _
                "IsNull(1, 0) Else 0 End) AS qQty" + CStr(wi_Ctr)
            
      'Loc to Prcs-Getting LocCd from LocPrcs if Prcs selected by user
      If OptLoc(1).Value = True Then
      'sv.70 getting multiple location for a process
      Dim wRsLocForPrcs As MDORowSet, wLocForPrcs As String
      Set wRsLocForPrcs = moCn.OpenResultset("select LocCd from Loc where LocCoCd='" + IIF(ADC("UOYN2") = "Y", ctSelfCoCd, IIF(wi_ScopeCos = 1, ws_ScoCnd, ctSelfCoCd)) + "' and LocPrcs='" + wReqStr + "'")
      wLocForPrcs = ""
      If wRsLocForPrcs.RecCount > 0 Then
        wRsLocForPrcs.MoveFirst
        Do While Not (wRsLocForPrcs.BOF Or wRsLocForPrcs.EOF)
          wLocForPrcs = wLocForPrcs + IIF(wLocForPrcs <> "", ",", "") + "'" + wRsLocForPrcs!LocCd + "'"
          wRsLocForPrcs.MoveNext
        Loop
      Else
        wLocForPrcs = "''"
      End If
      'sv.70 wLocForPrcs added for multiple location
      ws_fldlst = ws_fldlst + ",Sum(Case When IsNull(" + IIF(ADC("UOYN2") = "N", "BLoc", "LocPrntCd") + ", '') in(" + wLocForPrcs + ") and '" + ADC("UOYN7") + "'='Y' Then " + _
                "IsNull(BQty, 0) When IsNull(" + IIF(ADC("UOYN2") = "N", "BLoc", "LocPrntCd") + ", '') in(" + wLocForPrcs + ") and '" + ADC("UOYN7") + "'='N' Then " + _
                "IsNull(1, 0) Else 0 End) AS qQty" + CStr(wi_Ctr)
      
      Else
      ws_fldlst = ws_fldlst + ",Sum(Case When IsNull(" + IIF(ADC("UOYN2") = "N", "BLoc", "LocPrntCd") + ", '')='" + wReqStr + "' and '" + ADC("UOYN7") + "'='Y' Then " + _
                "IsNull(BQty, 0) When IsNull(" + IIF(ADC("UOYN2") = "N", "BLoc", "LocPrntCd") + ", '')='" + wReqStr + "' and '" + ADC("UOYN7") + "'='N' Then " + _
                "IsNull(1, 0) Else 0 End) AS qQty" + CStr(wi_Ctr)
      End If
                      
      GRP_REP.SetFormula mRep, "hQty" + CStr(wi_Ctr), "'" + wReqStr + "'"
      
      '*** For Setting The Location Code Key ***
      SetRepText mRep, "TxtLocCd" + CStr(wi_Ctr), wReqStr + " -"
      ws_LocDesc = Find(wRsLocDesc, "LocCd", wReqStr, "LocDesc")
      If ws_LocDesc <> False Then
        SetRepText mRep, "TxtLocDesc" + CStr(wi_Ctr), ws_LocDesc
      Else
        SetRepText mRep, "TxtLocDesc" + CStr(wi_Ctr), ""
      End If
      '*** For Setting The Location Code Key ***
      
    End If
    wi_Ctr = wi_Ctr + 1
  Loop
  FldLst = ws_fldlst
End Function
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
' **** Zubin 211 **** '
  Select Case UCase(IdName)
    Case Is = UCase("UOBCHRFR")
      'If ADC("UoBagTyp") = "N" Then
      ADC("UOBCHRFR") = ""
    Case Is = UCase("UoYN6")
      If ADC("UoYn") = "Y" Then
        ChkBoxArr(6).Enabled = True
      Else
        ChkBoxArr(6).Enabled = False
        ADC("UoYn6") = "N"
      End If
  End Select
' **** Zubin 211 **** '
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

' ****** Manali 3.5.0 - 12/11/08 - Radio Butttons added
Private Sub OptDelDt_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN3") = "Y"
Case Is = 1
  ADC("UoYN3") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 12/11/08 - Radio Butttons added

' ****** Manali 3.5.0 - 12/11/08 - Radio Butttons added
Private Sub OptDelDt_GotFocus(Index As Integer)
  DispMsg FraOptDelDt.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 12/11/08 - Radio Butttons added

Private Sub OptLoc_GotFocus(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UOSEQCDFR") = ""
  ADC("UoLocTypFr") = "P"
  ADC("UOPMCDFR") = ""
  ADC("UOPMCDTO") = ""
  ADC("UOPMCDSEL") = ""
Case Is = 1
  ADC("UOSEQCDFR") = ""
  ADC("UoLocTypFr") = "R"
  ADC("UOPMCDFR") = ""
  ADC("UOPMCDTO") = ""
  ADC("UOPMCDSEL") = ""
End Select

End Sub

Private Sub OptQty_Click(Index As Integer)
'Uma
Select Case Index
Case Is = 0
  ADC("UoYN7") = "Y"
Case Is = 1
  ADC("UoYN7") = "N"
End Select
End Sub

Private Sub OptQty_GotFocus(Index As Integer)
'uma
  DispMsg FraOptQty.ToolTipText, etInfo
End Sub

'***
Private Sub tab_REP_Click(PreviousTab As Integer)
  fra_tabrep(TAB_REP.Tab).Enabled = True
  fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  ADC("UOUSRCD") = ADC.UsrCd
  ADC("UOMNUCD") = ADC.MenuCd
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = False Then
      ADC("UoCoCdFr") = gs_CoCd: ADC("UoCoCdTo") = gs_CoCd: ADC("UoCoCdSel") = ""
  End If
 '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
 
 '**Geeta*Emr208*Default Show picture is always no
 ADC("UoYN1") = "N"
 ADC("UOYN2") = "N"
 ADC("UoYN7") = "Y"   'uma
 
 'Manoj
  If UCase(ADC.MenuCd) = UCase("RepWIWPrd") Or UCase(ADC.MenuCd) = UCase("RepWIWLongPrd") Then
    ADC("UOYN3") = "N"
    ADC("UoDelDt1Fr") = MWLib.EmptyDate
    ADC("UoDelDt1To") = MWLib.EmptyDate
  End If
  
  '*******Bhavna Emr 5.0.2 diamond status default 'N'
      ADC("UoYN6") = "N"
      If ADC("UoYn") = "Y" Then
        ChkBoxArr(6).Enabled = True
      Else
        ChkBoxArr(6).Enabled = False
        ADC("UoYn6") = "N"
      End If
'Manoj

  '*** Jay 2.14Next ***
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
  '*** Jay 2.14Next ***

'*** Jay 2.13(CT) ***
'  '*** (Jen 2.11)
'  If moCn.GetFldVal("Select hFluteBagYN from Head where HCoCd= '" + adc("UoCoCdFr") + "' and " + _
'                                                    "HCd= '" + ctSelfCmCd + "' ") = "N" Then
'    LblBagTyp.Visible = False
'    adc("UoBagTyp").Visible = False
'    LblUoBChrFr.Visible = False
'    adc("UoBChrFr").Visible = False
'    adc("UoBagTyp") = "N": adc("UoBChrFr") = ""
'  End If
'  '*** (Jen 2.11)
'*** Jay 2.13(CT) ***

  ' ****** Manali 3.5.0 - 12/11/08 - Radio Butttons added
  If ADC("UoYN3") = "Y" Then OptDelDt(0).Value = True
  If ADC("UoYN3") = "N" Then OptDelDt(1).Value = True
  ' ****** Manali 3.5.0 - 12/11/08 - Radio Butttons added

  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  If ADC("UoYN") = "" Then ADC("UoYN") = "N"
  If ADC("UoYN1") = "" Then ADC("UoYN1") = "N"
  If ADC("UoYN4") = "" Then ADC("UoYN4") = "N"
  If ADC("UoYN5") = "" Then ADC("UoYN5") = "N"
  ' *******Bhavna 5.0.2 - 14/07/15
  If ADC("UoYN6") = "" Then ADC("UoYN6") = "N"
  
  If ADC("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If ADC("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If ADC("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  If ADC("UoYN5") = "Y" Then ChkBoxArr(5).Value = Checked Else ChkBoxArr(5).Value = Unchecked
  ' *******Bhavna 5.0.2 - 14/07/15
  If ADC("UoYN6") = "Y" Then ChkBoxArr(6).Value = Checked Else ChkBoxArr(6).Value = Unchecked
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
'uma
  If ADC("UoYN7") = "Y" Then OptQty(0).Value = True
  If ADC("UoYN7") = "N" Then OptQty(1).Value = True
  
  If ADC("UOLOCTYPFR") = "P" Then
    OptLoc(0).Value = True
  ElseIf ADC("UOLOCTYPFR") = "R" Then
    OptLoc(1).Value = True
  End If

End Sub
Private Sub DispCoNm()
'mRep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  '***** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + gs_CoCd + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
    mRep.wCoCd.SetText ws_HName
    mRep.TxtCoCdFr.SetText ADC("UOCOCDFR")
    mRep.TxtCoCdTo.Suppress = False: mRep.TxtCoCdSel.Suppress = False
    mRep.TxtCoCdToLbl.Suppress = False: mRep.TxtCoCdSelLbl.Suppress = False
    mRep.TxtCoCdTo.SetText ADC("UOCOCDTO")
    mRep.TxtCoCdSel.SetText ADC("UOCOCDSEL")
  Else
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + ADC("UoCoCdFr") + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
    mRep.wCoCd.SetText ws_HName
    mRep.TxtCoCdFr.SetText ADC("UOCOCDFR")
    mRep.TxtCoCdTo.Suppress = True: mRep.TxtCoCdSel.Suppress = True
    mRep.TxtCoCdToLbl.Suppress = True: mRep.TxtCoCdSelLbl.Suppress = True
  End If
  '***** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  mRep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    mRep.RHMain.UnderlaySection = True: mRep.wCoCd.Suppress = True
  Else
    mRep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
'******* Manali - Before 3.2.2
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  mRep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  mRep.wSrvrTime.SetText ws_DtTm

'****** Manali 3.2.2 - User, date and time in Reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  mRep.TxtRepUsrDtTime.SetText Trim(ADC.UsrCd) + " / " + ws_Dt + " / " + ws_Tm

  '22 Customer From & To, wiw sequence code & description added to header
  mRep.TxtCustFr.SetText ADC("UoCmCdFr"): mRep.TxtCustTo.SetText ADC("UoCmCdTo")
  Dim wSeqDesc As String
  wSeqDesc = moCn.GetFldVal("select vPDesc from vParam where vPCoCd ='" + IIF(ws_MultiCoMod = True, gs_CoCd, ADC("UoCoCdFr")) + _
                "' and vPTyp ='" + gs_PTyp + "' and vPMCd ='" + ADC("UoSeqCdFr") + "'")

  mRep.TxtSeqDesc.SetText ADC("UoSeqCdFr") + " - " + wSeqDesc

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
  Set mRep = Nothing
  Set Rep1 = Nothing
  '*** (09/08/05)
End Sub
