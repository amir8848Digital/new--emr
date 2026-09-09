VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpFgList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finished Goods Listing"
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
      Left            =   11730
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   9630
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4395
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5895
      TabIndex        =   89
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   90
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
         TabIndex        =   91
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
      Left            =   7110
      TabIndex        =   88
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   85
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
      TabIndex        =   81
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
      TabIndex        =   84
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   83
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpFgList.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpFgList.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   87
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   82
               Top             =   360
               Width           =   8145
               _ExtentX        =   14367
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   86
            Top             =   360
            Width           =   15105
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1935
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   147
               Top             =   6390
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
               Index           =   3
               Left            =   15
               TabIndex        =   146
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Stock Rate"
               Top             =   6750
               Width           =   2100
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Only Pnd Fg            "
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
               Left            =   15
               TabIndex        =   78
               Tag             =   "ShowPndFg"
               ToolTipText     =   "Check To Show Only Pending FG Records"
               Top             =   6060
               Width           =   2130
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Box Det        "
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
               Left            =   15
               TabIndex        =   76
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show Box Details"
               Top             =   5775
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show 2nd Line       "
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
               Left            =   15
               TabIndex        =   74
               Tag             =   "ShowLine"
               ToolTipText     =   "Check To Show 2nd Line"
               Top             =   5490
               Width           =   2115
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   5415
               TabIndex        =   7
               ToolTipText     =   "Enter To Finished Goods Year"
               Top             =   630
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOFGYYTO"
               IdName          =   "UOFGYYTO"
               CmpStr          =   "FdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2715
               TabIndex        =   4
               ToolTipText     =   "Enter From Finished Goods Year"
               Top             =   630
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOFGYYFR"
               IdName          =   "UOFGYYFR"
               CmpStr          =   "FdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1935
               TabIndex        =   3
               ToolTipText     =   "Enter From Finished Goods Tc"
               Top             =   630
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOFGTCFR"
               IdName          =   "UOFGTCFR"
               CmpStr          =   "FdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   3135
               TabIndex        =   5
               ToolTipText     =   "Enter From Finished Goods Character"
               Top             =   630
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFGCHRFR"
               IdName          =   "UOFGCHRFR"
               CmpStr          =   "FdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   4635
               TabIndex        =   6
               ToolTipText     =   "Enter To Finished Goods Tc"
               Top             =   630
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOFGTCTO"
               IdName          =   "UOFGTCTO"
               CmpStr          =   "FdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   5835
               TabIndex        =   8
               ToolTipText     =   "Enter To Finished Goods Character"
               Top             =   630
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFGCHRTO"
               IdName          =   "UOFGCHRTO"
               CmpStr          =   "FdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   5535
               TabIndex        =   12
               ToolTipText     =   "Enter To Finished Goods Serial"
               Top             =   915
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFDSRTO"
               IdName          =   "UOFDSRTO"
               CmpStr          =   "FdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   2835
               TabIndex        =   10
               ToolTipText     =   "Enter From Finished Goods Serial"
               Top             =   915
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFDSRFR"
               IdName          =   "UOFDSRFR"
               CmpStr          =   "FdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   1935
               TabIndex        =   9
               ToolTipText     =   "Enter From Finished Goods Number"
               Top             =   915
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFGNOFR"
               IdName          =   "UOFGNOFR"
               CmpStr          =   "FdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   4635
               TabIndex        =   11
               ToolTipText     =   "Enter To Finished Goods Number"
               Top             =   915
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFGNOTO"
               IdName          =   "UOFGNOTO"
               CmpStr          =   "FdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   4635
               TabIndex        =   14
               ToolTipText     =   "Enter To Finished Goods Transaction Date"
               Top             =   1200
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "FgDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   1935
               TabIndex        =   13
               ToolTipText     =   "Enter From Finished Goods Transaction Date"
               Top             =   1200
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "FgDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5415
               TabIndex        =   31
               ToolTipText     =   "Enter To Prd Order Year"
               Top             =   2340
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "FdPrdOdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2715
               TabIndex        =   28
               ToolTipText     =   "Enter From Prd Order Year"
               Top             =   2340
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "FdPrdOdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1935
               TabIndex        =   27
               ToolTipText     =   "Enter From Prd Order Tc"
               Top             =   2340
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "FdPrdOdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3135
               TabIndex        =   29
               ToolTipText     =   "Enter From Prd Order Character"
               Top             =   2340
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "FdPrdOdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4635
               TabIndex        =   30
               ToolTipText     =   "Enter To Prd Order Tc"
               Top             =   2340
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "FdPrdOdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   5835
               TabIndex        =   32
               ToolTipText     =   "Enter To Prd Order Character"
               Top             =   2340
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "FdPrdOdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5535
               TabIndex        =   36
               ToolTipText     =   "Enter To Prd Order Serial"
               Top             =   2625
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "FdPrdOdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   2835
               TabIndex        =   34
               ToolTipText     =   "Enter From Prd Order Serial"
               Top             =   2625
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "FdPrdOdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1935
               TabIndex        =   33
               ToolTipText     =   "Enter From Prd Order Number"
               Top             =   2625
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "FdPrdOdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   4635
               TabIndex        =   35
               ToolTipText     =   "Enter To Prd Order Number"
               Top             =   2625
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "FdPrdOdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   4635
               TabIndex        =   18
               ToolTipText     =   "Enter To Bag Year"
               Top             =   1485
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYTO"
               IdName          =   "UOBYYTO"
               CmpStr          =   "FdBYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1935
               TabIndex        =   15
               ToolTipText     =   "Enter From Bag Year"
               Top             =   1485
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYFR"
               IdName          =   "UOBYYFR"
               CmpStr          =   "FdBYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   2355
               TabIndex        =   16
               ToolTipText     =   "Enter From Bag Character"
               Top             =   1485
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRFR"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "FdBChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   3135
               TabIndex        =   17
               ToolTipText     =   "Enter From Bag Number"
               Top             =   1485
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOFR"
               IdName          =   "UOBNOFR"
               CmpStr          =   "FdBNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   5055
               TabIndex        =   19
               ToolTipText     =   "Enter To Bag Character"
               Top             =   1485
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRTO"
               IdName          =   "UOBCHRTO"
               CmpStr          =   "FdBChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   5835
               TabIndex        =   20
               ToolTipText     =   "Enter To Bag Number"
               Top             =   1485
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOTO"
               IdName          =   "UOBNOTO"
               CmpStr          =   "FdBNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   7320
               TabIndex        =   26
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2055
               Width           =   7530
               _ExtentX        =   13282
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "FdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1935
               TabIndex        =   24
               ToolTipText     =   "Enter From Design Code"
               Top             =   2055
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "FdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4635
               TabIndex        =   38
               ToolTipText     =   "Enter To Prd Customer"
               Top             =   2910
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "FdPrdCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   7320
               TabIndex        =   39
               ToolTipText     =   "Enter Prd Customer Selection"
               Top             =   2910
               Width           =   7530
               _ExtentX        =   13282
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "FdPrdCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1935
               TabIndex        =   37
               ToolTipText     =   "Enter From Prd Customer"
               Top             =   2910
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "FdPrdCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   4635
               TabIndex        =   68
               ToolTipText     =   "Enter To Export Number"
               Top             =   4905
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOTO"
               IdName          =   "UOINEXPNOTO"
               CmpStr          =   "FdInExpNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   7320
               TabIndex        =   69
               ToolTipText     =   "Enter Export Number Selection"
               Top             =   4905
               Width           =   7530
               _ExtentX        =   13282
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOINEXPNOSEL"
               IdName          =   "UOINEXPNOSEL"
               CmpStr          =   "FdInExpNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1935
               TabIndex        =   67
               ToolTipText     =   "Enter From Export Number"
               Top             =   4920
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOFR"
               IdName          =   "UOINEXPNOFR"
               CmpStr          =   "FdInExpNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   5415
               TabIndex        =   44
               ToolTipText     =   "Enter To Exp Order Year"
               Top             =   3195
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOEXPODYYTO"
               IdName          =   "UOEXPODYYTO"
               CmpStr          =   "FdExpOdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   2715
               TabIndex        =   41
               ToolTipText     =   "Enter From Exp Order Year"
               Top             =   3195
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOEXPODYYFR"
               IdName          =   "UOEXPODYYFR"
               CmpStr          =   "FdExpOdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1935
               TabIndex        =   40
               ToolTipText     =   "Enter From Exp Order Tc"
               Top             =   3195
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOEXPODTCFR"
               IdName          =   "UOEXPODTCFR"
               CmpStr          =   "FdExpOdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   3135
               TabIndex        =   42
               ToolTipText     =   "Enter From Exp Order Character"
               Top             =   3195
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOEXPODCHRFR"
               IdName          =   "UOEXPODCHRFR"
               CmpStr          =   "FdExpOdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4635
               TabIndex        =   43
               ToolTipText     =   "Enter To Exp Order Tc"
               Top             =   3195
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOEXPODTCTO"
               IdName          =   "UOEXPODTCTO"
               CmpStr          =   "FdExpOdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   5835
               TabIndex        =   45
               ToolTipText     =   "Enter To Exp Order Character"
               Top             =   3195
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOEXPODCHRTO"
               IdName          =   "UOEXPODCHRTO"
               CmpStr          =   "FdExpOdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   2835
               TabIndex        =   47
               ToolTipText     =   "Enter From Exp Order Serial"
               Top             =   3495
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOEXPODSRFR"
               IdName          =   "UOEXPODSRFR"
               CmpStr          =   "FdExpOdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   1935
               TabIndex        =   46
               ToolTipText     =   "Enter From Exp Order Number"
               Top             =   3495
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOEXPODNOFR"
               IdName          =   "UOEXPODNOFR"
               CmpStr          =   "FdExpOdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   7320
               TabIndex        =   52
               ToolTipText     =   "Enter Exp Customer Selection"
               Top             =   3780
               Width           =   7530
               _ExtentX        =   13282
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOEXPCMCDSEL"
               IdName          =   "UOEXPCMCDSEL"
               CmpStr          =   "FdExpCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   1935
               TabIndex        =   50
               ToolTipText     =   "Enter From Exp Customer"
               Top             =   3780
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOEXPCMCDFR"
               IdName          =   "UOEXPCMCDFR"
               CmpStr          =   "FdExpCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   4635
               TabIndex        =   56
               ToolTipText     =   "Enter To Packing List Year"
               Top             =   4065
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
               Index           =   36
               Left            =   1935
               TabIndex        =   53
               ToolTipText     =   "Enter From Packing List Year"
               Top             =   4065
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
               Left            =   2355
               TabIndex        =   54
               ToolTipText     =   "Enter From Packing List Character"
               Top             =   4065
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
               Left            =   3135
               TabIndex        =   55
               ToolTipText     =   "Enter From Packing List Number"
               Top             =   4065
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
               Index           =   58
               Left            =   2715
               TabIndex        =   60
               ToolTipText     =   "Enter From Invoice Year"
               Top             =   4350
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "FdInYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   59
               Left            =   1935
               TabIndex        =   59
               ToolTipText     =   "Enter From Invoice Tc"
               Top             =   4350
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "FdInTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   3135
               TabIndex        =   61
               ToolTipText     =   "Enter From Invoice Character"
               Top             =   4350
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "FdInChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   4635
               TabIndex        =   62
               ToolTipText     =   "Enter To Invoice Tc"
               Top             =   4350
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOINTCTO"
               IdName          =   "UOINTCTO"
               CmpStr          =   "FdInTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   1935
               TabIndex        =   65
               ToolTipText     =   "Enter From Invoice Number"
               Top             =   4635
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "FdInNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   64
               Left            =   4635
               TabIndex        =   66
               ToolTipText     =   "Enter To Invoice Number"
               Top             =   4635
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOTO"
               IdName          =   "UOINNOTO"
               CmpStr          =   "FdInNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   65
               Left            =   3135
               TabIndex        =   73
               ToolTipText     =   "Show Second Line (Y/N) ?"
               Top             =   5490
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   3135
               TabIndex        =   75
               ToolTipText     =   "Show Box Detail (Yes / No)"
               Top             =   5775
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
               Index           =   66
               Left            =   3135
               TabIndex        =   77
               ToolTipText     =   "Show Pending/ All Fg records (Yes / No)"
               Top             =   6060
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   68
               Left            =   4635
               TabIndex        =   71
               ToolTipText     =   "Enter To FgSubLoc Location"
               Top             =   5190
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOFGSUBLOCTO"
               IdName          =   "UOFGSUBLOCTO"
               CmpStr          =   "FdSubLoc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   69
               Left            =   7320
               TabIndex        =   72
               ToolTipText     =   "Enter FgSubLoc Location Selection"
               Top             =   5190
               Width           =   7530
               _ExtentX        =   13282
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOFGSUBLOCSEL"
               IdName          =   "UOFGSUBLOCSEL"
               CmpStr          =   "FdSubLoc In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   1935
               TabIndex        =   70
               ToolTipText     =   "Enter From FgSubLoc Location"
               Top             =   5205
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOFGSUBLOCFR"
               IdName          =   "UOFGSUBLOCFR"
               CmpStr          =   "FdSubLoc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   72
               Left            =   7320
               TabIndex        =   23
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   1770
               Width           =   7530
               _ExtentX        =   13282
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
               CmpStr          =   "DmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   73
               Left            =   1935
               TabIndex        =   21
               ToolTipText     =   "Enter From Design Category"
               Top             =   1770
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
               Left            =   4635
               TabIndex        =   25
               ToolTipText     =   "Enter To Design Code"
               Top             =   2055
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "FdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   4635
               TabIndex        =   22
               ToolTipText     =   "Enter To Design Category"
               Top             =   1770
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
               Index           =   13
               Left            =   5535
               TabIndex        =   49
               ToolTipText     =   "Enter To Exp Order Serial"
               Top             =   3495
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOEXPODSRTO"
               IdName          =   "UOEXPODSRTO"
               CmpStr          =   "FdExpOdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   4635
               TabIndex        =   48
               ToolTipText     =   "Enter To Exp Order Number"
               Top             =   3495
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOEXPODNOTO"
               IdName          =   "UOEXPODNOTO"
               CmpStr          =   "FdExpOdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4635
               TabIndex        =   51
               ToolTipText     =   "Enter To Exp Customer"
               Top             =   3780
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOEXPCMCDTO"
               IdName          =   "UOEXPCMCDTO"
               CmpStr          =   "FdExpCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   5055
               TabIndex        =   57
               ToolTipText     =   "Enter To Packing List Character"
               Top             =   4065
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
               Index           =   56
               Left            =   5835
               TabIndex        =   58
               ToolTipText     =   "Enter To Packing List Number"
               Top             =   4065
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
               Index           =   57
               Left            =   5415
               TabIndex        =   63
               ToolTipText     =   "Enter To Invoice Year"
               Top             =   4350
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOINYYTO"
               IdName          =   "UOINYYTO"
               CmpStr          =   "FdInYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   5835
               TabIndex        =   64
               ToolTipText     =   "Enter To Invoice Character"
               Top             =   4350
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRTO"
               IdName          =   "UOINCHRTO"
               CmpStr          =   "FdInChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   74
               Left            =   4635
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   360
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "FdCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   7320
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   360
               Width           =   8520
               _ExtentX        =   15028
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "FdCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   76
               Left            =   1935
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   360
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "FdCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   240
               Index           =   89
               Left            =   2460
               TabIndex        =   148
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   6720
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
               Index           =   51
               Left            =   15
               TabIndex        =   149
               Top             =   6390
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
               Index           =   7
               Left            =   5475
               TabIndex        =   145
               Top             =   3495
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
               Index           =   38
               Left            =   5775
               TabIndex        =   144
               Top             =   4065
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
               Index           =   40
               Left            =   4995
               TabIndex        =   143
               Top             =   4065
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
               Index           =   43
               Left            =   5355
               TabIndex        =   142
               Top             =   4350
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
               Index           =   45
               Left            =   5775
               TabIndex        =   141
               Top             =   4350
               Width           =   105
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
               Index           =   50
               Left            =   60
               TabIndex        =   140
               Top             =   1785
               Width           =   1365
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
               TabIndex        =   139
               Top             =   5205
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
               Index           =   49
               Left            =   60
               TabIndex        =   138
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label LblShowPndFg 
               BackStyle       =   0  'Transparent
               Caption         =   "Only Pnd Fg (Y/N)"
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
               Left            =   60
               TabIndex        =   137
               ToolTipText     =   "Location"
               Top             =   6060
               Visible         =   0   'False
               Width           =   1995
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Box Det (Y/N)"
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
               Index           =   48
               Left            =   60
               TabIndex        =   136
               ToolTipText     =   "Location"
               Top             =   5775
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.Label Lbl2ndLine 
               BackStyle       =   0  'Transparent
               Caption         =   "2nd Line (Y/N)"
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
               Left            =   60
               TabIndex        =   135
               ToolTipText     =   "Location"
               Top             =   5490
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Vch Tc/Yy/Chr"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   47
               Left            =   60
               TabIndex        =   134
               Top             =   4350
               Width           =   1935
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Vch No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   46
               Left            =   60
               TabIndex        =   133
               Top             =   4635
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
               Index           =   44
               Left            =   3075
               TabIndex        =   132
               Top             =   4350
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
               Index           =   42
               Left            =   2655
               TabIndex        =   131
               Top             =   4350
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Index           =   41
               Left            =   60
               TabIndex        =   130
               Top             =   4065
               Width           =   1905
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
               Index           =   39
               Left            =   2295
               TabIndex        =   129
               Top             =   4065
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
               Index           =   37
               Left            =   3075
               TabIndex        =   128
               Top             =   4065
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Exp Customer"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   127
               Top             =   3780
               Width           =   1455
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Exp Ord Tc/Yy/Chr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   60
               TabIndex        =   126
               Top             =   3210
               Width           =   1935
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Exp Ord No/Sr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   125
               Top             =   3495
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
               Index           =   16
               Left            =   5775
               TabIndex        =   124
               Top             =   3195
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
               Index           =   15
               Left            =   3075
               TabIndex        =   123
               Top             =   3195
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
               Index           =   13
               Left            =   5355
               TabIndex        =   122
               Top             =   3195
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
               Index           =   8
               Left            =   2655
               TabIndex        =   121
               Top             =   3195
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
               Index           =   6
               Left            =   2775
               TabIndex        =   120
               Top             =   3495
               Width           =   105
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
               Left            =   60
               TabIndex        =   119
               Top             =   4920
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Customer"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   118
               Top             =   2910
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
               TabIndex        =   117
               Top             =   2055
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
               Index           =   21
               Left            =   5775
               TabIndex        =   116
               Top             =   1485
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
               Index           =   20
               Left            =   3075
               TabIndex        =   115
               Top             =   1485
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
               Index           =   19
               Left            =   4995
               TabIndex        =   114
               Top             =   1485
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
               Index           =   18
               Left            =   2295
               TabIndex        =   113
               Top             =   1485
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Yy/Chr/No"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   112
               Top             =   1500
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Ord Tc/Yy/Chr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   111
               Top             =   2340
               Width           =   1905
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Ord No/Sr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   110
               Top             =   2625
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
               Left            =   5775
               TabIndex        =   109
               Top             =   2340
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
               Left            =   3075
               TabIndex        =   108
               Top             =   2340
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
               Left            =   5355
               TabIndex        =   107
               Top             =   2340
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
               Left            =   2655
               TabIndex        =   106
               Top             =   2340
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
               Left            =   5475
               TabIndex        =   105
               Top             =   2625
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
               Index           =   31
               Left            =   2775
               TabIndex        =   104
               Top             =   2625
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   103
               Top             =   1215
               Width           =   1605
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
               TabIndex        =   102
               Top             =   645
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
               TabIndex        =   101
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
               Index           =   28
               Left            =   5775
               TabIndex        =   100
               Top             =   630
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
               Left            =   3075
               TabIndex        =   99
               Top             =   630
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
               Left            =   5355
               TabIndex        =   98
               Top             =   630
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
               Left            =   2655
               TabIndex        =   97
               Top             =   630
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
               Left            =   5475
               TabIndex        =   96
               Top             =   915
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
               Left            =   2775
               TabIndex        =   95
               Top             =   915
               Width           =   105
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
               TabIndex        =   94
               Top             =   30
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
               Left            =   4635
               TabIndex        =   93
               Top             =   30
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
               Left            =   1935
               TabIndex        =   92
               Top             =   30
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpFgList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepFbList
Dim Rep1 As New EmrRepFrList
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ws_MultiCoMod As Boolean

Private Sub SetGroupSort()
  Select Case UCase(adc.MenuCd)
  Case Is = UCase("RepFbLst"), UCase("RepFwbLst")
    GRP_REP.Add "Fg Vch No", "FdCoCd+FdTc+FdYy+FdChr+str(FdNo)", "FdCoCd+'/'+FdTc+'/'+FdYy+'/'+FdChr+'/'+LTrim(Str(FdNo))", "", "", "Fg Dt", "", "", ""
    '###  '6.1
    GRP_REP.Add "Fg Dt", "convert(varchar(8), FgDt, 112)", "convert(varchar(8), FgDt, " + CStr(DtFmtSlahVal()) + ")", "", "", "", "", "", ""
    '###
    GRP_REP.Add "Bag No", "FdBYy+FdBChr+str(FdBNo)", "FdBYy+'/'+FdBChr+'/'+LTrim(Str(FdBNo))", "", "", "Fg Dt,Fg Vch No", "", "", ""
    '**Geeta***Emr208***
      GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "", "", "", "(Select Substring(PDesc,1,120) From Param Where PTyp= 'DMCTG' and PmCd= DmCtg and PSCd= '') "
    '**Geeta***Emr208***
    'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp=FdDmTcTyp in DsgMst where condition
    GRP_REP.Add "Design Cd", "FdDmCd", "", "", "", "Design Ctg", "", "", "(Select Substring(DmDesc,1,120) From DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey=FdPrtKey and ", "") + " DmTcTyp=FdDmTcTyp and DmCd= FdDmCd and DmSz= '') "
      If ws_MultiCoMod = True Then
        GRP_REP.Add "Company Code", "FdCoCd", "FdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=FdCoCd) "
      End If

    GRP_REP.Add "Prd Ord No", "FdPrdOdTc+FdPrdOdYy+FdPrdOdChr+str(FdPrdOdNo)", "FdPrdOdTc+'/'+FdPrdOdYy+'/'+FdPrdOdChr+'/'+LTrim(Str(FdPrdOdNo))", "", "", "Prd Customer", "", "", "'('+FdPrdCmCd+ ')' "
'???? check
    GRP_REP.Add "Prd Ord Sr", "FdPrdOdTc+FdPrdOdYy+FdPrdOdChr+str(FdPrdOdNo)+str(FdPrdOdsr)", "FdPrdOdTc+'/'+FdPrdOdYy+'/'+FdPrdOdChr+'/'+LTrim(Str(FdPrdOdNo))+'/'+LTrim(Str(FdPrdOdSr))", "", "", "Design Cd,Prd Customer,Prd Order No", "", "", "'('+FdPrdCmCd+ ')' "
      '******Geeta*****************Emr206****05/04/04*******
      '**********If UserLevel>6 then don't display the Customer Name
    ' ***** Manali 3.8.0 - EmrFrmMDI.TXT_UCD.text Replaced by gs_UsrCd
    GRP_REP.Add "Prd Customer", "FdPrdCmCd", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end ) From CustMst Where CmCtg= 'C' and CmCd= FdPrdCmCd) "
    'GRP_REP.Add "Prd Customer", "FdPrdCmCd", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end ) From CustMst Where CmCtg= 'C' and CmCd= FdPrdCmCd) "
    GRP_REP.Add "Exp Ord No", "FdExpOdTc+FdExpOdYy+FdExpOdChr+str(FdExpOdNo)", "FdExpOdTc+'/'+FdPrdOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo))", "", "", "Exp Customer", "", "", "'('+FdExpCmCd+ ')' "
'???? check
    GRP_REP.Add "Exp Ord Sr", "FdExpOdTc+FdExpOdYy+FdExpOdChr+str(FdExpOdNo)+str(FdExpOdSr)", "FdExpOdTc+'/'+FdPrdOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo))+'/'+LTrim(Str(FdExpOdSr))", "", "", "Design Cd,Exp Customer,Exp Order No", "", "", "'('+FdExpCmCd+ ')' "
    
    ' ***** Manali 3.8.0 - EmrFrmMDI.TXT_UCD.text Replaced by gs_UsrCd
    GRP_REP.Add "Exp Customer", "FdExpCmCd", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= FdExpCmCd) "
    'GRP_REP.Add "Exp Customer", "FdExpCmCd", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= FdExpCmCd) "
    GRP_REP.Add "Pck List", "FdPlYy+FdPlChr+str(FdPlNo)", "FdPlYy+'/'+FdPlChr+'/'+LTrim(Str(FdPlNo))", "", "", "", "", "", ""
    GRP_REP.Add "Inv Vch No", "FdInTc+FdInYy+FdInChr+str(FdInNo)", "FdInTc+'/'+FdInYy+'/'+FdInChr+'/'+LTrim(Str(FdInNo))", "", "", "Inv Exp No", "", "", ""
    GRP_REP.Add "Inv Exp No", "FdInExpNo", "", "", "", "Inv Vch No", "", "", ""
  Case Is = UCase("RepFrLst"), UCase("RepFwrLst")
    GRP_REP.Add "Fg Vch No", "FdCoCd+FdTc+FdYy+FdChr+str(FdNo)", "FdCoCd+'/'+FdTc+'/'+FdYy+'/'+FdChr+'/'+LTrim(Str(FdNo))", "wFgNo", "hFgNo", "Fg Dt", "", "", ""
    '###  '6.1
    GRP_REP.Add "Fg Dt", "convert(varchar(8), FgDt, 112)", "convert(varchar(8), FgDt, " + CStr(DtFmtSlahVal()) + ")", "FgDt", "hFgDt", "", "", "", ""
    '###
    'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp=FdDmTcTyp in DsgMst where condition
    GRP_REP.Add "Design Cd", "FdDmCd", "", "FdDmCd", "hFdDmCd", "", "", "", "(Select Substring(DmDesc,1,120) From DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey=FdPrtKey and ", "") + " DmTcTyp= FdDmTcTyp and DmCd= FdDmCd and DmSz= '') "
    GRP_REP.Add "Exp Ord No", "FdExpOdTc+FdExpOdYy+FdExpOdChr+str(FdExpOdNo)", "FdExpOdTc+'/'+FdPrdOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo))", "wExpOrdNo,FdExpCmCd", "hExpOrdNo,hFdExpCmCd", "Exp Customer", "", "", "'('+FdExpCmCd+ ')' "
'???? check
    GRP_REP.Add "Exp Ord Sr", "FdExpOdTc+FdExpOdYy+FdExpOdChr+str(FdExpOdNo)+str(FdExpOdSr)", "FdExpOdTc+'/'+FdPrdOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo))+'/'+LTrim(Str(FdExpOdSr))", "wExpOrdSr,FdExpCmCd", "hExpOrdSr,hFdExpCmCd", "Design Cd,Exp Customer,Exp Order Sr", "", "", "'('+FdExpCmCd+ ')' "
    
    ' ***** Manali 3.8.0 - EmrFrmMDI.TXT_UCD.text Replaced by gs_UsrCd
    GRP_REP.Add "Exp Customer", "FdExpCmCd", "", "FdExpCmCd", "hFdExpCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end ) From CustMst Where CmCtg= 'C' and CmCd= FdExpCmCd) "
    'GRP_REP.Add "Exp Customer", "FdExpCmCd", "", "FdExpCmCd", "hFdExpCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end ) From CustMst Where CmCtg= 'C' and CmCd= FdExpCmCd) "
    GRP_REP.Add "Pck List", "FdPlYy+FdPlChr+str(FdPlNo)", "FdPlYy+'/'+FdPlChr+'/'+LTrim(Str(FdPlNo))", "wPckNo", "hPckNo", "", "", "", ""
    GRP_REP.Add "Inv Vch No", "FdInTc+FdInYy+FdInChr+str(FdInNo)", "FdInTc+'/'+FdInYy+'/'+FdInChr+'/'+LTrim(Str(FdInNo))", "wInNo", "hInNo", "Inv Exp No", "", "", ""
    GRP_REP.Add "Inv Exp No", "FdInExpNo", "", "FdInExpNo", "hFdInExpNo", "Inv Vch No", "", "", ""
  End Select
End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  '*** Report Sql ***
' Zubin 212
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  FdTc, FdYy, FdChr, FdNo, FdSr, FgDt, FdBYy, FdBChr, FdBNo,DmCtg,FdDmCd, FdSfx, FdDmSz,
'  FdQty, FdGrWt, FdPrdOdTc, FdPrdOdYy, FdPrdOdChr, FdPrdOdNo, FdPrdOdSr, FdPrdCmCd,
'  FdExpCmCd, FdExpOdTc, FdExpOdYy, FdExpOdChr, FdExpOdNo, FdExpOdSr,FdPlYy, FdPlChr,
'  FdPlNo, FdInTc, FdInYy, FdInChr, FdInNo, FdInExpNo,
'  FrSrNo, FrRmCd, FrLotNo, FrRmSz, FrRmStkRt, PValue3 as qPtrYN, FrRmQty, FrRmWt, 0 as qMul,FdDesc,
'  FdSubLoc
'  From Fg, Fgd, FgRm, Param,DsgMst where 1= 2
  '*** Report Sql ***   '*******Geeta*****New Field DmCtg added***Emr208 in The Query, Scope,Group
  
  '*** This Report has To Be Taken Always With The Detail Level ***
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String
  Dim wPncCnd As String, wStr As String
  
  ' ******** Manali 3.03 - 09/07/08 - RmSzDesc
  Dim wRmSzDesc As String
  ' ******** Manali 3.03 - 09/07/08 - RmSzDesc
  
  Dim wAddlCnd As String '****** Sachin 3.02.0 - New fields in where clause
  
  Set Rep = Nothing
  Set Rep1 = Nothing
  
  If UCase(adc.MenuCd) = UCase("RepFbLst") Then gs_FgTcTyp = "FB"
  If UCase(adc.MenuCd) = UCase("RepFwbLst") Then gs_FgTcTyp = "FWB"
  If UCase(adc.MenuCd) = UCase("RepFrLst") Then gs_FgTcTyp = "FR"
  If UCase(adc.MenuCd) = UCase("RepFwrLst") Then gs_FgTcTyp = "FWR"
  
  Select Case UCase(adc.MenuCd)
  Case UCase("RepFbLst")
    Set adc.RepSource = Rep
    Set mRep = Rep
  Case UCase("RepFwbLst")
    Set adc.RepSource = Rep
    Set mRep = Rep
  Case UCase("RepFrLst")
    Set adc.RepSource = Rep1
    Set mRep = Rep1
  Case UCase("RepFwrLst")
    Set adc.RepSource = Rep1
    Set mRep = Rep1
  End Select
  
  Call DispCoNm
  Call SetMwName(mRep)   ' ***** Manali 3.6.1 - Microway footer
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
  
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then
    Rep.DET2.Suppress = True: Rep.PH2.Suppress = True
    Rep.BoxHd.Suppress = True: Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True: Rep.DetSec2.Suppress = True
  
    Rep1.BoxHd.Suppress = True: Rep1.BoxDet.Suppress = True
    Rep1.BoxFt.Suppress = True: Rep1.DetSec2.Suppress = True
  End If
  
  ' ***** Manali 3.03 - 09/07/08 - RmSzDesc
  wRmSzDesc = "case when (Select PDesc225 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RisRmCtg, '') and PSCd= RmSCtg)= 'Y' then IsNull(RisRmSzDesc, '') Else '' End "
  ' ***** Manali 3.03 - 09/07/08 - RmSzDesc
  
  If adc("UoYN2") = "Y" Then
    wPncCnd = " and FdQty- FdDespQty> 0 "
    wStr = " (case when FdQty> 0 then round((IsNull(FrRmQty, 0)+0.0)* (FdQty- FdDespQty)/FdQty , 0) else 0 end) as FrRmQty, " + _
           "(case when FdQty> 0 then IsNull(FrRmWt, 0)* (FdQty- FdDespQty)/FdQty else 0 end) as FrRmWt, "
  Else
    wPncCnd = " "
    wStr = " IsNull(FrRmQty, 0) as FrRmQty, IsNull(FrRmWt, 0) as FrRmWt, "
  End If
  
  '****** Sachin 3.02.0 - New Fields in Where Clause
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " FdDt >= '" + Trim(adc("UODTFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " FdDt <= '" + Trim(adc("UODTTO")) + "'", "")
  
  ' Zubin 212
  '****** Sachin 3.02 - Id fields in Joins
  ' ***** Manali 3.03 - 09/07/08 - RmSzDesc added and Join with RmIdSz added
  'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp=FdDmTcTyp in DsgMst join
  wSqlStrg = "Select " + grpflds + ", " + _
             "FdCoCd, FdTc, FdYy, FdChr, FdNo, FdSr, FgDt, FdBYy, FdBChr, FdBNo,DmCtg,FdDmCd, FdSfx, " + _
             "FdDmSz, FdQty, FdGrWt, FdPrdOdTc, FdPrdOdYy, FdPrdOdChr, FdPrdOdNo, " + _
             "FdPrdOdSr, FdPrdCmCd,FdExpCmCd, FdExpOdTc, FdExpOdYy, FdExpOdChr, FdExpOdNo, " + _
             "FdExpOdSr, FdPlYy, FdPlChr, FdPlNo, FdInTc, FdInYy, FdInChr, FdInNo, " + _
             "FdInExpNo, IsNull(FrSrNo, 0) as FrSrNo, IsNull(FrRmCd, '') as FrRmCd, IsNull(FrLotNo, '') as FrLotNo, IsNull(FrRmSz, 0) as FrRmSz, IsNull(FrRmSz2, 0) as FrRmSz2, IsNull(FrRmSz3, 0) as FrRmSz3, " + wRmSzDesc + " as qRmSzDesc, " + _
             "IsNull(FrRmStkRt, 0) as FrRmStkRt, IsNull((Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RmCtg, '') and PSCd= IsNull(RmSCtg, '')), '') as qPtrYN, " + _
             wStr + "IsNull((Case when F.LocTyp='F' and FrFrRmDc='D' then 1 " + _
             "when F.LocTyp='F' and FrFrRmDc='C' then -1 " + _
             "when T.LocTyp='F' and FrToRmDc='D' then 1 " + _
             "when T.LocTyp='F' and FrToRmDc='C' then -1 " + _
             "Else 0 end), 0) as qMul, FdDesc, FdSubLoc " + _
             "From Fgd join Fg on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and FgTc= FdTc and FgYy= FdYy " + _
             " and FgChr= FdChr and FgNo= FdNo " + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey=FdPrtKey", "") + _
             " left outer join FgRm on FrFdIdNo=FdIdNo and FrCoCd= FdCoCd and FrTc= FdTc and FrYy= FdYy " + _
             " and FrChr= FdChr and FrNo= FdNo and FrSr= FdSr " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=FdPrtKey", "") + _
             " left outer join RmMst on RmCd= FrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=FrPrtKey", "") + _
             " left outer join Loc F on F.LocCoCd= FrCoCd and F.LocCd= FrFrRmLoc " + _
             " left outer join Loc T on T.LocCoCd= FrCoCd and T.LocCd= FrToRmLoc " + _
             " Join DsgMst On DmIdNo=FdDmIdNo and DmTcTyp=FdDmTcTyp and DmCd=FdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=FdPrtKey ", "") + _
             "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
             " where " + IIF(gs_Partition = ctCurrPrtn, " FdPrtKey='" + ctCurrPrtn + "' and ", "") + " (Select PValue from Param where PTyp= 'TC' and PMCd= FdTc)= '" + _
             gs_FgTcTyp + "' " + _
             wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + wPncCnd                     '*****Geeta***Emr208***DsgMst Join & DmCtg Field
  
  
'  wSqlStrg = "Select " + grpflds + ", " + _
'             "FdTc, FdYy, FdChr, FdNo, FdSr, FgDt, FdBYy, FdBChr, FdBNo, FdDmCd, FdSfx, " + _
'             "FdDmSz, FdQty, FdGrWt, FdPrdOdTc, FdPrdOdYy, FdPrdOdChr, FdPrdOdNo, " + _
'             "FdPrdOdSr, FdPrdCmCd,FdExpCmCd, FdExpOdTc, FdExpOdYy, FdExpOdChr, FdExpOdNo, " + _
'             "FdExpOdSr, FdPlYy, FdPlChr, FdPlNo, FdInTc, FdInYy, FdInChr, FdInNo, " + _
'             "FdInExpNo, IsNull(FrSrNo, 0), IsNull(FrRmCd, ''), IsNull(FrRmSz, 0), " + _
'             "IsNull(FrRmStkRt, 0), IsNull(FrRmQty, 0), IsNull(FrRmWt, 0), " + _
'             "IsNull((Case when F.LocTyp='F' and FrFrRmDc='D' then 1 " + _
'             "when F.LocTyp='F' and FrFrRmDc='C' then -1 " + _
'             "when T.LocTyp='F' and FrToRmDc='D' then 1 " + _
'             "when T.LocTyp='F' and FrToRmDc='C' then -1 " + _
'             "Else 0 end), 0) as qMul " + _
'             "From Fgd join Fg on FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo " + _
'             "left outer join FgRm on FrTc= FdTc and FrYy= FdYy and FrChr= FdChr " + _
'             "and FrNo= FdNo and FrSr= FdSr join Loc F on F.LocCd= FrFrRmLoc " + _
'             "join Loc T on T.LocCd= FrToRmLoc " + wCnd
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added
  
  If adc("UoYN") = "N" Then
    Rep.PH2.Suppress = True
    Rep.DET2.Suppress = True
  End If
  
  If adc("UoYN1") = "N" Then
    Rep.BoxHd.Suppress = True
    Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True
    Rep.DetSec2.Suppress = True
  
    Rep1.BoxHd.Suppress = True
    Rep1.BoxDet.Suppress = True
    Rep1.BoxFt.Suppress = True
    Rep1.DetSec2.Suppress = True
  End If
  
  ' ***** Manali 3.03 - 09/07/08 - RmSzDesc
  
  ' GRP_REP.SetFormula Rep, "wFrRmSz", "If {rdo.qRmSzDesc} <> '' Then {rdo.qRmSzDesc} " + _
                                     "Else (If {rdo.qPtrYn}= 'Y' Then ToText({rdo.FrRmSz}, 4) Else ToText({rdo.FrRmSz}, 3))"
  'GRP_REP.SetFormula Rep, "wRmSz", "If {rdo.qRmSzDesc} <> '' Then {rdo.qRmSzDesc} "
 
  GRP_REP.SetFormula Rep, "wRmSz", "If {rdo.qRmSzDesc} <> '' Then {rdo.qRmSzDesc} " + _
                                      "Else CustRmSzDisp ({rdo.FrRmSz},{rdo.FrRmSz2},{rdo.FrRmSz3}) "
  
 
 
  'KC.3- StkRt-when Show Stock Rate=YES then Stock Rate is visible in report
  'if it is NO then Stock Rate is invisible in report
 
  If adc("UoYN3") = "Y" Then
    Rep.hTdRmStkRt.SetText "Stock Rt"
    Rep.FldRmStkRt.Suppress = False
  Else
    Rep.hTdRmStkRt.SetText ""
    Rep.FldRmStkRt.Suppress = True
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.FrRmCd}", 8)
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     ' Tag = "ShowLine"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1     ' Tag = "ShowBoxDet"
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  Case Is = 2     ' Tag = "ShowPndFg"
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  'StkRt-To select this option(Show Stock rate) either Seo user can give their password or
  ' can enter SEOPWD password otherwise this option can't not select
  Case Is = 3
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
' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added

' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added

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
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)

  Call InitProp(Me)
  If ws_MultiCoMod = True Then
    Set adc.FirNKeyCtl = adc("UoCoCdFr")
    adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
  Else
    Set adc.FirNKeyCtl = adc("UoFgTcFr")
    adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
  End If
  

'    Set adc.FirNKeyCtl = adc("UoFgTcFr")
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    'indigo.164 sketch design tctype and order tc added
    gs_DmTcTyp = "DM','SM"
    gs_OmTcTyp = "SO','OS"

    gs_CmCtg = "C"
    If UCase(adc.MenuCd) = UCase("RepFbLst") Then gs_FgTcTyp = "FB"
    If UCase(adc.MenuCd) = UCase("RepFwbLst") Then gs_FgTcTyp = "FWB"
    If UCase(adc.MenuCd) = UCase("RepFrLst") Then gs_FgTcTyp = "FR"
    If UCase(adc.MenuCd) = UCase("RepFwrLst") Then gs_FgTcTyp = "FWR"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    If UCase(adc.MenuCd) = UCase("RepFrLst") Or UCase(adc.MenuCd) = UCase("RepFwrLst") Then
      adc("UoYN") = ""
      Lbl2ndLine.Visible = False: adc("UoYN").Visible = False: adc("UoYN").Enabled = False
    End If
    
    If UCase(adc.MenuCd) = UCase("RepFwrLst") Or UCase(adc.MenuCd) = UCase("RepFwbLst") Then
      adc("UoYN2") = ""
      LblShowPndFg.Visible = False: adc("UoYN2").Visible = False: adc("UoYN2").Enabled = False
    End If
      
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  If UCase(adc.MenuCd) = UCase("RepFbLst") Then gs_FgTcTyp = "FB"
  If UCase(adc.MenuCd) = UCase("RepFwbLst") Then gs_FgTcTyp = "FWB"
  If UCase(adc.MenuCd) = UCase("RepFrLst") Then gs_FgTcTyp = "FR"
  If UCase(adc.MenuCd) = UCase("RepFwrLst") Then gs_FgTcTyp = "FWR"
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"

  If UCase(adc.MenuCd) = UCase("RepFbLst") Then gs_FgTcTyp = "FB"
  If UCase(adc.MenuCd) = UCase("RepFwbLst") Then gs_FgTcTyp = "FWB"
  If UCase(adc.MenuCd) = UCase("RepFrLst") Then gs_FgTcTyp = "FR"
  If UCase(adc.MenuCd) = UCase("RepFwrLst") Then gs_FgTcTyp = "FWR"
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
 If ws_MultiCoMod = False Then _
      adc("UoCoCdFr") = gs_CoCd: adc("UoCoCdTo") = gs_CoCd: adc("UoCoCdSel") = "": adc("UoCoCdFr").Enabled = False
 
 ' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
 If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
 If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
 If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
 ' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
  'StkRt-Set Show StockRate option is No
  adc("UoYN3") = "N"
  
End Sub
Private Sub DispCoNm()
'    Rep.wCoCd.SetText gs_CoNm
'   Rep1.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = True, gs_CoCd, adc("UOCOCDFR")) + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
  Rep.wCoCd.SetText ws_HName
  Rep1.wCoCd.SetText ws_HName
  
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
'  Rep.wSrvrDate.SetText ws_DtTm
'  Rep1.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  Rep.wSrvrTime.SetText ws_DtTm
'  Rep1.wSrvrTime.SetText ws_DtTm
  
   If ws_MultiCoMod = True Then
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdTo.SetText adc("UOCOCDTO"): Rep.TxtCoCdSel.SetText adc("UOCOCDSEL")
  Else
      Rep.TxtCoCdTo.Suppress = True: Rep.TxtCoCdSel.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True: Rep.TxtCoCdSelLbl.Suppress = True
  End If
  Rep.TxtCoCdFr.SetText adc("UOCOCDFR")
 
  ' ******** Manali 3.2.2 - User date and time in reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  Rep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  Rep1.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm

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
    Set Rep1 = Nothing
    '*** (09/08/05)
End Sub
'StkRt-if password is change to blank, Show StockRate check box will turn into blank
Private Sub TxtPwd_Change()
    If Trim(TxtPwd.text) = "" Then
        ChkBoxArr(3).Value = Unchecked
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

