VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpTxnLst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Listing"
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
      TabIndex        =   84
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
      TabIndex        =   85
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
      TabIndex        =   94
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   95
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
         TabIndex        =   96
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
      TabIndex        =   93
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   90
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
      TabIndex        =   86
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
      TabIndex        =   89
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   88
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpTxnLst.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpTxnLst.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   92
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   87
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
            TabIndex        =   91
            Top             =   360
            Width           =   15105
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
               Index           =   0
               Left            =   4080
               TabIndex        =   83
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Stock Rate"
               Top             =   8400
               Width           =   2100
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1515
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   82
               Top             =   8400
               Width           =   2340
            End
            Begin VB.Frame FraOptAck 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1560
               TabIndex        =   143
               ToolTipText     =   "Show Open Or Closed Or All Bags"
               Top             =   7920
               Width           =   3465
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
                  TabIndex        =   79
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
                  TabIndex        =   80
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   945
               End
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
                  TabIndex        =   81
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.Frame FraOptAmt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1560
               TabIndex        =   139
               ToolTipText     =   "Select Whether To Show Stock Amount Or Purchase Amount"
               Top             =   7440
               Width           =   5595
               Begin VB.OptionButton OptAmt 
                  Caption         =   "Stock Amt "
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
                  TabIndex        =   77
                  ToolTipText     =   "Select Whether To Show Stock Amount Or Purchase Amount"
                  Top             =   0
                  Width           =   1635
               End
               Begin VB.OptionButton OptAmt 
                  Caption         =   "Purchase Amt"
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
                  Left            =   2100
                  TabIndex        =   78
                  ToolTipText     =   "Select Whether To Show Stock Amount Or Purchase Amount"
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4005
               TabIndex        =   44
               ToolTipText     =   "Enter To Design Category"
               Top             =   3975
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
               Left            =   6510
               TabIndex        =   45
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   3975
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
               Left            =   1515
               TabIndex        =   43
               ToolTipText     =   "Enter From Design Category"
               Top             =   3975
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
               Left            =   4005
               TabIndex        =   50
               ToolTipText     =   "Enter To Design Code"
               Top             =   4545
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
               Left            =   6510
               TabIndex        =   51
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   4545
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
               Left            =   1515
               TabIndex        =   49
               ToolTipText     =   "Enter From Design Code"
               Top             =   4545
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
               CmpStr          =   "TDt <="
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
               CmpStr          =   "TDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   4785
               TabIndex        =   56
               ToolTipText     =   "Enter To Order Year"
               Top             =   4830
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "BOdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2295
               TabIndex        =   53
               ToolTipText     =   "Enter From Order Year"
               Top             =   4830
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "BOdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1515
               TabIndex        =   52
               ToolTipText     =   "Enter From Order Tc"
               Top             =   4830
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "BOdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   2715
               TabIndex        =   54
               ToolTipText     =   "Enter From Order Character"
               Top             =   4830
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "BOdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   4005
               TabIndex        =   55
               ToolTipText     =   "Enter To Order Tc"
               Top             =   4830
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "BOdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   5205
               TabIndex        =   57
               ToolTipText     =   "Enter To Order Character"
               Top             =   4830
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "BOdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   1515
               TabIndex        =   58
               ToolTipText     =   "Enter From Order Number"
               Top             =   5115
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "BOdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   4005
               TabIndex        =   60
               ToolTipText     =   "Enter To Order Number"
               Top             =   5115
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "BOdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4005
               TabIndex        =   35
               ToolTipText     =   "Enter To By Worker"
               Top             =   3405
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKTO"
               IdName          =   "UOWRKTO"
               CmpStr          =   "Td.TdByWrk <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   6510
               TabIndex        =   36
               ToolTipText     =   "Enter By Worker Selection"
               Top             =   3405
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOWRKSEL"
               IdName          =   "UOWRKSEL"
               CmpStr          =   "Td.TdByWrk In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1515
               TabIndex        =   34
               ToolTipText     =   "Enter From By Worker"
               Top             =   3405
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKFR"
               IdName          =   "UOWRKFR"
               CmpStr          =   "Td.TdByWrk >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4005
               TabIndex        =   47
               ToolTipText     =   "Enter To Design Production Category"
               Top             =   4260
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
               Left            =   6510
               TabIndex        =   48
               ToolTipText     =   "Enter Design Production Category Selection"
               Top             =   4260
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
               Left            =   1515
               TabIndex        =   46
               ToolTipText     =   "Enter From Design Production Category"
               Top             =   4260
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
               Index           =   24
               Left            =   4785
               TabIndex        =   5
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   525
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
               CmpStr          =   "Td.TdYy <="
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
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
               CmpStr          =   "Td.TdYy >="
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
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
               CmpStr          =   "Td.TdTc >="
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
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
               CmpStr          =   "Td.TdChr >="
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
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
               CmpStr          =   "Td.TdTc <="
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
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
               CmpStr          =   "Td.TdChr <="
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
               DataField       =   "UOTDSRTO"
               IdName          =   "UOTDSRTO"
               CmpStr          =   "Td.TdSr <="
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
               DataField       =   "UOTDSRFR"
               IdName          =   "UOTDSRFR"
               CmpStr          =   "Td.TdSr >="
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
               DataField       =   "UOTNOFR"
               IdName          =   "UOTNOFR"
               CmpStr          =   "Td.TdNo >="
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
               DataField       =   "UOTNOTO"
               IdName          =   "UOTNOTO"
               CmpStr          =   "Td.TdNo <="
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
               Index           =   26
               Left            =   4890
               TabIndex        =   61
               ToolTipText     =   "Enter To Order Serial"
               Top             =   5115
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "BOdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   2400
               TabIndex        =   59
               ToolTipText     =   "Enter From Order Serial"
               Top             =   5115
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "BOdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   4005
               TabIndex        =   39
               ToolTipText     =   "Enter To Bag Year"
               Top             =   3690
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYTO"
               IdName          =   "UOBYYTO"
               CmpStr          =   "Td.TdBYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1515
               TabIndex        =   37
               ToolTipText     =   "Enter From Bag Year"
               Top             =   3690
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYFR"
               IdName          =   "UOBYYFR"
               CmpStr          =   "Td.TdBYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   1920
               TabIndex        =   38
               ToolTipText     =   "Enter From Bag Character"
               Top             =   3690
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOBCHRFR"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "Td.TdBChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4425
               TabIndex        =   40
               ToolTipText     =   "Enter To Bag Character"
               Top             =   3690
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOBCHRTO"
               IdName          =   "UOBCHRTO"
               CmpStr          =   "Td.TdBChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   2715
               TabIndex        =   41
               ToolTipText     =   "Enter From Bag Number"
               Top             =   3690
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOFR"
               IdName          =   "UOBNOFR"
               CmpStr          =   "Td.TdBNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   5235
               TabIndex        =   42
               ToolTipText     =   "Enter To Bag Number"
               Top             =   3690
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOTO"
               IdName          =   "UOBNOTO"
               CmpStr          =   "Td.TdBNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   4005
               TabIndex        =   63
               ToolTipText     =   "Enter To Rejection Code"
               Top             =   5400
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOREJCDTO"
               IdName          =   "UOREJCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   6510
               TabIndex        =   64
               ToolTipText     =   "Enter Rejection Code Selection"
               Top             =   5400
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOREJCDSEL"
               IdName          =   "UOREJCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1515
               TabIndex        =   62
               ToolTipText     =   "Enter From Rejection Code"
               Top             =   5400
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOREJCDFR"
               IdName          =   "UOREJCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   4005
               TabIndex        =   66
               ToolTipText     =   "Enter To Supplier"
               Top             =   5685
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
               TabIndex        =   67
               ToolTipText     =   "Enter Supplier Selection"
               Top             =   5685
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
               TabIndex        =   65
               ToolTipText     =   "Enter From Supplier"
               Top             =   5685
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
               TabIndex        =   26
               ToolTipText     =   "Enter To Location"
               Top             =   2520
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   6510
               TabIndex        =   27
               ToolTipText     =   "Enter Location Selection"
               Top             =   2520
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
               TabIndex        =   25
               ToolTipText     =   "Enter From Location"
               Top             =   2520
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   59
               Left            =   4005
               TabIndex        =   72
               ToolTipText     =   "Enter To Loss Date"
               Top             =   6540
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
               TabIndex        =   71
               ToolTipText     =   "Enter From Loss Date"
               Top             =   6540
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
               Index           =   61
               Left            =   4005
               TabIndex        =   69
               ToolTipText     =   "Enter To Bill Of Entry Number "
               Top             =   5970
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
               TabIndex        =   70
               ToolTipText     =   "Enter Bill Of Entry Number Selection"
               Top             =   5970
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
               TabIndex        =   68
               ToolTipText     =   "Enter From Bill Of Entry Number "
               Top             =   5970
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBOEFR"
               IdName          =   "UOBOEFR"
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
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "Td.TdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   64
               Left            =   5475
               TabIndex        =   74
               ToolTipText     =   "Enter Voucher Currency  US$[Y]/Rs[N]  and ' ' for ALL records in US$ "
               Top             =   6885
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
               Index           =   65
               Left            =   1515
               TabIndex        =   75
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   7170
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOBAGTYP"
               IdName          =   "UOBAGTYP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   66
               Left            =   1515
               TabIndex        =   31
               ToolTipText     =   "Enter From Lot No."
               Top             =   3105
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataType        =   4
               DataField       =   "UOLOTNOFR"
               IdName          =   "UOLOTNOFR"
               CmpStr          =   "Rm.TdLotNo>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   67
               Left            =   4005
               TabIndex        =   32
               ToolTipText     =   "Enter To Lot No."
               Top             =   3110
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOLOTNOTO"
               IdName          =   "UOLOTNOTO"
               CmpStr          =   "Rm.TdLotNo<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   68
               Left            =   6510
               TabIndex        =   33
               ToolTipText     =   "Enter Lot No. Selection"
               Top             =   3110
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOLOTNOSEL"
               IdName          =   "UOLOTNOSEL"
               CmpStr          =   "Rm.TdLotNo In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   69
               Left            =   1515
               TabIndex        =   73
               ToolTipText     =   "Enter Currency Code Whose Vouchers You Want To View Or Leave It Blank To View All Vouchers In Base Currency"
               Top             =   6840
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOCURCDFR"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   7800
               TabIndex        =   76
               ToolTipText     =   "Enter Y to Show Stock Amt, N to Show Purchase Amt"
               Top             =   7440
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
               Index           =   71
               Left            =   4005
               TabIndex        =   20
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1950
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
               Top             =   1950
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
               Top             =   1950
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
               Top             =   2235
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
               Top             =   2235
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
               Top             =   2235
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
               Index           =   74
               Left            =   7800
               TabIndex        =   142
               ToolTipText     =   "Show Yes Or No Or All Acknowledge"
               Top             =   7920
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
               Left            =   6525
               TabIndex        =   144
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   8370
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   423
               Enabled         =   0   'False
               DataField       =   "UOYN3"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   4005
               TabIndex        =   29
               ToolTipText     =   "Enter Location To Range"
               Top             =   2805
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UORECLOCTO"
               IdName          =   "UORECLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   76
               Left            =   6510
               TabIndex        =   30
               ToolTipText     =   "Enter Location Selection"
               Top             =   2805
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UORECLOCSEL"
               IdName          =   "UORECLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   80
               Left            =   1515
               TabIndex        =   28
               ToolTipText     =   "Enter Location From Range"
               Top             =   2805
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UORECLOCFR"
               IdName          =   "UORECLOCFR"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Receiving Loc"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   148
               Top             =   2805
               Width           =   1425
            End
            Begin VB.Label LblBagSlh4 
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
               Left            =   5160
               TabIndex        =   147
               Top             =   3645
               Width           =   105
            End
            Begin VB.Label LblBagSlh3 
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
               Left            =   2640
               TabIndex        =   146
               Top             =   3645
               Width           =   105
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
               Left            =   60
               TabIndex        =   145
               Top             =   8400
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
               Left            =   60
               TabIndex        =   141
               ToolTipText     =   "Location"
               Top             =   7800
               Visible         =   0   'False
               Width           =   1485
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
               TabIndex        =   140
               Top             =   1950
               Width           =   1425
            End
            Begin VB.Label LblStkVal 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Amount"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   465
               Left            =   60
               TabIndex        =   138
               ToolTipText     =   "Location"
               Top             =   7470
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
               Left            =   60
               TabIndex        =   137
               ToolTipText     =   "Location"
               Top             =   6840
               Width           =   1485
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
               TabIndex        =   136
               ToolTipText     =   "Location"
               Top             =   7170
               Width           =   1035
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
               TabIndex        =   135
               Top             =   3060
               Width           =   1425
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
               Left            =   4020
               TabIndex        =   134
               ToolTipText     =   "Location"
               Top             =   6885
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
               TabIndex        =   133
               Top             =   240
               Width           =   1335
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
               TabIndex        =   132
               Top             =   5970
               Width           =   1425
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
               TabIndex        =   131
               Top             =   6540
               Width           =   1425
            End
            Begin VB.Label LblLsLoc 
               BackStyle       =   0  'Transparent
               Caption         =   "Issuing Loc"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Top             =   2535
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
               TabIndex        =   129
               Top             =   5685
               Width           =   1425
            End
            Begin VB.Label LblRejCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Rejection Cd"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Top             =   5400
               Width           =   1425
            End
            Begin VB.Label LblBagSlh1 
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
               TabIndex        =   127
               Top             =   3645
               Width           =   105
            End
            Begin VB.Label LblBagSlh2 
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
               Left            =   4365
               TabIndex        =   126
               Top             =   3645
               Width           =   105
            End
            Begin VB.Label LblBagNo 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   12960
               TabIndex        =   125
               Top             =   480
               Visible         =   0   'False
               Width           =   1605
            End
            Begin VB.Label LblBagYyChr 
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
               Left            =   60
               TabIndex        =   124
               Top             =   3690
               Width           =   1365
            End
            Begin VB.Label LblOrdSlh6 
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
               Left            =   4830
               TabIndex        =   123
               Top             =   5115
               Width           =   105
            End
            Begin VB.Label LblOrdSlh5 
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
               Left            =   2340
               TabIndex        =   122
               Top             =   5115
               Width           =   105
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
               TabIndex        =   121
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
               TabIndex        =   120
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
               TabIndex        =   119
               Top             =   2235
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
               TabIndex        =   118
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
               TabIndex        =   117
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
               TabIndex        =   116
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
               TabIndex        =   115
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
               TabIndex        =   114
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
               TabIndex        =   113
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
               TabIndex        =   112
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
               TabIndex        =   111
               Top             =   810
               Width           =   105
            End
            Begin VB.Label LblPrdCtg 
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
               Left            =   60
               TabIndex        =   110
               Top             =   4260
               Width           =   1425
            End
            Begin VB.Label LblByWrk 
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
               Left            =   60
               TabIndex        =   109
               Top             =   3405
               Width           =   1425
            End
            Begin VB.Label LblOrdTcYyChr 
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
               Left            =   60
               TabIndex        =   108
               Top             =   4830
               Width           =   1725
            End
            Begin VB.Label LblOrdNo 
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
               Left            =   60
               TabIndex        =   107
               Top             =   5115
               Width           =   1605
            End
            Begin VB.Label LblOrdSlh4 
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
               Left            =   5145
               TabIndex        =   106
               Top             =   4830
               Width           =   105
            End
            Begin VB.Label LblOrdSlh2 
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
               TabIndex        =   105
               Top             =   4830
               Width           =   105
            End
            Begin VB.Label LblOrdSlh3 
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
               Left            =   4725
               TabIndex        =   104
               Top             =   4830
               Width           =   105
            End
            Begin VB.Label LblOrdSlh1 
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
               Left            =   2235
               TabIndex        =   103
               Top             =   4830
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
               TabIndex        =   102
               Top             =   1095
               Width           =   1425
            End
            Begin VB.Label LblDmCd 
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
               Left            =   60
               TabIndex        =   101
               Top             =   4545
               Width           =   1425
            End
            Begin VB.Label LblDmCtg 
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
               Left            =   60
               TabIndex        =   100
               Top             =   3975
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
               TabIndex        =   99
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
               TabIndex        =   98
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
               TabIndex        =   97
               Top             =   0
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpTxnLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add IssRecLoc-Receiving Loc added as new, for Issue Loc we used Loss Loc which is already
'exist in the form.Loss Loc used only for Recovery transaction. so Loss Loc label
'and From,To,Selection txt box property are changed.

Option Explicit
Dim Rep1 As New EmrRepTxnLstOne
Dim Rep2 As New EmrRepTxnLstTwo
Dim Rep3 As New EmrRepTxnLstThree
Dim mRep As CRAXDRT.Report
Dim moCn As MwfLib.MDOConnection '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean, ws_FrmCaption As String
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  GRP_REP.Add "VoucherNo", "TTc+TYy+TChr+str(TNo)", "TTc+'/'+TYy+'/'+TChr+'/'+LTrim(Str(TNo))", "", "", "Vch date", "", "", ""
  '6.1
  GRP_REP.Add "Vch Date", "convert(char(8),TDt,112)", "convert(varchar, TDt, " + CStr(DtFmtSlahYyyyVal()) + ")", "", "", "", "", "", ""
' **** Zubin 211 **** '
'''''  If gs_TxnTcTyp = "DT" Or gs_TxnTcTyp = "BS" Or gs_TxnTcTyp = "MLT" Or _
'''''    gs_TxnTcTyp = "PR" Or gs_TxnTcTyp = "CNV" Or gs_TxnTcTyp = "REC" Then
  If gs_TxnTcTyp = "DT" Or gs_TxnTcTyp = "BS" Or gs_TxnTcTyp = "MLT" Or _
    gs_TxnTcTyp = "PR" Or gs_TxnTcTyp = "CNV" Or gs_TxnTcTyp = "REC" Or gs_TxnTcTyp = "BM" Then
' **** Zubin 211 **** '
    GRP_REP.Add "Rm Ctg", "IsNull(RmCtg, '')", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = IsNull(RmCtg, '')) "
    GRP_REP.Add "Rm SubCtg", "IsNull(RmCtg+RmSCtg, '')", "'('+ RTrim(IsNull(RmCtg, ''))+ ') '+ IsNull(RmSCtg, '')", "", "", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = IsNull(RmCtg, '') and PSCd= IsNull(RmSCtg, ''))"
    GRP_REP.Add "Rm Code", "IsNull(RmCtg+RmSCtg+RmCd, '')", "'('+ RTrim(IsNull(RmCtg, ''))+'/'+ IsNull(RmSCtg, '')+ ') '+ IsNull(RmCd, '')", "", "", "Rm Ctg,Rm SubCtg", "", "", "IsNull(SubString(RmDesc, 1, 30), '')"
    
    GRP_REP.Add "ByLoc", "Td.TdByLoc", "", "", "", "", "", "", "(Select LocDesc From Loc Where LocCoCd= Td.TdCoCd and LocCd= Td.TdByLoc) "
    
    '4.1.2
    GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "", "", "Rm Ctg,Rm SubCtg", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
  End If
  
  If gs_TxnTcTyp = "BV" Then GRP_REP.Add "Worker", "Td.TdWrk", "", "", "", "", "", "", "(Select vPDesc From vParam Where vPCoCd = Td.TdCoCd and vPTyp= 'WORK' and vPMCd= Td.TdWrk) "
  
' **** Zubin 211 **** '
'''''  If gs_TxnTcTyp = "DT" Or gs_TxnTcTyp = "BS" Or gs_TxnTcTyp = "MLT" Or _
'''''    gs_TxnTcTyp = "BV" Or gs_TxnTcTyp = "REJ" Then
  If gs_TxnTcTyp = "DT" Or gs_TxnTcTyp = "BS" Or gs_TxnTcTyp = "MLT" Or _
    gs_TxnTcTyp = "BV" Or gs_TxnTcTyp = "REJ" Or gs_TxnTcTyp = "BM" Then
' **** Zubin 211 **** '
    If gs_TxnTcTyp <> "BV" Then GRP_REP.Add "ByWorker", "Td.TdByWrk", "", "", "", "", "", "", "(Select vPDesc From vParam Where vPCoCd = Td.TdCoCd and vPTyp= 'WORK' and vPMCd= Td.TdByWrk) "
    GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
    GRP_REP.Add "Prd Ctg", "DmPrdCtg", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'PRDCTG' and PMCd= DmPrdCtg) "
    GRP_REP.Add "Design Code", "DmCd", "", "", "", "Design Ctg,Prd Ctg,(DsgCd)Colour", "", "", "Substring(DmDesc,1,120) "
    GRP_REP.Add "(DsgCd)Colour", "OdDmCd+OdDmCol", "'('+OdDmCd+') '+OdDmCol", "", "", "Design Ctg,Prd Ctg,Design Code", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
    GRP_REP.Add "OrderNo", "BOdTc+BOdYy+BOdChr+str(BOdNo)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "", "", "", "", "", ""
  End If
  
  ' **** Zubin 211 **** '
'''''  If gs_TxnTcTyp = "DT" Or gs_TxnTcTyp = "BS" Or gs_TxnTcTyp = "BV" Then
    If gs_TxnTcTyp = "DT" Or gs_TxnTcTyp = "BS" Or gs_TxnTcTyp = "BV" Or gs_TxnTcTyp = "BM" Then
  ' **** Zubin 211 **** '
       GRP_REP.Add "BagNo", "Td.TdBYy+Td.TdBChr+str(Td.TdBNo)", "Td.TdBYy+'/'+Td.TdBChr+'/'+LTrim(Str(Td.TdBNo))", "", "", "Design Ctg,Prd Ctg,Design Code,Bag Type,Bag Character", "", "", ""
    End If
  
  ' Zubin 212 (gs_TxnTcTyp = "BM" added)
  If gs_TxnTcTyp = "DT" Or gs_TxnTcTyp = "BM" Then
   ' GRP_REP.Add "Rm Sv/Ln", "str(Rm.TdRmSz,6,3)", "", "TdRmSz", "hTdRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
     '****************Geeta**************Emr206*********14/04/04
     Dim wsPtrYN As String
     wsPtrYN = "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg)"
     'GRP_REP.Add "Rm Sv/Ln", "case when " + wsPtrYN + "='Y' then str(Rm.TdRmSz,7,4) else str(Rm.TdRmSz,6,3) end ", "", "TdRmSz", "hTdRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
     ' ****** Manali 3.03 - RmSzDesc added
'     GRP_REP.Add "Rm Sv/Ln", "case when " + wsPtrYN + "='Y' then str(Rm.TdRmSz,7,4) else IsNull(RisRmSzDesc, str(Rm.TdRmSz,6,3)) end ", "", "wRmSzPtr", "hTdRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
     
     GRP_REP.Add "Rm Sv/Sz", "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("Rm.TdRmSz", "Rm.TdRmSz2", "Rm.TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("Rm.TdRmSz", "Rm.TdRmSz2", "Rm.TdRmSz3") + " else RisRmSzDesc end) end ", "", "wRmSz", "hTdRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
 
     '*************
     
     
    GRP_REP.Add "Rm StkRt", "str(Rm.TdRmStkRt,9,2)", "", "TdRmStkRt", "hTdRmStkRt", "Rm Ptr,Rm Sv/Ln,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
    If moCn.GetFldVal("Select HPtrYN from Head where HCoCd ='" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then _
      ' GRP_REP.Add "Rm Ptr", "str(Rm.TdRmSz,6,3)", "", "TdRmSz", "hTdRmSz", "Rm Sv/Ln,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
      '****************Geeta**************Emr206*********14/04/04
      ' GRP_REP.Add "Rm Ptr", "case when " + wsPtrYN + "='Y' then str(Rm.TdRmSz,7,4) else str(Rm.TdRmSz,6,3) end ", "", "TdRmSz", "hTdRmSz", "Rm Sv/Ln,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
      ' *** Manali 3.03 - RmSzDesc added
      GRP_REP.Add "Rm Ptr", "case when " + wsPtrYN + "='Y' then str(Rm.TdRmSz,7,4) else IsNull(RisRmSzDesc, str(Rm.TdRmSz,6,3)) end ", "", "wRmSzPtr", "hTdRmSz", "Rm Sv/Ln,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
'      GRP_REP.Add "Rm Ptr", "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("Rm.TdRmSz", "Rm.TdRmSz2", "Rm.TdRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("Rm.TdRmSz", "Rm.TdRmSz2", "Rm.TdRmSz3") + ") end ", "", "wRmSzPtr", "hTdRmSz", "Rm Sv/Ln,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
    End If
    ' GRP_REP.Add "(RmCd) Sv/Ln", "RmCd+str(Rm.TdRmSz,6,3)", "'('+RmCd+') '+LTrim(str(Rm.TdRmSz,6,3))", "RmCtg,RmSCtg,RmCd,TdRmSz", "hRmCtg,hRmSCtg,hRmCd,hTdRmSz", "Rm Code,Rm Ctg,Rm SubCtg,(RmCd) StkRt,Rm Sv/Ln,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
    '****************Geeta**************Emr206*********14/04/04
    'GRP_REP.Add "(RmCd) Sv/Ln", "RmCd+case when " + wsPtrYN + "='Y' then LTrim(str(Rm.TdRmSz,7,4)) else LTrim(str(Rm.TdRmSz,6,3)) end", "'('+RmCd+') '+case when " + wsPtrYN + "='Y' then LTrim(str(Rm.TdRmSz,7,4)) else LTrim(str(Rm.TdRmSz,6,3)) end", "RmCtg,RmSCtg,RmCd,TdRmSz", "hRmCtg,hRmSCtg,hRmCd,hTdRmSz", "Rm Code,Rm Ctg,Rm SubCtg,(RmCd) StkRt,Rm Sv/Ln,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
    ' ***** Manali 3.03 - 08/07/08
'    GRP_REP.Add "(RmCd) Sv/Ln", "RmCd+case when " + wsPtrYN + "='Y' then LTrim(str(Rm.TdRmSz,7,4)) else LTrim(IsNull(RisRmSzDesc, str(Rm.TdRmSz,6,3))) end", "'('+RmCd+') '+case when " + wsPtrYN + "='Y' then LTrim(str(Rm.TdRmSz,7,4)) else LTrim(IsNull(RisRmSzDesc, str(Rm.TdRmSz,6,3))) end", "RmCtg,RmSCtg,RmCd,wRmSzPtr", "hRmCtg,hRmSCtg,hRmCd,hTdRmSz", "Rm Code,Rm Ctg,Rm SubCtg,(RmCd) StkRt,Rm Sv/Ln,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
    GRP_REP.Add "(RmCd) Sv/Sz", "RmCd+case when " + wsPtrYN + "='Y' then LTrim(" + GetRmSzDBFldNm("Rm.TdRmSz", "Rm.TdRmSz2", "Rm.TdRmSz3") + ") else LTrim((case when RisRmSzDesc is null then " + GetRmSzDBFldNm("Rm.TdRmSz", "Rm.TdRmSz2", "Rm.TdRmSz3") + " else RisRmSzDesc end)) end", "'('+RmCd+') '+case when " + wsPtrYN + "='Y' then LTrim(" + GetRmSzDispFldNm("Rm.TdRmSz", "Rm.TdRmSz2", "Rm.TdRmSz3") + ") else LTrim(IsNull(RisRmSzDesc, " + GetRmSzDispFldNm("Rm.TdRmSz", "Rm.TdRmSz2", "Rm.TdRmSz3") + ")) end", "RmCtg,RmSCtg,RmCd,wRmSz", "hRmCtg,hRmSCtg,hRmCd,hTdRmSz", "Rm Code,Rm Ctg,Rm SubCtg,(RmCd) StkRt,Rm Sv/Ln,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
    GRP_REP.Add "(RmCd) StkRt", "RmCd+str(Rm.TdRmStkRt,9,2)", "'('+RmCd+') '+LTrim(str(Rm.TdRmStkRt,9,2))", "RmCtg,RmSCtg,RmCd,TdRmStkRt", "hRmCtg,hRmSCtg,hRmCd,hTdRmStkRt", "Rm Code,Rm Ctg,Rm SubCtg,(RmCd) Sv/Ln,Rm Sv/Ln,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  End If
  
  If gs_TxnTcTyp = "REJ" Then
    GRP_REP.Add "Rej Code", "Td.TdRjCd", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'REJCD' and PMCd= Td.TdRjCd) "
    GRP_REP.Add "ByLoc", "TdByLoc", "", "", "", "", "", "", "(Select LocDesc From Loc Where LocCoCd= TdCoCd and LocCd= TdByLoc) "
    'IG.147 -OrdSr addded in group sort
    GRP_REP.Add "OrderSr", "BCoCd+BOdTc+BOdYy+BOdChr+str(BOdNo)+str(BOdSr)", "BCoCd+'/'+BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))", "", "", "", "", "", "'('+OdOmCmCd+')'"
  End If
  If gs_TxnTcTyp = "PR" Then
    GRP_REP.Add "Supplier", "TSuppCd", "", "", "", "", "", "", "(Select CmName From CustMst Where CmCtg= 'S' and CmCd= TSuppCd) "
    GRP_REP.Add "BOE No", "TBOENo", "", "", "", "", "", "", "TBOEDt"
  End If
  If gs_TxnTcTyp = "REC" Then
    GRP_REP.Add "Loss Loc", "TLsLoc", "", "", "", "", "", "", "(Select LocDesc From Loc Where LocCoCd= TCoCd and LocCd= TLsLoc) "
  End If
  
  ' **** Zubin 211 **** '
  If gs_TxnTcTyp = "DT" Or gs_TxnTcTyp = "BS" Or gs_TxnTcTyp = "BV" Or gs_TxnTcTyp = "BM" Or gs_TxnTcTyp = "REJ" _
     Or gs_TxnTcTyp = "MLT" Then
    '*** Jay 2.13(CT) ***
    GRP_REP.Add "Bag Type", "(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end)", "(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end)", "", "", "Bag Typ+Chr", "", "", "(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end)"
    GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
    GRP_REP.Add "Bag Typ+Chr", "(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end)+BChr", "(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end)+'/ ' + BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
    'GRP_REP.Add "Bag Type", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end )", "", "", "Bag Typ+Chr", "", "", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)"
    'GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
    'GRP_REP.Add "Bag Typ+Chr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+BChr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+'/ ' + BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
    '*** Jay 2.13(CT) ***
  End If
  ' **** Zubin 211 **** '
  
  ' Zubin 212
  If moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UoCoCdFr") + "' and HLotNoYN= 'Y' ") Then
    If gs_TxnTcTyp = "DT" Or gs_TxnTcTyp = "BS" Or gs_TxnTcTyp = "BM" Or gs_TxnTcTyp = "MLT" Then
      GRP_REP.Add "RM Lot No.", "Rm.TdLotNo", "", "", "", "", "", "", "Rm.TdLotNo"
    ElseIf gs_TxnTcTyp = "PR" Or gs_TxnTcTyp = "REC" Or gs_TxnTcTyp = "CNV" Then
      GRP_REP.Add "RM Lot No.", "TdLotNo", "", "", "", "", "", "", "TdLotNo"
    End If
  End If
  ' Zubin 212
  
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  '*** SQL for DT, BS, MLT
  ' Zubin 212
  ' *** manali 3.03 - qRmSzDesc added
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'TTc, Tyy, TChr, TNo, TDt, TFrRmLoc, TToRmLoc, TFrBLoc, TdSr, TdBYy, TdBChr, TdBNo,
  'TdBDc, TdBQty, TdPtQty, TdBGrWt, TdFrBLoc, TdToBLoc, TdRjDc, TdRjQty, TdRjPtQty,
  'TdRjCd, TdWrk, TdAck, TdPrdYN, TdByLoc as qBagByLoc, TdByWrk as qBagByWrk,
  'TdDesc as qBagDesc, TdSrNo, TdFrRmLoc, TdFrRmDc, RmCtg, TdRmCd, TdLotNo, TdRmSz, RisRmSzDesc as qRmSzDesc, TdRmStkRt,
  'PValue3 as qPtrYN, TdRmQty, TdRmWt, TdToRmLoc, TdByLoc As qRmByLoc, TdByWrk As qRmByWrk,
  'TdDustWt, TdDesc As qRmDesc, LocTyp as qBMLocTyp From Txn, Txnd, RmMst, Bag, DsgMst, Loc,
  'Param Where 1 = 2
  
  '*** SQL for BV, REJ
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'TTc, Tyy, TChr, TNo, TDt, TdSr, TdBYy, TdBChr, TdBNo, TdBDc, TdBQty, TdPtQty, TdBGrWt,
  'TdFrBLoc, TdToBLoc, TdRjCd, TdWrk, TdAck, TdByLoc as qBagByLoc, TdByWrk as qBagByWrk
  'From Txn, Txnd, Bag, DsgMst Where 1 = 2
  
  '*** SQL for PR, REC, CNV, 'MV'
  ' Zubin 212
  ' **** Manali 3.03 - 27/06/08 - qRmStkAmt included for Purchase
  ' *** manali 3.03 - qRmSzDesc added
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'TTc, Tyy, TChr, Tno, TDt, TFrRmLoc, TToRmLoc, TLsLoc, TLsFrDt, TLsToDt, TGldinLs,
  'TGldRec, TDustWt, TSuppCd, TBillNo,  TBillDt, TBOENo, TBOEDt, TCnvRt, TNetAmt as qTNetAmt,
  'TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TAddAmt3 as qTAddAmt3, TAddAmt4 as qTAddAmt4,
  'TTotAmt as qTTotAmt, TdSr, TdFrRmLoc, TdFrRmDc, RmCtg, RmSCtg, RmPurityWt, TdRmCd , TdLotNo, TdRmSz, RisRmSzDesc as qRmSzDesc, TdRmStkRt,
  'PValue3 as qPtrYN, TdRmQty, TdRmWt, TdToRmLoc, TdPurRt as qTdPurRt, TdPurAmt as qTdPurAmt, 0 as qRmStkAmt, TdDustWt
  'From Txn, Txnd, RmMst, Param Where 1 = 2
  '*** Report Sql ***
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wi_Opt As Integer
  Dim wPurStr As String, wTTotAmtStr As String
  
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
  Dim ws_BagJoin As String
  ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                    IIF(adc("UoBagTyp") <> "", " and PValue= '" + adc("UoBagTyp") + "' ", " ")
'  If UCase(adc("UoBagTyp")) = "N" Then
'    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'N' "
'  ElseIf UCase(adc("UoBagTyp")) = "Y" Then
'    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'Y' "
'  Else
'    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' "
'  End If
' **** Zubin 211 **** '
'*** Jay 2.13(CT) ***
  Set Rep1 = Nothing
  Set Rep2 = Nothing
  Set Rep3 = Nothing
  
  If UCase(adc.MenuCd) = UCase("RepTxnDT") Then gs_TxnTcTyp = "DT"
  If UCase(adc.MenuCd) = UCase("RepTxnBS") Then gs_TxnTcTyp = "BS"
  If UCase(adc.MenuCd) = UCase("RepTxnMLT") Then gs_TxnTcTyp = "MLT"
  If UCase(adc.MenuCd) = UCase("RepTxnBV") Then gs_TxnTcTyp = "BV"
  If UCase(adc.MenuCd) = UCase("RepTxnREJ") Then gs_TxnTcTyp = "REJ"
  If UCase(adc.MenuCd) = UCase("RepTxnPR") Then gs_TxnTcTyp = "PR"
  If UCase(adc.MenuCd) = UCase("RepTxnREC") Then gs_TxnTcTyp = "REC"
  If UCase(adc.MenuCd) = UCase("RepTxnCNV") Then gs_TxnTcTyp = "CNV"
  If UCase(adc.MenuCd) = UCase("RepTxnMV") Then gs_TxnTcTyp = "MV"
  ' **** Zubin 211 **** '
  If UCase(adc.MenuCd) = UCase("RepTxnBM") Then gs_TxnTcTyp = "BM"
  ' **** Zubin 211 **** '
  
  Select Case UCase(gs_TxnTcTyp)
  ' **** Zubin 211 **** '
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
  '*** If Detail option is not selected then report is not shown ***
  If wDetPos = 0 And UCase(gs_TxnTcTyp) <> "PR" Then DispMsg "Detail Option Has To Be Selected For This Report", etError
  
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
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " Rm.TdRmCtg >= '" + Trim(adc("UORMCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGTO")) <> "", " Rm.TdRmCtg <= '" + Trim(adc("UORMCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGSEL")) <> "", " Rm.TdRmCtg In (" + Trim(adc("UORMCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " Rm.TdRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " Rm.TdRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " Rm.TdRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " Td.TdDt >= '" + Trim(adc("UODTFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " Td.TdDt <= '" + Trim(adc("UODTTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  ' **** Manali 3.03 - 08/07/08 - RmSzDesc
  wRmSzDesc = "(Case when (Select PDesc225 from Param where PTyp='RMSCTG' and PMCd=IsNull(RisRmCtg, '') and PSCd=RmSCtg) = 'Y' Then IsNull(RisRmSzDesc, '') Else '' End) "
  ' **** Manali 3.03 - 08/07/08 - RmSzDesc
  
  Select Case UCase(gs_TxnTcTyp)
    ' **** Zubin 211 **** '
  Case "DT", "BS", "MLT", "BM"
                            
    ' **** Manali 3.03 - 08/07/08 - RmSzDesc
    wRmIdSzJoin = "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=Rm.TdRmSz "
    ' **** Manali 3.03 - 08/07/08 - RmSzDesc
    
    ' Zubin 212
    '****** Sachin 3.02 - Id fields In Joins
    ' *** Manali 3.03 - 07/08/07 - qRmSzDesc added and Join with RmIdSz added
    wSqlStrg = "Select " + grpflds + ", " + _
               "TTc, Tyy, TChr, TNo, TDt, TFrRmLoc, TToRmLoc, TFrBLoc, Td.TdSr, Td.TdBYy, " + _
               "Td.TdBChr, Td.TdBNo, Td.TdBDc, Td.TdBQty, Td.TdPtQty, Td.TdBGrWt, " + _
               "Td.TdFrBLoc, Td.TdToBLoc, Td.TdRjDc, Td.TdRjQty, Td.TdRjPtQty, " + _
               "Td.TdRjCd, Td.TdWrk, Td.TdAck, Td.TdPrdYN, Td.TdByLoc as qBagByLoc, Td.TdByWrk as qBagByWrk, " + _
               "Td.TdDesc as qBagDesc, IsNull(Rm.TdSrNo, 0) as TdSrNo, IsNull(Rm.TdFrRmLoc, '') as TdFrRmLoc, " + _
               "IsNull(Rm.TdFrRmDc, '') as TdFrRmDc, " + _
               "IsNull(RmCtg, '') as RmCtg, IsNull(Rm.TdRmCd, '') as TdRmCd, IsNull(Rm.TdLotNo, '') as TdLotNo, IsNull(Rm.TdRmSz, 0) as TdRmSz, IsNull(Rm.TdRmSz2, 0) as TdRmSz2, IsNull(Rm.TdRmSz3, 0) as TdRmSz3, " + wRmSzDesc + " as qRmSzDesc, " + _
               "IsNull(Rm.TdRmStkRt, 0) as TdRmStkRt, IsNull((Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RmCtg, '') and PSCd= IsNull(RmSCtg, '')), '') as qPtrYN, " + _
               "IsNull(Rm.TdRmQty, 0) as TdRmQty, IsNull(Rm.TdRmWt, 0) as TdRmWt, " + _
               "IsNull(Rm.TdToRmLoc, '') as TdToRmLoc, IsNull(Rm.TdByLoc, '') as qRmByLoc, " + _
               "IsNull(Rm.TdByWrk, '') as qRmByWrk, " + _
               "IsNull(Rm.TdDustWt, 0) as TdDustWt, IsNull(Rm.TdDesc, '') as qRmDesc, (Select LocTyp from Loc where " + _
               "LocCoCd = IsNull(Rm.TdCoCd, '') and LocCd= IsNull(Rm.TdToRmLoc, '')) as qBMLocTyp " + _
               "From Txn Join Txnd Td on Td.TdTIdNo=TIdNo and Td.TdCoCd= TCoCd and Td.TdTc= TTc and Td.TdYy= TYy " + _
               "and Td.TdChr= TChr and Td.TdNo= TNo and Td.TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " and Td.TdPrtKey=TPrtKey ", "") + _
               "Left Outer Join Txnd Rm on Rm.TdCoCd= Td.TdCoCd and Rm.TdTc= Td.TdTc and Rm.TdYy= Td.TdYy " + _
               "and Rm.TdChr= Td.TdChr and Rm.TdNo= Td.TdNo and Rm.TdSr= Td.TdSr and Rm.TdSrNo<> 0 " + IIF(gs_Partition = ctCurrPrtn, " and Rm.TdPrtKey=Td.TdPrtKey ", "") + _
               "Left Outer Join RmMst on RmCd= Rm.TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=Rm.TdPrtKey ", "") + _
               "Join Bag on BIdNo=Td.TdBIdNo and BCoCd= Td.TdCoCd and BYy= Td.TdBYy and BChr= Td.TdBChr and BNo= Td.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=Td.TdPrtKey ", "") + ws_BagJoin + _
               "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
               "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
               "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
               wRmIdSzJoin + "where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' and ", "") + " (Select PValue from Param where PTyp= 'TC' and PMCd= TTc)= '" + gs_TxnTcTyp + "' " + _
               wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "")
    ' **** Zubin 211 **** '
  Case "BV", "REJ"
    ' **** Zubin 211 (ws_BagJoin added) **** '
    
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
    wAddlCnd = ""
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " Td.TdDt >= '" + Trim(adc("UODTFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " Td.TdDt <= '" + Trim(adc("UODTTO")) + "'", "")
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
    'IG.147 -Design Cd, Order Tc,yy,chr,No.,Sr are added
    '****** Sachin 3.02 - Id fields In Join
    wSqlStrg = "Select " + grpflds + ", " + _
               "TTc, Tyy, TChr, TNo, TDt, TdSr, TdBYy, TdBChr, TdBNo, TdBDc, TdBQty, " + _
               "TdPtQty, TdBGrWt , TdFrBLoc, TdToBLoc, TdRjCd, TdWrk, TdAck, TdByLoc as qBagByLoc, " + _
               "TdByWrk as qBagByWrk, BOdDmCd, BCoCd, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr   " + _
               "From Txn Join Txnd Td on Td.TdCoCd= TCoCd and Td.TdTc= TTc and Td.TdYy= TYy " + _
               "and Td.TdChr= TChr and Td.TdNo= TNo and Td.TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " and Td.TdPrtKey=TPrtKey ", "") + _
               "Join Bag on BCoCd= Td.TdCoCd and BYy= Td.TdBYy and BChr= Td.TdBChr and BNo= Td.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=Td.TdPrtKey ", "") + _
               ws_BagJoin + _
               "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and " + _
               "OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
               "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
               "where " + IIF(gs_Partition = ctCurrPrtn, "  TPrtKey='" + ctCurrPrtn + "' and ", "") + " (Select PValue from Param where PTyp= 'TC' and PMCd= TTc)= '" + gs_TxnTcTyp + "' " + _
               wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "")
    ' **** Zubin 211 (ws_BagJoin added) **** '
  Case "PR", "REC", "CNV", "MV"
    '************Geeta****************Emr207*******************
    ' ****** Sachin 2.13.0 Correction ******
    Dim ws_BaseCurCd As String, ws_Join As String, wAck As String
        ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                      "HCoCd= '" + adc("UoCoCdFr") + "'")
  
        If adc("UoCurCdFr") <> "" Then
          wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "TCurrCd='" + adc("UoCurCdFr") + "'"
          wPurStr = "TdPurRt as qTdPurRt, TdPurAmt qTdPurAmt "
          wTTotAmtStr = "TNetAmt as qTNetAmt, TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TAddAmt3 as qTAddAmt3, " + _
                         " TAddAmt4 as qTAddAmt4, TTotAmt as qTTotAmt,"
        Else
          wPurStr = " (case when TMulDiv= 'M' then TdPurRt/ (case when TCnvRt<> 0 then TCnvRt else 1 end) else TdPurRt* TCnvRt end) as qTdPurRt, (case when TMulDiv= 'M' then TdPurAmt/TCnvRt else TdPurAmt*TCnvRt end) qTdPurAmt "

          wTTotAmtStr = "(case when TMulDiv= 'M' then TNetAmt/ (case when TCnvRt<> 0 then TCnvRt else 1 end) Else TNetAmt* TCnvRt End) as qTNetAmt, " + _
                        " (case when TMulDiv= 'M' then TAddAmt1/ (case when TCnvRt<> 0 then TCnvRt else 1 end) Else TAddAmt1* TCnvRt End) as qTAddAmt1, " + _
                        " (case when TMulDiv= 'M' then TAddAmt2/ (case when TCnvRt<> 0 then TCnvRt else 1 end) Else TAddAmt2* TCnvRt End) as qTAddAmt2, " + _
                        " (case when TMulDiv= 'M' then TAddAmt3/ (case when TCnvRt<> 0 then TCnvRt else 1 end) Else TAddAmt3* TCnvRt End) as qTAddAmt3, " + _
                        " (case when TMulDiv= 'M' then TAddAmt4/ (case when TCnvRt<> 0 then TCnvRt else 1 end) Else TAddAmt4* TCnvRt End) as qTAddAmt4, " + _
                        " (case when TMulDiv= 'M' then TTotAmt/ (case when TCnvRt<> 0 then TCnvRt else 1 end) Else TTotAmt* TCnvRt End) as qTTotAmt, "
        End If
        
        ws_Join = " Join (Select PMCd as qPMCd, (case when " + IIF(adc("UoCurCdFr") = "", "1", "0") + " = 1 then 'Y' " + _
                  "else 'N' end) as qBaseCur from Param where PTyp= 'CURNCY') Cur on qPMCd= TCurrCd "
    ' ****** Sachin 2.13.0 Correction ******
    'MMAck-Voucher Ack Y/N added, blank means all voucher
    wAck = IIF(adc("UoYN2") = "Y", " and TdAck= 'Y' ", IIF(adc("UoYN2") = "N", " and TdAck= 'N' ", ""))
    
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
    wAddlCnd = ""
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " Td.TdRmCtg >= '" + Trim(adc("UORMCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGTO")) <> "", " Td.TdRmCtg <= '" + Trim(adc("UORMCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGSEL")) <> "", " Td.TdRmCtg In (" + Trim(adc("UORMCTGSEL")) + ")", "")
    
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " Td.TdRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " Td.TdRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " Td.TdRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
    
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " Td.TdDt >= '" + Trim(adc("UODTFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " Td.TdDt <= '" + Trim(adc("UODTTO")) + "'", "")
    
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
            
    ' ****** Sachin Before 2.13.0 ******
    '    If UCase(gs_TxnTcTyp) = "PR" Then
    '      If adc("UoYN") <> "" Then
    '        wCnd = wCnd + IIF(wCnd <> "", " and ", "") '+ "TUSDYn='" + adc("UoYN") + "'"
    '        wPurStr = "TdPurRt as qTdPurRt, TdPurAmt as qTdPurAmt "
    '        wTTotAmtStr = "TNetAmt as qTNetAmt, TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TAddAmt3 as qTAddAmt3, " + _
    '                      " TAddAmt4 as qTAddAmt4, TTotAmt as qTTotAmt,"
    '      Else
    '
    '        wPurStr = "TdPurRt as qTdPurRt, TdPurAmt as qTdPurAmt "
    '        wTTotAmtStr = "TNetAmt as qTNetAmt, TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TAddAmt3 as qTAddAmt3, " + _
    '                      " TAddAmt4 as qTAddAmt4, TTotAmt as qTTotAmt,"
    ' ****** Sachin Before 2.13.0 ******
      
        'wPurStr = " (case when TUSDYn='N' then TdPurRt/TCnvRt else TdPurRt end ) as qTdPurRt, " + _
                  " (case when TUSDYn='N' then TdPurAmt/TCnvRt else TdPurAmt end ) as qTdPurAmt "
        'wTTotAmtStr = "(case when TUSDYn='N' then TNetAmt/TCnvRt else TNetAmt end ) as qTNetAmt," + _
                      "(case when TUSDYn='N' then TAddAmt1/TCnvRt else TAddAmt1 end ) as qTAddAmt1," + _
                      "(case when TUSDYn='N' then TAddAmt2/TCnvRt else TAddAmt2 end ) as qTAddAmt2, " + _
                      "(case when TUSDYn='N' then TAddAmt3/TCnvRt else TAddAmt3 end ) as qTAddAmt3, " + _
                      "(case when TUSDYn='N' then TAddAmt4/TCnvRt else TAddAmt4 end ) as qTAddAmt4," + _
                      "(case when TUSDYn='N' then TTotAmt/TCnvRt else TTotAmt end ) as qTTotAmt,"
    ' ****** Sachin Before 2.13.0 ******
    '      End If
    '    Else
    '       wPurStr = "TdPurRt as qTdPurRt, TdPurAmt qTdPurAmt "
    '       wTTotAmtStr = "TNetAmt as qTNetAmt, TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TAddAmt3 as qTAddAmt3, " + _
    '                     " TAddAmt4 as qTAddAmt4, TTotAmt as qTTotAmt,"
    '    End If
    '************Geeta****************Emr207*******************
    '*********************Replaced  TdPurRt, TdPurAmt with  wPurStr*********
    '*********************Replaced  TNetAmt, TAddAmt1, TAddAmt2, TAddAmt3, TAddAmt4, TTotAmt, with  wTTotAmtStr*********
    ' Zubin 212
    '****** Sachin 3.02 - Id fields In Joins
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    ' Manali 3.03 27/06/08 - Stk amt included in query for Purchase
    ' **** Manali 3.03 - 08/07/08 - RmSzDesc included below and join with RmIdSz
    ' ***** Manali 3.8.0 - TdRtByQw added
    '578.43 pure wt added to X
    wSqlStrg = "Select " + grpflds + ", " + _
               "TTc, TYy, TChr, TNo, TDt, TFrRmLoc, TToRmLoc, TLsLoc, TLsFrDt, " + _
               "TLsToDt, TGldinLs, TGldRec, TDustWt, TSuppCd, TBillNo,  TBillDt, TBOENo, " + _
               "TBOEDt, TCnvRt," + wTTotAmtStr + _
               "TdSr, TdFrRmLoc, TdFrRmDc, RmCtg, RmSCtg, (case when RmCtg= 'X' then RmBasePurityWt else RmPurityWt end)/ " + _
               "(case when RmCtg= 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
               "      when RmCtg= 'P' then " + CStr(wd_BasePltRmPurity) + " " + _
               "      when RmCtg= 'S' then " + CStr(wd_BaseSlvRmPurity) + " " + _
               "      when RmCtg= 'L' then " + CStr(wd_BasePdRmPurity) + " " + _
               "      when RmCtg= 'X' then IsNull((Select RmPurityWt from Param, RmMst Base where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "      RmCd= PValue and PTyp= 'RMCTG' and PMCd=IsNull((Select RmCtg From RmMst Rm where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "      RmPrtKey='C' and RmCd= RmMst.RmBaseCd), '') ), 1) " + _
               " else 1 end) as RmPurityWt, " + _
               "TdRmCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, " + wRmSzDesc + " as qRmSzDesc, TdRmStkRt, " + _
               "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
               "TdRmQty, TdRmWt, TdToRmLoc, TdRtByQw, " + wPurStr + ", " + _
               "cast((TdRmWt * TdRmStkRt) as decimal (16,2)) as qRmStkAmt, TdDustWt, TdDesc,TdAck From Txn " + _
               "Join Txnd Td on Td.TdTIdNo=TIdNo and Td.TdCoCd= TCoCd and Td.TdTc= TTc and Td.TdYy= TYy " + _
               "and Td.TdChr= TChr and Td.TdNo= TNo and Td.TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " and Td.TdPrtKey=TPrtKey ", "") + _
               "Join RmMst on RmCd= Td.TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=Td.TdPrtKey ", "") + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
               "where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' and ", "") + " (Select PValue from Param where PTyp= 'TC' and PMCd= TTc)= '" + gs_TxnTcTyp + "' " + _
               wCnd + wAck + IIF(wAddlCnd <> "", " And " + wAddlCnd, "")
  End Select
  
             
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  '*********Geeta************Emr207***
  If UCase(gs_TxnTcTyp) = "PR" Then
    '****** Sachin Before 2.13.0 ******
    'If adc("UoYN") = "" Or adc("UoYN") = "Y" Then
    '  mRep.TxtPRCurCd.SetText "In US$"
    'ElseIf adc("UoYN") = "N" Then
    '  mRep.TxtPRCurCd.SetText "In RS"
    'End If
    '****** Sachin Before 2.13.0 ******
    If adc("UoCurCdFr") = "" Then
      mRep.TxtPRCurCd.SetText "Purchase Currency (" + ws_BaseCurCd + ")"
    Else
      mRep.TxtPRCurCd.SetText "Purchase Currency (" + adc("UoCurCdFr") + ")"
    End If
    
  ElseIf UCase(gs_TxnTcTyp) = "REC" Or UCase(gs_TxnTcTyp) = "CNV" Then
    mRep.TxtPRCurCd.SetText ""
  End If
  '*********Geeta************Emr207***
 
  mRep.TxtHead.SetText ws_FrmCaption  '****** Sachin 2.14.0
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
'  If gs_TxnTcTyp = "MV" Then wi_Opt = 4
  If gs_TxnTcTyp = "MV" Then wi_Opt = 1
  
  Select Case wi_Opt
  Case Is = 1
    mRep.Hd2.Suppress = True
    mRep.Hd3.Suppress = True
    mRep.Hd4.Suppress = True
    mRep.DET2.Suppress = True
    mRep.DET3.Suppress = True
    mRep.DET4.Suppress = True
    mRep.G1F2.Suppress = True
    mRep.G1F3.Suppress = True
    mRep.G1F4.Suppress = True
    mRep.G2F2.Suppress = True
    mRep.G2F3.Suppress = True
    mRep.G2F4.Suppress = True
    mRep.RF2.Suppress = True
    mRep.RF3.Suppress = True
    mRep.RF4.Suppress = True
    If wDetPos = 1 Or wNonePos = 2 Then mRep.G1F1.Suppress = True: mRep.G2F1.Suppress = True
    If wDetPos = 2 Or wNonePos = 3 Then mRep.G1F1.Suppress = True
    'sri
    If wDetPos = 0 And (gs_TxnTcTyp = "PR" Or gs_TxnTcTyp = "MV") Then
        With mRep
        '.hVchNo.Suppress = True:
        .DET1.Suppress = True:  .h1TdSr.Suppress = True: .h1TDt.Suppress = True
        .h1TSuppCd.Suppress = True: .h1TBoeNo.Suppress = True: .h1TBoeDt.Suppress = True
        .h1TCnvRt.Suppress = True: .h1TdRmFrLoc.Suppress = True: .h1TdRmCd.Suppress = True: .h1TdRmDC.Suppress = True: .h1LotNo.Suppress = True
        .h1TdRmSz.Suppress = True: .h1TdRmStkRt.Suppress = True: .h1TdRmQW.Suppress = True: .h1TdPurRt.Suppress = True: .h1TdRmToLoc.Suppress = True
        Dim wGrpStr As String
        Dim wArr() As String
       wGrpStr = GRP_REP.GrpOpts
       wArr() = Split(wGrpStr, ",")
        .h1VchNo.SetText wArr(2)
        End With
    End If
    If (gs_TxnTcTyp <> "PR" And gs_TxnTcTyp <> "MV") Then mRep.G3F.Suppress = True
  Case Is = 2
    mRep.Hd1.Suppress = True
    mRep.Hd3.Suppress = True
    mRep.Hd4.Suppress = True
    mRep.DET1.Suppress = True
    mRep.DET3.Suppress = True
    mRep.DET4.Suppress = True
    mRep.G1F1.Suppress = True
    mRep.G1F3.Suppress = True
    mRep.G1F4.Suppress = True
    mRep.G2F1.Suppress = True
    mRep.G2F3.Suppress = True
    mRep.G2F4.Suppress = True
    mRep.RF1.Suppress = True
    mRep.RF3.Suppress = True
    mRep.RF4.Suppress = True
    If wDetPos = 1 Or wNonePos = 2 Then mRep.G1F2.Suppress = True: mRep.G2F2.Suppress = True
    If wDetPos = 2 Or wNonePos = 3 Then mRep.G1F2.Suppress = True
  Case Is = 3
    mRep.Hd1.Suppress = True
    mRep.Hd2.Suppress = True
    mRep.Hd4.Suppress = True
    mRep.DET1.Suppress = True
    mRep.DET2.Suppress = True
    mRep.DET4.Suppress = True
    mRep.G1F1.Suppress = True
    mRep.G1F2.Suppress = True
    mRep.G1F4.Suppress = True
    mRep.G2F1.Suppress = True
    mRep.G2F2.Suppress = True
    mRep.G2F4.Suppress = True
    mRep.G3F.Suppress = True
    mRep.RF1.Suppress = True
    mRep.RF2.Suppress = True
    mRep.RF4.Suppress = True
    If wDetPos = 1 Or wNonePos = 2 Then mRep.G1F3.Suppress = True: mRep.G2F3.Suppress = True
    If wDetPos = 2 Or wNonePos = 3 Then mRep.G1F3.Suppress = True
  Case Is = 4
    mRep.Hd1.Suppress = True
    mRep.Hd2.Suppress = True
    mRep.Hd3.Suppress = True
    mRep.DET1.Suppress = True
    mRep.DET2.Suppress = True
    mRep.DET3.Suppress = True
    mRep.G1F1.Suppress = True
    mRep.G1F2.Suppress = True
    mRep.G1F3.Suppress = True
    mRep.G2F1.Suppress = True
    mRep.G2F2.Suppress = True
    mRep.G2F3.Suppress = True
    mRep.G3F.Suppress = True
    mRep.RF1.Suppress = True
    mRep.RF2.Suppress = True
    mRep.RF3.Suppress = True
    If wDetPos = 1 Or wNonePos = 2 Then mRep.G1F4.Suppress = True: mRep.G2F4.Suppress = True
    If wDetPos = 2 Or wNonePos = 3 Then mRep.G1F4.Suppress = True
  End Select
  
  ' Zubin 212
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UOCOCDFR") + "' and HLotNoYN= 'Y' ") Then
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
  ' Zubin 212
  
  ' ***** Manali 3.03 - 08/07/08 - RmSzDesc
'  GRP_REP.SetFormula mRep, "wRmSzPtr", "If {rdo.qRmSzDesc} <> '' Then {rdo.qRmSzDesc} " + _
                                      "Else (If {rdo.qPtrYn} = 'Y' Then ToText({rdo.TdRmSz}, 4) Else ToText({rdo.TdRmSz}, 3))"
  GRP_REP.SetFormula mRep, "wRmSz", "If {rdo.qRmSzDesc} <> '' Then {rdo.qRmSzDesc} " + _
                                      "Else CustRmSzDisp ({rdo.TdRmSz},{rdo.TdRmSz2},{rdo.TdRmSz3}) "
  
  ' ***** Manali 3.03 - 08/07/08 - RmSzDesc
  
  ' ***** Manali 3.3.0
  If adc("UoYN1") = "Y" Then
    GRP_REP.SetFormula Rep3, "wAmt", "{rdo.qRmStkAmt}"
    Rep3.HPurAmt.SetText "Stk Amt"
  Else
    GRP_REP.SetFormula Rep3, "wAmt", "{rdo.qTdPurAmt}"
  End If
 
  
  'KC.3- StkRt-when Show Stock Rate=YES then Stock Rate is visible in report
  'if it is NO then Stock Rate is invisible in report
 
  If adc("UoYN3") = "Y" Then
    Rep1.hTdRmStkRt.SetText "Stock Rt":   Rep1.hTdRmStkRt2.SetText "Stock Rt":  Rep1.hTdRmStkRt3.SetText "Stock Rt"
    Rep1.FldRmStkRt1.Suppress = False:    Rep1.FldRmStkRt2.Suppress = False:    Rep1.FldRmStkRt3.Suppress = False
    Rep3.h1TdRmStkRt.SetText "Stock Rt":  Rep3.h2TdRmStkRt.SetText "Stock Rt"
    Rep3.h3TdRmStkRt.SetText "Stock Rt":  Rep3.h4TdRmStkRt.SetText "Stock Rt"
    Rep3.FldRmStkRt1.Suppress = False:    Rep3.FldRmStkRt2.Suppress = False
    Rep3.FldRmStkRt3.Suppress = False:    Rep3.FldRmStkRt4.Suppress = False
    
  Else
    Rep1.hTdRmStkRt.SetText "":         Rep1.hTdRmStkRt2.SetText "":      Rep1.hTdRmStkRt3.SetText ""
    Rep1.FldRmStkRt1.Suppress = True:   Rep1.FldRmStkRt2.Suppress = True: Rep1.FldRmStkRt3.Suppress = True
    Rep3.h1TdRmStkRt.SetText "":        Rep3.h2TdRmStkRt.SetText ""
    Rep3.h3TdRmStkRt.SetText "":        Rep3.h4TdRmStkRt.SetText ""
    Rep3.FldRmStkRt1.Suppress = True:   Rep3.FldRmStkRt2.Suppress = True
    Rep3.FldRmStkRt3.Suppress = True:   Rep3.FldRmStkRt4.Suppress = True
    If adc("UoYN1") = "Y" Then
      Rep3.HPurAmt.SetText "":        Rep3.wAmt1.Suppress = True
      Rep3.g3PurAmt1.Suppress = True: Rep3.g2PurAmt1.Suppress = True
      Rep3.g1PurAmt1.Suppress = True: Rep3.gPurAmt1.Suppress = True
    End If
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula mRep, "wRmCdFontSz", GetRmcdFontSz("{rdo.TdRmCd}", IIF(gs_TxnTcTyp = "PR" Or gs_TxnTcTyp = "MV", 6, 7))
      
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
  Set adc.FirNKeyCtl = IIF(ws_MultiCoMod = True, adc("UoCoCdFr"), adc("UoTTcFr"))
      adc("UOCOCDFR").Enabled = ws_MultiCoMod
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******

  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_LocTyp = " "
  gb_RmCtgFor = False
  gb_RmZ = False
  gs_DmTcTyp = "DM"
  gs_OmTcTyp = "SO"
  gs_CmCtg = "S"
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  
  '************Geeta****************Emr207*******************
    'LblUSDYN.Visible = False: adc("UoYN").Visible = False
  '************Geeta****************Emr207*******************
  Select Case UCase(adc.MenuCd)
' **** Zubin 211 **** '
'''''  Case Is = UCase("RepTxnDT"), UCase("RepTxnBS"), UCase("RepTxnMLT")
  Case Is = UCase("RepTxnDT"), UCase("RepTxnBS"), UCase("RepTxnMLT"), UCase("RepTxnBM")
' **** Zubin 211 **** '
    '*** General Properties ***
    If UCase(adc.MenuCd) = UCase("RepTxnDT") Then gs_TxnTcTyp = "DT"
    If UCase(adc.MenuCd) = UCase("RepTxnBS") Then gs_TxnTcTyp = "BS"
    If UCase(adc.MenuCd) = UCase("RepTxnMLT") Then gs_TxnTcTyp = "MLT"
    ' **** Zubin 211 **** '
    If UCase(adc.MenuCd) = UCase("RepTxnBM") Then gs_TxnTcTyp = "BM"
    ' **** Zubin 211 **** '
    
    LblRejCd.Visible = False: LblSuppCd.Visible = False: LblBOENo.Visible = False
    'IssRecLoc-LblLsLoc visible property chnaged from False to True
    LblLsLoc.Visible = True: LblLsDt.Visible = False
    LblStkVal.Visible = False    ' ***** Manali 3.3.0
    
    adc("UoRejCdFr").Visible = False: adc("UoRejCdTo").Visible = False: adc("UoRejCdSel").Visible = False
    adc("UoCmCdFr").Visible = False: adc("UoCmCdTo").Visible = False: adc("UoCmCdSel").Visible = False
    adc("UoBOEFr").Visible = False: adc("UoBOETo").Visible = False: adc("UoBOESel").Visible = False
    'IssRecLoc-IssLoc visible property chnaged from False to True
    'adc("UoIssLocFr").Visible = False: adc("UoIssLocTo").Visible = False: adc("UoIssLocSel").Visible = False
    adc("UoLsDtFr").Visible = False: adc("UoLsDtTo").Visible = False
    adc("UoYN1").Visible = False      ' ***** Manali 3.3.0
    FraOptAmt.Visible = False: FraOptAmt.Enabled = False ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
    FraOptAck.Visible = False: FraOptAck.Enabled = False  'MMAck-Disable visible property of Ack option buttons
    
    adc("UoRejCdFr").Enabled = False: adc("UoRejCdTo").Enabled = False: adc("UoRejCdSel").Enabled = False
    adc("UoCmCdFr").Enabled = False: adc("UoCmCdTo").Enabled = False: adc("UoCmCdSel").Enabled = False
    adc("UoBOEFr").Enabled = False: adc("UoBOETo").Enabled = False: adc("UoBOESel").Enabled = False
    'IssRecLoc-IssLoc ENABLE property chnaged from False to True
    'adc("UoIssLocFr").Enabled = False: adc("UoIssLocTo").Enabled = False: adc("UoIssLocSel").Enabled = False
    adc("UoLsDtFr").Enabled = False: adc("UoLsDtTo").Enabled = False
    LblCurrCd.Visible = False: adc("UOCURCDFR").Visible = False
    adc("UoYN1").Enabled = False      ' ***** Manali 3.3.0
    '*** General Properties ***
    ' Zubin 212
    adc("UoLotNoFr").CmpStr = "Rm.TdLotNo>= ": adc("UoLotNoTo").CmpStr = "Rm.TdLotNo<= ": adc("UoLotNoSel").CmpStr = "Rm.TdLotNo In "
    ' Zubin 212
    'IssRecLoc- .CmpStr applied to Issue & Receiving Loc
    adc("UORECLOCFR").CmpStr = " (CASE WHEN Rm.TdFrRmDc='D' THEN Rm.TdFrRmLoc ELSE Rm.TdToRmLoc END) >= "
    adc("UORECLOCTO").CmpStr = " (CASE WHEN Rm.TdFrRmDc='D' THEN Rm.TdFrRmLoc ELSE Rm.TdToRmLoc END) <= "
    adc("UORECLOCSEL").CmpStr = " (CASE WHEN Rm.TdFrRmDc='D' THEN Rm.TdFrRmLoc ELSE Rm.TdToRmLoc END) In "
    adc("UOISSLOCFR").CmpStr = " (CASE WHEN Rm.TdFrRmDc='C' THEN Rm.TdFrRmLoc ELSE Rm.TdToRmLoc END) >= "
    adc("UOISSLOCTO").CmpStr = " (CASE WHEN Rm.TdFrRmDc='C' THEN Rm.TdFrRmLoc ELSE Rm.TdToRmLoc END) <= "
    adc("UOISSLOCSEL").CmpStr = " (CASE WHEN Rm.TdFrRmDc='C' THEN Rm.TdFrRmLoc ELSE Rm.TdToRmLoc END) In "
    
  Case Is = UCase("RepTxnBV"), UCase("RepTxnREJ")
    '*** General Properties ***
    LblRmCtg.Visible = False: LblRmSCtg.Visible = False: LblRmCd.Visible = False
    LblRejCd.Visible = False: LblSuppCd.Visible = False: LblBOENo.Visible = False
    'IssRecLoc-LblLsLoc visible property chnaged from False to True
    LblLsLoc.Visible = True: LblLsDt.Visible = False
    LblStkVal.Visible = False    ' ***** Manali 3.3.0
    adc("UoRmCtgFr").Visible = False: adc("UoRmCtgTo").Visible = False: adc("UoRmCtgSel").Visible = False
    adc("UoRmSCtgFr").Visible = False: adc("UoRmSCtgTo").Visible = False: adc("UoRmSCtgSel").Visible = False
    adc("UoRmCdFr").Visible = False: adc("UoRmCdTo").Visible = False: adc("UoRmCdSel").Visible = False
    adc("UoRejCdFr").Visible = False: adc("UoRejCdTo").Visible = False: adc("UoRejCdSel").Visible = False
    adc("UoCmCdFr").Visible = False: adc("UoCmCdTo").Visible = False: adc("UoCmCdSel").Visible = False
    adc("UoBOEFr").Visible = False: adc("UoBOETo").Visible = False: adc("UoBOESel").Visible = False
    'IssRecLoc-IssLoc visible property chnaged from False to True
    'adc("UoIssLocFr").Visible = False: adc("UoIssLocTo").Visible = False: adc("UoIssLocSel").Visible = False
    adc("UoLsDtFr").Visible = False: adc("UoLsDtTo").Visible = False
    adc("UoYN1").Visible = False      ' ***** Manali 3.3.0
    FraOptAmt.Visible = False: FraOptAmt.Enabled = False ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
    FraOptAck.Visible = False: FraOptAck.Enabled = False  'MMAck-Disable visible property of Ack option buttons
    
    adc("UoRmCtgFr").Enabled = False: adc("UoRmCtgTo").Enabled = False: adc("UoRmCtgSel").Enabled = False
    adc("UoRmSCtgFr").Enabled = False: adc("UoRmSCtgTo").Enabled = False: adc("UoRmSCtgSel").Enabled = False
    adc("UoRmCdFr").Enabled = False: adc("UoRmCdTo").Enabled = False: adc("UoRmCdSel").Enabled = False
    adc("UoRejCdFr").Enabled = False: adc("UoRejCdTo").Enabled = False: adc("UoRejCdSel").Enabled = False
    adc("UoCmCdFr").Enabled = False: adc("UoCmCdTo").Enabled = False: adc("UoCmCdSel").Enabled = False
    adc("UoBOEFr").Enabled = False: adc("UoBOETo").Enabled = False: adc("UoBOESel").Enabled = False
    'IssRecLoc-IssLoc ENABLE property chnaged from False to True
    'adc("UoIssLocFr").Enabled = False: adc("UoIssLocTo").Enabled = False: adc("UoIssLocSel").Enabled = False
    adc("UoLsDtFr").Enabled = False: adc("UoLsDtTo").Enabled = False
    adc("UoYN1").Enabled = False      ' ***** Manali 3.3.0
    
    adc("UoRmCtgFr").CmpStr = "": adc("UoRmCtgTo").CmpStr = "": adc("UoRmCtgSel").CmpStr = ""
    adc("UoRmSCtgFr").CmpStr = "": adc("UoRmSCtgTo").CmpStr = "": adc("UoRmSCtgSel").CmpStr = ""
    adc("UoRmCdFr").CmpStr = "": adc("UoRmCdTo").CmpStr = "": adc("UoRmCdSel").CmpStr = ""
    LblCurrCd.Visible = False: adc("UOCURCDFR").Visible = False
    
    ' Zubin 212
    LblLotNo.Visible = False
    adc("UoLotNoFr").Visible = False: adc("UoLotNoTo").Visible = False: adc("UoLotNoSel").Visible = False
    adc("UoLotNoFr").CmpStr = "": adc("UoLotNoTo").CmpStr = "": adc("UoLotNoSel").CmpStr = ""
    adc("UoLotNoFr") = "": adc("UoLotNoTo") = "": adc("UoLotNoSel") = ""
    ' Zubin 212
    '*** General Properties ***
  
    If UCase(adc.MenuCd) = UCase("RepTxnBV") Then gs_TxnTcTyp = "BV"
    '*** Only For Rejection ***
    If UCase(adc.MenuCd) = UCase("RepTxnREJ") Then
      gs_TxnTcTyp = "REJ"
      adc("UoRejCdFr").CmpStr = "Td.TdRjCd>= ": adc("UoRejCdTo").CmpStr = "Td.TdRjCd<= ": adc("UoRejCdSel").CmpStr = "Td.TdRjCd In "
      LblRejCd.Visible = True
      adc("UoRejCdFr").Visible = True: adc("UoRejCdTo").Visible = True: adc("UoRejCdSel").Visible = True
      adc("UoRejCdFr").Enabled = True: adc("UoRejCdTo").Enabled = True: adc("UoRejCdSel").Enabled = True
    End If
    '*** Only For Rejection ***
    'IssRecLoc- .CmpStr applied to Issue & Receiving Loc
    adc("UORECLOCFR").CmpStr = " TdToBLoc >= "
    adc("UORECLOCTO").CmpStr = " TdToBLoc <= "
    adc("UORECLOCSEL").CmpStr = " TdToBLoc In "
    adc("UOISSLOCFR").CmpStr = " TdFrBLoc >= "
    adc("UOISSLOCTO").CmpStr = " TdFrBLoc <= "
    adc("UOISSLOCSEL").CmpStr = " TdFrBLoc In "
        
  Case Is = UCase("RepTxnPR"), UCase("RepTxnREC"), UCase("RepTxnCNV"), UCase("RepTxnMV")
    '*** General Properties ***
    LblByWrk.Visible = False: LblBagYyChr.Visible = False: LblBagNo.Visible = False
    LblDmCtg.Visible = False: LblPrdCtg.Visible = False: LblDmCd.Visible = False
    LblOrdTcYyChr.Visible = False: LblOrdNo.Visible = False
    LblBagSlh1.Visible = False: LblBagSlh2.Visible = False: LblBagSlh3.Visible = False: LblBagSlh4.Visible = False
    LblOrdSlh1.Visible = False: LblOrdSlh2.Visible = False: LblOrdSlh3.Visible = False
    LblOrdSlh4.Visible = False: LblOrdSlh5.Visible = False: LblOrdSlh6.Visible = False
    LblRejCd.Visible = False: LblSuppCd.Visible = False: LblBOENo.Visible = False
    'IssRecLoc-LblLsLoc visible property chnaged from False to True
    LblLsLoc.Visible = True: LblLsDt.Visible = False
    LblStkVal.Visible = False       ' ***** Manali 3.3.0
    adc("UoWrkFr").Visible = False: adc("UoWrkTo").Visible = False: adc("UoWrkSel").Visible = False
    adc("UoBYyFr").Visible = False: adc("UoBYyTo").Visible = False
    adc("UoBChrFr").Visible = False: adc("UoBChrTo").Visible = False
    adc("UoBNoFr").Visible = False: adc("UoBNoTo").Visible = False
    adc("UoDmCtgFr").Visible = False: adc("UoDmCtgTo").Visible = False: adc("UoDmCtgSel").Visible = False
    adc("UoDmPrdCtgFr").Visible = False: adc("UoDmPrdCtgTo").Visible = False: adc("UoDmPrdCtgSel").Visible = False
    adc("UoDmCdFr").Visible = False: adc("UoDmCdTo").Visible = False: adc("UoDmCdSel").Visible = False
    adc("UoOmTcFr").Visible = False: adc("UoOmTcTo").Visible = False
    adc("UoOmYyFr").Visible = False: adc("UoOmYyTo").Visible = False
    adc("UoOmChrFr").Visible = False: adc("UoOmChrTo").Visible = False
    adc("UoOmNoFr").Visible = False: adc("UoOmNoTo").Visible = False
    adc("UoOdSrFr").Visible = False: adc("UoOdSrTo").Visible = False
    adc("UoRejCdFr").Visible = False: adc("UoRejCdTo").Visible = False: adc("UoRejCdSel").Visible = False
    adc("UoCmCdFr").Visible = False: adc("UoCmCdTo").Visible = False: adc("UoCmCdSel").Visible = False
    adc("UoBOEFr").Visible = False: adc("UoBOETo").Visible = False: adc("UoBOESel").Visible = False
    'IssRecLoc-IssLoc visible property chnaged from False to True
    'adc("UoIssLocFr").Visible = False: adc("UoIssLocTo").Visible = False: adc("UoIssLocSel").Visible = False
    adc("UoLsDtFr").Visible = False: adc("UoLsDtTo").Visible = False
     adc("UoYN1").Visible = False      ' ***** Manali 3.3.0
  
    adc("UoWrkFr").Enabled = False: adc("UoWrkTo").Enabled = False: adc("UoWrkSel").Enabled = False
    adc("UoBYyFr").Enabled = False: adc("UoBYyTo").Enabled = False
    adc("UoBChrFr").Enabled = False: adc("UoBChrTo").Enabled = False
    adc("UoBNoFr").Enabled = False: adc("UoBNoTo").Enabled = False
    adc("UoDmCtgFr").Enabled = False: adc("UoDmCtgTo").Enabled = False: adc("UoDmCtgSel").Enabled = False
    adc("UoDmPrdCtgFr").Enabled = False: adc("UoDmPrdCtgTo").Enabled = False: adc("UoDmPrdCtgSel").Enabled = False
    adc("UoDmCdFr").Enabled = False: adc("UoDmCdTo").Enabled = False: adc("UoDmCdSel").Enabled = False
    adc("UoOmTcFr").Enabled = False: adc("UoOmTcTo").Enabled = False
    adc("UoOmYyFr").Enabled = False: adc("UoOmYyTo").Enabled = False
    adc("UoOmChrFr").Enabled = False: adc("UoOmChrTo").Enabled = False
    adc("UoOmNoFr").Enabled = False: adc("UoOmNoTo").Enabled = False
    adc("UoOdSrFr").Enabled = False: adc("UoOdSrTo").Enabled = False
    adc("UoRejCdFr").Enabled = False: adc("UoRejCdTo").Enabled = False: adc("UoRejCdSel").Enabled = False
    adc("UoCmCdFr").Enabled = False: adc("UoCmCdTo").Enabled = False: adc("UoCmCdSel").Enabled = False
    adc("UoBOEFr").Enabled = False: adc("UoBOETo").Enabled = False: adc("UoBOESel").Enabled = False
    'IssRecLoc-IssLoc ENABLE property chnaged from False to True
    'adc("UoIssLocFr").Enabled = False: adc("UoIssLocTo").Enabled = False: adc("UoIssLocSel").Enabled = False
    adc("UoLsDtFr").Enabled = False: adc("UoLsDtTo").Enabled = False
    adc("UoYN1").Enabled = False      ' ***** Manali 3.3.0
    FraOptAmt.Visible = False: FraOptAmt.Enabled = False ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  
    adc("UoWrkFr").CmpStr = "": adc("UoWrkTo").CmpStr = "": adc("UoWrkSel").CmpStr = ""
    adc("UoBYyFr").CmpStr = "": adc("UoBYyTo").CmpStr = ""
    adc("UoBChrFr").CmpStr = "": adc("UoBChrTo").CmpStr = ""
    adc("UoBNoFr").CmpStr = "": adc("UoBNoTo").CmpStr = ""
    adc("UoDmCtgFr").CmpStr = "": adc("UoDmCtgTo").CmpStr = "": adc("UoDmCtgSel").CmpStr = ""
    adc("UoDmPrdCtgFr").CmpStr = "": adc("UoDmPrdCtgTo").CmpStr = "": adc("UoDmPrdCtgSel").CmpStr = ""
    adc("UoDmCdFr").CmpStr = "": adc("UoDmCdTo").CmpStr = "": adc("UoDmCdSel").CmpStr = ""
    adc("UoOmTcFr").CmpStr = "": adc("UoOmTcTo").CmpStr = ""
    adc("UoOmYyFr").CmpStr = "": adc("UoOmYyTo").CmpStr = ""
    adc("UoOmChrFr").CmpStr = "": adc("UoOmChrTo").CmpStr = ""
    adc("UoOmNoFr").CmpStr = "": adc("UoOmNoTo").CmpStr = ""
    adc("UoOdSrFr").CmpStr = "": adc("UoOdSrTo").CmpStr = ""
    LblCurrCd.Visible = False: adc("UOCURCDFR").Visible = False
    ' **** Zubin 211 **** '
    LblBagTyp.Visible = False: adc("UoBagTyp").Visible = False
    ' **** Zubin 211 **** '
    ' Zubin 212
    adc("UoLotNoFr").CmpStr = "TdLotNo>= ": adc("UoLotNoTo").CmpStr = "TdLotNo<= ": adc("UoLotNoSel").CmpStr = "TdLotNo In "
  ' Zubin 212
    '*** General Properties ***
    If UCase(adc.MenuCd) <> UCase("RepTxnREC") Then _
    FraOptAck.Visible = True: FraOptAck.Enabled = True: LblAck.Visible = True  'MMAck-Enable visible property of Ack label & option buttons
    'IssRecLoc- .CmpStr applied to Issue & Receiving Loc
    adc("UORECLOCFR").CmpStr = " (CASE WHEN TdFrRmDc='D' THEN TdFrRmLoc ELSE TdToRmLoc END) >= "
    adc("UORECLOCTO").CmpStr = " (CASE WHEN TdFrRmDc='D' THEN TdFrRmLoc ELSE TdToRmLoc END) <= "
    adc("UORECLOCSEL").CmpStr = " (CASE WHEN TdFrRmDc='D' THEN TdFrRmLoc ELSE TdToRmLoc END) In "
    adc("UOISSLOCFR").CmpStr = " (CASE WHEN TdFrRmDc='C' THEN TdFrRmLoc ELSE TdToRmLoc END) >= "
    adc("UOISSLOCTO").CmpStr = " (CASE WHEN TdFrRmDc='C' THEN TdFrRmLoc ELSE TdToRmLoc END) <= "
    adc("UOISSLOCSEL").CmpStr = " (CASE WHEN TdFrRmDc='C' THEN TdFrRmLoc ELSE TdToRmLoc END) In "
    
    '*** Only For Purchase ***
    If UCase(adc.MenuCd) = UCase("RepTxnPR") Then
      gs_TxnTcTyp = "PR"
      adc("UoCmCdFr").CmpStr = "TSuppCd>= ": adc("UoCmCdTo").CmpStr = "TSuppCd<= ": adc("UoCmCdSel").CmpStr = "TSuppCd In "
      adc("UoBOEFr").CmpStr = "TBOENo>= ": adc("UoBOETo").CmpStr = "TBOENo<= ": adc("UoBOESel").CmpStr = "TBOENo In "
      LblSuppCd.Visible = True: LblBOENo.Visible = True
      adc("UoCmCdFr").Visible = True: adc("UoCmCdTo").Visible = True: adc("UoCmCdSel").Visible = True
      adc("UoBOEFr").Visible = True: adc("UoBOETo").Visible = True: adc("UoBOESel").Visible = True
      adc("UoCmCdFr").Enabled = True: adc("UoCmCdTo").Enabled = True: adc("UoCmCdSel").Enabled = True
      adc("UoBOEFr").Enabled = True: adc("UoBOETo").Enabled = True: adc("UoBOESel").Enabled = True
      LblCurrCd.Visible = True: adc("UOCURCDFR").Visible = True
      ' ***** Manali 3.3.0
      LblStkVal.Visible = True: 'adc("UoYN1").Visible = True:  adc("UoYN1").Enabled = True  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
      ' ***** Manali 3.3.0
      FraOptAmt.Visible = True: FraOptAmt.Enabled = True ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
      '************Geeta****************Emr207*******************
        'LblUSDYN.Visible = True: adc("UoYN").Visible = True
      '************Geeta****************Emr207*******************
    End If
    '*** Only For Purchase ***
    '*** Only For Recovery ***
    If UCase(adc.MenuCd) = UCase("RepTxnREC") Then
      gs_TxnTcTyp = "REC"
      adc("UoIssLocFr").CmpStr = "TLsLoc>= ": adc("UoIssLocTo").CmpStr = "TLsLoc<= ": adc("UoIssLocSel").CmpStr = "TLsLoc In "
      adc("UoLsDtFr").CmpStr = "TLsFrDt>= ": adc("UoLsDtTo").CmpStr = "TLsToDt<= "
      LblLsLoc.Visible = True: LblLsDt.Visible = True
      adc("UoIssLocFr").Visible = True: adc("UoIssLocTo").Visible = True: adc("UoIssLocSel").Visible = True
      adc("UoLsDtFr").Visible = True: adc("UoLsDtTo").Visible = True
      adc("UoIssLocFr").Enabled = True: adc("UoIssLocTo").Enabled = True: adc("UoIssLocSel").Enabled = True
      adc("UoLsDtFr").Enabled = True: adc("UoLsDtTo").Enabled = True
      'IssRecLoc-LblLsLoc caption chnaged
      LblLsLoc.Caption = "Loss Loc"
    End If
    '*** Only For Recovery ***
    If UCase(adc.MenuCd) = UCase("RepTxnCNV") Then gs_TxnTcTyp = "CNV"
    If UCase(adc.MenuCd) = UCase("RepTxnMV") Then gs_TxnTcTyp = "MV"
  End Select
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  ' Zubin 212
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
  adc("UoLsDtFr") = "01/01/1980"
  adc("UoLsDtTo") = "01/01/1980"
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(adc("UOCOCDFR"), ws_FrmCaption)
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
 
' **** Zubin 211 **** '
  If (UCase(adc.MenuCd) = UCase("RepTxnDT") Or UCase(adc.MenuCd) = UCase("RepTxnBS") _
      Or UCase(adc.MenuCd) = UCase("RepTxnMLT") Or UCase(adc.MenuCd) = UCase("RepTxnBV") _
      Or UCase(adc.MenuCd) = UCase("RepTxnREJ") Or UCase(adc.MenuCd) = UCase("RepTxnBM")) Then
    If adc("UoBagTyp") = "" Then adc("UoBagTyp") = "P"
  Else
    adc("UoBagTyp") = ""
  End If
' **** Zubin 211 **** '

 '*** Jay 2.13(CT) ***
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + ctSelfCoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    adc("UoBagTyp").Visible = False
    adc("UoBagTyp") = "P"
  End If
  Set wHeadChk = Nothing
  '*** Jay 2.13(CT) ***
  
'*** Jay 2.13(CT) ***
'  '*** (Jen 2.11)
'  If moCn.GetFldVal("Select hFluteBagYN from Head where HCoCd= '" + adc("UoCoCdFr") + "' and " + _
'                                                    "HCd= '" + ctSelfCmCd + "' ") = "N" Then
'    LblBagTyp.Visible = False
'    adc("UoBagTyp").Visible = False
'    adc("UoBagTyp") = "N"
'  End If
'  '*** (Jen 2.11)
'*** Jay 2.13(CT) ***

  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  If adc("UoYN1") = "Y" Then OptAmt(0).Value = True
  If adc("UoYN1") = "N" Then OptAmt(1).Value = True
  'MMAck-Ack option button selection changing as per value in UOYN2 textbox
  If adc("UoYN2") = "Y" Then
    OptAck(0).Value = True
  ElseIf adc("UoYN2") = "N" Then
    OptAck(1).Value = True
  ElseIf adc("UoYN2") = "" Then
    OptAck(2).Value = True
  End If
  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  'StkRt-Set Show StockRate option is No
  adc("UoYN3") = "N"
  
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  If UCase(adc.MenuCd) = UCase("RepTxnDT") Then gs_TxnTcTyp = "DT"
  If UCase(adc.MenuCd) = UCase("RepTxnBS") Then gs_TxnTcTyp = "BS"
  If UCase(adc.MenuCd) = UCase("RepTxnMLT") Then gs_TxnTcTyp = "MLT"
  If UCase(adc.MenuCd) = UCase("RepTxnBV") Then gs_TxnTcTyp = "BV"
  If UCase(adc.MenuCd) = UCase("RepTxnREJ") Then gs_TxnTcTyp = "REJ"
  If UCase(adc.MenuCd) = UCase("RepTxnPR") Then gs_TxnTcTyp = "PR"
  If UCase(adc.MenuCd) = UCase("RepTxnREC") Then gs_TxnTcTyp = "REC"
  If UCase(adc.MenuCd) = UCase("RepTxnCNV") Then gs_TxnTcTyp = "CNV"
  If UCase(adc.MenuCd) = UCase("RepTxnMV") Then gs_TxnTcTyp = "MV"
' **** Zubin 211 **** '
  If UCase(adc.MenuCd) = UCase("RepTxnBM") Then gs_TxnTcTyp = "BM"
' **** Zubin 211 **** '
  
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
  If UCase(adc.MenuCd) = UCase("RepTxnDT") Then gs_TxnTcTyp = "DT"
  If UCase(adc.MenuCd) = UCase("RepTxnBS") Then gs_TxnTcTyp = "BS"
  If UCase(adc.MenuCd) = UCase("RepTxnMLT") Then gs_TxnTcTyp = "MLT"
  If UCase(adc.MenuCd) = UCase("RepTxnBV") Then gs_TxnTcTyp = "BV"
  If UCase(adc.MenuCd) = UCase("RepTxnREJ") Then gs_TxnTcTyp = "REJ"
  If UCase(adc.MenuCd) = UCase("RepTxnPR") Then gs_TxnTcTyp = "PR"
  If UCase(adc.MenuCd) = UCase("RepTxnREC") Then gs_TxnTcTyp = "REC"
  If UCase(adc.MenuCd) = UCase("RepTxnCNV") Then gs_TxnTcTyp = "CNV"
  If UCase(adc.MenuCd) = UCase("RepTxnMV") Then gs_TxnTcTyp = "MV"
' **** Zubin 211 **** '
  If UCase(adc.MenuCd) = UCase("RepTxnBM") Then gs_TxnTcTyp = "BM"
' **** Zubin 211 **** '
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  gb_CoCdFor = True
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
  '*** Jay 3.01 **** [LotNo Err]
  gs_RmLocTyp = "": gs_Loc = ""
  '*** Jay 3.01 **** [LotNo Err]
   
  SetProp Me, IdName, When
End Sub

Private Sub OptAmt_Click(Index As Integer)
Select Case Index
Case Is = 0
    adc("UoYN1") = "Y"
Case Is = 1
    adc("UoYN1") = "N"
End Select
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

Private Sub OptAmt_GotFocus(Index As Integer)
  DispMsg FraOptAmt.ToolTipText, etInfo
End Sub

Private Sub OptAck_GotFocus(Index As Integer)
  DispMsg FraOptAck.ToolTipText, etInfo
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
  Set Rep1 = Nothing
  Set Rep2 = Nothing
  Set Rep3 = Nothing
  '*** (09/08/05)
End Sub
'StkRt-if password is change to blank, Show StockRate check box will turn into blank
Private Sub TxtPwd_Change()
    If Trim(TxtPwd.text) = "" Then
        ChkBoxArr(0).Value = Unchecked
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
Private Sub ChkBoxArr_Click(Index As Integer)
 'StkRt-To select this option(Show Stock rate) either Seo user can give their password or
  ' can enter SEOPWD password otherwise this option can't not select
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0
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

