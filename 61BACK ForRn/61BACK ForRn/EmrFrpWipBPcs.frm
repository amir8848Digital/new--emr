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
   ScaleHeight     =   11040
   ScaleWidth      =   20400
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   10830
      TabIndex        =   95
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
      TabIndex        =   94
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
      TabIndex        =   104
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   105
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
         TabIndex        =   106
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
      TabIndex        =   103
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   100
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
      TabIndex        =   96
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
      TabIndex        =   99
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   98
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
            TabIndex        =   101
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraInvNo 
               BorderStyle     =   0  'None
               Height          =   645
               Left            =   60
               TabIndex        =   167
               Top             =   7020
               Visible         =   0   'False
               Width           =   6945
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   83
                  Left            =   5205
                  TabIndex        =   72
                  ToolTipText     =   "Enter To Invoice Year"
                  Top             =   0
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataField       =   "UOINYYTO"
                  IdName          =   "UOINYYTO"
                  CmpStr          =   "isnull(IfYy,'') <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   84
                  Left            =   2265
                  TabIndex        =   69
                  ToolTipText     =   "Enter From Invoice Year"
                  Top             =   15
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataField       =   "UOINYYFR"
                  IdName          =   "UOINYYFR"
                  CmpStr          =   "isnull(IfYy,'') >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   85
                  Left            =   2685
                  TabIndex        =   70
                  ToolTipText     =   "Enter From Invoice Character"
                  Top             =   15
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataType        =   4
                  DataField       =   "UOINCHRFR"
                  IdName          =   "UOINCHRFR"
                  CmpStr          =   "isnull(IfChr,'') >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   86
                  Left            =   4425
                  TabIndex        =   71
                  ToolTipText     =   "Enter To Invoice Tc"
                  Top             =   0
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataField       =   "UOINTCTO"
                  IdName          =   "UOINTCTO"
                  CmpStr          =   "isnull(IfTc,'') <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   87
                  Left            =   5625
                  TabIndex        =   73
                  ToolTipText     =   "Enter To Invoice Character"
                  Top             =   0
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataType        =   4
                  DataField       =   "UOINCHRTO"
                  IdName          =   "UOINCHRTO"
                  CmpStr          =   "isnull(IfChr,'') <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   88
                  Left            =   4425
                  TabIndex        =   75
                  ToolTipText     =   "Enter To Invoice Number"
                  Top             =   285
                  Width           =   825
                  _ExtentX        =   1455
                  _ExtentY        =   503
                  MaxLength       =   0
                  DataType        =   1
                  DataField       =   "UOINNOTO"
                  IdName          =   "UOINNOTO"
                  CmpStr          =   "isnull(IfNo,0) <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   89
                  Left            =   1485
                  TabIndex        =   68
                  ToolTipText     =   "Enter From Invoice Tc"
                  Top             =   15
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataField       =   "UOINTCFR"
                  IdName          =   "UOINTCFR"
                  CmpStr          =   "isnull(IfTc,'') >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   90
                  Left            =   1485
                  TabIndex        =   74
                  ToolTipText     =   "Enter From Invoice Number"
                  Top             =   300
                  Width           =   825
                  _ExtentX        =   1455
                  _ExtentY        =   503
                  MaxLength       =   0
                  DataType        =   1
                  DataField       =   "UOINNOFR"
                  IdName          =   "UOINNOFR"
                  CmpStr          =   "isnull(IfNo,0) >="
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Inv Tc/Yy/Chr"
                  BeginProperty Font 
                     Name            =   "Arial"
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
                  Left            =   0
                  TabIndex        =   173
                  Top             =   15
                  Width           =   1365
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Inv No"
                  BeginProperty Font 
                     Name            =   "Arial"
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
                  Left            =   0
                  TabIndex        =   172
                  Top             =   300
                  Width           =   1395
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
                  Left            =   5565
                  TabIndex        =   171
                  Top             =   0
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
                  Left            =   2625
                  TabIndex        =   170
                  Top             =   15
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
                  Left            =   5145
                  TabIndex        =   169
                  Top             =   0
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
                  Left            =   2205
                  TabIndex        =   168
                  Top             =   15
                  Width           =   105
               End
            End
            Begin VB.Frame FraDelayDays 
               BorderStyle     =   0  'None
               Height          =   405
               Left            =   8880
               TabIndex        =   165
               ToolTipText     =   "Select Delay Days As"
               Top             =   7800
               Width           =   6225
               Begin VB.OptionButton OptDlyDays 
                  Caption         =   "By Opn Dt"
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
                  Index           =   3
                  Left            =   4920
                  TabIndex        =   174
                  ToolTipText     =   "Select Delay Days As "
                  Top             =   0
                  Width           =   1245
               End
               Begin VB.OptionButton OptDlyDays 
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
                  Height          =   285
                  Index           =   2
                  Left            =   3240
                  TabIndex        =   90
                  ToolTipText     =   "Select Delay Days As"
                  Top             =   0
                  Width           =   1605
               End
               Begin VB.OptionButton OptDlyDays 
                  Caption         =   "By Location"
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
                  Left            =   1680
                  TabIndex        =   89
                  ToolTipText     =   "Select Delay Days As"
                  Top             =   0
                  Width           =   1425
               End
               Begin VB.OptionButton OptDlyDays 
                  Caption         =   "Not Required"
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
                  TabIndex        =   88
                  ToolTipText     =   "Select Delay Days As "
                  Top             =   0
                  Width           =   1605
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show All Weights "
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
               TabIndex        =   87
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show All Weights"
               Top             =   8760
               Width           =   2085
            End
            Begin VB.Frame FraLoc 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1440
               TabIndex        =   156
               Top             =   680
               Width           =   4785
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
                  Left            =   3045
                  TabIndex        =   158
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
                  TabIndex        =   157
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
               TabIndex        =   150
               ToolTipText     =   "Show Bags In FG / INV / Both"
               Top             =   7905
               Width           =   3945
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
                  Left            =   2940
                  TabIndex        =   153
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
                  TabIndex        =   152
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
                  TabIndex        =   151
                  ToolTipText     =   "Show Bags In FG / INV / Both"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.Frame FraOptBagPcs 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1545
               TabIndex        =   142
               ToolTipText     =   "Show Normal Or Rejected Bag Pieces"
               Top             =   8190
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
                  TabIndex        =   84
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
                  TabIndex        =   85
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
               TabIndex        =   86
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   8475
               Width           =   1725
            End
            Begin VB.Frame FraOptBag 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1545
               TabIndex        =   141
               ToolTipText     =   "Show Open Or Closed Or All Bags"
               Top             =   7620
               Width           =   3945
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
                  TabIndex        =   81
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
                  TabIndex        =   82
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
                  Left            =   2940
                  TabIndex        =   83
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
               TabIndex        =   63
               Top             =   6165
               Width           =   2340
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4485
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
               Left            =   7485
               TabIndex        =   8
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1305
               Width           =   7550
               _ExtentX        =   13309
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4485
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
               Left            =   7485
               TabIndex        =   20
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2445
               Width           =   7550
               _ExtentX        =   13309
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
               Left            =   4485
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
               Left            =   7485
               TabIndex        =   23
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2730
               Width           =   7550
               _ExtentX        =   13309
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "BOdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   7485
               TabIndex        =   39
               ToolTipText     =   "Enter Customer Selection"
               Top             =   3885
               Width           =   7550
               _ExtentX        =   13309
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
               Left            =   5265
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
               Left            =   4485
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
               Left            =   5685
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
               Left            =   4485
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
               Left            =   4485
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
               Left            =   7485
               TabIndex        =   26
               ToolTipText     =   "Enter Karat Selection"
               Top             =   3015
               Width           =   7550
               _ExtentX        =   13309
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "OdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   4485
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
               Left            =   7485
               TabIndex        =   42
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   4170
               Width           =   7550
               _ExtentX        =   13309
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5370
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
               Left            =   7485
               TabIndex        =   11
               ToolTipText     =   "Enter FgSubLoc Location Selection"
               Top             =   1590
               Width           =   7550
               _ExtentX        =   13309
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOFGSUBLOCSEL"
               IdName          =   "UOFGSUBLOCSEL"
               CmpStr          =   "BFgSubLoc In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   4485
               TabIndex        =   62
               ToolTipText     =   "Show Bags which are Dalayed By Less  than or equal to Entered Days: (Enter zero to see all Bags)"
               Top             =   5865
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0"
               MaxLength       =   4
               DataType        =   1
               DataField       =   "UOBDLDAYSTO"
               IdName          =   "UOBDLDAYSTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2085
               TabIndex        =   93
               ToolTipText     =   "Show Picture? ( Enter 'Y'/'N')"
               Top             =   8475
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
               TabIndex        =   76
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   7020
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOBAGTYP"
               ReCalcParent    =   "UOBCHRFR,UOBCHR2FR,UOBCHR2TO,UOBCHR2SEL"
               IdName          =   "UOBAGTYP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   1800
               TabIndex        =   78
               ToolTipText     =   "Enter Bag Character From"
               Top             =   7020
               Visible         =   0   'False
               Width           =   105
               _ExtentX        =   185
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRFR"
               ReCalcOn        =   "UOYN2"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "BChr >= "
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
               Left            =   1560
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
               TabIndex        =   92
               ToolTipText     =   "Show Normal Or Rejected Pieces ?"
               Top             =   8190
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
               TabIndex        =   91
               ToolTipText     =   "Show Open Or Closed Or All Bags ? (Open - 'Y', Closed - 'N', All - Blank)"
               Top             =   7620
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
               TabIndex        =   56
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
               TabIndex        =   57
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
               Left            =   4485
               TabIndex        =   58
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
               Left            =   4905
               TabIndex        =   59
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
               Left            =   5685
               TabIndex        =   60
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
               TabIndex        =   64
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
               TabIndex        =   65
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
               Left            =   4485
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
               CmpStr          =   "BCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   7485
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   360
               Width           =   7550
               _ExtentX        =   13309
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "BCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   4485
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
               Left            =   7485
               TabIndex        =   5
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1020
               Width           =   7550
               _ExtentX        =   13309
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
               Left            =   2820
               TabIndex        =   44
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   4455
               Width           =   1655
               _ExtentX        =   2910
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   5770
               TabIndex        =   46
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   4455
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   7485
               TabIndex        =   47
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   4455
               Visible         =   0   'False
               Width           =   7550
               _ExtentX        =   13309
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "OdPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   4485
               TabIndex        =   52
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
               TabIndex        =   51
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
               Left            =   4485
               TabIndex        =   54
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
               TabIndex        =   53
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
               Left            =   4485
               TabIndex        =   49
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
               Left            =   7485
               TabIndex        =   50
               ToolTipText     =   "Enter Priority Code Selection"
               Top             =   4740
               Width           =   7550
               _ExtentX        =   13309
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
               TabIndex        =   48
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
               Left            =   7485
               TabIndex        =   14
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1875
               Visible         =   0   'False
               Width           =   7550
               _ExtentX        =   13309
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORECLOCSEL"
               IdName          =   "UORECLOCSEL"
               CmpStr          =   "IsNull(TdFrBLoc,BLoc) In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   4485
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
               Left            =   4485
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
               Left            =   4485
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
               Left            =   7485
               TabIndex        =   17
               ToolTipText     =   "Enter By Worker Selection"
               Top             =   2160
               Width           =   7550
               _ExtentX        =   13309
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
               Left            =   4485
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
               Left            =   1560
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
               TabIndex        =   61
               ToolTipText     =   "Show Bags which are Dalayed By More than or equal to Entered Days: (Enter zero to see all Bags)"
               Top             =   5880
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0"
               MaxLength       =   4
               DataType        =   1
               DataField       =   "UOBDLDAYS"
               IdName          =   "UOBDLDAYS"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1545
               TabIndex        =   55
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
               TabIndex        =   66
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
               Left            =   4485
               TabIndex        =   67
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
               TabIndex        =   154
               ToolTipText     =   "Show Bags In FG / INV / Both"
               Top             =   7905
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN2"
               ReCalcParent    =   "UOBCHRFR,UOBCHR2FR,UOBCHR2TO,UOBCHR2SEL"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   6960
               TabIndex        =   159
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
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   72
               Left            =   4485
               TabIndex        =   79
               ToolTipText     =   "Enter Bag Character To"
               Top             =   7320
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHR2TO"
               ReCalcOn        =   "UOYN2,UOBAGTYP"
               IdName          =   "UOBCHR2TO"
               CmpStr          =   "BChr <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   73
               Left            =   7530
               TabIndex        =   80
               ToolTipText     =   "Enter Bag Character Selection"
               Top             =   7320
               Width           =   7485
               _ExtentX        =   13203
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOBCHR2SEL"
               ReCalcOn        =   "UOYN2,UOBAGTYP"
               IdName          =   "UOBCHR2SEL"
               CmpStr          =   "BChr In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   74
               Left            =   1545
               TabIndex        =   77
               ToolTipText     =   "Enter Bag Character From"
               Top             =   7320
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHR2FR"
               ReCalcOn        =   "UOYN2,UOBAGTYP"
               IdName          =   "UOBCHR2FR"
               CmpStr          =   "BChr >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   1545
               TabIndex        =   43
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   4455
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
               Left            =   4495
               TabIndex        =   45
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   4455
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
               Index           =   77
               Left            =   2160
               TabIndex        =   163
               ToolTipText     =   "Show All Weights "
               Top             =   8760
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               IdName          =   "WSHOWALLWTS"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   78
               Left            =   14640
               TabIndex        =   166
               ToolTipText     =   "Show Delay Days As Not Required /By Location /By Worker"
               Top             =   8160
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UODLYDAYS"
               IdName          =   "UODLYDAYS"
            End
            Begin VB.Label LblDelayDays 
               BackStyle       =   0  'Transparent
               Caption         =   "Delay Days As"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   7440
               TabIndex        =   164
               Top             =   7800
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Note: Selecting the option as ‘Y’es , may slow down the processing of report."
               BeginProperty Font 
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
               Index           =   10
               Left            =   2520
               TabIndex        =   162
               Top             =   8720
               Width           =   7455
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
               Left            =   2760
               TabIndex        =   161
               Top             =   4455
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
               Left            =   5710
               TabIndex        =   160
               Top             =   4440
               Width           =   105
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
               TabIndex        =   155
               Top             =   7905
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
               TabIndex        =   149
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
               TabIndex        =   148
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
               TabIndex        =   147
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
               TabIndex        =   146
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
               TabIndex        =   145
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
               TabIndex        =   144
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
               TabIndex        =   143
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
               TabIndex        =   140
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
               TabIndex        =   139
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
               TabIndex        =   138
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
               TabIndex        =   137
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
               TabIndex        =   136
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
               Left            =   5625
               TabIndex        =   135
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
               Left            =   4845
               TabIndex        =   134
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
               TabIndex        =   133
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
               TabIndex        =   132
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
               Left            =   60
               TabIndex        =   131
               Top             =   7335
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
               TabIndex        =   130
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
               TabIndex        =   129
               ToolTipText     =   "Location"
               Top             =   8520
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
               TabIndex        =   128
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
               TabIndex        =   127
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
               TabIndex        =   126
               Top             =   7620
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
               TabIndex        =   125
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
               Left            =   5310
               TabIndex        =   124
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
               TabIndex        =   123
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
               TabIndex        =   122
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
               TabIndex        =   121
               Top             =   8190
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
               TabIndex        =   120
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
               TabIndex        =   119
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
               TabIndex        =   118
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
               Left            =   5625
               TabIndex        =   117
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
               TabIndex        =   116
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
               Left            =   5205
               TabIndex        =   115
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
               TabIndex        =   114
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
               TabIndex        =   113
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
               TabIndex        =   112
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
               TabIndex        =   111
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
               TabIndex        =   110
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
               Left            =   7485
               TabIndex        =   109
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
               Left            =   4455
               TabIndex        =   108
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
               TabIndex        =   107
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
            TabIndex        =   102
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   97
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
  If UCase(ADC.MenuCd) = UCase("RprWIP") Then
    GRP_REP.Add "Location", "BRepLoc", "", "BLoc", "hBLoc", "Company Code,Parent Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'P' and LocCd= BRepLoc) "
    GRP_REP.Add "Process", "LocPrcs", "", "BLoc", "hBLoc", "Company Code,Parent Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'R' and LocCd= Lc.LocPrcs) "
    GRP_REP.Add "(Location)Worker", "BRepLoc+BWrk", "'('+ RTrim(BRepLoc)+ ') '+ BWrk", "BLoc,BWrk", "hBLoc,hBWrk", "Parent Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'WORK' and vPMCd= BWrk) "
    GRP_REP.Add "(Process)Worker", "LocPrcs+BWrk", "'('+ RTrim(LocPrcs)+ ') '+ BWrk", "BLoc,BWrk", "hBLoc,hBWrk", "Parent Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'WORK' and vPMCd= BWrk) "
    
    GRP_REP.Add "OrderNo", "BRepOdTc+BRepOdYy+BRepOdChr+str(BRepOdNo)", "BRepOdTc+'/'+BRepOdYy+'/'+BRepOdChr+'/'+LTrim(Str(BRepOdNo))", "OmCmCd,OmPoNo", "hOmCmCd,hOmPoNo", "Company Code,Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
    GRP_REP.Add "OrderSr", "BRepOdTc+BRepOdYy+BRepOdChr+str(BRepOdNo)+str(BRepOdSr)", "BRepOdTc+'/'+BRepOdYy+'/'+BRepOdChr+'/'+LTrim(Str(BRepOdNo))+'/'+LTrim(Str(BRepOdSr))", "OmCmCd,OmPoNo", "hOmCmCd,hOmPoNo", "Company Code,OrderNo,Ord Kt,Design Cd,Design Ctg,(DsgCd)Colour,Customer Code,Customer PO,Prd Del Date,Priority,(OrdKt)Colour", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
    GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "Company Code", "", "", "BRepOdTc+'/'+BRepOdYy+'/'+BRepOdChr+'/'+LTrim(Str(BRepOdNo))"
    'vk.19 sub customer added
    GRP_REP.Add "Cust SUB PO", "OmCmCd+OmPONo+OdSubCust+OdPoNo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)+(case when OdPONo<> '' Or OdSubCust<>'' then '/'+LTrim(OdSubCust)+' : '+ LTrim(OdPONo) else '' end )", "OmCmCd,OmPoNo,OdPoNo", "hOmCmCd,hOmPONo,hOdPONo", "Customer Code,Customer PO", "", "", "BRepOdTc+'/'+BRepOdYy+'/'+BRepOdChr+'/'+LTrim(Str(BRepOdNo))+'/'+LTrim(Str(BRepOdSr))+(case when OdPONo<> '' Or OdSubCust<>'' then '/'+LTrim(OdSubCust)+' : '+ LTrim(OdPONo) else '' end )"
  Else
    GRP_REP.Add "Location", "BLoc", "", "BLoc", "hBLoc", "Company Code,Parent Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'P' and LocCd= BLoc) "
    GRP_REP.Add "Process", "LocPrcs", "", "BLoc", "hBLoc", "Company Code,Parent Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'R' and LocCd= Lc.LocPrcs) "
    GRP_REP.Add "(Location)Worker", "BLoc+BWrk", "'('+ RTrim(BLoc)+ ') '+ BWrk", "BLoc,BWrk", "hBLoc,hBWrk", "Parent Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'WORK' and vPMCd= BWrk) "
    GRP_REP.Add "(Process)Worker", "LocPrcs+BWrk", "'('+ RTrim(LocPrcs)+ ') '+ BWrk", "BLoc,BWrk", "hBLoc,hBWrk", "Parent Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'WORK' and vPMCd= BWrk) "
    
    GRP_REP.Add "OrderNo", "BOdTc+BOdYy+BOdChr+str(BOdNo)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "OmCmCd,OmPoNo", "hOmCmCd,hOmPoNo", "Company Code,Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
    GRP_REP.Add "OrderSr", "BOdTc+BOdYy+BOdChr+str(BOdNo)+str(BOdSr)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))", "OmCmCd,OmPoNo", "hOmCmCd,hOmPoNo", "Company Code,OrderNo,Ord Kt,Design Cd,Design Ctg,(DsgCd)Colour,Customer Code,Customer PO,Prd Del Date,Priority,(OrdKt)Colour", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
    GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "Company Code", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))"
    'vk.19 sub customer added
    GRP_REP.Add "Cust SUB PO", "OmCmCd+OmPONo+OdSubCust+OdPoNo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)+(case when OdPONo<> '' Or OdSubCust<>'' then '/'+LTrim(OdSubCust)+' : '+ LTrim(OdPONo) else '' end )", "OmCmCd,OmPoNo,OdPoNo", "hOmCmCd,hOmPONo,hOdPONo", "Customer Code,Customer PO", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))+(case when OdPONo<> '' Or OdSubCust<>'' then '/'+LTrim(OdSubCust)+' : '+ LTrim(OdPONo) else '' end )"
    
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
  If UCase(ADC.MenuCd) = UCase("RepWipBPcs") Or UCase(ADC.MenuCd) = UCase("RprWIP") Then
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
    
  ElseIf UCase(ADC.MenuCd) = UCase("RepWipFgBPcs") Then
    GRP_REP.Add "Pck List", "FdPlYy+FdPlChr+str(FdPlNo)", "(case when FdPlNo> 0 then FdPlYy+'/'+FdPlChr+'/'+LTrim(Str(FdPlNo)) else '' end)", "", "", "", "", "", ""
    GRP_REP.Add "FG / INV", " (Case when exists (Select 'x' From InvFgd where IfBIdNo = BIdNo and IfFdBYy = BYy and IfFdBChr = BChr and IfFdBNo = BNo and IfCoCd = BCoCd) then 'INV' else 'FG' end)", " (Case when exists (Select 'x' From InvFgd where IfBIdNo = BIdNo and IfFdBYy = BYy and IfFdBChr = BChr and IfFdBNo = BNo and IfCoCd = BCoCd) then 'INV' else 'FG' end)", "", "", "", "", "", ""
  '*** 2.11 (Jen)
    '3 InvNo grp-sort added
    GRP_REP.Add "InvNo", "isnull(IfTc+IfYy+IfChr+str(IfNo),'')", "isnull(IfTc+'/'+IfYy+'/'+IfChr+'/'+LTrim(Str(IfNo)),'')", "", "", "", "", "", ""

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
  wb_InBaseCur = IIF(UCase(ADC.MenuCd) = UCase("RepWipFgBPcs") And UCase(ADC("UoCurCdFr")) = "", True, False)
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + ADC("UoCoCdFr") + "'")
  ' Zubin 213
      
  Set Rep = Nothing
  Set ADC.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
    grpflds = GRP_REP.GrpFldLst
   
   '*** To check if Grp1, Grp2 is selected and also to check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer, wi_InvNoPos As Integer
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
    '3 getting InvNo group sort position
    If UCase(wGrp(i)) = UCase("InvNo") Then
      wi_InvNoPos = i
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
    If UCase(ADC("UoPMCdFr")) = "C" Then
      ' Zubin 213 (Bef 213)
      'wFrVal = "OdCstPrc*BQty as qSalVal "
      'Rep.TxtValTyp.SetText "(Value = Cost)"
      ' Zubin 213
      wFrVal = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdCstPrc/OmCnvFct else OdCstPrc*OmCnvFct end) ", "OdCstPrc") + "*BQty as qSalVal "
      Rep.TxtValTyp.SetText "(Value = Cost [" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + "])"
    ElseIf UCase(ADC("UoPMCdFr")) = "S" Then
      ' Zubin 213 (Bef 213)
      'wFrVal = "OdSalPrc*BQty as qSalVal "
      'Rep.TxtValTyp.SetText "(Value = Sale)"
      ' Zubin 213
      wFrVal = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc") + "*BQty as qSalVal "
      Rep.TxtValTyp.SetText "(Value = Sale [" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + "])"
    Else
      wFrVal = " 0.00 as qSalVal "
      Rep.TxtValTyp.SetText ""
      Rep.hSalVal.Suppress = True
      Rep.h1SalVal.Suppress = True
    End If
  ' ****** sachin 2.12 - 24/11/05 ****** '

  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    'Uma Location / Process selection
    If OptLoc(0).Value = True Then
      If UCase(ADC.MenuCd) = UCase("RprWIP") Then
        If ADC("UoIssLocFr") <> "" Then wCnd = wCnd + " and BRepLoc >= '" + ADC("UoIssLocFr") + "'"
        If ADC("UoIssLocTo") <> "" Then wCnd = wCnd + " and BRepLoc <= '" + ADC("UoIssLocTo") + "'"
        If ADC("UoIssLocSel") <> "" Then wCnd = wCnd + " and BRepLoc In (" + ADC("UoIssLocSel") + ")"
      Else
        If ADC("UoIssLocFr") <> "" Then wCnd = wCnd + " and BLoc >= '" + ADC("UoIssLocFr") + "'"
        If ADC("UoIssLocTo") <> "" Then wCnd = wCnd + " and BLoc <= '" + ADC("UoIssLocTo") + "'"
        If ADC("UoIssLocSel") <> "" Then wCnd = wCnd + " and BLoc In (" + ADC("UoIssLocSel") + ")"
      End If
      If ADC("UOPRNTCDFR") <> "" Then wCnd = wCnd + " and LocPrntCd >= '" + ADC("UOPRNTCDFR") + "'"
      If ADC("UOPRNTCDTO") <> "" Then wCnd = wCnd + " and LocPrntCd <= '" + ADC("UOPRNTCDTO") + "'"
      If ADC("UOPRNTCDSEL") <> "" Then wCnd = wCnd + " and LocPrntCd In (" + ADC("UOPRNTCDSEL") + ")"
      
    ElseIf OptLoc(1).Value = True Then
      If ADC("UoIssLocFr") <> "" Then wCnd = wCnd + " and LocPrcs >= '" + ADC("UoIssLocFr") + "'"
      If ADC("UoIssLocTo") <> "" Then wCnd = wCnd + " and LocPrcs <= '" + ADC("UoIssLocTo") + "'"
      If ADC("UoIssLocSel") <> "" Then wCnd = wCnd + " and LocPrcs In (" + ADC("UoIssLocSel") + ")"
            
      If ADC("UOPRNTCDFR") <> "" Then wCnd = wCnd + " and (Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs) >= '" + ADC("UOPRNTCDFR") + "'"
      If ADC("UOPRNTCDTO") <> "" Then wCnd = wCnd + " and (Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs) <= '" + ADC("UOPRNTCDTO") + "'"
      If ADC("UOPRNTCDSEL") <> "" Then wCnd = wCnd + " and (Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs) In (" + ADC("UOPRNTCDSEL") + ")"
            
    End If

  
  ' Zubin 213
  wCnd = wCnd + IIF(UCase(ADC.MenuCd) = UCase("RepWipFgBPcs") And wb_InBaseCur = False, _
                                " and OmCmCurCd= '" + ADC("UoCurCdFr") + "'", "")
  
  'wSelCnd = IIf(adc("UoBagPcs") = "N", " where (BQty<> 0 or BPtQty<> 0 or (BQty= 0 and BPtQty= 0 and BRjQty= 0 and BRjPtQty= 0 and Round(BGrWt, 3)<> 0)) ", " where (BRjQty<> 0 or BRjPtQty<> 0) ")
  '### wSelCnd = IIf(ADC("UoBagPcs") = "N", " where (BQty<> 0 or BPtQty<> 0) ", " where (BRjQty<> 0 or BRjPtQty<> 0) ")    '(This condition should have given the same result as above but it does not give)
  '### wSelCnd = IIf(ADC("UoBagPcs") = "N", " where (BRjQty= 0 and BRjPtQty= 0) ", " where (BRjQty<> 0 or BRjPtQty<> 0) ")
  
  wBagClsCnd = IIF(ADC("UoYN") = "Y", " and BCls= 'N' ", IIF(ADC("UoYN") = "N", " and BCls= 'Y' ", ""))
  
  ' urmi BDlDays 2.06
  wBagDelayCnd = IIF(ADC("UoBDlDays") = 0, "", "and datediff(dd, BRecvDt, '" + CStr(moCn.SrvrDate) + "') >= " + CStr(ADC("UoBDlDays")))
  '*****Geeta*****Emr206********
  wBagDelayCnd = wBagDelayCnd + IIF(ADC("UoBDlDaysTo") = 0, "", " and  datediff(dd, BRecvDt, '" + CStr(moCn.SrvrDate) + "') <= " + CStr(ADC("UoBDlDaysTo")))
  '*****Geeta*****Emr206********
  Dim ws_TdYyKeyStr As String, ws_InvFgdJn As String
  If UCase(ADC.MenuCd) = UCase("RepWipBPcs") Then
    wQty = IIF(ADC("UoBagPcs") = "N", " BQty as qQty", " BRjQty as qQty")
    wPtQty = IIF(ADC("UoBagPcs") = "N", " BPtQty as qPtQty", " BRjPtQty as qPtQty")
    wSelCnd = IIF(ADC("UoBagPcs") = "N", " where (BQty<> 0 or BPtQty<> 0 or (BQty= 0 and BPtQty= 0 and BRjQty= 0 and BRjPtQty= 0 and Round(BGrWt, 3)<> 0)) ", " where (BRjQty<> 0 or BRjPtQty<> 0) ")
    wSelCnd = wSelCnd + " and BLoc<> 'PFG' "
    If ADC("UoBagPcs") = "N" Then Rep.TxtHead.SetText "WIP Bag Pieces Report (Normal)"
    If ADC("UoBagPcs") = "R" Then Rep.TxtHead.SetText "WIP Bag Pieces Report (Rejected)"
    Rep.TxtLocFr.SetText ADC("UoIssLocFr")
    Rep.TxtLocTo.SetText ADC("UoIssLocTo")
    'Rep.TxtLocSel.SetText adc("UoIssLocSel") 22 Loc sel removed from header
    
    '*** Jay 2.13(CT) ***
    'Uma *** blank <> 'D
    'IG.88-BChr Fr,To & Sel added in Param Join to get BChr list
    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                    IIF(ADC("UoBagTyp") <> "", " and PValue= '" + ADC("UoBagTyp") + "' ", " and PValue <> 'D'") + _
                    IIF(ADC("UoBChr2Fr") <> "", " and PMCd>= '" + ADC("UoBChr2Fr") + "' ", " ") + _
                    IIF(ADC("UoBChr2To") <> "", " and PMCd<= '" + ADC("UoBChr2To") + "' ", " ") + _
                    IIF(ADC("UoBChr2Sel") <> "", " and PMCd IN (" + ADC("UoBChr2Sel") + ") ", " ")
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
  ElseIf UCase(ADC.MenuCd) = UCase("RepWipFgBPcs") Then
    wQty = " BQty as qQty"
    wPtQty = " BPtQty as qPtQty"
    '/*** Where had disappeared from 2.01 so and has replaces by 'where' again ***/
    'wSelCnd = " and BLoc= 'PFG' "
    wSelCnd = " where BLoc= 'PFG' "
    Rep.TxtHead.SetText "Fg Bag Pieces Report"
    Rep.TxthLocFr.SetText "Loc :": Rep.TxthLocTo.SetText "Bags In": 'Rep.TxthLocSel.SetText " "  22 Loc sel removed from header
    Rep.TxtLocFr.SetText "PFG"
    Rep.TxtLocTo.SetText IIF(ADC("UoYn2") = "Y", " FG ", IIF(ADC("UoYn2") = "N", " INV ", "FG & INV"))
    
    '*** 2.11 (Jen) (changed Jen 2.14 Next Patch 1)
    ws_FgJoin = " "
    ws_TdYyKeyStr = " BLstYyKey "
    '****** Sachin 3.02 s- Id fields in Joins
    If wi_PckLstPos > 0 Or ADC("UoFdPlYyFr") <> "" Or ADC("UoFdPlChrFr") <> "" Or ADC("UoFdPlNoFr") <> 0 Or _
       ADC("UoFdPlYyTo") <> "" Or ADC("UoFdPlChrTo") <> "" Or ADC("UoFdPlNoTo") <> 0 Then
      ws_FgJoin = " Join Fgd On FdBIdNo=BIdNo and FdCoCd= BCoCd and FdBYy= BYy and FdBChr= BChr and FdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=BPrtKey ", "") + _
                  "     and exists (select 'x' from Param where PTyp= 'TC' and PMCd= FdTc and PValue= 'FB') " + _
                  "     and (FdDespQty= 0 or " + _
                  "          exists (Select 'x' From InvFgd where IfFdIdNo=FdIdNo and IfCoCd= FdCoCd and IfFdTc= FdTc and IfFdYy= FdYy " + _
                  "                  and IfFdChr= FdChr and IfFdNo= FdNo and IfFdSr= FdSr and (Select PValue from Param Where PTyp='TC' and PMCd=IfTc)='" + ctInTcTyp + "' " + IIF(gs_Partition = ctCurrPrtn, " and IfPrtKey=FdPrtKey ", "") + ")) "
    End If
    
    '*** 2.11 (Jen) (changed Jen 2.14 Next Patch 1)
    
    If ADC("UoYn2") <> "" Then
        ws_FGINVChk = IIF(ADC("UoYn2") = "Y", " Not ", "") + _
            "exists (Select 'x' From InvFgd where IfBIdNo = BIdNo and " + _
            "IfFdBYy = BYy and IfFdBChr = BChr and IfFdBNo = BNo and IfCoCd = BCoCd)"
    End If
    
    '3 InvFgd join added to get InvNo. if InvNo scope or InvNo group sort selected then InvFgd join will be add
    If ADC("UoInTcFr") <> "" Or ADC("UoInYyFr") <> "" Or ADC("UoInChrFr") <> "" Or ADC("UoInNoFr") <> 0 Or _
       ADC("UoInTcTo") <> "" Or ADC("UoInYyTo") <> "" Or ADC("UoInChrTo") <> "" Or ADC("UoInNoTo") <> 0 Or wi_InvNoPos <> 0 Then
        ws_InvFgdJn = " left outer join InvFgd on IfBIdNo = BIdNo and IfFdBYy = BYy and IfFdBChr = BChr and IfFdBNo = BNo and IfCoCd = BCoCd and IfTc='IN' "
    End If
    
  'for Repair
  ElseIf UCase(ADC.MenuCd) = UCase("RprWIP") Then
    wQty = " BRepQty as qQty"
    'wQty = " OdOrdQty as qQty"
    wPtQty = " 0 as qPtQty"
    wSelCnd = " where BLoc= 'PFG' "
    If ADC("UoBagPcs") = "N" Then Rep.TxtHead.SetText "WIP Bag Pieces Report (Normal)"
    If ADC("UoBagPcs") = "R" Then Rep.TxtHead.SetText "WIP Bag Pieces Report (Rejected)"
    Rep.TxtLocFr.SetText ADC("UoIssLocFr")
    Rep.TxtLocTo.SetText ADC("UoIssLocTo")
    'Rep.TxtLocSel.SetText adc("UoIssLocSel")  22 Loc sel removed from header
        
    'IG.88-BChr Fr,To & Sel added in Param Join to get BChr list
    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                    IIF(ADC("UoBagTyp") <> "", " and PValue= '" + ADC("UoBagTyp") + "' ", " and PValue <> 'D'") + _
                    IIF(ADC("UoBChr2Fr") <> "", " and PMCd>= '" + ADC("UoBChr2Fr") + "' ", " ") + _
                    IIF(ADC("UoBChr2To") <> "", " and PMCd<= '" + ADC("UoBChr2To") + "' ", " ") + _
                    IIF(ADC("UoBChr2Sel") <> "", " and PMCd IN (" + ADC("UoBChr2Sel") + ") ", " ")
              
    ws_FgJoin = " "
    ws_TdYyKeyStr = " (Select Max(TdYyKey) From Txnd where TdCoCd=BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo and TdSrNo=0 and TdTc='BV') "
    
  End If
  '22 Customer From & To added to header
  Rep.TxtCustFr.SetText ADC("UoCmCdFr"): Rep.TxtCustTo.SetText ADC("UoCmCdTo")
  

  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(ADC("UOCMCDFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(ADC("UOCMCDTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(ADC("UOCMCDSEL")) + ")", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  'Rn.158 Wts for Wip / FG bags
  If UCase(ADC.MenuCd) = UCase("RepWipBPcs") Or UCase(ADC.MenuCd) = UCase("RepWipFgBPcs") Then
    Dim wAllWts As String, wDlyDays As String
    If ADC("wShowAllWts") = "Y" Then
      If UCase(ADC.MenuCd) = UCase("RepWipBPcs") Then
        wAllWts = "Isnull((select sum(TdRmWt *(Case when TdFrRmLoc= 'W' and TdFrRmDc= 'D' then 1 when TdFrRmLoc= 'W' and TdFrRmDc= 'C' then -1" + _
                  " when TdToRmLoc= 'W' and TdToRmDc= 'D' then 1 when TdToRmLoc= 'W' and TdToRmDc= 'C' then -1 Else 0 end) ) " + _
                  " from Txnd where TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo and TdSrNo<> 0 and " + _
                  " TdRmCtg in ('G', 'P', 'S', 'L', 'M') and TdPrtKey=BPrtKey ),0) as qMetWt, " + _
                  " Isnull((select sum(TdRmWt *(Case when TdFrRmLoc= 'W' and TdFrRmDc= 'D' then 1 when TdFrRmLoc= 'W' and TdFrRmDc= 'C' then -1 " + _
                  " when TdToRmLoc= 'W' and TdToRmDc= 'D' then 1 when TdToRmLoc= 'W' and TdToRmDc= 'C' then -1 Else 0 end) ) " + _
                  " from Txnd where TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo and TdSrNo<> 0 and " + _
                  " TdRmCtg in ('D','C') and TdPrtKey=BPrtKey ),0) as qDCWt, " + _
                  " Isnull((select sum(TdRmWt *(Case when TdFrRmLoc= 'W' and TdFrRmDc= 'D' then 1 when TdFrRmLoc= 'W' and TdFrRmDc= 'C' then -1 " + _
                  " when TdToRmLoc= 'W' and TdToRmDc= 'D' then 1 when TdToRmLoc= 'W' and TdToRmDc= 'C' then -1 Else 0 end) ) " + _
                  " from Txnd where TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo and TdSrNo<> 0 and " + _
                  " TdRmCtg in ('X') and TdPrtKey=BPrtKey ),0) as qXWt, " + _
                  " Isnull((select sum(((Case when TdFrRmLoc= 'W' and TdFrRmDc= 'D' then 1 when TdFrRmLoc= 'W' and TdFrRmDc= 'C' then -1  " + _
                  " when TdToRmLoc= 'W' and TdToRmDc= 'D' then 1 when TdToRmLoc= 'W' and TdToRmDc= 'C' then -1 Else 0 end) * TdRmWt ) * " + _
                  " (Case when TdRmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt else 0 end) / (case when TdRmCtg in ('G', 'P', 'S', 'L', 'M') then RmPure.qPurityWt else 1 end) ) " + _
                  " from Txnd join RmMst on RmCd= TdRmCd and RmPrtKey=TdPrtKey " + _
                  " Left Outer Join (Select PMCd as qPMCd, RmPurityWt as qPurityWt from Param join RmMst on RmCd= PValue where PTyp='RMCTG')RmPure On qPMCd= TdRmCtg " + _
                  " Where TdBIdNo = BIdNo And TdCoCd = BCoCd And TdBYy = BYy And TdBChr = BChr and TdBNo=BNo and TdSrNo<> 0 and TdPrtKey=BPrtKey ),0) as qPureWt "
                  
      Else
        '578-34 STR(FdKey, 7) changed to STR(FdKey, 8) in 8 place
        wAllWts = "Isnull((select sum(FrRmWt *(Case when FrFrRmLocTyp= 'F' and FrFrRmDc= 'D' then 1 when FrFrRmLocTyp= 'F' and FrFrRmDc= 'C' then -1 " + _
                  " when FrToRmLocTyp= 'F' and FrToRmDc= 'D' then 1 when FrToRmLocTyp= 'F' and FrToRmDc= 'C' then -1 Else 0 end) ) " + _
                  " from Fgd join FgRm on FrCoCd =FdCoCd and FrTc =FdTc and FrYy =FdYy and FrChr =FdChr and FrNo =FdNo and FrSr =FdSr and " + _
                  " FrFdIdNo =FdIdNo and FrPrtKey =FdPrtKey where FdBIdNo =BIdNo and FdCoCd= BCoCd and FdBYy=BYy and FdBChr=BChr and FdBNo=BNo and FdPrtKey=BPrtKey and exists(Select 'x' from Param where PTyp= 'TC' and PMCd= FdTc and PValue= 'FB') " + _
                  " and (FdYy + STR(FdKey, 8)) = (Select max(FdYy + STR(FdKey, 8)) from Fgd Fd1 Where Fd1.FdBIdNo=BIdNo and Fd1.FdCoCd = BCoCd And Fd1.FdBYy = BYy And Fd1.FdBChr = BChr " + _
                  " and Fd1.FdBNo= BNo and Fd1.FdPrtKey=BPrtKey ) and FrRmCtg in ('G', 'P', 'S', 'L', 'M') ),0) as qMetWt," + _
                  "Isnull((select sum(FrRmWt *(Case when FrFrRmLocTyp= 'F' and FrFrRmDc= 'D' then 1 when FrFrRmLocTyp= 'F' and FrFrRmDc= 'C' then -1 " + _
                  " when FrToRmLocTyp= 'F' and FrToRmDc= 'D' then 1 when FrToRmLocTyp= 'F' and FrToRmDc= 'C' then -1 Else 0 end) ) " + _
                  " from Fgd join FgRm on FrCoCd =FdCoCd and FrTc =FdTc and FrYy =FdYy and FrChr =FdChr and FrNo =FdNo and FrSr =FdSr and " + _
                  " FrFdIdNo =FdIdNo and FrPrtKey =FdPrtKey where FdBIdNo =BIdNo and FdCoCd= BCoCd and FdBYy=BYy and FdBChr=BChr and FdBNo=BNo and FdPrtKey=BPrtKey and exists(Select 'x' from Param where PTyp= 'TC' and PMCd= FdTc and PValue= 'FB') " + _
                  " and (FdYy + STR(FdKey, 8)) = (Select max(FdYy + STR(FdKey, 8)) from Fgd Fd1 Where Fd1.FdBIdNo=BIdNo and Fd1.FdCoCd = BCoCd And Fd1.FdBYy = BYy And Fd1.FdBChr = BChr " + _
                  " and Fd1.FdBNo= BNo and Fd1.FdPrtKey=BPrtKey ) and FrRmCtg in ('D','C') ),0) as qDCWt," + _
                  "Isnull((select sum(FrRmWt *(Case when FrFrRmLocTyp= 'F' and FrFrRmDc= 'D' then 1 when FrFrRmLocTyp= 'F' and FrFrRmDc= 'C' then -1 " + _
                  " when FrToRmLocTyp= 'F' and FrToRmDc= 'D' then 1 when FrToRmLocTyp= 'F' and FrToRmDc= 'C' then -1 Else 0 end) ) " + _
                  " from Fgd join FgRm on FrCoCd =FdCoCd and FrTc =FdTc and FrYy =FdYy and FrChr =FdChr and FrNo =FdNo and FrSr =FdSr and " + _
                  " FrFdIdNo =FdIdNo and FrPrtKey =FdPrtKey where FdBIdNo =BIdNo and FdCoCd= BCoCd and FdBYy=BYy and FdBChr=BChr and FdBNo=BNo and FdPrtKey=BPrtKey and exists(Select 'x' from Param where PTyp= 'TC' and PMCd= FdTc and PValue= 'FB') " + _
                  " and (FdYy + STR(FdKey, 8)) = (Select max(FdYy + STR(FdKey, 8)) from Fgd Fd1 Where Fd1.FdBIdNo=BIdNo and Fd1.FdCoCd = BCoCd And Fd1.FdBYy = BYy And Fd1.FdBChr = BChr " + _
                  " and Fd1.FdBNo= BNo and Fd1.FdPrtKey=BPrtKey ) and FrRmCtg in ('X') ),0) as qXWt, " + _
                  "Isnull((select sum(((Case when FrFrRmLocTyp= 'F' and FrFrRmDc= 'D' then 1 when FrFrRmLocTyp= 'F' and FrFrRmDc= 'C' then -1 " + _
                  " when FrToRmLocTyp= 'F' and FrToRmDc= 'D' then 1 when FrToRmLocTyp= 'F' and FrToRmDc= 'C' then -1 Else 0 end)* FrRmWt ) * " + _
                  " (Case when FrRmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt else 0 end) / (case when FrRmCtg in ('G', 'P', 'S', 'L', 'M') then RmPure.qPurityWt else 1 end) ) " + _
                  " from Fgd join FgRm on FrCoCd =FdCoCd and FrTc =FdTc and FrYy =FdYy and FrChr =FdChr and FrNo =FdNo and FrSr =FdSr and  FrFdIdNo =FdIdNo and FrPrtKey = FdPrtKey " + _
                  " join RmMst on RmCd= FrRmCd and RmPrtKey=FrPrtKey Left Outer Join (Select PMCd as qPMCd, RmPurityWt as qPurityWt from Param join RmMst on RmCd= PValue where PTyp='RMCTG')RmPure On qPMCd= FrRmCtg " + _
                  " where FdBIdNo =BIdNo and FdCoCd= BCoCd and FdBYy=BYy and FdBChr=BChr and FdBNo=BNo and FdPrtKey=BPrtKey and exists(Select 'x' from Param where PTyp= 'TC' and PMCd= FdTc and PValue= 'FB')  and (FdYy + STR(FdKey, 8)) = " + _
                  " (Select max(FdYy + STR(FdKey, 8)) from Fgd Fd1 Where Fd1.FdBIdNo=BIdNo and Fd1.FdCoCd = BCoCd And Fd1.FdBYy = BYy And Fd1.FdBChr = BChr  and Fd1.FdBNo= BNo and Fd1.FdPrtKey=BPrtKey ) ),0) as qPureWt "
                        
      End If
    Else
      wAllWts = "0.000 as qMetWt,0.000 as qDCWt,0.000 as qXWt, 0.000 as qPureWt "
    End If
    
    'Rn.158 calculating delay days and where condition for the same
    'if Yes then calculating Location wise delay,
    'if No, for FG report delay days based on FgSub location, for Wip report delay based on worker
    'if nothing then will not show delay days
    '578-29 UoYn3 changed as UoDlyDays. YN values changed as L,W
    If ADC("UoDlyDays") = "L" Then
      wDlyDays = "datediff(dd, BRecvDt, '" + CStr(moCn.SrvrDate) + "')"
      wBagDelayCnd = IIF(ADC("UoBDlDays") = 0, "", " and " + wDlyDays + " >= " + CStr(ADC("UoBDlDays")))
      wBagDelayCnd = wBagDelayCnd + IIF(ADC("UoBDlDaysTo") = 0, "", " and " + wDlyDays + " <= " + CStr(ADC("UoBDlDaysTo")))
    '578-29 UoYn3 changed as UoDlyDays
    ElseIf ADC("UoDlyDays") = "W" Then
      If UCase(ADC.MenuCd) = UCase("RepWipFgBPcs") Then
        wDlyDays = "datediff(dd, ((select max(FmDt) from Fmd join Fm on FmCoCd=FmdCoCd and FmTc=FmdTc and FmYy=FmdYy and FmChr =FmdChr and FmNo =FmdNo " + _
                  "where FmdCoCd= BCoCd and FmdBYy=BYy and FmdBChr=BChr and FmdBNo=BNo and FmdToFgSubLoc =BFgSubLoc )) , '" + CStr(moCn.SrvrDate) + "')"
      Else
        wDlyDays = "datediff(dd, (select min(TdDt) from Txnd where TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy=BYy and " + _
                 "TdBChr=BChr and TdBNo=BNo and TdSrNo= 0 and TdPrtKey=BPrtKey and TdFrBLoc =BLoc and TdWrk =BWrk) , '" + CStr(moCn.SrvrDate) + "')"
      End If
      wBagDelayCnd = IIF(ADC("UoBDlDays") = 0, "", " and " + wDlyDays + " >= " + CStr(ADC("UoBDlDays")))
      wBagDelayCnd = wBagDelayCnd + IIF(ADC("UoBDlDaysTo") = 0, "", " and " + wDlyDays + " <= " + CStr(ADC("UoBDlDaysTo")))
    '578-29 new option bag open date added
    ElseIf ADC("UoDlyDays") = "O" Then
      wDlyDays = "datediff(dd, BOpnDt, '" + CStr(moCn.SrvrDate) + "')"
      wBagDelayCnd = IIF(ADC("UoBDlDays") = 0, "", " and " + wDlyDays + " >= " + CStr(ADC("UoBDlDays")))
      wBagDelayCnd = wBagDelayCnd + IIF(ADC("UoBDlDaysTo") = 0, "", " and " + wDlyDays + " <= " + CStr(ADC("UoBDlDaysTo")))
      
    Else
      wDlyDays = 0
      wBagDelayCnd = ""
    End If
  End If
  
  
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
  If UCase(ADC.MenuCd) = UCase("RprWIP") Then
  'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp=BOdDmTcTyp in DsgMst join
  '11.d OdSfx added
  wSqlStrg = "Select " + grpflds + ", " + _
             "BCoCd, BYy, BChr, BNo, BRepOdTc as BOdTc,BRepOdYy as BOdYy ,BRepOdChr as BOdChr,BRepOdNo as BOdNo ,BRepOdSr as BOdSr," + _
             "DmTcTyp, DmSrcDsgCd, BOdDmCd, OdSfx, OdDmSz, DmCtg, OdKt, OdDmCol, OdPicNm, BRepLoc as BLoc, " + _
             IIF(OptLoc(0).Value = True, "LocPrntCd", "(Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs)") + " as LocPrntCd, BWrk, BRepGrWt as BGrWt, " + _
             "datediff(dd, BRecvDt, '" + CStr(moCn.SrvrDate) + "') as BDlDays, " + _
             wQty + "," + wFrVal + "," + wPtQty + ", OmCmCd, " + _
             "OmPoNo, LocPrcs, " + wAllWts + " From Bag " + ws_BagJoin + ws_FgJoin + _
             "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd and DmSz= ''  " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
             "Join Loc Lc on LocCocd= BCoCd and BRepLoc= LocCd " + _
             "Join OrdDsg on OdCoCd= BCoCd and OdTc= BRepOdTc and OdYy= BRepOdYy " + _
             "and OdChr= BRepOdChr and OdNo= BRepOdNo and OdSr= BRepOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
             "Join OrdMst on OmCoCd= BCoCd and OmTc= BRepOdTc and OmYy= BRepOdYy " + _
             "and OmChr= BRepOdChr and OmNo= BRepOdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=BPrtKey ", "") + _
             wSelCnd + IIF(ADC("UoYN") = "Y", " and BRepYN= 'Y' ", IIF(ADC("UoYN") = "N", " and BRepYN= 'N' ", " and BRepOdTc <>'' ")) + _
             wBagDelayCnd + _
             wCnd + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", "") + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
             IIF(ws_FGINVChk <> "", " And " + ws_FGINVChk, "")
  Else
  'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp=BOdDmTcTyp in DsgMst join
  '11.d OdSfx added
  '3 ws_InvFgdJn added
  wSqlStrg = "Select " + grpflds + ", " + _
             "BCoCd, BYy, BChr, BNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, DmTcTyp, DmSrcDsgCd, BOdDmCd, OdSfx, OdDmSz, DmCtg, " + _
             "OdKt, OdDmCol, OdPicNm," + IIF(OptLoc(0).Value = True, "BLoc", "LocPrcs") + " as BLoc, " + _
             IIF(OptLoc(0).Value = True, "LocPrntCd", "(Select L1.LocPrntCd From Loc L1 where L1.LocCoCd =BCoCd AND L1.LocCd =Lc.LocPrcs)") + " as LocPrntCd, BWrk, BGrWt, " + _
             wDlyDays + " as BDlDays, " + _
             wQty + "," + wFrVal + "," + wPtQty + ", OmCmCd, " + _
             "OmPoNo, LocPrcs, " + wAllWts + " From Bag " + ws_BagJoin + ws_FgJoin + _
             "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp=BOdDmTcTyp and DmCd= BOdDmCd and DmSz= ''  " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
             "Join Loc Lc on LocCocd= BCoCd and BLoc= LocCd " + _
             "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
             "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
             "Join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy " + _
             "and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=BPrtKey ", "") + _
             ws_InvFgdJn + _
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

  ADC.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  '***Geeta**Emr208 **Picture
  If ADC("UoYN1") = "N" Or wDetPos = 0 Then
    Rep.PICSEC.Suppress = True
  Else
    Rep.PICSEC.Suppress = False
  End If
  
  '***(Jen 2.11)
  '*** Jay 2.13(CT) ***
  'Uma *** new 'D' btypKK
  'IG.88-BChr Fr, To & Sel added in Bag Chr range to show in report
  If ADC("UoBagTyp") = "F" Then
    Rep.TxtBagTypAndChr.SetText "(Flute Bags" & IIF(ADC("UoBChr2Fr") <> "", ", FrChr = " & ADC("UoBChr2Fr"), "") & IIF(ADC("UoBChr2To") <> "", "  ToChr:" & ADC("UoBChr2To"), "") & ")"
  ElseIf ADC("UoBagTyp") = "P" Then
    Rep.TxtBagTypAndChr.SetText "(Primary Bags" & IIF(ADC("UoBChr2Fr") <> "", ", FrChr:" & ADC("UoBChr2Fr"), "") & IIF(ADC("UoBChr2To") <> "", "  ToChr:" & ADC("UoBChr2To"), "") & ")"
  ElseIf ADC("UoBagTyp") = "C" Then
    Rep.TxtBagTypAndChr.SetText "(Component Bags" & IIF(ADC("UoBChr2Fr") <> "", ", FrChr = " & ADC("UoBChr2Fr"), "") & IIF(ADC("UoBChr2To") <> "", "  ToChr:" & ADC("UoBChr2To"), "") & ")"
  ElseIf ADC("UoBagTyp") = "D" Then
    Rep.TxtBagTypAndChr.SetText "(PD Bags" & IIF(ADC("UoBChr2Fr") <> "", ", FrChr = " & ADC("UoBChr2Fr"), "") & IIF(ADC("UoBChr2To") <> "", "  ToChr:" & ADC("UoBChr2To"), "") & ")"
  End If
  'If adc("UoBagTyp") <> "" Then Rep.TxtBagTypAndChr.SetText IIF(adc("UoBagTyp") = "F", "(Flute Bags", "(Primary Bags") & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
  '*** Jay 2.13(CT) ***
  '***(Jen 2.11)
  'Rn.158 Show Wts=Yes then suppress the fields and section that are not related to weights
  'Added a new detail section for wts to avoid over writing wt fields with customer,Po details in detail section
  ' For this new detail section new heading section also created
  If ADC("wShowAllWts") = "Y" Then
    Rep.DET.Suppress = True
    Rep.PHDET.Suppress = True
    Rep.hSalVal.Suppress = True:  Rep.h1SalVal.Suppress = True
    Rep.gqSalVal1.Suppress = True:  Rep.g1qSalVal1.Suppress = True
    Rep.g2qSalVal1.Suppress = True: Rep.g3qSalVal1.Suppress = True
    If wDetPos = 0 Then
     Rep.PHDetWithWt.Suppress = True
     Rep.DetWithWt.Suppress = True
    End If
  
  Else
    'when Show Wt=No, not displaying wts fields
    Rep.DetWithWt.Suppress = True
    Rep.PHDetWithWt.Suppress = True
    Rep.qMetWt1.Suppress = True:  Rep.qPureWt1.Suppress = True:  Rep.qDCWt1.Suppress = True:   Rep.qXWt1.Suppress = True
    Rep.g1MetWt1.Suppress = True: Rep.g1PureWt1.Suppress = True: Rep.g1DCWt1.Suppress = True:  Rep.g1XWt1.Suppress = True
    Rep.g2MetWt1.Suppress = True: Rep.g2PureWt1.Suppress = True: Rep.g2DCWt1.Suppress = True:  Rep.g2XWt1.Suppress = True
    Rep.g3MetWt1.Suppress = True: Rep.g3PureWt1.Suppress = True: Rep.g3DCWt1.Suppress = True:  Rep.g3XWt1.Suppress = True
    Rep.hMetWt.Suppress = True:   Rep.hPureWt.Suppress = True:   Rep.hDCWt.Suppress = True:    Rep.hXWt.Suppress = True
    Rep.h1MetWt.Suppress = True:  Rep.h1PureWt.Suppress = True:  Rep.h1DCWt.Suppress = True:   Rep.h1XWt.Suppress = True
    
  End If
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******
  
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  'Rn.158 it is not allowing to show wt details within RTC time
  Case Is = 0     'Tag = "ShowAllWts"
    If .Value = Checked Then
      If WithinRTCRange("N") = False Then
        .Value = Unchecked: .Enabled = False
        ADC("wShowAllWts") = "N"
        MsgBox "Weights Cannot Be Taken During RTC Time. " & vbNewLine & "Contact Your EMR Admin."
        Exit Sub
      End If
      ADC("wShowAllWts") = "Y"
    Else
      ADC("wShowAllWts") = "N"
    End If
      
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

  '  Set adc.FirNKeyCtl = adc("UoCoCdFr")
  If UCase(ADC.MenuCd) = UCase("RepWipBPcs") Or UCase(ADC.MenuCd) = UCase("RprWIP") Then
    
    'adc("UoIssLocFr").CmpStr = "BLoc >= "
    'adc("UoIssLocTo").CmpStr = "BLoc <= "
    'adc("UoIssLocSel").CmpStr = "BLoc In "
    LblProcess.Visible = True
    
    FraLoc.Visible = True: FraLoc.Enabled = True
    
    ADC("UoIssLocFr").Visible = True: ADC("UoIssLocFr").Enabled = True
    ADC("UoIssLocTo").Visible = True: ADC("UoIssLocTo").Enabled = True
    ADC("UoIssLocSel").Visible = True: ADC("UoIssLocSel").Enabled = True
    ADC("UoFgSubLocFr").CmpStr = ""
    ADC("UoFgSubLocTo").CmpStr = ""
    ADC("UoFgSubLocSel").CmpStr = ""
    lblFgSubLoc.Visible = False
    ADC("UoFgSubLocFr").Visible = False: ADC("UoFgSubLocFr").Enabled = False
    ADC("UoFgSubLocTo").Visible = False: ADC("UoFgSubLocTo").Enabled = False
    ADC("UoFgSubLocSel").Visible = False: ADC("UoFgSubLocSel").Enabled = False
    '**********Bhavna src dsg cd for Fg bag pcs
    ADC("UoBOEFr").CmpStr = ""
    ADC("UoBOETo").CmpStr = ""
    LblSrcDsgCd.Visible = False
    ADC("UoBOEFr").Visible = False: ADC("UoBOEFr").Enabled = False
    ADC("UoBOETo").Visible = False: ADC("UoBOETo").Enabled = False
    '**********Bhavna src dsg cd for Fg bag pcs
    
    '*** (2.11) Jen
    LblPckLst.Visible = False
    ADC("UoFdPlYyFr").Visible = False: ADC("UoFdPlChrFr").Visible = False: ADC("UoFdPlNoFr").Visible = False
    ADC("UoFdPlYyTo").Visible = False: ADC("UoFdPlChrTo").Visible = False: ADC("UoFdPlNoTo").Visible = False
    LblPkLstDeLimit1.Visible = False: LblPkLstDeLimit2.Visible = False: LblPkLstDeLimit3.Visible = False: LblPkLstDeLimit4.Visible = False
    ADC("UoFdPlYyFr").CmpStr = "": ADC("UoFdPlChrFr").CmpStr = "": ADC("UoFdPlNoFr").CmpStr = ""
    ADC("UoFdPlYyTo").CmpStr = "": ADC("UoFdPlChrTo").CmpStr = "": ADC("UoFdPlNoTo").CmpStr = ""
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
      ADC("UoPmCdFr").Visible = False
      TxtPwd.Visible = False
      lblseopwd.Visible = False
      LblValTyp.Visible = False
    ' ****** sachin 2.12 - 24/11/05 ****** '
    ' Zubin 213
    ADC("UoCurCdFr").Visible = False
    LblCurrCd.Visible = False
    ' Zubin 213
    
  ElseIf UCase(ADC.MenuCd) = UCase("RepWipFgBPcs") Then
    ' **** Zubin 211 **** '
    '''''    Set adc.FirNKeyCtl = adc("UoWrkFr")
    'Set adc.FirNKeyCtl = adc("UoFgSubLocFr")
    
    ' **** Zubin 211 **** '
    ADC("UoIssLocFr").CmpStr = ""
    ADC("UoIssLocTo").CmpStr = ""
    ADC("UoIssLocSel").CmpStr = ""
    LblProcess.Visible = False
    
    FraLoc.Visible = False: FraLoc.Enabled = False
    ADC("UoIssLocFr").Visible = False: ADC("UoIssLocFr").Enabled = False
    ADC("UoIssLocTo").Visible = False: ADC("UoIssLocTo").Enabled = False
    ADC("UoIssLocSel").Visible = False: ADC("UoIssLocSel").Enabled = False
    lblFgSubLoc.Visible = True
    ADC("UoFgSubLocFr").Visible = True: ADC("UoFgSubLocFr").Enabled = True
    ADC("UoFgSubLocTo").Visible = True: ADC("UoFgSubLocTo").Enabled = True
    ADC("UoFgSubLocSel").Visible = True: ADC("UoFgSubLocSel").Enabled = True

    '*** (2.11) Jen
    LblPckLst.Visible = True
    ADC("UoFdPlYyFr").Visible = True: ADC("UoFdPlChrFr").Visible = True: ADC("UoFdPlNoFr").Visible = True
    ADC("UoFdPlYyTo").Visible = True: ADC("UoFdPlChrTo").Visible = True: ADC("UoFdPlNoTo").Visible = True
    LblPkLstDeLimit1.Visible = True: LblPkLstDeLimit2.Visible = True: LblPkLstDeLimit3.Visible = True: LblPkLstDeLimit4.Visible = True
    
    ADC("UoFdPlYyFr").CmpStr = "FdPlYy >=": ADC("UoFdPlChrFr").CmpStr = "FdPlChr >=": ADC("UoFdPlNoFr").CmpStr = "FdPlNo >="
    ADC("UoFdPlYyTo").CmpStr = "FdPlYy <=": ADC("UoFdPlChrTo").CmpStr = "FdPlChr <=": ADC("UoFdPlNoTo").CmpStr = "FdPlNo <="
    ADC("UoPrntCdFr").CmpStr = "LocPrntCd <=": ADC("UoPrntCdTo").CmpStr = "LocPrntCd <=": ADC("UoPrntCdSel").CmpStr = "LocPrntCd In "
    '*** (2.11) Jen
    
    '**********Bhavna src dsg cd for Fg bag pcs
    ADC("UoBOEFr").CmpStr = "DmSrcDsgCd >= "
    ADC("UoBOETo").CmpStr = "DmSrcDsgCd <= "
    LblSrcDsgCd.Visible = True
    ADC("UoBOEFr").Visible = True: ADC("UoBOEFr").Enabled = True
    ADC("UoBOETo").Visible = True: ADC("UoBOETo").Enabled = True
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
    ADC("UoBagPcs").Visible = False: ADC("UoBagPcs").Enabled = False
    ADC("UoBagPcs") = ""
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
    FraOptBagPcs.Enabled = False: FraOptBagPcs.Visible = False
    OptBagPcs(0).Value = False: OptBagPcs(1).Value = False
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
    
    FraOptBagIn.Enabled = True: FraOptBagIn.Visible = True
    LblBagIn.Visible = True
    
    LblBagTyp.Visible = False
    ADC("UoBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    ADC("UoBChr2Fr").Visible = False
    ADC("UoBChr2To").Visible = False 'IG.88- BChr not in FG WIP
    ADC("UoBChr2Sel").Visible = False
    
    '****** Sachin 2.14.0 - [Multi-Company]
    LblPrntCd.Visible = False
    'adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
    ADC("UoPrntCdFr").Visible = False: ADC("UoPrntCdTo").Visible = False: ADC("UoPrntCdSel").Visible = False
    ADC("UoPrntCdFr").CmpStr = "": ADC("UoPrntCdTo").CmpStr = "": ADC("UoPrntCdSel").CmpStr = ""
    '****** Sachin 2.14.0 - [Multi-Company]
    '3 InvNo scope visible only for fg bag pcs
    FraInvNo.Visible = True
    '578-29 bag open date is not visible for fg bag pcs
    OptDlyDays(3).Visible = False
  End If
  
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_LocTyp = "'P'"
  'gb_RmCtgFor = True
  'gb_RmZ = False
  
  'gs_TxnTcTyp = "BV"
  'indigo.164 sketch design tctype and order tc added
  gs_OmTcTyp = "SO','OS"
  gs_DmTcTyp = "DM','SM"

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
              "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + ADC.MenuCd + "' and " + _
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
    'indigo.164 sketch design tctype and order tc added
    gs_OmTcTyp = "SO','OS"
    gs_DmTcTyp = "DM','SM"

  '*** (Jen 2.11)
  Select Case UCase(IdName)
    'IG.88- BChr Fr, TO & Sel for help added
    Case Is = UCase("UoBChr2Fr"), UCase("UoBChr2To")
      If ADC.Mode = xNorm Then
        If ADC("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Bag Type.": Exit Sub
        HlpList.PMCd "BCHR", "'" + ADC("UoBagTyp") + "'": Exit Sub
      End If
    Case Is = UCase("UoBChr2Sel")
      If ADC.Mode = xNorm Then
        If ADC("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Bag Type.": Exit Sub
        HlpList.MultiSelect = True
        HlpList.PMCd "BCHR", "'" + ADC("UoBagTyp") + "'": Exit Sub
      End If
                      
    Case Is = UCase("UOISSLOCFR"), UCase("UOISSLOCTO"), UCase("UOISSLOCSEL"), UCase("UOPRNTCDFR"), UCase("UOPRNTCDTO"), UCase("UOPRNTCDSEL")
        If ADC("UOLOCTYPFR") = "R" Then gs_LocTyp = "'R'"
        If ADC("UOLOCTYPFR") = "P" Then gs_LocTyp = "'P'"

    Case Is = UCase("UOOMTCFR")
          If UCase(ADC.MenuCd) = UCase("RprWIP") Then HlpList.PMCd "TC", "'JT'":  Exit Sub
    Case Is = UCase("UOOMTCTO")
          If UCase(ADC.MenuCd) = UCase("RprWIP") Then HlpList.PMCd "TC", "'JT'":  Exit Sub
        
  End Select
  '*** (Jen 2.11)
  
 SetProp Me, IdName, When
   
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
' **** Zubin 211 **** '
  Select Case UCase(IdName)
'    Case Is = UCase("UoBChrFr")
      'If ADC("UoBagTyp") = "N" Then
    Case Is = UCase("UoBChr2Fr")
      ADC("UoBChr2Fr") = ""
    Case Is = UCase("UoBChr2To")
      ADC("UoBChr2To") = ""
    Case Is = UCase("UoBChr2Sel")
      ADC("UoBChr2Sel") = ""
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
  ADC("UoYN") = "Y"
Case Is = 1
  ADC("UoYN") = "N"
Case Is = 2
  ADC("UoYN") = ""
End Select
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptBagIn_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN2") = "Y"
Case Is = 1
  ADC("UoYN2") = "N"
Case Is = 2
  ADC("UoYN2") = ""
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

Private Sub OptLocation_Click(Index As Integer)

End Sub
'Rn.158 getting values as per option selected for delay days
Private Sub OptDlyDays_Click(Index As Integer)
'578-29 UoYn3 changed as UoDlyDays. YN,blank values modified as L,W,N
'case 3 is newly added for bag open date
  Select Case Index
  Case Is = 0
    'adc("UoDlyDays") = ""
    ADC("UoDlyDays") = "N"
  Case Is = 1
    'adc("UoDlyDays") = "Y"
    ADC("UoDlyDays") = "L"
  Case Is = 2
    'adc("UoDlyDays") = "N"
    ADC("UoDlyDays") = "W"
  Case Is = 3
    ADC("UoDlyDays") = "O"
  End Select
End Sub
'Rn.158 tooltip for delay days
Private Sub OptDlyDays_GotFocus(Index As Integer)
  DispMsg FraDelayDays.ToolTipText, etInfo
End Sub

Private Sub OptLoc_GotFocus(Index As Integer)
  DispMsg FraLoc.ToolTipText, etInfo
Select Case Index
Case Is = 0
  ADC("UOISSLOCFR") = "":   ADC("UOISSLOCTO") = "":   ADC("UOISSLOCsel") = ""
  ADC("UOPRNTCDFR") = "":   ADC("UOPRNTCDTO") = "":   ADC("UOPRNTCDSEL") = ""
  ADC("UoLocTypFr") = "P"
Case Is = 1
  ADC("UOISSLOCFR") = "":   ADC("UOISSLOCTO") = "":   ADC("UOISSLOCsel") = ""
  ADC("UOPRNTCDFR") = "":   ADC("UOPRNTCDTO") = "":   ADC("UOPRNTCDSEL") = ""
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
  If ws_MultiCoMod = False Then _
      ADC("UoCoCdFr") = gs_CoCd: ADC("UoCoCdTo") = gs_CoCd: ADC("UoCoCdSel") = "": ADC("UoCoCdFr").Enabled = False
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
  '**Geeta*Emr208*Default Show picture is always no
  ADC("UoYN1") = "N"
  ' ****** sachin 2.12 ****** '
  ADC("UoPmCdFr") = ""
  ' ****** sachin 2.12 ****** '

  '*** Jay 2.13(CT) ***
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + gs_CoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    ADC("UoBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    ADC("UoBChr2Fr").Visible = False: ADC("UoBChr2To").Visible = False: ADC("UoBChr2Sel").Visible = False
    ADC("UoBagTyp") = "P": ADC("UoBChr2Fr") = "": ADC("UoBChr2To") = "": ADC("UoBChr2Sel") = ""
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
  If ADC("UoYN") = "Y" Then
    OptBag(0).Value = True
  ElseIf ADC("UoYN") = "N" Then
    OptBag(1).Value = True
  ElseIf ADC("UoYN") = "" Then
    OptBag(2).Value = True
  End If
  
  If ADC("UoYN2") = "Y" Then
    OptBagIn(0).Value = True
  ElseIf ADC("UoYN2") = "N" Then
    OptBagIn(1).Value = True
  ElseIf ADC("UoYN2") = "" Then
    OptBagIn(2).Value = True
  End If
  
  If UCase(ADC.MenuCd) = UCase("RepWipBPcs") Then
    If ADC("UoBagPcs") = "N" Then OptBagPcs(0).Value = True
    If ADC("UoBagPcs") = "R" Then OptBagPcs(1).Value = True
  End If
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
  'Rn.158 default value for Show wt=No
  ADC("wShowAllWts") = "N"
  '578-29 UoYn3 changed as UoDlyDays. YN,blank values modified as L,W,N
  'OptDlyDays(3) is newly added for bag open date
  'If adc("UoDlyDays") = "N" Then
  If ADC("UoDlyDays") = "W" Then
    OptDlyDays(2).Value = True
  'ElseIf adc("UoDlyDays") = "Y" Then
  ElseIf ADC("UoDlyDays") = "L" Then
    OptDlyDays(1).Value = True
  'ElseIf adc("UoDlyDays") = "" Then
  ElseIf ADC("UoDlyDays") = "N" Then
    OptDlyDays(0).Value = True
  ElseIf ADC("UoDlyDays") = "O" Then
    OptDlyDays(3).Value = True
  End If
  '578-29 if delay days is blank then default value set as Not required
  If ADC("UoDlyDays") = "" Then OptDlyDays(0).Value = True: ADC("UoDlyDays") = "N"
  
  'Rn.158 Caption changed for Fg report
  If UCase(ADC.MenuCd) = UCase("RepWipFgBPcs") Then
    OptDlyDays(1).Caption = "In FG"
    OptDlyDays(2).Caption = "In FG SubLoc"
  End If
  
End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = False Or UCase(ADC.MenuCd) = UCase("RepWipFgBPcs"), ADC("UOCOCDFR"), gs_CoCd) + "' and " + _
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
  If ws_MultiCoMod = False Or UCase(ADC.MenuCd) = UCase("RepWipFgBPcs") Then
      Rep.TxtCoCdTo.Suppress = True ': Rep.TxtCoCdSel.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True ': Rep.TxtCoCdSelLbl.Suppress = True
  Else
      Rep.TxtCoCdToLbl.Suppress = False ': Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdTo.SetText ADC("UOCOCDTO") ': Rep.TxtCoCdSel.SetText ADC("UOCOCDSEL")
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
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ")
  If Trim(TxtPwd.text) = "" And ws_VldPwd <> "" Then ADC("UoPMCdFr") = "": Exit Sub
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
