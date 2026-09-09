VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpGldLs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gold Loss"
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
      Left            =   11280
      TabIndex        =   90
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
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5235
      TabIndex        =   95
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   96
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
         TabIndex        =   97
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
      Left            =   6450
      TabIndex        =   94
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   93
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
      TabIndex        =   91
      Top             =   9660
      Width           =   15180
      _ExtentX        =   26776
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
      TabIndex        =   92
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
         TabPicture(0)   =   "EmrFrpGldLs.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpGldLs.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   101
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraFgScp 
               BorderStyle     =   0  'None
               Height          =   1755
               Left            =   0
               TabIndex        =   131
               Top             =   4860
               Width           =   15075
               Begin VB.Frame FraOptBag 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   147
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   1470
                  Width           =   3705
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
                     Left            =   1290
                     TabIndex        =   80
                     ToolTipText     =   "Show Open Or Closed Or All Bags"
                     Top             =   0
                     Width           =   945
                  End
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
                     TabIndex        =   79
                     ToolTipText     =   "Show Open Or Closed Or All Bags"
                     Top             =   0
                     Width           =   1005
                  End
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   42
                  Left            =   5055
                  TabIndex        =   62
                  ToolTipText     =   "Enter To Invoice Year"
                  Top             =   0
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataType        =   4
                  DataField       =   "UOINYYTO"
                  IdName          =   "UOINYYTO"
                  CmpStr          =   "IfYy <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   43
                  Left            =   2385
                  TabIndex        =   59
                  ToolTipText     =   "Enter From Invoice Year"
                  Top             =   0
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataType        =   4
                  DataField       =   "UOINYYFR"
                  IdName          =   "UOINYYFR"
                  CmpStr          =   "IfYy >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   44
                  Left            =   1605
                  TabIndex        =   58
                  ToolTipText     =   "Enter From Invoice Tc"
                  Top             =   0
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataType        =   4
                  DataField       =   "UOINTCFR"
                  IdName          =   "UOINTCFR"
                  CmpStr          =   "IfTc >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   45
                  Left            =   2805
                  TabIndex        =   60
                  ToolTipText     =   "Enter From Invoice Character"
                  Top             =   0
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataType        =   4
                  DataField       =   "UOINCHRFR"
                  IdName          =   "UOINCHRFR"
                  CmpStr          =   "IfChr >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   46
                  Left            =   4275
                  TabIndex        =   61
                  ToolTipText     =   "Enter To Invoice Tc"
                  Top             =   0
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataType        =   4
                  DataField       =   "UOINTCTO"
                  IdName          =   "UOINTCTO"
                  CmpStr          =   "IfTc <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   47
                  Left            =   5475
                  TabIndex        =   63
                  ToolTipText     =   "Enter To Invoice Character"
                  Top             =   0
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataType        =   4
                  DataField       =   "UOINCHRTO"
                  IdName          =   "UOINCHRTO"
                  CmpStr          =   "IfChr <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   50
                  Left            =   1605
                  TabIndex        =   64
                  ToolTipText     =   "Enter From Invoice Number"
                  Top             =   285
                  Width           =   825
                  _ExtentX        =   1455
                  _ExtentY        =   503
                  MaxLength       =   0
                  DataType        =   1
                  DataField       =   "UOINNOFR"
                  IdName          =   "UOINNOFR"
                  CmpStr          =   "IfNo >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   52
                  Left            =   4275
                  TabIndex        =   65
                  ToolTipText     =   "Enter To Invoice Number"
                  Top             =   285
                  Width           =   825
                  _ExtentX        =   1455
                  _ExtentY        =   503
                  MaxLength       =   0
                  DataType        =   1
                  DataField       =   "UOINNOTO"
                  IdName          =   "UOINNOTO"
                  CmpStr          =   "IfNo <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   53
                  Left            =   1605
                  TabIndex        =   66
                  ToolTipText     =   "Enter From Invoice Date"
                  Top             =   570
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   3
                  DataField       =   "UODT1FR"
                  IdName          =   "UODT1FR"
                  CmpStr          =   "InDt >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   54
                  Left            =   4275
                  TabIndex        =   67
                  ToolTipText     =   "Enter To Invoice Date"
                  Top             =   570
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   3
                  DataField       =   "UODT1TO"
                  IdName          =   "UODT1TO"
                  CmpStr          =   "InDt <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   55
                  Left            =   5055
                  TabIndex        =   72
                  ToolTipText     =   "Enter To Finished Goods Year"
                  Top             =   870
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
                  Index           =   56
                  Left            =   2385
                  TabIndex        =   69
                  ToolTipText     =   "Enter From Finished Goods Year"
                  Top             =   870
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
                  Index           =   57
                  Left            =   1605
                  TabIndex        =   68
                  ToolTipText     =   "Enter From Finished Goods Tc"
                  Top             =   870
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
                  Index           =   58
                  Left            =   2805
                  TabIndex        =   70
                  ToolTipText     =   "Enter From Finished Goods Character"
                  Top             =   870
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
                  Index           =   59
                  Left            =   4275
                  TabIndex        =   71
                  ToolTipText     =   "Enter To Finished Goods Tc"
                  Top             =   870
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
                  Index           =   60
                  Left            =   5475
                  TabIndex        =   73
                  ToolTipText     =   "Enter To Finished Goods Character"
                  Top             =   870
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
                  Index           =   61
                  Left            =   5175
                  TabIndex        =   77
                  ToolTipText     =   "Enter To Finished Goods Serial"
                  Top             =   1155
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
                  Index           =   62
                  Left            =   2505
                  TabIndex        =   75
                  ToolTipText     =   "Enter From Finished Goods Serial"
                  Top             =   1155
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
                  Index           =   63
                  Left            =   1605
                  TabIndex        =   74
                  ToolTipText     =   "Enter From Finished Goods Number"
                  Top             =   1155
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
                  Index           =   64
                  Left            =   4275
                  TabIndex        =   76
                  ToolTipText     =   "Enter To Finished Goods Number"
                  Top             =   1155
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
                  Index           =   65
                  Left            =   5400
                  TabIndex        =   78
                  ToolTipText     =   "Show Open Or Closed Or All Bags ? (Open - 'Y', Closed - 'N', All - Blank)"
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  DataField       =   "UOYN4"
                  IdName          =   "UOYN4"
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
                  TabIndex        =   148
                  Top             =   1440
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
                  Index           =   35
                  Left            =   2445
                  TabIndex        =   146
                  Top             =   1155
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
                  Left            =   5115
                  TabIndex        =   145
                  Top             =   1155
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
                  Left            =   2325
                  TabIndex        =   144
                  Top             =   870
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
                  Left            =   4995
                  TabIndex        =   143
                  Top             =   870
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
                  Index           =   10
                  Left            =   2745
                  TabIndex        =   142
                  Top             =   870
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
                  Left            =   5415
                  TabIndex        =   141
                  Top             =   870
                  Width           =   105
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
                  Index           =   5
                  Left            =   60
                  TabIndex        =   140
                  Top             =   1170
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
                  TabIndex        =   139
                  Top             =   885
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
                  Index           =   19
                  Left            =   2325
                  TabIndex        =   138
                  Top             =   0
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
                  Left            =   4995
                  TabIndex        =   137
                  Top             =   0
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
                  Left            =   2745
                  TabIndex        =   136
                  Top             =   0
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
                  Index           =   23
                  Left            =   5415
                  TabIndex        =   135
                  Top             =   0
                  Width           =   105
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
                  Index           =   28
                  Left            =   60
                  TabIndex        =   134
                  Top             =   285
                  Width           =   1245
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
                  Index           =   30
                  Left            =   60
                  TabIndex        =   133
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.Label LblInDt 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Invoice Date"
                  BeginProperty Font 
                     Name            =   "Arial"
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
                  Top             =   570
                  Width           =   1605
               End
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
               Left            =   4260
               TabIndex        =   150
               ToolTipText     =   "Select Production or Process"
               Top             =   1440
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
               Left            =   1560
               TabIndex        =   149
               ToolTipText     =   "Select Production or Process"
               Top             =   1440
               Value           =   -1  'True
               Width           =   2085
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Allowed Loss %     "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Index           =   5
               Left            =   30
               TabIndex        =   57
               Tag             =   "ShowLs"
               ToolTipText     =   "Check To Allowed Loss % fro given Process"
               Top             =   5490
               Width           =   1770
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Loss %     "
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
               Left            =   20
               TabIndex        =   87
               Tag             =   "ShowLs"
               ToolTipText     =   "Enter SEO Pwd and Check To Show Ls %"
               Top             =   7275
               Width           =   1790
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Box Det            "
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
               TabIndex        =   83
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show Box Details"
               Top             =   6650
               Width           =   1770
            End
            Begin VB.Frame FraOptIOWt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1605
               TabIndex        =   130
               ToolTipText     =   "Select if to show In Wt Or Out Wt"
               Top             =   6960
               Width           =   3855
               Begin VB.OptionButton OptIOWt 
                  Caption         =   "Out Wt"
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
                  Left            =   2700
                  TabIndex        =   85
                  Top             =   0
                  Width           =   1725
               End
               Begin VB.OptionButton OptIOWt 
                  Caption         =   "In Wt"
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
                  Top             =   0
                  Width           =   1725
               End
            End
            Begin VB.Frame FraOptRepFmt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1605
               TabIndex        =   129
               ToolTipText     =   "Select Report Format based on recovery with Dust or Actual Or Recovery not rEQUIRED"
               Top             =   5190
               Width           =   7665
               Begin VB.OptionButton OptRepFmt 
                  Caption         =   "Recovery Not Reqd"
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
                  Left            =   5280
                  TabIndex        =   55
                  Top             =   0
                  Width           =   2175
               End
               Begin VB.OptionButton OptRepFmt 
                  Caption         =   "Recovery On Actual"
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
                  Left            =   2640
                  TabIndex        =   54
                  Top             =   0
                  Width           =   2175
               End
               Begin VB.OptionButton OptRepFmt 
                  Caption         =   "Recovery With Dust"
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
                  TabIndex        =   53
                  Top             =   0
                  Width           =   2175
               End
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1605
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   88
               Top             =   7680
               Width           =   2340
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   1605
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   225
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "Td.TdCoCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   1605
               TabIndex        =   3
               ToolTipText     =   "Enter From Transaction Date"
               Top             =   510
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "Td.TdDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   4275
               TabIndex        =   4
               ToolTipText     =   "Enter To Transaction Date"
               Top             =   510
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "Td.TdDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   6930
               TabIndex        =   5
               ToolTipText     =   "Enter Transaction Date Selection"
               Top             =   510
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTSEL"
               IdName          =   "UODTSEL"
               CmpStr          =   "TDt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   1605
               TabIndex        =   6
               ToolTipText     =   "Enter Raw Material Category"
               Top             =   795
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "Rm.RmCtg>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1605
               TabIndex        =   12
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   1725
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4275
               TabIndex        =   13
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   1725
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   6930
               TabIndex        =   14
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1725
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1605
               TabIndex        =   24
               ToolTipText     =   "Enter From Design Code"
               Top             =   2865
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
               Index           =   19
               Left            =   4275
               TabIndex        =   25
               ToolTipText     =   "Enter To Design Code"
               Top             =   2865
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
               Index           =   20
               Left            =   6930
               TabIndex        =   26
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2865
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
               Index           =   12
               Left            =   1605
               TabIndex        =   18
               ToolTipText     =   "Enter From Design Category"
               Top             =   2295
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
               Index           =   13
               Left            =   4275
               TabIndex        =   19
               ToolTipText     =   "Enter To Design Category"
               Top             =   2295
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
               Index           =   14
               Left            =   6930
               TabIndex        =   20
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2295
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
               Index           =   32
               Left            =   1605
               TabIndex        =   38
               ToolTipText     =   "Enter From Customer"
               Top             =   4005
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OdOmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   4275
               TabIndex        =   39
               ToolTipText     =   "Enter To Customer"
               Top             =   4005
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "OdOmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   6930
               TabIndex        =   40
               ToolTipText     =   "Enter Customer Selection"
               Top             =   4005
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OdOmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   2385
               TabIndex        =   31
               ToolTipText     =   "Enter From Order Year"
               Top             =   3435
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
               Index           =   28
               Left            =   5055
               TabIndex        =   34
               ToolTipText     =   "Enter To Order Year"
               Top             =   3435
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
               Index           =   24
               Left            =   1605
               TabIndex        =   30
               ToolTipText     =   "Enter From Order Tc"
               Top             =   3435
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
               Index           =   27
               Left            =   4275
               TabIndex        =   33
               ToolTipText     =   "Enter To Order Tc"
               Top             =   3435
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
               Index           =   26
               Left            =   2805
               TabIndex        =   32
               ToolTipText     =   "Enter From Order Character"
               Top             =   3435
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
               Index           =   29
               Left            =   5475
               TabIndex        =   35
               ToolTipText     =   "Enter To Order Character"
               Top             =   3435
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
               Index           =   30
               Left            =   1605
               TabIndex        =   36
               ToolTipText     =   "Enter From Order Number"
               Top             =   3720
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
               Index           =   31
               Left            =   4275
               TabIndex        =   37
               ToolTipText     =   "Enter To Order Number"
               Top             =   3720
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
               Index           =   9
               Left            =   1605
               TabIndex        =   15
               ToolTipText     =   "Enter From By Worker"
               Top             =   2010
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKFR"
               IdName          =   "UOWRKFR"
               CmpStr          =   "Td.TdByWrk >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   4275
               TabIndex        =   16
               ToolTipText     =   "Enter To By Worker"
               Top             =   2010
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKTO"
               IdName          =   "UOWRKTO"
               CmpStr          =   "Td.TdByWrk <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   6930
               TabIndex        =   17
               ToolTipText     =   "Enter By Worker Selection"
               Top             =   2010
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
               Index           =   21
               Left            =   1605
               TabIndex        =   27
               ToolTipText     =   "Enter From Karat"
               Top             =   3150
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
               Index           =   22
               Left            =   4275
               TabIndex        =   28
               ToolTipText     =   "Enter To Karat"
               Top             =   3150
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
               Index           =   23
               Left            =   6930
               TabIndex        =   29
               ToolTipText     =   "Enter Karat Selection"
               Top             =   3150
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
               Index           =   35
               Left            =   1605
               TabIndex        =   41
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   4290
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
               Index           =   36
               Left            =   4275
               TabIndex        =   42
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   4290
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
               Index           =   37
               Left            =   6930
               TabIndex        =   43
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   4290
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
               Index           =   15
               Left            =   1605
               TabIndex        =   21
               ToolTipText     =   "Enter From Design Production Category"
               Top             =   2580
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
               Index           =   16
               Left            =   4275
               TabIndex        =   22
               ToolTipText     =   "Enter To Design Production Category"
               Top             =   2580
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
               Index           =   17
               Left            =   6930
               TabIndex        =   23
               ToolTipText     =   "Enter Design Production Category Selection"
               Top             =   2580
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
               Index           =   38
               Left            =   2265
               TabIndex        =   51
               ToolTipText     =   "Show Box Detail (Yes / No)"
               Top             =   6600
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   9600
               TabIndex        =   52
               ToolTipText     =   "Enter 'Y' For Recovery With Dust, 'N' For Recovery On Actuals, Blank For Recovery Not Required"
               Top             =   5190
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
               Left            =   6885
               TabIndex        =   82
               ToolTipText     =   "Enter 'Y' For In Wt, 'N' For Out Wt"
               Top             =   6960
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
               Height          =   285
               Index           =   40
               Left            =   1605
               TabIndex        =   49
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   4860
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
               Index           =   77
               Left            =   3255
               TabIndex        =   50
               ToolTipText     =   "Enter Bag Character"
               Top             =   4860
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
               Index           =   73
               Left            =   4275
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   225
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "Td.TdCoCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   6930
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   225
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "Td.TdCoCd in"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   4275
               TabIndex        =   10
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1080
               Width           =   1700
               _ExtentX        =   2990
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDTO"
               IdName          =   "UOPRNTCDTO"
               CmpStr          =   "Lc.LocPrntCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   6930
               TabIndex        =   11
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1080
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
               CmpStr          =   "Lc.LocPrntCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   1605
               TabIndex        =   9
               ToolTipText     =   "Enter From Parent Code"
               Top             =   1080
               Width           =   1700
               _ExtentX        =   2990
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDFR"
               IdName          =   "UOPRNTCDFR"
               CmpStr          =   "Lc.LocPrntCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   2655
               TabIndex        =   86
               ToolTipText     =   "Enter 'Y' For Viewing Loss %, 'N' For Not Considering Loss %"
               Top             =   7275
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
               Index           =   69
               Left            =   2895
               TabIndex        =   45
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   4575
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   5565
               TabIndex        =   47
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   4575
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   6930
               TabIndex        =   48
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   4575
               Visible         =   0   'False
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "OdPoNo in"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   66
               Left            =   2625
               TabIndex        =   56
               ToolTipText     =   "Enter 'Y' For Viewing Allowed Loss %, 'N' For Not Considering Allowed Loss %"
               Top             =   5600
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
               Index           =   67
               Left            =   4275
               TabIndex        =   7
               ToolTipText     =   "Enter Raw Material Category"
               Top             =   795
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
               CmpStr          =   "Rm.RmCtg<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   68
               Left            =   6930
               TabIndex        =   8
               ToolTipText     =   "Enter Transaction Date Selection"
               Top             =   795
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UORMCTGSEL"
               IdName          =   "UORMCTGSEL"
               CmpStr          =   "Rm.RmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   90
               Left            =   6960
               TabIndex        =   151
               ToolTipText     =   "Show Production or Process Location"
               Top             =   1440
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
               Left            =   1605
               TabIndex        =   44
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   4575
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
               Index           =   74
               Left            =   4275
               TabIndex        =   46
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   4575
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
               Index           =   16
               Left            =   5505
               TabIndex        =   153
               Top             =   4575
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
               Index           =   12
               Left            =   2835
               TabIndex        =   152
               Top             =   4575
               Width           =   105
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
               Index           =   38
               Left            =   60
               TabIndex        =   128
               Top             =   4575
               Width           =   1425
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
               Index           =   29
               Left            =   60
               TabIndex        =   127
               Top             =   7680
               Width           =   1575
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
               TabIndex        =   126
               Top             =   1080
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
               TabIndex        =   125
               ToolTipText     =   "Location"
               Top             =   4860
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
               Left            =   2370
               TabIndex        =   124
               Top             =   4860
               Width           =   885
            End
            Begin VB.Label LblShowInWt 
               BackStyle       =   0  'Transparent
               Caption         =   "Show In Wt"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   40
               TabIndex        =   123
               Top             =   6960
               Width           =   1560
            End
            Begin VB.Label LblRepFmt 
               BackStyle       =   0  'Transparent
               Caption         =   "Report Format"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   122
               Top             =   5190
               Width           =   1530
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
               TabIndex        =   121
               Top             =   225
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Category"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   120
               ToolTipText     =   "Location"
               Top             =   795
               Width           =   1545
            End
            Begin VB.Label ALBL 
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
               Index           =   4
               Left            =   60
               TabIndex        =   119
               Top             =   2580
               Width           =   1425
            End
            Begin VB.Label LblProc 
               BackStyle       =   0  'Transparent
               Caption         =   "Process"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Top             =   1725
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
               TabIndex        =   117
               Top             =   2295
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
               TabIndex        =   116
               Top             =   2865
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
               TabIndex        =   115
               Top             =   4005
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
               Index           =   27
               Left            =   2325
               TabIndex        =   114
               Top             =   3435
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
               Left            =   4995
               TabIndex        =   113
               Top             =   3435
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
               Left            =   2745
               TabIndex        =   112
               Top             =   3435
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
               Left            =   5415
               TabIndex        =   111
               Top             =   3435
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Index           =   6
               Left            =   60
               TabIndex        =   110
               Top             =   3720
               Width           =   1605
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
               TabIndex        =   109
               Top             =   3435
               Width           =   1545
            End
            Begin VB.Label LblWrk 
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
               Left            =   60
               TabIndex        =   108
               Top             =   2010
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
               TabIndex        =   107
               Top             =   3150
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
               Index           =   0
               Left            =   60
               TabIndex        =   106
               Top             =   4290
               Width           =   1425
            End
            Begin VB.Label LblDt 
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
               Left            =   60
               TabIndex        =   105
               Top             =   510
               Width           =   1605
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
               Left            =   1605
               TabIndex        =   104
               Top             =   0
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
               Left            =   4275
               TabIndex        =   103
               Top             =   0
               Width           =   1275
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
               Left            =   6930
               TabIndex        =   102
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   99
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   100
               Top             =   360
               Width           =   8415
               _ExtentX        =   14843
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpGldLs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepGldLs
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
'  GRP_REP.Add "Process", Array("Td.TdByLoc", "max(Td.TdByLoc)"), "", "TdByLoc", "hTdByLoc", "", "", "", Array("(Select LocDesc From Loc Where LocTyp= 'P' and LocCd= Td.TdByLoc)", "(Select LocDesc From Loc Where LocTyp= 'P' and LocCd= max(Td.TdByLoc))")
'  GRP_REP.Add "(Process)Worker", Array("Td.TdByLoc+T d.TdByWrk", "max(Td.TdByLoc)+max(Td.TdByWrk)"), Array("'('+ RTrim(Td.TdByLoc)+ ') '+ Td.TdByWrk", "'('+ RTrim(max(Td.TdByLoc))+ ') '+ max(Td.TdByWrk)"), "TdByLoc,TdByWrk", "hTdByLoc,hTdByWrk", "Process", "", "", Array("(Select PDesc From Param Where PTyp= 'WORK' and PMCd= Td.TdByWrk)", "(Select PDesc From Param Where PTyp= 'WORK' and PMCd= max(Td.TdByWrk))")
  GRP_REP.Add "Design Ctg", Array("DmCtg", "max(DmCtg)"), "", "DmCtg", "hDmCtg", "", "", "", Array("(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) ", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) ")
  GRP_REP.Add "Prd Ctg", Array("DmPrdCtg", "max(DmPrdCtg)"), "", "DmPrdCtg", "hDmPrdCtg", "", "", "", Array("(Select PDesc from Param where PTyp= 'PRDCTG' and PMCd= DmPrdCtg) ", "(Select PDesc from Param where PTyp= 'PRDCTG' and PMCd= max(DmPrdCtg)) ")
  GRP_REP.Add "Design Code", Array("BOdDmCd", "max(BOdDmCd)"), "", "", "", "Design Ctg,Prd Ctg,(DsgCd)Colour", "", "", Array("Substring(DmDesc,1,120) ", "max(Substring(DmDesc,1,120)) ")
  GRP_REP.Add "(DsgCd)Colour", Array("OdDmCd+OdDmCol", "max(OdDmCd)+max(OdDmCol)"), Array("'('+OdDmCd+') '+OdDmCol", "'('+max(OdDmCd)+') '+max(OdDmCol)"), "", "", "Design Ctg,Prd Ctg,Design Code", "", "", Array("(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) ", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(OdDmCol)) ")
  'GRP_REP.Add "Ord Kt", Array("OdKt", "max(OdKt)"), "", "OdKt", "hOdKt", "", "", "", Array("(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt) ", "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(OdKt)) ")
  GRP_REP.Add "Ord Kt", Array("OdKt", "max(OdKt)"), "", "", "", "", "", "", Array("(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt) ", "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(OdKt)) ")
  GRP_REP.Add "OrderNo", Array("Td.TdCoCd+BOdTc+BOdYy+BOdChr+str(BOdNo)", "Td.TdCoCd+max(BOdTc)+max(BOdYy)+max(BOdChr)+str(max(BOdNo))"), Array("Td.TdCoCd+'/'+BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "Td.TdCoCd+'/'+max(BOdTc)+'/'+max(BOdYy)+'/'+max(BOdChr)+'/'+LTrim(Str(max(BOdNo)))"), "wOrdNo,OmCmCd,OmPoNo", "hOrdNo,hOmCmCd,hOmPoNo", "Company Code,Customer Code,Customer PO", "", "", Array("'('+OmCmCd+ '/'+ OmPONo+ ')' ", "'('+max(OmCmCd)+ '/'+ max(OmPONo)+ ')' ")
  'Array("BCoCd+'/'+BYy+'/'+BChr+'/'+LTrim(Str(BNo))", "max(BCoCd)+'/'+max(BYy)+'/'+max(BChr)+'/'+LTrim(Str(max(BNo)))"
  'Design Ctg, Prd Ctg, Design Code, (DsgCd)Color, Ord Kt, OrderNo, Customer Code, Customer PO, Cust SUB PO, FG Date, Fg Week, Fg Month, Fg Quarter, Inv Vch No, Inv Cust, Bag Type, Bag Character, Bag Typ+Chr'
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  ' Zubin 211 (Patch 03)
  'GRP_REP.Add "Customer Code", Array("OmCmCd", "max(OmCmCd)"), "", "OmCmCd", "hOmCmCd", "", "", "", Array("(Select CmName From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) ", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= max(OmCmCd)) ")
  GRP_REP.Add "Customer Code", Array("OmCmCd", "max(OmCmCd)"), "", "OmCmCd", "hOmCmCd", "", "", "", _
    Array("(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "') <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) ", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= max(OmCmCd)) ")
  
  GRP_REP.Add "Customer PO", Array("OmCmCd+OmPONo", "max(OmCmCd)+max(OmPONo)"), Array("'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "'('+ max(OmCmCd)+ ') '+ LTrim(max(OmPONo))"), "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "", "", "", Array("BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "max(BOdTc)+'/'+max(BOdYy)+'/'+max(BOdChr)+'/'+LTrim(Str(max(BOdNo)))")
  '*** Jay 3.2.0 [OdPoNo]
  'vk.19 sub customer added
  GRP_REP.Add "Cust SUB PO", Array("OmCmCd+OmPONo+OdSubCust+OdPoNo", "max(OmCmCd)+max(OmPONo)+max(OdSubCust)+max(OdPONo)"), Array("'('+ OmCmCd+ ') '+ LTrim(OmPONo)+(case when OdPONo<> '' or OdSubCust<>'' then '/'+ LTrim(OdSubCust) + ' : ' + LTrim(OdPONo) else '' end)", "'('+ max(OmCmCd)+ ') '+ LTrim(max(OmPONo))+(case when max(OdPONo)<> '' or max(OdSubCust)<>'' then '/'+ LTrim(max(OdSubCust)) + ' : ' + LTrim(max(OdPONo)) else '' end)"), "OmCmCd,OmPoNo,OdPoNo", "hOmCmCd,hOmPONo,hOdPoNo", "", "", "", Array("BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))+(case when OdPONo<> '' or OdSubCust<>'' then ' - '+ LTrim(OdSubCust) + ' : ' + LTrim(OdPONo) else '' end )", "max(BOdTc)+'/'+max(BOdYy)+'/'+max(BOdChr)+'/'+LTrim(Str(max(BOdNo)))+'/'+LTrim(Str(max(BOdSr)))+(case when Max(OdPONo)<> '' or Max(OdSubCust)<>'' then ' - '+ LTrim(Max(OdSubCust)) + ' : ' + LTrim(Max(OdPONo)) else '' end )")
  '*** Jay 3.2.0 [OdPoNo]
  
  '*** (Bef 3.2.0)
  'GRP_REP.Add "Date", Array("convert(char(10),TDt,112)", "convert(char(10),max(TDt),112)"), Array("convert(varchar, TDt, 3)", "convert(varchar, max(TDt), 3)"), "", "", "Quarter,Month,Week", "", "", ""
  'GRP_REP.Add "Week", Array("LTrim(Str(DatePart(yyyy,TDt)))+' - '+DateName(WW,TDt)", "LTrim(Str(DatePart(yyyy,max(TDt))))+' - '+DateName(WW,max(TDt))"), "", "", "", "Quarter,Month", "", "", ""
  'GRP_REP.Add "Month", Array("Str(DatePart(yyyy,TDt))+Str(DatePart(mm,TDt))", "Str(DatePart(yyyy,max(TDt)))+Str(DatePart(mm,max(TDt)))"), Array("LTrim(Str(DatePart(yyyy,TDt)))+' '+DateName(MM,TDt)", "LTrim(Str(DatePart(yyyy,max(TDt))))+' '+DateName(MM,max(TDt))"), "", "", "Quarter", "", "", ""
  'GRP_REP.Add "Quarter", Array("LTrim(Str(DatePart(yyyy,TDt)))+' - '+DateName(QQ,TDt)", "LTrim(Str(DatePart(yyyy,max(TDt))))+' - '+DateName(QQ,max(TDt))"), "", "", "", "", "", "", ""
  '*** (Bef 3.2.0)
  
  
  'GRP_REP.Add "Week", Array("DateName(WW,TDt)", "DateName(WW,max(TDt))"), "", "", "", "Quarter,Month", "", "", ""
  'GRP_REP.Add "Quarter", Array("DateName(QQ,TDt)", "DateName(QQ,max(TDt))"), "", "", "", "", "", "", ""
  'GRP_REP.Add "Month", Array("DateName(MM,TDt)", "DateName(MM,max(TDt))"), "", "", "", "Quarter", "", "", ""
  '*** These Groupings are below the Bag Level ###


  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ,Process,OrderNo,OrderSr,Customer PO,FgSubLoc ******
  If ws_MultiCoMod = True Then
    GRP_REP.Add "Company Code", "Td.TdCoCd", "Td.TdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=Td.TdCoCd) "
    If UCase(adc.MenuCd) <> UCase("RepMetLsFg") Then
      GRP_REP.Add "Loc Prnt Code", Array("Lc.LocPrntCd", "max(Lc.LocPrntCd)"), Array("Lc.LocPrntCd", "max(Lc.LocPrntCd)"), "", "", "", "", "", Array("(Select LocDesc from Loc Ln where LocCoCd= '" + ctSelfCoCd + "' and LocCd=Lc.LocPrntCd) ", "(Select LocDesc from Loc Ln where LocCoCd= '" + ctSelfCoCd + "' and LocCd=Max(Lc.LocPrntCd)) ")
    End If
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******
  
    ' **** Manali 3.8.0
  If UCase(adc.MenuCd) = UCase("RepMetLsFg") Then
    '6.1
    GRP_REP.Add "Fg Date", Array("convert(char(10),FdDt,112)", "convert(char(10),max(FdDt),112)"), Array("convert(varchar, FdDt, " + CStr(DtFmtSlahVal()) + ")", "convert(varchar, max(FdDt), " + CStr(DtFmtSlahVal()) + ")"), "", "", "Quarter,Month,Week", "", "", ""
    GRP_REP.Add "Fg Week", Array("LTrim(Str(DatePart(yyyy,FdDt)))+' - '+DateName(WW,FdDt)", "LTrim(Str(DatePart(yyyy,max(FdDt))))+' - '+DateName(WW,max(FdDt))"), "", "", "", "Quarter,Month", "", "", ""
    GRP_REP.Add "Fg Month", Array("Str(DatePart(yyyy,FdDt))+Str(DatePart(mm,FdDt))", "Str(DatePart(yyyy,max(FdDt)))+Str(DatePart(mm,max(FdDt)))"), Array("LTrim(Str(DatePart(yyyy,FdDt)))+' '+DateName(MM,FdDt)", "LTrim(Str(DatePart(yyyy,max(FdDt))))+' '+DateName(MM,max(FdDt))"), "", "", "Quarter", "", "", ""
    GRP_REP.Add "Fg Quarter", Array("LTrim(Str(DatePart(yyyy,FdDt)))+' - '+DateName(QQ,FdDt)", "LTrim(Str(DatePart(yyyy,max(FdDt))))+' - '+DateName(QQ,max(FdDt))"), "", "", "", "", "", "", ""

    GRP_REP.Add "Inv Vch No", Array("IsNull(InCoCd+InTc+InYy+InChr+str(InNo), '')", "IsNull(max(InCoCd)+max(InTc)+max(InYy)+max(InChr)+str(max(InNo)), '')"), Array("IsNull(InCoCd+'/'+InTc+'/'+InYy+'/'+InChr+'/'+LTrim(Str(InNo)), '')", "IsNull(max(InCoCd)+'/'+max(InTc)+'/'+max(InYy)+'/'+max(InChr)+'/'+LTrim(Str(max(InNo))), '')"), "", "", "Company Code,Inv Cust", "", "", Array("IsNull(InExpNo, '') ", "IsNull(max(InExpNo), '')")

    GRP_REP.Add "Inv Cust", Array("InCmCd", "max(InCmCd)"), Array("InCmCd", "max(InCmCd)"), "", "", "", "", "", _
        Array("(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "') <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= InCmCd) ", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= max(InCmCd)) ")
        
    
    GRP_REP.Add "RmCtg", Array("Rm.RmCtg", "max(Rm.RmCtg)"), "", "", "", "", "", "", ""

  Else
    GRP_REP.Add "Location", Array("Td.TdCoCd+Td.TdByLoc", "Td.TdCoCd+max(Td.TdByLoc)"), Array("'['+Td.TdCoCd+'] '+Td.TdByLoc", "'['+Td.TdCoCd+'] '+ max(Td.TdByLoc)"), "", "", "Company Code,Loc Prnt Code", "", "", Array("(Select LocDesc From Loc Where LocCoCd= Td.TdCoCd and LocTyp= 'P' and LocCd= Td.TdByLoc)", "(Select LocDesc From Loc with (index(PK_Loc)) Where LocCoCd= max(Td.TdCoCd) and LocTyp= 'P' and LocCd= max(Td.TdByLoc))")
    GRP_REP.Add "Process", Array("Td.TdCoCd+Lc.LocPrcs", "Td.TdCoCd+max(Lc.LocPrcs)"), Array("'['+Td.TdCoCd+'] '+Lc.LocPrcs", "'['+Td.TdCoCd+'] '+ max(Lc.LocPrcs)"), "", "", "Company Code,Loc Prnt Code", "", "", Array("(Select LocDesc From Loc Where LocCoCd= Td.TdCoCd and LocTyp= 'R' and LocCd= Lc.LocPrcs)", "(Select LocDesc From Loc with (index(PK_Loc)) Where LocCoCd= max(Td.TdCoCd) and LocTyp= 'R' and LocCd= max(Lc.LocPrcs))")
    GRP_REP.Add "(Location)Worker", Array("Td.TdByLoc+Td.TdByWrk", "max(Td.TdByLoc)+max(Td.TdByWrk)"), Array("'('+ RTrim(Td.TdByLoc)+ ') '+ Td.TdByWrk", "'('+ RTrim(max(Td.TdByLoc))+ ') '+ max(Td.TdByWrk)"), "", "", "Process", "", "", Array("(Select vPDesc From vParam Where vPCoCd = Td.TdCoCd and vPTyp= 'WORK' and vPMCd= Td.TdByWrk)", "(Select vPDesc From vParam Where vPCoCd = max(Td.TdCoCd) and vPTyp= 'WORK' and vPMCd= max(Td.TdByWrk))")
    GRP_REP.Add "(Process)Worker", Array("Lc.LocPrcs+Td.TdByWrk", "max(Lc.LocPrcs)+max(Td.TdByWrk)"), Array("'('+ RTrim(Lc.LocPrcs)+ ') '+ Td.TdByWrk", "'('+ RTrim(max(Lc.LocPrcs))+ ') '+ max(Td.TdByWrk)"), "", "", "Process", "", "", Array("(Select vPDesc From vParam Where vPCoCd = Td.TdCoCd and vPTyp= 'WORK' and vPMCd= Td.TdByWrk)", "(Select vPDesc From vParam Where vPCoCd = max(Td.TdCoCd) and vPTyp= 'WORK' and vPMCd= max(Td.TdByWrk))")
    
    GRP_REP.Add "RmCtg", Array("TdRmCtg", "max(Tr.TdRmCtg)"), "", "", "", "", "", "", ""
    '*** These Groupings are below the Bag Level ###
    '*** (3.2.0)  '6.1
    GRP_REP.Add "Date", Array("convert(char(10),Td.TdDt,112)", "convert(char(10),max(Td.TdDt),112)"), Array("convert(varchar, Td.TdDt, " + CStr(DtFmtSlahVal()) + ")", "convert(varchar, max(Td.TdDt), " + CStr(DtFmtSlahVal()) + ")"), "", "", "Quarter,Month,Week", "", "", ""
    GRP_REP.Add "Week", Array("LTrim(Str(DatePart(yyyy,Td.TdDt)))+' - '+DateName(WW,Td.TdDt)", "LTrim(Str(DatePart(yyyy,max(Td.TdDt))))+' - '+DateName(WW,max(Td.TdDt))"), "", "", "", "Quarter,Month", "", "", ""
    GRP_REP.Add "Month", Array("Str(DatePart(yyyy,Td.TdDt))+Str(DatePart(mm,Td.TdDt))", "Str(DatePart(yyyy,max(Td.TdDt)))+Str(DatePart(mm,max(Td.TdDt)))"), Array("LTrim(Str(DatePart(yyyy,Td.TdDt)))+' '+DateName(MM,Td.TdDt)", "LTrim(Str(DatePart(yyyy,max(Td.TdDt))))+' '+DateName(MM,max(Td.TdDt))"), "", "", "Quarter", "", "", ""
    GRP_REP.Add "Quarter", Array("LTrim(Str(DatePart(yyyy,Td.TdDt)))+' - '+DateName(QQ,Td.TdDt)", "LTrim(Str(DatePart(yyyy,max(Td.TdDt))))+' - '+DateName(QQ,max(Td.TdDt))"), "", "", "", "", "", "", ""
    '*** (3.2.0)
    
  '*** Jay 2.13(CT) ***
    GRP_REP.Add "Bag Type", Array("(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) ", "(case Max(PValue) when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end)"), _
                          Array("(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) ", "(case Max(PValue) when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end)"), "", "", "Bag Typ+Chr", "", "", Array("(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)", "(case Max(PValue) when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end)")
    GRP_REP.Add "Bag Character", Array("BChr", "max(BChr)"), Array("BChr", "max(BChr)"), "", "", "Bag Type,Bag Typ+Chr", "", "", Array("(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')", "(Select PDesc from param where PTyp ='BCHR' and PMCd= max(BChr) and PSCd= '')")
    GRP_REP.Add "Bag Typ+Chr", Array("(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) +BChr", "(case Max(PValue) when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) +max(BChr)"), _
                               Array("(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) + '/ ' + BChr", "(case Max(PValue) when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end)  + '/ ' + max(BChr)"), "", "", "Bag Type,Bag Character", "", "", Array("(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')", "(Select PDesc from param where PTyp ='BCHR' and PMCd= max(BChr) and PSCd= '')")
    
  End If
End Sub
Private Sub adc_setreprecsource()
  Dim grpflds0 As String, grpflds1 As String, wrepcnd As String, wCnd As String
  Dim wSqlStrg1 As String, wSqlStrg2 As String, wSqlStrg3 As String, wSqlStrg4 As String
  Dim wSqlStrg5 As String, wSqlStrg As String, wSqlStrg6 As String, wBasicWt As String
  Dim wSqlStrg7 As String, wSqlStrg2Grp As String
  Dim wSqlStrgForSubReport As String

  Dim ws_JoinLsRec As String
  Dim ws_BagJoin As String
  Dim wf_BaseGldRmPurity As Single
  Dim wf_BasePltRmPurity As Single
  Dim wf_BaseSlvRmPurity As Single
  Dim wf_BasePdRmPurity As Single
  Dim wAddlCnd As String   ' ****** Sachin 3.02 - New Fields in Where Clause
  Dim ws_CoCnd As String   ' ****** Sachin 2.14 - [Multi-Company]
  Dim ws_RmCtgwithNullStr, ws_RmCtgStr As String
  Dim ws_BagCnd As String   ' **** Manali 3.8.0
  Dim wb_ProcessGrpSel As Boolean     ' ***** Manali 3.10.0 - 31/03/12 - Show Allowed Ls% in report Only When Process Grps Selected
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  
  Call GetCnd
  wb_ProcessGrpSel = (adc("UoYN5") = "Y") And _
                    (InStr(1, UCase$("," + GRP_REP.GrpOpts + ","), UCase$(",Process,")) > 0 Or _
                    InStr(1, UCase$("," + GRP_REP.GrpOpts + ","), UCase$(",(Process)Worker,")) > 0)
  
  If adc("UoYn1") = "" Then
    wBasicWt = "LrcDustWt"
    GRP_REP.SetFormula Rep, "wDustWt", "if Round({rdo.qDustWt}, 3) <> 0 then 0 else 0"
    Rep.h1DustWt.Suppress = True
    Rep.h1PureRec.Suppress = True
    Rep.h1PureLs.Suppress = True
    Rep.h1AftRecPrc.Suppress = True
    Rep.h2DustWt.Suppress = True
    Rep.h2PureRec.Suppress = True
    Rep.h2PureLs.Suppress = True
    Rep.h2AftRecPrc.Suppress = True
    If wb_ProcessGrpSel Then
      Rep.h1DustWt.Suppress = False
      Rep.h2DustWt.Suppress = False
      Rep.h1DustWt.SetText "Allowed Ls%"
      Rep.h2DustWt.SetText "Allowed Ls%"
    End If
  Else
    Rep.h1DustWt.Suppress = False
    Rep.h1PureRec.Suppress = False
    Rep.h1PureLs.Suppress = False
    Rep.h1AftRecPrc.Suppress = False
    Rep.h2DustWt.Suppress = False
    Rep.h2PureRec.Suppress = False
    Rep.h2PureLs.Suppress = False
    Rep.h2AftRecPrc.Suppress = False
    If adc("UOYN1") = "Y" Then
      wBasicWt = "LrcDustWt"
      GRP_REP.SetFormula Rep, "wDustWt", "if Round({rdo.qDustWt}, 3) <> 0 then {rdo.TdDustWt} else 0"
      Rep.h1DustWt.SetText "Dust Wt"
      Rep.h2DustWt.SetText "Dust Wt"
    ElseIf adc("UOYN1") = "N" Then
      wBasicWt = "LrcActLs"
      GRP_REP.SetFormula Rep, "wDustWt", "if Round({rdo.qDustWt}, 3) <> 0 then {@wPureLsWt} else 0"
      Rep.h1DustWt.SetText "Actual" + Chr(13) + "Loss"
      Rep.h2DustWt.SetText "Actual" + Chr(13) + "Loss"
    End If
    Rep.PaperOrientation = crLandscape
  End If
  grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)
      
  '*** To check if Detail is selected ***
  Dim wGrp(3) As String, i As Integer, wDetPos As Integer, wFirstNonePos As Integer, wGrpNum As String, wLimit As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  wFirstNonePos = 0
  wLimit = 0
  For i = 1 To 3
    If wGrp(i) = "(Detail)" Then
      wDetPos = i   ': Exit For
    End If
    If wGrp(i) = "(None)" Or wGrp(i) = "(Grand)" Then
      wFirstNonePos = i: Exit For
    End If
  Next i
  
  '### Note TdByLoc,TdByWrk,OdKt fields place in G4-DET, will come from the last record of the Bag. OK.
    
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then
    Rep.BoxHd.Suppress = True: Rep.BoxDet.Suppress = True: Rep.BoxFt.Suppress = True
    Rep.DetSec2.Suppress = True
    Rep.SubRepLsDet.Suppress = True: Rep.LineSec.Suppress = True
    wGrpNum = "4"
    wLimit = 4 - wFirstNonePos
  Else
    wGrpNum = "3"
    wLimit = 3 - wDetPos
  End If
  
  ' **** Manali 3.8.0
  If UCase(adc.MenuCd) = UCase("RepMetLsFg") Then
    Rep.BoxHd.Suppress = True: Rep.BoxDet.Suppress = True: Rep.BoxFt.Suppress = True
    Rep.DET.Suppress = True
  End If
  
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_wLsWt", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_wPureLsWt", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_Prc", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_GldWt", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_PrnGldWt", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_ResGldWt", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_PureGldWt", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_PrnPureGldWt", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_ResPureGldWt", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_Qty", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_PrnQty", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_ResQty", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_wDustWt", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_wPureRec", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_wPureLs", "0"
  GRP_REP.SetFormula Rep, "g" + wGrpNum + "_wPrcntRec", "0"
    
  If wFirstNonePos > 0 Then
    For i = 1 To wLimit
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_wLsWt", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_wPureLsWt", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_Prc", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_GldWt", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_PrnGldWt", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_ResGldWt", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_PureGldWt", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_PrnPureGldWt", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_ResPureGldWt", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_Qty", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_PrnQty", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_ResQty", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_wDustWt", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_wPureRec", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_wPureLs", "0"
      GRP_REP.SetFormula Rep, "g" + CStr(i) + "_wPrcntRec", "0"
    Next i
  End If
    
  adc("UoRmCtgFr").CmpStr = "": adc("UoRmCtgTo").CmpStr = "": adc("UoRmCtgSel").CmpStr = ""
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", wrepcnd, "")
  
  If OptLoc(0).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and Td.TdByLoc >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Td.TdByLoc <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Td.TdByLoc In (" + adc("UoIssLocSel") + ")"
  ElseIf OptLoc(1).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and Lc.LocPrcs >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Lc.LocPrcs <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Lc.LocPrcs In (" + adc("UoIssLocSel") + ")"
  End If
  If adc("UOYN1") <> "" Then
    wSqlStrg6 = " IsNull(LR1.LrcRecWt, IsNull(LR2.LrcRecWt, 0)) as qRecWt "
    wSqlStrg7 = " IsNull(LR1." + wBasicWt + ", IsNull(LR2." + wBasicWt + ", 0)) as qDustWt "
    ws_JoinLsRec = " Left outer join LsRec LR1 on LR1.LrcCoCd = TdCoCd " + _
                   "and Substring(LTrim(Rtrim(Str(LR1.LrcYyMM))), 5, 2) = month(TdDt) " + _
                   "and Substring(LTrim(Rtrim(Str(LR1.LrcYyMM))), 1, 4) = year(TdDt) " + _
                   "and LR1.LrcLoc = TdByLoc and LR1.LrcRmCd = TdRmCd and LR1.LrcRmCtg = Rm.RmCtg " + _
                   "and convert(decimal(16, 3), LR1." + wBasicWt + ")<> 0 " + _
                   "Left outer join LsRec LR2 on LR2.LrcCoCd = TdCoCd " + _
                   "and Substring(LTrim(Rtrim(Str(LR2.LrcYyMM))), 5, 2) = month(TdDt) " + _
                   "and Substring(LTrim(Rtrim(Str(LR2.LrcYyMM))), 1, 4) = year(TdDt) " + _
                   "and LR2.LrcLoc = TdByLoc and LR2.LrcRmCd = '' and LR2.LrcRmCtg = Rm.RmCtg " + _
                   "and convert(decimal(16, 3), LR2." + wBasicWt + ")<> 0  "
  Else
    wSqlStrg6 = " 0 as qRecWt "
    wSqlStrg7 = " 0 as qDustWt "
    ws_JoinLsRec = " "
  End If
  
    
  wf_BaseGldRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst " + _
                       "on " + IIF(gs_Partition = ctCurrPrtn, "RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'G'), 1)")
  wf_BasePltRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst " + _
                       "on " + IIF(gs_Partition = ctCurrPrtn, "RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'P'), 1)")
  wf_BaseSlvRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst " + _
                       "on " + IIF(gs_Partition = ctCurrPrtn, "RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'S'), 1)")
  wf_BasePdRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst " + _
                       "on " + IIF(gs_Partition = ctCurrPrtn, "RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'L'), 1)")
  ws_BagJoin = " Join Param with (Index (PK_Param)) On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                    IIF(adc("UoBagTyp") <> "", " and PValue= '" + adc("UoBagTyp") + "' ", " ") + _
                    IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
      wAddlCnd = IIF(adc("UoRmCtgFr") <> "", " and Td.TdRmCtg>= '" + adc("UoRmCtgFr") + "'", "")
      wAddlCnd = wAddlCnd + IIF(adc("UoRmCtgTo") <> "", " and Td.TdRmCtg<='" + adc("UoRmCtgTo") + "'", "")
      wAddlCnd = wAddlCnd + IIF(adc("UoRmCtgSel") <> "", " and Td.TdRmCtg in (" + adc("UoRmCtgSel") + ")", "")
      
                    
  If UCase(adc.MenuCd) = UCase("RepMetLsFg") Then
    ws_BagCnd = IIF(adc("UoYn4") <> "", " and BCls='" + adc("UoYn4") + "' ", "")
    wSqlStrg1 = "Select " + grpflds1 + ", " + _
                   "'P' as qTag, TdBYy, TdBChr, TdBNo,TdRmCtg,  " + _
                   "Sum((Case When (TdToRmLocTyp= 'L' Or TdFrRmLocTyp= 'L') and not(TdFrRmLocTyp= 'L' and TdToRmLocTyp= 'L')  Then " + _
                   "                     (Case when TdToRmLocTyp= 'L' and TdToRmDc= 'D' then 1 " + _
                   "                     when TdToRmLocTyp= 'L' and TdToRmDc= 'C' then -1 " + _
                   "                     when TdFrRmLocTyp= 'L' and TdFrRmDc= 'D' then 1 " + _
                   "                     when TdFrRmLocTyp= 'L' and TdFrRmDc= 'C' then -1 Else 0 end) Else 0 End)* TdRmWt) as qLsWt, " + _
                   "1 as qMul, " + _
                   "Sum((Case When (TdToRmLocTyp= 'L' Or TdFrRmLocTyp= 'L') and not(TdFrRmLocTyp= 'L' and TdToRmLocTyp= 'L')  Then " + _
                   "                     (Case when TdToRmLocTyp= 'L' and TdToRmDc= 'D' then 1 " + _
                   "                     when TdToRmLocTyp= 'L' and TdToRmDc= 'C' then -1 " + _
                   "                     when TdFrRmLocTyp= 'L' and TdFrRmDc= 'D' then 1 " + _
                   "                     when TdFrRmLocTyp= 'L' and TdFrRmDc= 'C' then -1 Else 0 end) Else 0 End)* " + _
                   " TdRmWt* RmPurityWt/ (case when Rm.RmCtg= 'G' then " + CStr(wf_BaseGldRmPurity) + " when Rm.RmCtg= 'P' then " + CStr(wf_BasePltRmPurity) + " when Rm.RmCtg= 'S' then " + CStr(wf_BaseSlvRmPurity) + " when Rm.RmCtg= 'L' then " + CStr(wf_BasePdRmPurity) + " else 1 end)) as qPureLsWt, " + _
                    " sum(TdRmWt *  (Case when TdFrRmLoc='W' and TdFrRmDc='D' then 1 " + _
                   "                 when TdFrRmLoc='W' and TdFrRmDc='C' then -1 " + _
                   "                 when TdToRmLoc='W' and TdToRmDc='D' then 1 " + _
                   "                 when TdToRmLoc='W' and TdToRmDc='C' then -1 " + _
                   "                 Else 0 end)) as qBGldWt, " + _
                   "sum((Case when TdFrRmLoc='W' and TdFrRmDc='D' then 1 when TdFrRmLoc='W' and TdFrRmDc='C' then -1 when TdToRmLoc='W' and TdToRmDc='D' then 1 when TdToRmLoc='W' and TdToRmDc='C' then -1 Else 0 end) * TdRmWt * RmPurityWt " + _
                   "/ (case when IsNull(Rm.RmCtg, '')= 'G' then " + CStr(wf_BaseGldRmPurity) + " when Rm.RmCtg= 'P' then " + CStr(wf_BasePltRmPurity) + " when Rm.RmCtg= 'S' then " + CStr(wf_BaseSlvRmPurity) + " when Rm.RmCtg= 'L' then " + CStr(wf_BasePdRmPurity) + " else 1 end)) as qPureBGldWt "
     wSqlStrg1 = wSqlStrg1 + " ,'01/01/80' as TDt,'' as qRmCd, " + _
                " '' as qRmDc, " + _
                " '' as TdByWrk, '' as LocPrntCd,'' as TdByLoc, '' as qLsLoc, 0 as LocAllowLsPer, '' TdTc,'' TdYy,'' TdChr, " + _
                " 0 as TdNo,0 as  TdSr, '' as TdYyKey, max(DmCtg) as DmCtg,max(DmPrdCtg) as DmPrdCtg, max(BOdDmCd) as BOdDmCd ,Max(OdKt) as OdKt,Max(BOdTc) BOdTc,Max(BOdYy) as BOdYy,Max(BOdChr) BOdChr,Max(BOdNo) as BOdNo, Max(OmCmCd) as OmCmCd, Max(OmPoNo) as OmPoNo, " + _
                " max(BQty) as TdBQty, 0.0 as TdDustWt, " + wSqlStrg6 + ", " + wSqlStrg7
                
      wCnd = IIF(gs_Partition = ctCurrPrtn, " FdPrtKey= '" + ctCurrPrtn + "' " + IIF(wCnd <> "", " and ", ""), "") + wCnd
      wSqlStrg2Grp = " Group By TdCoCd, TdBYy, TdBChr, TdBNo, TdRmCtg   "
      '578-34 Str(FdKey,7) changed to Str(FdKey,8), Str(M.FdKey,7) changed to Str(M.FdKey,8)
      wSqlStrg2 = " From Fgd " + _
                  " Join OrdDsg  on OdCoCd= FdCoCd and OdTc= FdPrdOdTc " + _
                  " and OdYy= FdPrdOdYy and OdChr= FdPrdOdChr and OdNo= FdPrdOdNo and OdSr= FdPrdOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=FdPrtKey ", "") + _
                  " Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc " + _
                  " and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
                  " Join Bag  on BIdNo= FdBIdNo and BCoCd= FdCoCd and BYy= FdBYy and BChr= FdBChr and BNo= FdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey= FdPrtKey ", "") + _
                  " Join Txnd Td On TdBIdNo=BIdNo and TdCoCd=BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey= BPrtKey ", "") + _
                  " Join RmMst Rm on RmCd=TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=TdPrtKey ", "") + _
                  " Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd " + _
                  " and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
                  " Left Outer Join InvFgd On IfCoCd=FdCoCd and IfFdBYy=FdBYy and IfFdBChr=FdBChr and IfFdBNo=FdBNo and IfFdIdNo=FdIdNo and (Select PValue from Param Where PTyp='TC' and PMCd=IfTc)='" + ctInTcTyp + "' " + IIF(gs_Partition = ctCurrPrtn, " and IfPrtKey=FdPrtKey ", "") + _
                  " Left Outer Join InvHd On InCoCd=IfCoCd and InTc=IfTc and InYy=IfYy and InChr=IfChr and InNo=IfNo and InIdNo=IfInIdNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IfPrtKey ", "") + _
                  " Where " + wCnd + IIF(wCnd <> "", " and ", "") + _
                  " /* TdRmCtg= '" + adc("UoRmCtgFr") + "' and  TdRmCtg In ('G','P','S','L')  and */  BLoc='PFG' " + wAddlCnd + _
                  " and FdYy+Str(FdKey,8)=(Select Max(M.FdYy+Str(M.FdKey,8)) From Fgd M Where M.FdCoCd=BCoCd and M.FdBYy=BYy and M.FdBChr=BChr and M.FdBNo=BNo and M.FdBIdNo=BIdNo " + IIF(gs_Partition = ctCurrPrtn, " and M.FdPrtKey=BPrtKey ", "") + ") " + _
                  " and TdSrNo > 0 " + ws_BagCnd
      
  Else
  
    '****** Sachin 2.14.0 - [Multi-Company] - [27-11-2006] - Loc Join Lc added below ******
    '/** selecting Loss where TdRmFrLoc = 'L' or TdRmToLoc = 'L' (and not both = 'L') **/
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    ' ***** Manali 3.9.0 - TdYyKey added
    ' ***** Manali 3.9.0 - txn Join (commnted earlier) uncommented now - as query was slow in case date scope selected - Date compare strings changed to TDt from td.TdDt
    ' ***** Manali 3.10.0 - 31/03/12 - Show Allowed Ls% in report Only When Process Grps Selected
    wSqlStrg1 = "Select " + grpflds0 + ", " + _
                "'D' as qTag, TdCoCd, TdBYy, TdBChr, TdBNo,TdRmCtg, TdRmWt as qLsWt, " + _
                "(Case when L2.LocTyp= 'L' and TdToRmDc= 'D' then 1 " + _
                "when L2.LocTyp= 'L' and TdToRmDc= 'C' then -1 " + _
                "when L1.LocTyp= 'L' and TdFrRmDc= 'D' then 1 " + _
                "when L1.LocTyp= 'L' and TdFrRmDc= 'C' then -1 " + _
                "Else 0 end) as qMul, " + _
                "    /* TdRmWt*RmPurityWt/IsNull((Select RmPurityWt from Param join RmMst on " + _
                "       RmCd= PValue where PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) as qPureLsWt, " + _
                "    */ " + _
                "TdRmWt* RmPurityWt/ (case when Rm.RmCtg= 'G' then " + CStr(wf_BaseGldRmPurity) + " when Rm.RmCtg= 'P' then " + CStr(wf_BasePltRmPurity) + " " + _
                "when Rm.RmCtg= 'S' then " + CStr(wf_BaseSlvRmPurity) + " when Rm.RmCtg= 'L' then " + CStr(wf_BasePdRmPurity) + " else 1 end) as qPureLsWt,  " + _
                "0.0 as qBGldWt, 0.0 as qPureBGldWt, TdDt as TDt, TdRmCd as qRmCd, " + _
                "(Case when L2.LocTyp= 'L' then TdToRmDc when L1.LocTyp= 'L' then TdFrRmDc Else '' end) as qRmDc, " + _
                "TdByWrk, Lc.LocPrntCd as LocPrntCd," + IIF(OptLoc(0).Value = True, "TdByLoc", "Lc.LocPrcs") + " as TdByLoc, (Case when L1.LocTyp= 'L' then TdFrRmLoc " + _
                "when L2.LocTyp= 'L' then TdToRmLoc Else '' end) as qLsLoc, " + _
                IIF(wb_ProcessGrpSel, " (SELECT LocAllowLsPer FROM Loc WHERE LocCoCd=Lc.LocCoCd AND LOCCD=Lc.LocPrcs) ", " 0 ") + " as LocAllowLsPer, " + _
                "TdTc, TdYy, TdChr, " + _
                "TdNo, TdSr, TdYyKey, DmCtg, DmPrdCtg, BOdDmCd, OdKt, BOdTc, BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo, " + _
                "(Select TdBQty from Txnd where " + IIF(gs_Partition = ctCurrPrtn, " TdPrtKey=Td.TdPrtKey and ", "") + " TdCoCd= Td.TdCoCd " + _
                "and TdTc= Td.TdTc and TdYy= Td.TdYy and TdChr= Td.TdChr and TdNo= Td.TdNo " + _
                "and TdSr= Td.TdSr and TdSrNo= 0) as TdBQty, TdDustWt, " + wSqlStrg6 + ", " + wSqlStrg7
                
    ' **** Zubin 211 (ws_BagJoin added) **** '
    '****** Sachin 3.02 - Id fields In Joins
    
    wCnd = IIF(gs_Partition = ctCurrPrtn, " Td.TdPrtKey= '" + ctCurrPrtn + "' " + IIF(wCnd <> "", " and ", ""), "") + wCnd
    
    wSqlStrg2 = " /*From Txnd Td*/ From Txn " + _
                "Join Txnd Td on TdTIdNo= TIdNo and TdCoCd= TCoCd and TdTc= TTc " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=TPrtKey ", "") + _
                "and TdYy= TYy and TdChr= TChr and TdNo= TNo and TdSrNo> 0 " + _
                "Join RmMst Rm on RmCd=TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=TdPrtKey ", "") + _
                "join Loc L1  on L1.LocCoCd= TdCoCd and L1.LocCd= TdFrRmLoc " + _
                "join Loc L2  on L2.LocCoCd= TdCoCd and L2.LocCd= TdToRmLoc " + _
                "join Loc Lc  on Lc.LocCoCd= TdCoCd and Lc.LocCd= TdByLoc " + _
                "Join Bag  on BIdNo= TdBIdNo and BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey= TdPrtKey ", "") + _
                ws_BagJoin + _
                "Join OrdDsg  on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc " + _
                "and OdYy= BOdYy and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
                "Join OrdMst on OmIdNo=BOmIdNo and OmCoCd= BCoCd and OmTc= BOdTc " + _
                "and OmYy= BOdYy and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=BPrtKey ", "") + _
                "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd " + _
                " and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + ws_JoinLsRec + _
                " Where " + wCnd + _
                "  /* TdRmCtg= '" + adc("UoRmCtgFr") + "'   TdRmCtg In ('G','P','S','L') */ " + wAddlCnd + _
                " and TdSrNo > 0 and (L2.LocTyp= 'L' or L1.LocTyp= 'L') and (TdToRmLocTyp= 'L' Or TdFrRmLocTyp= 'L') " + _
                " and not(L1.LocTyp= 'L' and L2.LocTyp= 'L') and not(TdFrRmLocTyp= 'L' and TdToRmLocTyp= 'L') "
                
                '" Where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' and ", "") + ws_CoCnd + IIF(Trim(ws_CoCnd) <> "", " And ", "") + _
                wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + _
                " and RmCtg= '" + adc("UoRmCtgFr") + "' And TdRmCtg= '" + adc("UoRmCtgFr") + "'" + _
                " and (L2.LocTyp= 'L' or L1.LocTyp= 'L') And (TdToRmLocTyp= 'L' Or TdFrRmLocTyp= 'L') " + _
                " and not(L1.LocTyp= 'L' and L2.LocTyp= 'L') and not(TdFrRmLocTyp= 'L' and TdToRmLocTyp= 'L') "
                
    ' **** Zubin 211 (ws_BagJoin added) **** '
    
    '/*** selecting Prd Records ***/
    '/* (qBGldWt) Gold weight in the Bag as of TdYyKey of PrdYN='Y' */
    '/* (qBPureGldWt) Pure Gold weight in the Bag as of TdYyKey of PrdYN='Y' */
    
    '*** (Jen 2.14)
    If adc("UoYN3") = "Y" Then
    '*** (Jen 2.14)
    
      '****** Sachin 2.14.0 - [Multi-Company] - [27-11-2006] - Loc Join Lc added below ******
      ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
      ' ***** Manali 3.9.0 - TdYyKey added
    ' ***** Manali 3.9.0 - txn Join (commnted earlier) uncommented now - as query was slow in case date scope selected - Date compare strings changed to TDt from td.TdDt
      wSqlStrg3 = "Select " + grpflds1 + ", " + _
                  "'P' as qTag, Td.TdCoCd as TdCoCd, max(Td.TdBYy) as TdBYy, max(Td.TdBChr) as TdBChr, " + _
                  "max(Td.TdBNo) as TdBNo,Tr.TdRmCtg as TdRmCtg, 0.0 as qLsWt, 0 as qMul, 0.0 as qPureLsWt, " + _
                  "Sum(isnull(Tr.TdRmWt* (Case when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='D' then 1 " + _
                  "when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='C' then -1 " + _
                  "when Tr.TdToRmLoc='W' and Tr.TdToRmDc='D' then 1 " + _
                  "when Tr.TdToRmLoc='W' and Tr.TdToRmDc='C' then -1 Else 0 end),0)) as qBGldWt, " + _
                  "Sum(isnull(Tr.TdRmWt* (Rm.RmPurityWt)* " + _
                  "(Case when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='D' then 1 " + _
                  "when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='C' then -1 " + _
                  "when Tr.TdToRmLoc='W' and Tr.TdToRmDc='D' then 1 " + _
                  "when Tr.TdToRmLoc='W' and Tr.TdToRmDc='C' then -1 Else 0 end)/ " + _
                  "(case when isnull(Rm.RmCtg, '')= 'G' then " + CStr(wf_BaseGldRmPurity) + " when isnull(Rm.RmCtg, '')= 'P' then " + CStr(wf_BasePltRmPurity) + " " + _
                  "when isnull(Rm.RmCtg, '')= 'S' then " + CStr(wf_BaseSlvRmPurity) + " when isnull(Rm.RmCtg, '')= 'L' then " + CStr(wf_BasePdRmPurity) + " else 1 end) " + _
                  "/* IsNull(PureRm.RmPurityWt, 1) */  ,0)) as qBPureGldWt, "
      wSqlStrg3 = wSqlStrg3 + " max(Td.TdDt) as TDt, '' as qRmCd, '' as qRmDc, max(Td.TdByWrk) as TdByWrk, " + _
                  "max(Lc.LocPrntCd) as LocPrntCd," + IIF(OptLoc(0).Value = True, "max(Td.TdByLoc)", "max(Lc.LocPrcs)") + " as TdByLoc, '' as qLsLoc, " + _
                  IIF(wb_ProcessGrpSel, " max(Lc.LocAllowLsPer) ", " 0 ") + " as LocAllowLsPer, " + _
                  "Td.TdTc as TdTc, (Td.TdYy) as TdYy, " + _
                  "(Td.TdChr) as TdChr, (Td.TdNo) as TdNo, (Td.TdSr) as TdSr, max(Td.TdyyKey) as TdYyKey, max(DmCtg) as DmCtg, " + _
                  "max(DmPrdCtg) as DmPrdCtg, Max (BOdDmCd) as BOdDmCd, Max(OdKt) as OdKt, " + _
                  "Max(BOdTc) as BOdTc, Max(BOdYy) as BOdYy, Max(BOdChr) as BOdChr, " + _
                  "Max(BOdNo) as BOdNo, Max(OmCmCd) as OmCmCd, Max(OmPoNo) as OmPoNo, " + _
                  "max(Td.TdBQty) as TdBQty, 0.0 as TdDustWt, 0.0 as LrcRecWt, 0.0 as LrcDustWt "
      
      ' **** Zubin 211 (ws_BagJoin added) ****
      
      '4.1.2
      adc("UoRmCtgFr").CmpStr = "": adc("UoRmCtgTo").CmpStr = "": adc("UoRmCtgSel").CmpStr = ""
      wrepcnd = adc.RepCond
      wCnd = IIF(wrepcnd <> "", wrepcnd, "")
    
  
  If OptLoc(0).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and Td.TdByLoc >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Td.TdByLoc <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Td.TdByLoc In (" + adc("UoIssLocSel") + ")"
  ElseIf OptLoc(1).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and Lc.LocPrcs >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Lc.LocPrcs <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Lc.LocPrcs In (" + adc("UoIssLocSel") + ")"
  End If
      wCnd = IIF(gs_Partition = ctCurrPrtn, " Td.TdPrtKey= '" + ctCurrPrtn + "' " + IIF(wCnd <> "", " and ", ""), "") + wCnd
      wCnd = wCnd '+ IIF(wCnd <> "", " and ", "") + " Rm.RmCtg In ('G','P','S','L')"
      
      
      wAddlCnd = IIF(adc("UoRmCtgFr") <> "", "  IsNull(Tr.TdRmCtg,'')>= '" + adc("UoRmCtgFr") + "'", "")
      If adc("UoRmCtgTo") <> "" Then wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "", " and ", "") + " IsNull(Tr.TdRmCtg,'')<='" + adc("UoRmCtgTo") + "'"
      If adc("UoRmCtgSel") <> "" Then wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "", " and ", "") + " IsNull(Tr.TdRmCtg,'') in (" + adc("UoRmCtgSel") + ")"
      If wAddlCnd <> "" Then _
            wAddlCnd = " and (IsNull(Tr.TdRmCtg,'')=''  Or (" + wAddlCnd + " )) "
                     
     
      '****** Sachin 3.02 - Id fields in Joins
      wSqlStrg4 = "/*From Txnd Td*/  From Txn  " + _
                  "Join Txnd Td  on Td.TdTIdNo=TIdNo and Td.TdCoCd= TCoCd and Td.TdTc= TTc " + _
                  "and Td.TdYy= TYy and Td.TdChr= TChr and Td.TdNo= TNo and Td.TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " and Td.TDPrtKey=TPrtKey", "") + _
                  " Join Loc Lc on LocCoCd=Td.TdCoCd And LocCd=Td.TdByLoc " + _
                  "Join Bag on BIdNo=Td.TdBIdNo and BCoCd= Td.TdCoCd and BYy= Td.TdBYy and BChr= Td.TdBChr and BNo= Td.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=Td.TdPrtKey", "") + _
                  ws_BagJoin + _
                  " Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc " + _
                  "and OdYy= BOdYy and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", " ") + _
                  "Join OrdMst on OmCoCd=BCoCd and OmTc= BOdTc " + _
                  "and OmYy= BOdYy and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=BPrtKey ", " ") + _
                  "Join DsgMst on DmIdNo=BDmIdNo and  DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", " ") + _
                  "left outer Join Txnd Tr on Tr.TdBIdNo=Td.TdBIdNo and Tr.TdCoCd= Td.TdCoCd " + _
                  "and Tr.TdBYy= Td.TdBYy and Tr.TdBChr= Td.TdBChr and Tr.TdBNo = Td.TdBNo " + _
                  "and Tr.TdYyKey <= Td.TdYyKey and Tr.TdSrNo> 0 " + IIF(gs_Partition = ctCurrPrtn, " and Tr.TdPrtKey=Td.TdPrtKey ", " ") + _
                  "left outer Join RmMst Rm  on Rm.RmCd= Tr.TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and Rm.RmPrtKey=Tr.TdPrtKey ", " ") + _
                  "     /* left outer Join Param  on PTyp= 'RMCTG' and PMCd= Rm.RmCtg " + _
                  "        left outer Join RmMst PureRm  on PureRm.RmCd= PValue " + _
                  "     */ " + _
                  "Where " + wCnd + IIF(wCnd <> "", " and ", "") + _
                  " Td.TdSrNo= 0 and Td.TdPrdYN= 'Y' " + wAddlCnd + _
                  " /* And  isnull(Rm.RmCtg, '" + adc("UoRmCtgFr") + "')= '" + adc("UoRmCtgFr") + "'  and (Rm.RmCtg Is Null Or Rm.RmCtg In ('G','P','S','L')) */ " + _
                  "Group by Td.TdCoCd, Td.TdTc, Td.TdYy, Td.TdChr, Td.TdNo, Td.TdSr,Tr.TdRmCtg, Td.TdBYy, " + _
                  "Td.TdBChr, Td.TdBNo "
                  
    End If
  End If
  
  If UCase(adc.MenuCd) = UCase("RepMetLsFg") Then
    wSqlStrg5 = " Order By TdCoCd, TdBYy, TdBChr, TdBNo  " + _
                "OPTION (/*FORCE ORDER,*/ MAXDOP 1) "
  Else
    wSqlStrg5 = " Order By TdCoCd, TdYyKey, TdByy, TdBChr, TdBNo " + _
                "OPTION (MAXDOP 1) "
  
  End If
  
    '*** (Jen 2.14)
    If adc("UoYN3") = "Y" Then
    '*** (Jen 2.14)
      '6.1
      wSqlStrg = "set dateformat " + DtFmtStr() + " " + wSqlStrg1 + wSqlStrg2 + " Union All " + wSqlStrg3 + wSqlStrg4 + wSqlStrg5
    
    '*** (Jen 2.14)
    Else  '6.1
      wSqlStrg = "set dateformat " + DtFmtStr() + " " + wSqlStrg1 + wSqlStrg2 + wSqlStrg2Grp + wSqlStrg5
      
      
      Rep.h1AftRecPrc.Suppress = True: Rep.h2AftRecPrc.Suppress = True
      Rep.g1wPrcntRec.Suppress = True: Rep.g2wPrcntRec.Suppress = True
      Rep.g3wPrcntRec.Suppress = True: Rep.g4wPrcntRec.Suppress = True
      Rep.gwPrcntRec.Suppress = True
      Rep.h1LossPrc.Suppress = True: Rep.h2LossPrc.Suppress = True
      Rep.g1Prc.Suppress = True: Rep.g2Prc.Suppress = True
      Rep.g3Prc.Suppress = True: Rep.g4Prc.Suppress = True
      Rep.gPrc.Suppress = True
      Rep.h1PureGldWt.Suppress = True: Rep.h2PureGldWt.Suppress = True
      Rep.g1PrnIoPureGldWt.Suppress = True: Rep.g2PrnIoPureGldWt.Suppress = True
      Rep.g3PrnIoPureGldWt.Suppress = True: Rep.g4PrnIoPureGldWt.Suppress = True
      Rep.gPrnIoPureGldWt.Suppress = True
      Rep.h1InOutGldWt.Suppress = True: Rep.h2InOutGldWt.Suppress = True
      Rep.g1PrnIoGldWt.Suppress = True: Rep.g2PrnIoGldWt.Suppress = True
      Rep.g3PrnIoGldWt.Suppress = True: Rep.g4PrnIoGldWt.Suppress = True
      Rep.gPrnIoGldWt.Suppress = True
      
      Rep.h1InOutGldWt1.Suppress = True: Rep.h2InOutGldWt1.Suppress = True
      Rep.g1PrnIoGldWt1.Suppress = True: Rep.g2PrnIoGldWt1.Suppress = True
      Rep.g3PrnIoGldWt1.Suppress = True: Rep.g4PrnIoGldWt1.Suppress = True
      Rep.gPrnIoGldWt1.Suppress = True
    End If
  If UCase(adc.MenuCd) = UCase("RepMetLsFg") Then
  '6.1
      wSqlStrg = "set dateformat " + DtFmtStr() + " " + wSqlStrg1 + wSqlStrg2 + wSqlStrg2Grp + wSqlStrg5
      wSqlStrg1 = wSqlStrg1 + "," + "  (Case When (TdToRmLocTyp= 'L' Or TdFrRmLocTyp= 'L') and not(TdFrRmLocTyp= 'L' and TdToRmLocTyp= 'L') then (case when TdToRmLocTyp='L' then TdToRmLoc when TdFrRmLocTyp = 'L' then TdFrRmLoc else '' end) else '' end) as qLossLoc "
            wSqlStrgForSubReport = "      and TdSrNo > 0" + _
                                  "      and ((TdToRmLocTyp= 'L' Or TdFrRmLocTyp= 'L') and not(TdFrRmLocTyp= 'L' and TdToRmLocTyp= 'L')) " + _
                                  "      and TdYyKey <=(Select Max(FdByyKey) From Fgd M Where M.FdCoCd=BCoCd and M.FdBYy=BYy and M.FdBChr=BChr and M.FdBNo=BNo and M.FdBIdNo=BIdNo  and M.FdPrtKey=BPrtKey )" + _
                                  "      Group By TdCoCd, TdBYy, TdBChr, TdBNo, TdRmCtg, " + _
                                  "      (Case When (TdToRmLocTyp= 'L' Or TdFrRmLocTyp= 'L') and not(TdFrRmLocTyp= 'L' and TdToRmLocTyp= 'L') then (case when TdToRmLocTyp='L' then TdToRmLoc when TdFrRmLocTyp = 'L' then TdFrRmLoc else '' end) else '' end) "
            '6.1
            wSqlStrgForSubReport = "set dateformat " + DtFmtStr() + " " + wSqlStrg1 + wSqlStrg2 + wSqlStrgForSubReport + wSqlStrg5
  
    Call moCn.RepRes(Rep.Subreport1.OpenSubreport, wSqlStrgForSubReport)     '****** Sachin 3.01 added ctMaxDopOpt
  End If
  
  
  adc.RepRecSource = wSqlStrg

  Rep.TxtFrDt.SetText (adc("UoDtFr"))
  Rep.TxtToDt.SetText (adc("UoDtTo"))
  
  If UCase(adc.MenuCd) = UCase("RepMetLsFg") Then
    Rep.hTdByLoc.Suppress = True
    Rep.hTdByWrk.Suppress = True
    If adc("UoYN") <> "Y" Then
      Rep.SubRepLsDet.Suppress = True
    End If
  Else
    Rep.SubRepLsDet.Suppress = True
    If adc("UoYN") = "N" Then
      Rep.BoxHd.Suppress = True
      Rep.BoxDet.Suppress = True
      Rep.BoxFt.Suppress = True
      Rep.DET.Suppress = True
      Rep.FldDet2BagNo.Suppress = False
      Rep.FldDet2ByLoc.Suppress = False
      Rep.FldDet2ByWrk.Suppress = False
      Rep.FldDet2Kt.Suppress = False
      Rep.FldDet2TdRmCtg.Suppress = False
      Rep.FldDet2BOdDmCd.Suppress = False
    Else
      Rep.FldDet2BagNo.Suppress = True
      Rep.FldDet2ByLoc.Suppress = True
      Rep.FldDet2ByWrk.Suppress = True
      Rep.FldDet2Kt.Suppress = True
      Rep.FldDet2TdRmCtg.Suppress = True
      Rep.FldDet2BOdDmCd.Suppress = True
    End If
  End If
  
  CRV_REP.DisplayGroupTree = False
  
  Dim wMetDesc As String
  'wMetDesc = IIF(adc("UoRmCtgFr") = "G", "Gld", IIF(adc("UoRmCtgFr") = "P", "Plt", IIF(adc("UoRmCtgFr") = "S", "Slv", IIF(adc("UoRmCtgFr") = "L", "Pld", ""))))
  wMetDesc = "Met"
  'Or UCase(adc.MenuCd) = UCase("RepMetLsFg")
  If UCase$(adc("UOYN2")) = "Y" Then
    GRP_REP.SetFormula Rep, "g4_PrnIoGldWt", "If Round({@g4_PrnGldWt}, 3) <> 0 Then {@g4_PrnGldWt} + {@g4_wLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g4_PrnIoPureGldWt", "If Round({@g4_PrnPureGldWt}, 3) <> 0 Then {@g4_PrnPureGldWt} + {@g4_wPureLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g3_PrnIoGldWt", "If Round({@g3_PrnGldWt}, 3) <> 0 Then {@g3_PrnGldWt} + {@g3_wLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g3_PrnIoPureGldWt", "If Round({@g3_PrnPureGldWt}, 3) <> 0 Then {@g3_PrnPureGldWt} + {@g3_wPureLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g2_PrnIoGldWt", "If Round({@g2_PrnGldWt}, 3) <> 0 Then {@g2_PrnGldWt} + {@g2_wLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g2_PrnIoPureGldWt", "If Round({@g2_PrnPureGldWt}, 3) <> 0 Then {@g2_PrnPureGldWt} + {@g2_wPureLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g1_PrnIoGldWt", "If Round({@g1_PrnGldWt}, 3) <> 0 Then {@g1_PrnGldWt} + {@g1_wLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g1_PrnIoPureGldWt", "If Round({@g1_PrnPureGldWt}, 3) <> 0 Then {@g1_PrnPureGldWt} + {@g1_wPureLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g_PrnIoGldWt", "If Round({@g_PrnGldWt}, 3) <> 0 Then {@g_PrnGldWt} + {@g_wLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g_PrnIoPureGldWt", "If Round({@g_PrnPureGldWt}, 3) <> 0 Then {@g_PrnPureGldWt} + {@g_wPureLsWt} Else 0"
    
    
    
    GRP_REP.SetFormula Rep, "g4_PrnIoGldWt1", "{@g4_PrnGldWt}"
    GRP_REP.SetFormula Rep, "g3_PrnIoGldWt1", "{@g3_PrnGldWt}"
    GRP_REP.SetFormula Rep, "g2_PrnIoGldWt1", "{@g2_PrnGldWt}"
    GRP_REP.SetFormula Rep, "g1_PrnIoGldWt1", "{@g1_PrnGldWt}"
    GRP_REP.SetFormula Rep, "g_PrnIoGldWt1", "{@g_PrnGldWt}"
    
    
   
    
    
    ' **** Zubin 211 **** '
    Rep.h1InOutGldWt.SetText "In " + wMetDesc + " Wt"  '"In Gld Wt"
    Rep.h2InOutGldWt.SetText "In " + wMetDesc + " Wt"  '"In Gld Wt"
    ' **** Zubin 211 **** '
     Rep.h1InOutGldWt1.SetText "Out " + wMetDesc + " Wt"
    Rep.h2InOutGldWt1.SetText "Out " + wMetDesc + " Wt"
    
  ElseIf UCase$(adc("UOYN2")) = "N" Then
    GRP_REP.SetFormula Rep, "g4_PrnIoGldWt", "{@g4_PrnGldWt}"
    GRP_REP.SetFormula Rep, "g4_PrnIoPureGldWt", "{@g4_PrnPureGldWt}"
    GRP_REP.SetFormula Rep, "g3_PrnIoGldWt", "{@g3_PrnGldWt}"
    GRP_REP.SetFormula Rep, "g3_PrnIoPureGldWt", "{@g3_PrnPureGldWt}"
    GRP_REP.SetFormula Rep, "g2_PrnIoGldWt", "{@g2_PrnGldWt}"
    GRP_REP.SetFormula Rep, "g2_PrnIoPureGldWt", "{@g2_PrnPureGldWt}"
    GRP_REP.SetFormula Rep, "g1_PrnIoGldWt", "{@g1_PrnGldWt}"
    GRP_REP.SetFormula Rep, "g1_PrnIoPureGldWt", "{@g1_PrnPureGldWt}"
    GRP_REP.SetFormula Rep, "g_PrnIoGldWt", "{@g_PrnGldWt}"
    GRP_REP.SetFormula Rep, "g_PrnIoPureGldWt", "{@g_PrnPureGldWt}"
    
    
    GRP_REP.SetFormula Rep, "g4_PrnIoGldWt1", "If Round({@g4_PrnGldWt}, 3) <> 0 Then {@g4_PrnGldWt} + {@g4_wLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g3_PrnIoGldWt1", "If Round({@g3_PrnGldWt}, 3) <> 0 Then {@g3_PrnGldWt} + {@g3_wLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g2_PrnIoGldWt1", "If Round({@g2_PrnGldWt}, 3) <> 0 Then {@g2_PrnGldWt} + {@g2_wLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g1_PrnIoGldWt1", "If Round({@g1_PrnGldWt}, 3) <> 0 Then {@g1_PrnGldWt} + {@g1_wLsWt} Else 0"
    GRP_REP.SetFormula Rep, "g_PrnIoGldWt1", "If Round({@g_PrnGldWt}, 3) <> 0 Then {@g_PrnGldWt} + {@g_wLsWt} Else 0"
    
    
    
'    If UCase(adc.MenuCd) = UCase("RepMetLsFg") Then
'      Rep.h1InOutGldWt.SetText wMetDesc + " Wt"
'      Rep.h2InOutGldWt.SetText wMetDesc + " Wt"
'      Rep.h1PureGldWt.SetText "Pure " + wMetDesc + " Wt"
'      Rep.h2PureGldWt.SetText "Pure " + wMetDesc + " Wt"
'    Else
      ' **** Zubin 211 **** '
      Rep.h1InOutGldWt.SetText "Out " + wMetDesc + " Wt"  '"Out Gld Wt"
      Rep.h2InOutGldWt.SetText "Out " + wMetDesc + " Wt"  '"Out Gld Wt"
      ' **** Zubin 211 **** '
'    End If
        
     Rep.h1InOutGldWt1.SetText "In " + wMetDesc + " Wt"
      Rep.h2InOutGldWt1.SetText "In " + wMetDesc + " Wt"

  End If
  
'*** Jay 2.13(CT) ***
  If adc("UoBagTyp") = "F" Then
    Rep.TxtBagTypAndChr.SetText "(Flute Bags" & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
  ElseIf adc("UoBagTyp") = "P" Then
    Rep.TxtBagTypAndChr.SetText "(Primary Bags" & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
  ElseIf adc("UoBagTyp") = "C" Then
    Rep.TxtBagTypAndChr.SetText "(Component Bags" & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
  End If
'  '***(Jen 2.11)
'  If adc("UoBagTyp") <> "" Then Rep.TxtBagTypAndChr.SetText IIF(adc("UoBagTyp") = "Y", "(Flute Bags", "(Primary Bags") & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
'  '***(Jen 2.11)
'*** Jay 2.13(CT) ***

' ***** Manali 3.6.0 - 01/09/09 - Gold Loss Report Name and Menu Code Changed To Metal Loss For Users
Rep.TxtHead.SetText moCn.GetFldVal("Select PDesc from Param where PTyp='RMCTG' and PMCd='" + adc("UoRmCtgFr") + "'") + " Loss Report" + IIF(UCase(adc.MenuCd) = UCase("RepMetLsFg"), " (FG)", "")

ShowRepSummFld Rep.Fldg1AllowLsPer, Rep.Fldg2AllowLsPer, Rep.Fldg3AllowLsPer, Rep.FldgAllowLsPer, GRP_REP.GrpOpts, "Process,(Process)Worker", , True


Call DispReport(adc) '****** Sachin 3.02 - Display Report ******

End Sub

' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  'Select Case .Tag
  Select Case Index
  Case Is = 0       '"ShowBoxDet"    ' Index = 0
'    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
    If .Value = Checked Then
      If UCase(adc.MenuCd) = UCase("RepMetLsFg") Then
        If adc("UoYN2") = "N" And adc("UoYN3") = "Y" Then
          adc("UoYN") = "Y"
        Else
          .Value = Unchecked: adc("UoYN") = "N": DispMsg "Box Detail Can Be Viewed When Out Wt and Show Loss% Options Selected", etError: Exit Sub
        End If
      Else
        adc("UoYN") = "Y"
      End If
    Else
      adc("UoYN") = "N"
    End If
    
  Case Is = 3       '"ShowLs"        ' Index = 1
    If .Value = Checked Then
      If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
          "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
          "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
          "(Select PValue1 from Param where " + _
          "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ) )") Then
          .Value = Unchecked: adc("UoYN3") = "N": DispMsg "Enter SEO Password to View Loss %", etError: Exit Sub
       Else
          adc("UoYN3") = "Y"
       End If
    Else
      If UCase(adc.MenuCd) = UCase("RepMetLsFg") Then
        If adc("UoYN") = "Y" Then
          .Value = Checked: adc("UoYN3") = "Y": DispMsg "Show Loss % Should Be Selected When Box Detail Option is selected", etError: Exit Sub
        Else
          adc("UoYN3") = "N"
        End If
      Else
        adc("UoYN3") = "N"
      End If
    End If
    
  Case Is = 5
    If .Value = Checked Then
      If adc("UoYN1") = "" Then
        adc("UoYN5") = "Y"
      Else
        .Value = Unchecked: adc("UoYN5") = "N": DispMsg "Allowed Loss % Cannot Be Viewed When Recovery Option Selected", etError: Exit Sub
      End If
    Else
      adc("UoYN5") = "N"
    End If
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
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = False
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
    Set adc.FirNKeyCtl = adc("UoCoCdFr")
  Else
    Set adc.FirNKeyCtl = adc("UODTFR")
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  Call InitProp(Me)

  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
      adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
      LblPrntCd.Visible = True: adc("UOPRNTCDFR").Visible = True: adc("UOPRNTCDTO").Visible = True: adc("UOPRNTCDSEL").Visible = True
  Else
      adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
      LblPrntCd.Visible = False: adc("UOPRNTCDFR").Visible = False: adc("UOPRNTCDTO").Visible = False: adc("UOPRNTCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_LocTyp = "'P'"
  gb_RmCtgFor = True
  'gb_RmZ = False
  gs_DmTcTyp = "DM"
  'gs_TxnTcTyp = "BV"
  gs_OmTcTyp = "SO"
  gs_CmCtg = "C"
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  ' **** Manali 3.8.0
  If UCase(adc.MenuCd) = UCase("RepMetLsFg") Then
    LblPrntCd.Visible = False: adc("UOPRNTCDFR").Enabled = False: adc("UOPRNTCDFR").Visible = False
    adc("UOPRNTCDTo").Enabled = False: adc("UOPRNTCDTo").Visible = False
    adc("UOPRNTCDSel").Enabled = False: adc("UOPRNTCDSel").Visible = False
    LblProc.Visible = False: adc("UOISSLOCFR").Enabled = False: adc("UOISSLOCFr").Visible = False
    adc("UOISSLOCTo").Enabled = False: adc("UOISSLOCTo").Visible = False
    adc("UOISSLOCSel").Enabled = False: adc("UOISSLOCSel").Visible = False
    LblWrk.Visible = False: adc("UOWRKFR").Enabled = False: adc("UOWRKFR").Visible = False
    adc("UOWRKTo").Enabled = False: adc("UOWRKTo").Visible = False
    adc("UOWRKSel").Enabled = False: adc("UOWRKSel").Visible = False
    LblBagTyp.Visible = False: adc("UOBAGTYP").Enabled = False: adc("UOBAGTYP").Visible = False
    LblUoBChrFr.Visible = False: adc("UOBCHRFR").Enabled = False: adc("UOBCHRFR").Visible = False
    'LblShowInWt.Visible = False: FraOptIOWt.Enabled = False: FraOptIOWt.Visible = False
    LblRepFmt.Visible = False: FraOptRepFmt.Enabled = False: FraOptRepFmt.Visible = False
    'ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False
    LblDt.Caption = "Fg Date"
    ChkBoxArr(5).Enabled = False: ChkBoxArr(5).Visible = False
    
    OptLoc(0).Visible = False:  OptLoc(1).Visible = False
    
 Else
    FraFgScp.Enabled = False: FraFgScp.Visible = False
  End If
  ' **** Manali 3.8.0
  
  Call SetGroupSort
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_LocTyp = "'P'"
  gb_CoCdFor = False '****** Sachin 2.14.0 - [Multi-Company]

  '*** (Jen 2.11)
  Select Case UCase(IdName)
    Case Is = UCase("UoBChrFr")
      If adc.Mode = xNorm Then
        If adc("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Flute Bag Y/N.": Exit Sub
        HlpList.PMCd "BCHR", "'" + adc("UoBagTyp") + "'": Exit Sub
      End If
    Case Is = UCase("UOISSLOCFR"), UCase("UOISSLOCTO"), UCase("UOISSLOCSEL")
        If adc("UOLOCTYPFR") = "R" Then gs_LocTyp = "'R'"
        If adc("UOLOCTYPFR") = "P" Then gs_LocTyp = "'P'"
     
  End Select
  '*** (Jen 2.11)

  SetProp Me, IdName, When
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gb_CoCdFor = False '****** Sachin 2.14.0 - [Multi-Company]
  
  '*** (Jen 2.14)  (This code is required if the value in the field can be 'Y' only after entering SEO Pwd)
  Select Case UCase(IdName)
  Case Is = UCase("UoYN3")
    If pv_NewValue = "Y" Then
      If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
          "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
          "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
          "(Select PValue1 from Param where " + _
          "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ) )") Then _
          Cancel = True: ErrMsg = "Enter SEO Password to View Loss %": Exit Sub
    End If
  End Select
  '*** (Jen 2.14)  (This code is required if the value in the field can be 'Y' only after entering SEO Pwd)
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
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

Private Sub OptBag_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN4") = "N"
Case Is = 1
  adc("UoYN4") = "Y"
Case Is = 2
  adc("UoYN4") = ""
End Select

End Sub

Private Sub OptBag_GotFocus(Index As Integer)
  DispMsg FraOptBag.ToolTipText, etInfo
End Sub

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptIOWt_Click(Index As Integer)
Select Case Index
Case Is = 0
  If UCase(adc.MenuCd) = UCase("RepMetLsFg") Then
    If adc("UoYN") = "Y" Then
      OptIOWt(1).Value = True: adc("UoYN2") = "N": DispMsg "In Wt Cannot Be Selected When Box Detail Option is Selected", etError: Exit Sub
    Else
      adc("UoYN2") = "Y"
    End If
  Else
    adc("UoYN2") = "Y"
  End If
Case Is = 1
  adc("UoYN2") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptIOWt_GotFocus(Index As Integer)
  DispMsg FraOptIOWt.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

Private Sub OptLoc_GotFocus(Index As Integer)
Select Case Index
Case Is = 0
  adc("UOISSLOCFR") = "":   adc("UOISSLOCTO") = "":   adc("UOISSLOCsel") = ""
  adc("UoLocTypFr") = "P"
Case Is = 1
  adc("UOISSLOCFR") = "":   adc("UOISSLOCTO") = "":   adc("UOISSLOCsel") = ""
  adc("UoLocTypFr") = "R"
End Select

End Sub

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptRepFmt_Click(Index As Integer)
Select Case Index
Case Is = 0
  If adc("UoYN5") = "Y" Then OptRepFmt(2).Value = True: adc("UoYN1") = "": DispMsg "Recovery Option Cannot Be Selected When Show Ls% Option Selecte", etError: Exit Sub
  adc("UoYN1") = "Y"
Case Is = 1
  If adc("UoYN5") = "Y" Then OptRepFmt(2).Value = True: adc("UoYN1") = "": DispMsg "Recovery Option Cannot Be Selected When Show Ls% Option Selecte", etError: Exit Sub
  adc("UoYN1") = "N"
Case Is = 2
  adc("UoYN1") = ""
End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptRepFmt_GotFocus(Index As Integer)
  DispMsg FraOptRepFmt.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

Private Sub TxtPwd_Validate(Cancel As Boolean)
  
  '*** (Jen 2.14) (This code is required if the value in the field can be 'Y' only after entering SEO Pwd)
  Dim ws_VldPwd As String
  ws_VldPwd = moCn.GetFldVal("select isnull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + adc.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ")
  If Trim(TxtPwd.text) = "" And ws_VldPwd <> "" Then adc("UoYN3") = "N": ChkBoxArr(3).Value = Unchecked: Exit Sub
  If Trim(TxtPwd.text) <> "" And Trim(TxtPwd.text) <> ws_VldPwd Then
      Cancel = True:  DispMsg "Invalid Password / User Does Not Have Seo Rights ", etError
  End If
  '*** (Jen 2.14) (This code is required if the value in the field can be 'Y' only after entering SEO Pwd)
  
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
'    LblPrimFluteBags.Visible = False
'    adc("UoBagTyp").Visible = False
'    LblUoBChrFr.Visible = False
'    adc("UoBChrFr").Visible = False
'    adc("UoBagTyp") = "": adc("UoBChrFr") = ""
'  End If
'  '*** (Jen 2.11)
'*** Jay 2.13(CT) ***
  
  '*** (Jen 2.14)
  adc("UoYN3") = "N"   '*** (This code is required if the value in the field can be 'Y' only after entering SEO Pwd)
  'If adc("UoYN3") = "" Then adc("UoYN3") = "N"
  '*** (Jen 2.14)
  
  If adc("UoYN") = "" And adc("UoYN").Visible = True Then adc("UoYN") = "N"
  adc("UoYN5") = "N"
  
  ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  If adc("UoYN1") = "Y" Then
    OptRepFmt(0).Value = True
  ElseIf adc("UoYN1") = "N" Then
    OptRepFmt(1).Value = True
  ElseIf adc("UoYN1") = "" Then
    OptRepFmt(2).Value = True
  End If
  
  If adc("UoYN2") = "Y" Then
    OptIOWt(0).Value = True
  ElseIf adc("UoYN2") = "N" Then
    OptIOWt(1).Value = True
'  ElseIf adc("UoYN2") = "" Then
'    OptIOWt(1).Value = True
  End If
  ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  
  If adc("UoYN4") = "Y" Then
    OptBag(0).Value = True
  ElseIf adc("UoYN4") = "N" Then
    OptBag(1).Value = True
  ElseIf adc("UoYN4") = "" Then
    OptBag(2).Value = True
  End If
  
  ' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN3") = "Y" Then ChkBoxArr(3).Value = Checked Else ChkBoxArr(3).Value = Unchecked
  If adc("UoYN5") = "Y" Then ChkBoxArr(5).Value = Checked Else ChkBoxArr(5).Value = Unchecked
  ' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
  
  If adc("UoDt1Fr").Visible = False Or adc("UoDt1Fr") = "" Then adc("UoDt1Fr") = "01/01/80"
  If adc("UoDt1To").Visible = False Or adc("UoDt1To") = "" Then adc("UoDt1To") = "01/01/80"
  If adc("UOLOCTYPFR") = "P" Then
    OptLoc(0).Value = True
  ElseIf adc("UOLOCTYPFR") = "R" Then
    OptLoc(1).Value = True
  End If

End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = True, gs_CoCd, adc("UOCOCDFR")) + "' and " + _
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
  If ws_MultiCoMod = True Then
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdTo.SetText adc("UOCOCDTO"): Rep.TxtCoCdSel.SetText adc("UOCOCDSEL")
  Else
      Rep.TxtCoCdTo.Suppress = True: Rep.TxtCoCdSel.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True: Rep.TxtCoCdSelLbl.Suppress = True
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
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
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
  Cancel = atxt(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
  '*** (09/08/05)
  Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
  Set Rep = Nothing
  '*** (09/08/05)
End Sub

Private Sub GetCnd()
If UCase(adc.MenuCd) = UCase("RepMetLsFg") Then
  adc("UoCoCdFr").CmpStr = "FdCoCd >=": adc("UoCoCdTo").CmpStr = "FdCoCd <="
  adc("UoCoCdSel").CmpStr = "FdCoCd In"
  adc("UoDtFr").CmpStr = " FdDt >=": adc("UoDtTo").CmpStr = " FdDt <="
  adc("UoDtSel").CmpStr = " FdDt In "
  adc("UoInTcFr").CmpStr = " IfTc >=": adc("UoInTcTo").CmpStr = " IfTc <="
  adc("UoInYyFr").CmpStr = " IfYy >=": adc("UoInYyTo").CmpStr = " IfYy <="
  adc("UoInChrFr").CmpStr = " IfChr >=": adc("UoInChrTo").CmpStr = " IfChr <="
  adc("UoInNoFr").CmpStr = " IfNo >=": adc("UoInNoTo").CmpStr = " IfNo <="
  adc("UoFgTcFr").CmpStr = " FdTc >=": adc("UoFgTcTo").CmpStr = " FdTc <="
  adc("UoFgYyFr").CmpStr = " FdYy >=": adc("UoFgYyTo").CmpStr = " FdYy <="
  adc("UoFgChrFr").CmpStr = " FdChr >=": adc("UoFgChrTo").CmpStr = " FdChr <="
  adc("UoFgNoFr").CmpStr = " FdNo >=": adc("UoFgNoTo").CmpStr = " FdNo <="
  adc("UoFdSrFr").CmpStr = " FdSr >=": adc("UoFdSrTo").CmpStr = " FdSr <="
  adc("UoDt1Fr").CmpStr = " InDt >=": adc("UoDt1To").CmpStr = " InDt <="
  adc("UoPrntCdFr").CmpStr = "": adc("UoPrntCdTo").CmpStr = "": adc("UoPrntCdSel").CmpStr = ""
  adc("UoIssLocFr").CmpStr = "": adc("UoIssLocTo").CmpStr = "": adc("UoIssLocSel").CmpStr = ""
  adc("UoWrkFr").CmpStr = "": adc("UoWrkTo").CmpStr = "": adc("UoWrkSel").CmpStr = ""
Else
  adc("UoInTcFr").CmpStr = "": adc("UoInTcTo").CmpStr = ""
  adc("UoInYyFr").CmpStr = "": adc("UoInYyTo").CmpStr = ""
  adc("UoInChrFr").CmpStr = "": adc("UoInChrTo").CmpStr = ""
  adc("UoInNoFr").CmpStr = "": adc("UoInNoTo").CmpStr = ""
  adc("UoFgTcFr").CmpStr = "": adc("UoFgTcTo").CmpStr = ""
  adc("UoFgYyFr").CmpStr = "": adc("UoFgYyTo").CmpStr = ""
  adc("UoFgChrFr").CmpStr = "": adc("UoFgChrTo").CmpStr = ""
  adc("UoFgNoFr").CmpStr = "": adc("UoFgNoTo").CmpStr = ""
  adc("UoFdSrFr").CmpStr = "": adc("UoFdSrTo").CmpStr = ""
  adc("UoDt1Fr").CmpStr = "": adc("UoDt1To").CmpStr = ""

End If
End Sub


