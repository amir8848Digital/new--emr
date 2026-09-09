VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpJtVchAna 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher Analysis"
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
   Begin MwfCtl.MWCTL_BTN1 CmdSc 
      Height          =   390
      Left            =   2280
      TabIndex        =   139
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9680
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   688
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Det S&cope"
      BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11640
      TabIndex        =   88
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
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5745
      TabIndex        =   92
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   93
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
         TabIndex        =   94
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
      Left            =   6960
      TabIndex        =   91
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
      TabIndex        =   87
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
         Left            =   15
         TabIndex        =   79
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpJtVchAna.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpJtVchAna.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   84
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   85
               Top             =   360
               Width           =   5595
               _ExtentX        =   9869
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   80
            Top             =   330
            Width           =   15105
            Begin VB.Frame FraShowProfYN 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1680
               TabIndex        =   148
               ToolTipText     =   "Option to Show Profit Yes/No"
               Top             =   8715
               Width           =   1665
               Begin VB.OptionButton OptProfYN 
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
                  Left            =   900
                  TabIndex        =   150
                  ToolTipText     =   "Option to Show Profit Yes/No"
                  Top             =   0
                  Width           =   645
               End
               Begin VB.OptionButton OptProfYN 
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
                  TabIndex        =   149
                  ToolTipText     =   "Option to Show Profit Yes/No"
                  Top             =   0
                  Width           =   735
               End
            End
            Begin VB.Frame FraRmScp 
               BorderStyle     =   0  'None
               Height          =   1710
               Left            =   0
               TabIndex        =   123
               Top             =   6435
               Width           =   15105
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   35
                  Left            =   4800
                  TabIndex        =   61
                  ToolTipText     =   "Enter Raw Material Code To Range"
                  Top             =   570
                  Width           =   2505
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  MaxLength       =   16
                  DataField       =   "UORMCDTO"
                  IdName          =   "UORMCDTO"
                  CmpStr          =   "OrRmCd <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   38
                  Left            =   7860
                  TabIndex        =   62
                  ToolTipText     =   "Enter Raw Material Code Selection"
                  Top             =   570
                  Width           =   7170
                  _ExtentX        =   12647
                  _ExtentY        =   503
                  MaxLength       =   255
                  DataField       =   "UORMCDSEL"
                  IdName          =   "UORMCDSEL"
                  CmpStr          =   "OrRmCd In"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   45
                  Left            =   4800
                  TabIndex        =   58
                  ToolTipText     =   "Enter Raw Material Sub Category To Range"
                  Top             =   285
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataField       =   "UORMSCTGTO"
                  IdName          =   "UORMSCTGTO"
                  CmpStr          =   "OrRmSCtg <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   46
                  Left            =   7860
                  TabIndex        =   59
                  ToolTipText     =   "Enter Raw Material Sub Category Selection"
                  Top             =   285
                  Width           =   7170
                  _ExtentX        =   12647
                  _ExtentY        =   503
                  MaxLength       =   255
                  DataField       =   "UORMSCTGSEL"
                  IdName          =   "UORMSCTGSEL"
                  CmpStr          =   "OrRmSCtg In"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   47
                  Left            =   1755
                  TabIndex        =   54
                  ToolTipText     =   "Enter From Raw Material Category"
                  Top             =   0
                  Width           =   465
                  _ExtentX        =   820
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataField       =   "UORMCTGFR"
                  IdName          =   "UORMCTGFR"
                  CmpStr          =   "OrRmCtg >= "
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   50
                  Left            =   1755
                  TabIndex        =   57
                  ToolTipText     =   "Enter Raw Material Sub Category From Range"
                  Top             =   285
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataField       =   "UORMSCTGFR"
                  IdName          =   "UORMSCTGFR"
                  Data            =   "ABCDW"
                  CmpStr          =   "OrRmSCtg >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   51
                  Left            =   1755
                  TabIndex        =   60
                  ToolTipText     =   "Enter Raw Material Code From Range"
                  Top             =   570
                  Width           =   2505
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  MaxLength       =   16
                  DataField       =   "UORMCDFR"
                  IdName          =   "UORMCDFR"
                  CmpStr          =   "OrRmCd >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   52
                  Left            =   4800
                  TabIndex        =   64
                  ToolTipText     =   "Enter To Sale Rate"
                  Top             =   855
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  MaxLength       =   0
                  DataType        =   2
                  DataField       =   "UOSALRTTO"
                  IdName          =   "UOSALRTTO"
                  CmpStr          =   "OrSalRt <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   78
                  Left            =   1755
                  TabIndex        =   63
                  ToolTipText     =   "Enter From Sale Rate"
                  Top             =   855
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  MaxLength       =   0
                  DataType        =   2
                  DataField       =   "UOSALRTFR"
                  IdName          =   "UOSALRTFR"
                  CmpStr          =   "OrSalRt >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   79
                  Left            =   4800
                  TabIndex        =   55
                  ToolTipText     =   "Enter To Raw Material Category"
                  Top             =   0
                  Width           =   465
                  _ExtentX        =   820
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataField       =   "UORMCTGTO"
                  IdName          =   "UORMCTGTO"
                  CmpStr          =   "OrRmCtg <= "
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   81
                  Left            =   7860
                  TabIndex        =   56
                  ToolTipText     =   "Enter Raw Material Category Selection"
                  Top             =   0
                  Width           =   7170
                  _ExtentX        =   12647
                  _ExtentY        =   503
                  MaxLength       =   255
                  DataField       =   "UORMCTGSEL"
                  IdName          =   "UORMCTGSEL"
                  CmpStr          =   "OrRmCtg In "
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   53
                  Left            =   4800
                  TabIndex        =   66
                  ToolTipText     =   "Enter To Raw Material Len/Sv"
                  Top             =   1140
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#0.000"
                  MaxLength       =   6
                  DataType        =   2
                  DataField       =   "UOLN1TO"
                  IdName          =   "UOLN1TO"
                  CmpStr          =   "OrLn1 <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   54
                  Left            =   1755
                  TabIndex        =   65
                  ToolTipText     =   "Enter From Raw Material Len/Sv"
                  Top             =   1140
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#0.000"
                  MaxLength       =   6
                  DataType        =   2
                  DataField       =   "UOLN1FR"
                  IdName          =   "UOLN1FR"
                  CmpStr          =   "OrLn1 >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   56
                  Left            =   4800
                  TabIndex        =   68
                  ToolTipText     =   "Enter To Raw Material Breadth"
                  Top             =   1425
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  MaxLength       =   5
                  DataField       =   "UOLN2TO"
                  IdName          =   "UOLN2TO"
                  CmpStr          =   "OrLn2 <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   57
                  Left            =   1755
                  TabIndex        =   67
                  ToolTipText     =   "Enter From Raw Material Breadth"
                  Top             =   1425
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  MaxLength       =   5
                  DataField       =   "UOLN2FR"
                  IdName          =   "UOLN2FR"
                  CmpStr          =   "OrLn2 >="
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Len/Sv"
                  BeginProperty Font 
                     Name            =   "Arial"
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
                  Left            =   90
                  TabIndex        =   134
                  Top             =   1140
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Breadth"
                  BeginProperty Font 
                     Name            =   "Arial"
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
                  Left            =   90
                  TabIndex        =   133
                  Top             =   1425
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Rm Rate"
                  BeginProperty Font 
                     Name            =   "Arial"
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
                  Left            =   90
                  TabIndex        =   127
                  Top             =   855
                  Width           =   1425
               End
               Begin VB.Label ALBL 
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
                  Index           =   25
                  Left            =   90
                  TabIndex        =   126
                  ToolTipText     =   "Location"
                  Top             =   0
                  Width           =   1425
               End
               Begin VB.Label ALBL 
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
                  Index           =   24
                  Left            =   90
                  TabIndex        =   125
                  Top             =   285
                  Width           =   1425
               End
               Begin VB.Label ALBL 
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
                  Index           =   23
                  Left            =   90
                  TabIndex        =   124
                  Top             =   570
                  Width           =   1425
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Include Sec SalCtg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   460
               Index           =   0
               Left            =   60
               TabIndex        =   43
               Tag             =   "SalCtg2,3"
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   4830
               Width           =   1905
            End
            Begin VB.Frame FraLabScp 
               BorderStyle     =   0  'None
               Height          =   870
               Left            =   0
               TabIndex        =   135
               Top             =   6435
               Width           =   15105
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   58
                  Left            =   4800
                  TabIndex        =   75
                  ToolTipText     =   "Enter To Labour Rate"
                  Top             =   570
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  MaxLength       =   0
                  DataType        =   2
                  DataField       =   "UOCSTRTTO"
                  IdName          =   "UOCSTRTTO"
                  CmpStr          =   "OlSalRt <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   59
                  Left            =   1755
                  TabIndex        =   74
                  ToolTipText     =   "Enter From Labour Rate"
                  Top             =   570
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  MaxLength       =   0
                  DataType        =   2
                  DataField       =   "UOCSTRTFR"
                  IdName          =   "UOCSTRTFR"
                  CmpStr          =   "OlSalRt >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   60
                  Left            =   4800
                  TabIndex        =   70
                  ToolTipText     =   "Enter To Labour Main Code"
                  Top             =   0
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataField       =   "UOPMCDTO"
                  IdName          =   "UOPMCDTO"
                  CmpStr          =   "OlMCd <= "
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   61
                  Left            =   1755
                  TabIndex        =   69
                  ToolTipText     =   "Enter From Labour Main Code"
                  Top             =   0
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataField       =   "UOPMCDFR"
                  IdName          =   "UOPMCDFR"
                  CmpStr          =   "OlMCd >= "
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   62
                  Left            =   7860
                  TabIndex        =   71
                  ToolTipText     =   "Enter Labour Main Code Selection"
                  Top             =   0
                  Width           =   7170
                  _ExtentX        =   12647
                  _ExtentY        =   503
                  MaxLength       =   255
                  DataField       =   "UOPMCDSEL"
                  IdName          =   "UOPMCDSEL"
                  CmpStr          =   "OlMCd In"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   63
                  Left            =   4800
                  TabIndex        =   73
                  ToolTipText     =   "Enter To Labour Sub Code"
                  Top             =   285
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataField       =   "UOLSCDTO"
                  IdName          =   "UOLSCDTO"
                  CmpStr          =   "OlSCd <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   64
                  Left            =   1755
                  TabIndex        =   72
                  ToolTipText     =   "Enter From Labour Sub Code"
                  Top             =   285
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataField       =   "UOLSCDFR"
                  IdName          =   "UOLSCDFR"
                  CmpStr          =   "OlSCd >="
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lab Rate"
                  BeginProperty Font 
                     Name            =   "Arial"
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
                  Left            =   90
                  TabIndex        =   138
                  Top             =   570
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lab Main Cd"
                  BeginProperty Font 
                     Name            =   "Arial"
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
                  Left            =   90
                  TabIndex        =   137
                  Top             =   0
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lab Sub Cd"
                  BeginProperty Font 
                     Name            =   "Arial"
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
                  Left            =   90
                  TabIndex        =   136
                  Top             =   285
                  Width           =   1425
               End
            End
            Begin VB.Frame FraOptProfPer 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   6435
               TabIndex        =   128
               ToolTipText     =   "Profit Percentage on Sales Price/ Cost price"
               Top             =   8715
               Width           =   1785
               Begin VB.OptionButton OptProfPer 
                  Caption         =   "Cost"
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
                  TabIndex        =   130
                  ToolTipText     =   "Profit Percentage on Sales Price/ Cost price"
                  Top             =   0
                  Width           =   735
               End
               Begin VB.OptionButton OptProfPer 
                  Caption         =   "Sales"
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
                  Left            =   900
                  TabIndex        =   129
                  ToolTipText     =   "Profit Percentage on Sales Price/ Cost price"
                  Top             =   0
                  Width           =   885
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Add Value   "
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
               Left            =   4800
               TabIndex        =   8
               Tag             =   "AddVal3"
               ToolTipText     =   "Check To Add Value"
               Top             =   1680
               Width           =   1410
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Add Value"
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
               Left            =   4800
               TabIndex        =   6
               Tag             =   "AddVal2"
               ToolTipText     =   "Add Value"
               Top             =   1395
               Width           =   1410
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
               Index           =   3
               Left            =   75
               TabIndex        =   78
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   8430
               Width           =   1875
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4800
               TabIndex        =   38
               ToolTipText     =   "Enter To Design Category"
               Top             =   4245
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
               Index           =   0
               Left            =   7860
               TabIndex        =   39
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   4245
               Width           =   7170
               _ExtentX        =   12647
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
               Left            =   1755
               TabIndex        =   37
               ToolTipText     =   "Enter From Design Category"
               Top             =   4245
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
               Left            =   4800
               TabIndex        =   45
               ToolTipText     =   "Enter To Design Code"
               Top             =   5295
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
               Left            =   7860
               TabIndex        =   46
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   5295
               Width           =   7170
               _ExtentX        =   12647
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "OdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1755
               TabIndex        =   44
               ToolTipText     =   "Enter From Design Code"
               Top             =   5295
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
               Left            =   4800
               TabIndex        =   32
               ToolTipText     =   "Enter To Customer"
               Top             =   3675
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
               Left            =   7860
               TabIndex        =   33
               ToolTipText     =   "Enter Customer Selection"
               Top             =   3675
               Width           =   7170
               _ExtentX        =   12647
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
               Left            =   1755
               TabIndex        =   31
               ToolTipText     =   "Enter From Customer"
               Top             =   3675
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
               Index           =   15
               Left            =   4800
               TabIndex        =   52
               ToolTipText     =   "Enter To Karat"
               Top             =   6150
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
               Left            =   7860
               TabIndex        =   53
               ToolTipText     =   "Enter Karat Selection"
               Top             =   6150
               Width           =   7170
               _ExtentX        =   12647
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "IdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   1755
               TabIndex        =   51
               ToolTipText     =   "Enter From Karat"
               Top             =   6150
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
               Index           =   11
               Left            =   4800
               TabIndex        =   41
               ToolTipText     =   "Enter To Design Sales Category"
               Top             =   4530
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGTO"
               IdName          =   "UODMSALCTGTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   7860
               TabIndex        =   42
               ToolTipText     =   "Enter Design Sales Category Selection"
               Top             =   4530
               Width           =   7170
               _ExtentX        =   12647
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMSALCTGSEL"
               IdName          =   "UODMSALCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   1755
               TabIndex        =   40
               ToolTipText     =   "Enter From Design Sales Category"
               Top             =   4530
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGFR"
               IdName          =   "UODMSALCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   4800
               TabIndex        =   48
               ToolTipText     =   "Enter To Design Size"
               Top             =   5580
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
               CmpStr          =   "OdDmSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1755
               TabIndex        =   47
               ToolTipText     =   "Enter From Design Size"
               Top             =   5580
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "OdDmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   4800
               TabIndex        =   50
               ToolTipText     =   "Enter To Suffix"
               Top             =   5865
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOODSFXTO"
               IdName          =   "UOODSFXTO"
               CmpStr          =   "OdSfx <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   1755
               TabIndex        =   49
               ToolTipText     =   "Enter From Suffix"
               Top             =   5865
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOODSFXFR"
               IdName          =   "UOODSFXFR"
               CmpStr          =   "OdSfx >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5685
               TabIndex        =   18
               ToolTipText     =   "Enter To Order Serial"
               Top             =   2250
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
               Index           =   23
               Left            =   2655
               TabIndex        =   16
               ToolTipText     =   "Enter From Order Serial"
               Top             =   2250
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
               Index           =   19
               Left            =   1755
               TabIndex        =   15
               ToolTipText     =   "Enter From Order Number"
               Top             =   2250
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
               Index           =   22
               Left            =   4800
               TabIndex        =   17
               ToolTipText     =   "Enter To Order Number"
               Top             =   2250
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
               Index           =   24
               Left            =   6405
               TabIndex        =   82
               ToolTipText     =   "Show Invoice Or Order Customer ('Y' for Invoice / 'N' For Order)"
               Top             =   1395
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
               Index           =   49
               Left            =   1755
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   540
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
               Index           =   10
               Left            =   2955
               TabIndex        =   11
               ToolTipText     =   "Enter From Order Character"
               Top             =   1965
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
               Index           =   5
               Left            =   2535
               TabIndex        =   10
               ToolTipText     =   "Enter From Order Year"
               Top             =   1965
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
               Index           =   17
               Left            =   5985
               TabIndex        =   14
               ToolTipText     =   "Enter To Order Character"
               Top             =   1965
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
               Index           =   1
               Left            =   5565
               TabIndex        =   13
               ToolTipText     =   "Enter To Order Year"
               Top             =   1965
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
               Index           =   13
               Left            =   4800
               TabIndex        =   12
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1965
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
               Index           =   6
               Left            =   1755
               TabIndex        =   9
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1965
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
               Index           =   55
               Left            =   1755
               TabIndex        =   76
               ToolTipText     =   "Enter Currency Code Whose Invoice You Wish To View Or Leave It Blank If You Wish To View All Invoices In Base Currency"
               Top             =   8145
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   73
               Left            =   4800
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   540
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
               Index           =   75
               Left            =   7860
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   540
               Width           =   7170
               _ExtentX        =   12647
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "OdCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   76
               Left            =   2505
               TabIndex        =   81
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   8430
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
               Index           =   80
               Left            =   6405
               TabIndex        =   83
               ToolTipText     =   "Show invoice Date/ Printng Dat (Yes / No)"
               Top             =   1680
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
               Index           =   3
               Left            =   1755
               TabIndex        =   3
               ToolTipText     =   "Enter Customer Category"
               Top             =   825
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCTGFR"
               IdName          =   "UOCMCTGFR"
               CmpStr          =   "OmCmCtg ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1755
               TabIndex        =   4
               ToolTipText     =   "Enter Tc 1"
               Top             =   1110
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTC"
               ReCalcParent    =   "UOOMTCFR"
               IdName          =   "UOOMTC"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   1755
               TabIndex        =   5
               ToolTipText     =   "Enter Tc 2"
               Top             =   1395
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMTC1"
               ReCalcParent    =   "UOOMTCFR"
               IdName          =   "UOOMTC1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1755
               TabIndex        =   7
               ToolTipText     =   "Enter Tc 3"
               Top             =   1680
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMTC2"
               ReCalcParent    =   "UOOMTCFR"
               IdName          =   "UOOMTC2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   4800
               TabIndex        =   20
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   2535
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
               Left            =   1755
               TabIndex        =   19
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   2535
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
               Index           =   18
               Left            =   4800
               TabIndex        =   30
               ToolTipText     =   "Enter To Po Date"
               Top             =   3390
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODT1TO"
               IdName          =   "UODT1TO"
               CmpStr          =   "OmPoDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   1755
               TabIndex        =   29
               ToolTipText     =   "Enter From Po Date"
               Top             =   3390
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODT1FR"
               IdName          =   "UODT1FR"
               CmpStr          =   "OmPoDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   4800
               TabIndex        =   22
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   2820
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
               Index           =   25
               Left            =   7860
               TabIndex        =   23
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   2820
               Width           =   7170
               _ExtentX        =   12647
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1755
               TabIndex        =   21
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   2820
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
               Index           =   27
               Left            =   6100
               TabIndex        =   27
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   3105
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
               Index           =   29
               Left            =   7860
               TabIndex        =   28
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   3105
               Visible         =   0   'False
               Width           =   7170
               _ExtentX        =   12647
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "OdPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   77
               Left            =   3040
               TabIndex        =   25
               ToolTipText     =   "Enter From Sub Purchase Order "
               Top             =   3105
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
               Index           =   4
               Left            =   9120
               TabIndex        =   132
               ToolTipText     =   "Show Value Or Weight Figures ('Y' for Value / 'N' For Weight)"
               Top             =   8715
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
               Index           =   68
               Left            =   4800
               TabIndex        =   35
               ToolTipText     =   "Enter To Sales Executive"
               Top             =   3960
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
               Index           =   69
               Left            =   7860
               TabIndex        =   36
               ToolTipText     =   "Enter Sales Executive Selection"
               Top             =   3960
               Width           =   7170
               _ExtentX        =   12647
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSALPERSEL"
               IdName          =   "UOSALPERSEL"
               CmpStr          =   "OmCmSalPer In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   1755
               TabIndex        =   34
               ToolTipText     =   "Enter From Sales Executive"
               Top             =   3960
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
               Index           =   65
               Left            =   1755
               TabIndex        =   24
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   3105
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
               Index           =   66
               Left            =   4800
               TabIndex        =   26
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   3105
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
               Index           =   67
               Left            =   2400
               TabIndex        =   145
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   4920
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               IdName          =   "WADDALLSALCTG"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   118
               Left            =   6795
               TabIndex        =   77
               ToolTipText     =   "Enter Divided By Amount "
               Top             =   8160
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "UOVALDIVBY"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   3720
               TabIndex        =   151
               ToolTipText     =   "Show Value Or Weight Figures ('Y' for Value / 'N' For Weight)"
               Top             =   8715
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN5"
               IdName          =   "UOYN5"
            End
            Begin VB.Label LblProfitYN 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Profit"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   147
               Top             =   8715
               Width           =   1185
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Value Divided By"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   4800
               TabIndex        =   146
               Top             =   8160
               Width           =   1785
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
               Index           =   36
               Left            =   2985
               TabIndex        =   144
               Top             =   3105
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
               Left            =   6030
               TabIndex        =   143
               Top             =   3105
               Width           =   105
            End
            Begin VB.Label LblSalExec 
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
               Left            =   90
               TabIndex        =   142
               Top             =   3960
               Width           =   1845
            End
            Begin VB.Label LblProfPer 
               BackStyle       =   0  'Transparent
               Caption         =   "Profit % On"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4770
               TabIndex        =   131
               Top             =   8715
               Width           =   1185
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "PO Dt"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   90
               TabIndex        =   122
               Top             =   3390
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch Dt"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   90
               TabIndex        =   121
               Top             =   2520
               Width           =   1485
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
               Left            =   90
               TabIndex        =   120
               Top             =   2820
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "SUB PO No"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   90
               TabIndex        =   119
               Top             =   3105
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tc 3"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   90
               TabIndex        =   118
               Top             =   1680
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tc 2"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   90
               TabIndex        =   117
               Top             =   1395
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tc 1"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   90
               TabIndex        =   116
               Top             =   1110
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Ctg"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   90
               TabIndex        =   115
               Top             =   825
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               Index           =   175
               Left            =   90
               TabIndex        =   114
               Top             =   8145
               Width           =   1005
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
               Index           =   12
               Left            =   90
               TabIndex        =   113
               Top             =   1965
               Width           =   1785
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
               Index           =   0
               Left            =   2475
               TabIndex        =   112
               Top             =   1965
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
               Index           =   4
               Left            =   2895
               TabIndex        =   111
               Top             =   1965
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
               Index           =   3
               Left            =   5505
               TabIndex        =   110
               Top             =   1965
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
               Left            =   5925
               TabIndex        =   109
               Top             =   1965
               Width           =   105
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
               Index           =   21
               Left            =   90
               TabIndex        =   108
               Top             =   540
               Width           =   1335
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
               Index           =   20
               Left            =   7860
               TabIndex        =   107
               Top             =   60
               Width           =   1575
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
               Index           =   11
               Left            =   90
               TabIndex        =   106
               Top             =   2250
               Width           =   1485
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
               Left            =   5625
               TabIndex        =   105
               Top             =   2250
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
               Left            =   2595
               TabIndex        =   104
               Top             =   2250
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Suffix"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   90
               TabIndex        =   103
               Top             =   5865
               Width           =   1425
            End
            Begin VB.Label LblDmSz 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Size"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   102
               Top             =   5580
               Width           =   1365
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
               Left            =   90
               TabIndex        =   101
               Top             =   6150
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
               Left            =   90
               TabIndex        =   100
               Top             =   3675
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
               Left            =   90
               TabIndex        =   99
               Top             =   5295
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
               Left            =   90
               TabIndex        =   98
               Top             =   4245
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sal Ctg"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   90
               TabIndex        =   97
               Top             =   4530
               Width           =   1365
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
               Left            =   1755
               TabIndex        =   96
               Top             =   60
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
               Left            =   4800
               TabIndex        =   95
               Top             =   60
               Width           =   1275
            End
         End
      End
      Begin VB.Frame FraSc 
         Height          =   2925
         Left            =   1560
         TabIndex        =   140
         Top             =   6480
         Visible         =   0   'False
         Width           =   3945
         Begin MwfCtl.MWCTL_FLX GrdDsgAna 
            Height          =   2430
            Left            =   120
            TabIndex        =   141
            Top             =   240
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   4286
            Cols            =   4
            colname1        =   "UDANASR"
            heading1        =   "Sr. "
            datafld1        =   "UdAnaSr"
            datatype1       =   4
            recalcparent1   =   "GrdDsgAna(UDANACDFR),GrdDsgAna(UDANACDTO)"
            maxlength1      =   2
            tooltiptext1    =   "Enter Analysis Sr. Number"
            colname2        =   "UDANACDFR"
            heading2        =   "Fr Ana Cd      "
            datafld2        =   "UdAnaCdFr"
            datatype2       =   4
            recalcon2       =   "GrdDsgAna(UDANASR)"
            maxlength2      =   8
            tooltiptext2    =   "Enter From Analysis Code"
            colname3        =   "UDANACDTO"
            heading3        =   "To Ana Cd      "
            datafld3        =   "UdAnaCdTo"
            datatype3       =   4
            recalcon3       =   "GrdDsgAna(UDANASR)"
            maxlength3      =   8
            tooltiptext3    =   "Enter To Analysis Code"
         End
      End
   End
End
Attribute VB_Name = "EmrFrpJtVchAna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepJtVchAna
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Dim ws_MultiCoMod As Boolean


Private Sub SetGroupSort()

  If ws_MultiCoMod = True Then _
    GRP_REP.Add "Company Code", "OdCoCd", "OdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=OdCoCd) "
    GRP_REP.Add "Sales Exec", "OmCmSalPer", "", "", "", "", "", "", "(Select PDesc from Param Where PTyp = 'SP' and PMCd= OmCmSalPer) "
  If UCase(ADC.MenuCd) = UCase("RepJtVchAnaRm") Or UCase(ADC.MenuCd) = UCase("RepJtVchAnaLab") Then
  
    GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "Sales Ctg", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= DmCtg) "
    GRP_REP.Add "Sales Ctg", "DmSalCtg", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and PmCd=DmCtg and PSCd= DmSalCtg) "
    GRP_REP.Add "Design Code", "OdDmCd", "", "", "", "Design Ctg,Sales Ctg", "", "", "DmDesc "
    GRP_REP.Add "Design Size", "OdDmSz", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= OdDmSz)"
    GRP_REP.Add "Suffix", "OdSfx", "", "", "", "", "", "", ""
    GRP_REP.Add "Cust/Supp", "OmCmCd", "", "", "", "", "", "", " case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd End "
    GRP_REP.Add "VchNo", "OdCoCd+OdTc+OdYy+OdChr+str(OdNo)", "OdCoCd+'/'+OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))", "", "", "Cust/Supp", "", "", ""
    GRP_REP.Add "Karat", "OdKt", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = OdKt) "
    GRP_REP.Add "(DsgCd)Colour", "OdDmCd+OdDmCol", "'('+OdDmCd+')'+OdDmCol", "", "", "Design Ctg,Sales Ctg,Design Code,Design Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
    GRP_REP.Add "Design Colour", "OdDmCol", "", "", "", "(DsgCd)Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
        
    If UCase(ADC.MenuCd) = UCase("RepJtVchAnaRm") Then
      GRP_REP.Add "Rm Ctg", "OrRmCtg", "", "OrRmCtg", "hRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = OrRmCtg) "
      GRP_REP.Add "Rm SubCtg", "OrRmCtg+OrRmSCtg", "'('+ RTrim(OrRmCtg)+ ') '+ OrRmSCtg", "OrRmCtg,OrRmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = OrRmCtg and PSCd= OrRmSCtg) "
      GRP_REP.Add "Rm Code", "OrRmCtg+OrRmSCtg+OrRmCd", "'('+ RTrim(OrRmCtg)+'/'+ OrRmSCtg+ ') '+ OrRmCd", "OrRmCtg,OrRmSCtg,OrRmCd", "hRmCtg,hRmSCtg,hOrRmCd", "Rm Ctg,Rm SubCtg", "", "", "SubString(RmDesc, 1, 30) "
    End If
    If UCase(ADC.MenuCd) = UCase("RepJtVchAnaLab") Then
      GRP_REP.Add "Lab Main Cd", Array("OlMCd", "'SET'"), "", "OlMCd", "hLabMCd", "", "", "", Array("(Select PDesc From Param Where PTyp= 'LABMCD' and PMCd= OlMCd)", "(Select PDesc From Param Where PTyp= 'LABMCD' and PMCd= 'SET')")
      GRP_REP.Add "Lab Sub Cd", Array("OlMCd+OlSCd", "'SET'+Ltrim(RTrim(OrSetSCd))"), Array("'('+OlMCd+')'+OlSCd", "'(SET)'+OrSetSCd"), "OlMCd,OlSCd", "hLabMCd,hLabSCd", "Lab Main Cd", "", "", Array("(Select PDesc From Param Where PTyp= 'LABSCD' and PMCd= OlMCd and PSCd = OlSCd) ", "(Select PDesc From Param Where PTyp= 'LABSCD' and PMCd= 'SET' and PSCd = OrSetSCd) ")
    End If
        
  Else
    GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "Sales Ctg", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= DmCtg) "
    'uni.4 DmSalCtg replaced by qDmSalCtgRepl,qDmSalCtgDescRepl to get either SalCtg,SalCTg2 or SalCTg3
    GRP_REP.Add "Sales Ctg", "||qDmSalCtgRepl||", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and PmCd=DmCtg and PSCd= ||qDmSalCtgDescRepl|| ) "
    GRP_REP.Add "Design Code", "OdDmCd", "", "OdDmCd", "hOdDmCd", "Design Ctg,Sales Ctg", "", "", "DmDesc "
    GRP_REP.Add "Design Size", "OdDmSz", "", "OdDmSz", "hOdDmSz", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= OdDmSz)"
    GRP_REP.Add "Suffix", "OdSfx", "", "OdSfx", "hOdSfx", "", "", "", ""
    GRP_REP.Add "Cust/Supp", "OmCmCd", "", "", "", "", "", "", " case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd End "
    GRP_REP.Add "VchNo", "OdCoCd+OdTc+OdYy+OdChr+str(OdNo)", "OdCoCd+'/'+OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))", "", "", "Cust/Supp", "", "", ""
    GRP_REP.Add "Karat", "OdKt", "", "OdKt", "hOdKt", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = OdKt) "
    GRP_REP.Add "(DsgCd)Colour", "OdDmCd+OdDmCol", "'('+OdDmCd+')'+OdDmCol", "OdDmCd", "hOdDmCd", "Design Ctg,Sales Ctg,Design Code,Design Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
    GRP_REP.Add "Design Colour", "OdDmCol", "", "", "", "(DsgCd)Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
     
  End If
  
  '3.11.0
    GRP_REP.Add "Inv ExpNo", "IsNull(InExpNo,'')", "", "", "", "VchNo,Cust/Supp", "", "", ""
  
End Sub

Private Sub adc_setreprecsource()
  '*** Report Sql ***
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, InExpNo, DmCtg, IdDmCd, IdSfx, IdDmSz, IdKt, InDt,
  'InCmCd, IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr, DmSalCtg, IdVaCtg,
  'IdQty, IdIFob, IrRmIVal as qMetVal, IrRmIVal as qDCVal, IrRmIVal as qXVal,
  'IrSetIVal as qSetVal, iLiVal as qLabVal, IrRmIWt as qGrWt, IrRmIWt as qMetWt,
  'IrRmIWt As qDCWt, IrRmIWt As qXWt,
  'OdPicNm
  'from InvDsg, InvHd, DsgMst, InvRm, RmMst, InvLab, OrdDsg where 1= 2
  '*** Report Sql ***
  
  Dim grpflds As String, wSqlStrg As String, wStr1 As String
  Dim grpflds0 As String, grpflds1 As String
  Dim wrepcnd As String, wCnd As String, wQtyStr As String
  Dim wTcCnd As String

  Set Rep = Nothing
  Set ADC.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
  
  grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)

  Dim wb_InBaseCur As Boolean, ws_BaseCurCd As String
  '578.22 if currency code not given it takes base currency from head else takes from scope
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + ADC("UoCoCdFr") + "'")
  If ADC("UoCurCdFr") = "" Then
    wb_InBaseCur = True
    Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
  Else
    wb_InBaseCur = False
    Rep.TxtValIn.SetText "Value In " + ADC("UoCurCdFr")
  End If
  
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  
  wTcCnd = IIF(ADC("UoOmTc") <> "", " '" + ADC("UoOmTc") + "' ", "")
  wTcCnd = wTcCnd + IIF(ADC("UoOmTc1") <> "", IIF(wTcCnd <> "", ", ", "") + "'" + ADC("UoOmTc1") + "' ", "")
  wTcCnd = wTcCnd + IIF(ADC("UoOmTc2") <> "", IIF(wTcCnd <> "", ", ", "") + "'" + ADC("UoOmTc2") + "' ", "")
  wTcCnd = IIF(wTcCnd <> "", " and OdTc In (" + wTcCnd + ") ", "")
  
  
  wCnd = wCnd + wTcCnd
  '" and OdTc In ('" + adc("UoOmTc") + "', '" + adc("UoOmTc1") + "', '" + adc("UoOmTc2") + "') "
  wCnd = wCnd + IIF(wb_InBaseCur = False, " and CmCurCd= '" + ADC("UoCurCdFr") + "'", "")
  
  ' ***** Manali 3.10.0 - 17/04/12 - Show Only JT transctions
  wCnd = wCnd + " and (Select PValue From Param Where PTyp='TC' and PMCd=OmTc)='" + ctJtTcTyp + "' "
  
  'uni.4 where condition for SalCtg only for Rm and Lab voucher analysis
  If UCase(ADC.MenuCd) <> UCase("RepJtVchAnaDsg") Then
    If ADC("UoDmSalCtgFr") <> "" Then wCnd = wCnd + " and DmSalCtg>='" + ADC("UoDmSalCtgFr") + "' "
    If ADC("UoDmSalCtgTo") <> "" Then wCnd = wCnd + " and DmSalCtg<='" + ADC("UoDmSalCtgTo") + "' "
    If ADC("UoDmSalCtgSel") <> "" Then wCnd = wCnd + " and DmSalCtg in(" + ADC("UoDmSalCtgSel") + ") "
  End If
  
  '*** To check if 'Detail' group is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 3
    If wGrp(i) = "(Detail)" Then wDetPos = i: Exit For
  Next i
  
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then Rep.PicSec.Suppress = True

  If wTcCnd <> "" Then
    'wTcCnd = IIF(adc("UoOmTc") <> "", " '" + adc("UoOmTc") + "' ", "")
    wQtyStr = IIF(ADC("UoOmTc1") <> "", " (OdTc='" + ADC("UoOmTc1") + "' and '" + ADC("UoYn1") + "'='N') ", "")
    wQtyStr = wQtyStr + IIF(ADC("UoOmTc2") <> "", IIF(wQtyStr <> "", " Or ", "") + " (OdTc='" + ADC("UoOmTc2") + "' and '" + ADC("UoYn2") + "'='N') ", "")

    wQtyStr = IIF(wQtyStr <> "", " (Case When " + wQtyStr + " Then -1 Else 1 End) ", "")

    wQtyStr = IIF(wQtyStr = "", "1", wQtyStr)
    'wQtyStr = " (Case When (OdTc='" + adc("UoOmTc1") + "' and '" + adc("UoYn1") + "'='N') or (OdTc='" + adc("UoOmTc1") + "' and '" + adc("UoYn2") + "'='N') Then -1 Else 1 End) "
  Else
    wQtyStr = IIF(wQtyStr = "", "1", wQtyStr)
  End If
  
  Dim wDmSalCtgCnd As String, ws_DmSalCtg As String
  If UCase(ADC.MenuCd) = UCase("RepJtVchAnaDsg") Then
    'Uni.4 - SalCtg scope will check all of 3 Sal Ctg(SalCtg, SalCtg2 & SalCtg3).
    'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is
    'calling function to get where clause to join DsgMst
    If ADC("UoDmSalCtgFr") <> "" Or ADC("UoDmSalCtgTo") <> "" Or ADC("UoDmSalCtgSel") <> "" Then
      wDmSalCtgCnd = GetDmSalCtgWhrCnd(ADC("UoDmSalCtgFr"), ADC("UoDmSalCtgTo"), ADC("UoDmSalCtgSel"), ADC("wAddAllSalCtg"))
    End If
    
    'uni.4 selecting Dmsalctg either from SalCtg,SalCtg2 or SalCtg3
    'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is
    'calling function to get SalCtg value to show in detail section and group sort
    ws_DmSalCtg = GetDmSalCtg(ADC("UoDmSalCtgFr"), ADC("UoDmSalCtgTo"), ADC("UoDmSalCtgSel"), ADC("wAddAllSalCtg"), "")

  End If
  
  Dim wValDivBy As Double
  'sv.30 when divied by amount is blank then default value is 1
  wValDivBy = IIF(ADC("UoValDivBy") = 0, 1, ADC("UoValDivBy"))
      
  wSqlStrg = "Select " + grpflds + ", " + _
          "OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OmDt, OmCmCd, OmPoNo, OmPoDt, OdPicNm, "

  '3.11.0 - InvHd Join added
  If UCase(ADC.MenuCd) = UCase("RepJtVchAnaDsg") Then
    ' ***** Manali  3.10.0 - 14/07/12 - Lab Report
    'uni.4 wDmSalCtgCnd added to DsgMst join, to get SalCtg value from scope.
    'this is to avoid SalCtg scope condition in where clause
    'sv.30 wValDivBy added to costprice, sale price
    '578.22 conversion factor added to costprice & sale price
    wStr1 = " OdDmCd, OdSfx, OdDmSz, OdDmStkNo, OdDmStkYy, OdDmStkChr, OdPoNo, DmCtg,  " + _
            " Convert(Decimal(14,1), " + wQtyStr + "* OdOrdQty) as OdOrdQty, 0.0 as OdExpQty, " + _
            " Convert(Decimal(14,3), " + IIF(wb_InBaseCur, "(Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc") + "/" + CStr(wValDivBy) + " ) OdSalPrc," + _
            " Convert(Decimal(14,3), " + IIF(wb_InBaseCur, "(Case when OmMulDiv= 'M' then OdCstPrc/OmCnvFct else OdCstPrc*OmCnvFct end) ", "OdCstPrc") + "/" + CStr(wValDivBy) + ") OdCstPrc," + _
            " '' as OrRmCtg, '' as OrRmSCtg, '' as OrRmCd, 0.0 as OrLn1, 0.0 as OrLn2, 0.0 as OrRmPtr, 0 as OrQty, 0.0 as OrWt, 0.0 as qWtInGms, 0.0 as qPureWt, 0.0 as OrSalRt, '' RMQW, " + _
            " 0.0 as OrSalVal, '' as OlMCd, '' as OlSCd, 0.0 as OlSalRt, '' as OlQW,  0 as OlQty, 0.0 as OdLabAsWt, 0.0 as OlSalVal, IsNull(InExpNo,'') as InExpNo " + _
            " From OrdDsg " + _
            " Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo and OmIdNo=OdOmIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
            " Join CustMst On CmCtg=OmCmCtg and CmCd=OmCmCd " + _
            " Left Outer Join InvHd On OdCoCd=InCoCd And OdTc=InTc And OdYy=InYy And Odchr=InChr And OdNo=InNo" + _
            " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + wDmSalCtgCnd + _
            " And (Select Count(*) From UsrOptDsgAna " + _
            "Join DsgAna on " + IIF(gs_Partition = ctCurrPrtn, " DaPrtKey='" + ctCurrPrtn + "' and ", "") + " DaAnaSr = UdAnaSr " + _
            "where (DaAnaCd between UdAnaCdFr and UdAnaCdTo) and DaTcTyp = DmTcTyp and DaCd= DmCd and DaSz= DmSz and " + _
            "UdUsrCd = '" + ADC.UsrCd + "' and UdMnuCd= '" + ADC.MenuCd + "')= " & (GrdDsgAna.Rows - 1) & " " + _
            wCnd + IIF(gs_Partition = ctCurrPrtn, " and OMPrtKey='" + ctCurrPrtn + "' ", "")

  ElseIf UCase(ADC.MenuCd) = UCase("RepJtVchAnaRm") Then
    ' ***** Manali  3.10.0 - 14/07/12 - Lab Report
    'sv.8 used Prdpurity for pure wt calculations, existing RtPurity replaced with PrdPurity
    'sv.30 wValDivBy added to OrSalRt, OrSalVal
    '578.22 conversion factor added to OrSalRt & OrSalVal
    wStr1 = " '' as OdDmCd, '' as OdSfx, '' as OdDmSz, 0 as OdDmStkNo, '' as OdDmStkYy, '' as OdDmStkChr, '' as OdPoNo, '' as DmCtg, 0.0 as OdOrdQty, 0.0 OdSalPrc, 0.0 as OdExpQty, 0.0 as OdCstPrc, " + _
            " OrRmCtg, OrRmSCtg, OrRmCd, OrLn1, OrLn2, OrRmPtr, " + _
            " Convert(int, " + wQtyStr + "* OdOrdQty * OrQty) as OrQty, Convert(Decimal(14,3), " + wQtyStr + "* OdOrdQty * OrWt) as OrWt, " + _
            " Convert(Decimal(14,3), " + wQtyStr + "* OdOrdQty * OrWt/(case when OrRmCtg in ('D', 'C') then 5 else 1 end)) as qWtInGms,   " + _
            " Convert(Decimal(14,3), " + wQtyStr + "* OdOrdQty * OrWt/(case when OrRmCtg in ('D', 'C') then 5 else 1 end) " + _
            " *(case when OrRmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt Else 0 end) )" + _
            "/(case when OrRmCtg in ('G', 'P', 'S', 'L', 'M') then " + _
                          "IsNull((Select RmPurityWt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= OrRmCtg), 1) " + _
            " else 1 end) as qPureWt, " + _
            " Convert(Decimal(14,3), " + IIF(wb_InBaseCur, "(Case when OmMulDiv= 'M' then OrSalRt/OmCnvFct else OrSalRt*OmCnvFct end) ", "OrSalRt") + "/" + CStr(wValDivBy) + ") OrSalRt, RmQW, " + _
            " Convert(Decimal(14,2), " + wQtyStr + "* OdOrdQty *" + IIF(wb_InBaseCur, "(Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + "/" + CStr(wValDivBy) + ") as OrSalVal, " + _
            " '' as OlMCd, '' as OlSCd, 0.0 as OlSalRt, '' as OlQW,  0 as OlQty, 0.0 as OdLabAsWt, 0.0 as OlSalVal " + _
            " From OrdRm " + _
            " Join OrdDsg On OdCoCd=OrCoCd and OdTc=OrTc and OdYy=OrYy and OdChr=orChr and OdNo=OrNo and OdSr=OrSr and OdIdNo=OrOdIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OrPrtKey ", "") + _
            " Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo and OmIdNo=OdOmIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
            " Join CustMst On CmCtg=OmCmCtg and CmCd=OmCmCd " + _
            " Join RmMst On RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + _
            "  Left Outer Join InvHd On OdCoCd=InCoCd And OdTc=InTc And OdYy=InYy And Odchr=InChr And OdNo=InNo " + _
            " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
            wCnd + IIF(gs_Partition = ctCurrPrtn, " and OMPrtKey='" + ctCurrPrtn + "' ", "")
  
  ' ***** Manali  3.10.0 - 14/07/12 - Lab Report
  ElseIf UCase(ADC.MenuCd) = UCase("RepJtVchAnaLab") Then
    
    Dim wStrLab As String, wStrSet As String
    Dim wSqlStrLab1 As String, wSqlStrLab2 As String
    
    wSqlStrLab1 = " Select " + grpflds0 + ", " + _
          " OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OmDt, OmCmCd, OmPoNo, OmPoDt, OdPicNm, "
    
    wSqlStrLab2 = " Select " + grpflds1 + ", " + _
          " OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OmDt, OmCmCd, OmPoNo, OmPoDt, OdPicNm, "
    
    wStr1 = " '' as OdDmCd, '' as OdSfx, '' as OdDmSz, 0 as OdDmStkNo, '' as OdDmStkYy, '' as OdDmStkChr, '' as OdPoNo, '' as DmCtg, 0.0 as OdOrdQty, 0.0 OdSalPrc, 0.0 as OdExpQty, 0.0 as OdCstPrc, " + _
            " '' as OrRmCtg, '' as OrRmSCtg, '' as OrRmCd, 0.0 as OrLn1, 0.0 as OrLn2, 0.0 as OrRmPtr, 0 as OrQty, " + _
            "0.0 as OrWt, 0.0 as qWtInGms, 0.0 as qPureWt, 0.0 as OrSalRt, '' as RmQW , 0.0 as OrSalVal, "
    'sv.30 wValDivBy added to OlSalRt, OlSalVal
    '578.22 conversion factor added to OlSalRt & OlSalVal
    wStrLab = " OlMCd, OlSCd, (" + IIF(wb_InBaseCur, "(Case when OmMulDiv= 'M' then OlSalRt/OmCnvFct else OlSalRt*OmCnvFct end) ", "OlSalRt") + "/" + CStr(wValDivBy) + ") as OlSalRt, OlQW, Convert(int, " + wQtyStr + "* OdOrdQty * OlQty) as OlQty, " + _
            " Convert(Decimal(14,3), " + wQtyStr + "* OdOrdQty * OdLabAsWt) as OdLabAsWt, " + _
            " Convert(Decimal(14,2), " + wQtyStr + "* OdOrdQty * " + IIF(wb_InBaseCur, "(Case when OmMulDiv= 'M' then OlSalVal/OmCnvFct else OlSalVal*OmCnvFct end) ", "OlSalVal") + "/" + CStr(wValDivBy) + ") OlSalVal " + _
            "From OrdLab " + _
            "Join OrdDsg On OdCoCd=OlCoCd and OdTc=OlTc and OdYy=OlYy and OdChr=OlChr and OdNo=OlNo and OdSr=OlSr and OdIdNo=OlOdIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OlPrtKey ", "") + _
            "Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo and OmIdNo=OdOmIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
            "Join CustMst On CmCtg=OmCmCtg and CmCd=OmCmCd " + _
            " Left Outer Join InvHd On OdCoCd=InCoCd And OdTc=InTc And OdYy=InYy And Odchr=InChr And OdNo=InNo " + _
            "join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
            wCnd + IIF(gs_Partition = ctCurrPrtn, " and OMPrtKey='" + ctCurrPrtn + "' ", "")
    'sv.30 wValDivBy added to OrSetSalRt, OrSetSalVal
    '578.22 conversion factor added to OrSetSalRt & OrSetSalVal
    wStrSet = " 'SET' as OlMCd, OrSetSCd as OlSCd, (" + IIF(wb_InBaseCur, "(Case when OmMulDiv= 'M' then OrSetSalRt/OmCnvFct else OrSetSalRt*OmCnvFct end) ", "OrSetSalRt") + "/" + CStr(wValDivBy) + ") as OlSalRt, 'Q' as OlQW, Convert(int, " + wQtyStr + "* OdOrdQty * OrQty) as OlQty, 0.0 as OdLabAsWt, " + _
            " Convert(Decimal(14,2), " + wQtyStr + "* OdOrdQty * " + IIF(wb_InBaseCur, "(Case when OmMulDiv= 'M' then OrSetSalVal/OmCnvFct else OrSetSalVal*OmCnvFct end) ", "OrSetSalVal") + "/" + CStr(wValDivBy) + ")  as OlSalVal " + _
            " From OrdRm " + _
            "   Join OrdDsg On OdCoCd=OrCoCd and OdTc=OrTc and OdYy=OrYy and OdChr=orChr and OdNo=OrNo and OdSr=OrSr and OdIdNo=OrOdIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OrPrtKey ", "") + _
            "   Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo and OmIdNo=OdOmIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
            "   Join CustMst On CmCtg=OmCmCtg and CmCd=OmCmCd " + _
            "   Left Outer Join InvHd On OdCoCd=InCoCd And OdTc=InTc And OdYy=InYy And Odchr=InChr And OdNo=InNo " + _
            " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
            wCnd + IIF(gs_Partition = ctCurrPrtn, " and OMPrtKey='" + ctCurrPrtn + "' ", "") + " and OrRmCtg In ('D', 'C') "
 
  End If
  'uni.4 replacing DmSalCtg value and its description in group sort of Sales Ctg
  wSqlStrg = Replace(wSqlStrg, "||qDmSalCtgRepl||", ws_DmSalCtg)
  wSqlStrg = Replace(wSqlStrg, "||qDmSalCtgDescRepl||", ws_DmSalCtg)
  
  If (UCase(ADC.MenuCd) = UCase("RepJtVchAnaLab")) Then
    ADC.RepRecSource = wSqlStrLab1 + wStr1 + wStrLab + " Union All " + wSqlStrLab2 + wStr1 + wStrSet + _
                  " Order By OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr " + ctMaxDopOpt
  Else
    ADC.RepRecSource = wSqlStrg + wStr1 + ctMaxDopOpt
  End If
  
  If UCase(ADC.MenuCd) = UCase("RepJtVchAnaDsg") Then
    'Rep.HOdDmId.Suppress = False: Rep.HOdDmSz.Suppress = False: Rep.hOdDmCd.Suppress = False
    'Rep.HOdSfx.Suppress = False: Rep.HDmCtg.Suppress = False: Rep.hOdSalPrc.Suppress = False
    'Rep.hOdCstPrc.Suppress = False: Rep.hProfit.Suppress = False: Rep.hProfitPer.Suppress = False
    'Rep.hOdOrdQty.Suppress = False: Rep.hOdSalVal.Suppress = False
    'Rep.gHOdOrdQty.Suppress = False: Rep.ghOdSalVal.Suppress = False
    'Rep.ghProfit.Suppress = False: Rep.ghProfPer.Suppress = False

    Rep.hRmCtg.Suppress = True: Rep.hRmSCtg.Suppress = True: Rep.hRmCd.Suppress = True
    Rep.hOrLn1.Suppress = True: Rep.hOrLn2.Suppress = True: Rep.hOrRmPtr.Suppress = True
    Rep.hOrQty.Suppress = True: Rep.HOrWt.Suppress = True: Rep.hVal.Suppress = True
    Rep.hWtInGms.Suppress = True: Rep.hPureWt.Suppress = True: Rep.hOrSalRt.Suppress = True
    Rep.gHOrQty.Suppress = True: Rep.gHOrWt.Suppress = True: Rep.gHWtInGms.Suppress = True
    Rep.gHPureWt.Suppress = True: Rep.gHVal.Suppress = True
    
    Rep.hLabMCd.Suppress = True: Rep.hLabSCd.Suppress = True      ' ***** Manali  3.10.0 - 14/07/12 - Lab Report
    
    GRP_REP.SetFormula Rep, "wVal", "0"
'    GRP_REP.SetFormula Rep, "wVal", "{rdo.OdSalPrc}*{rdo.OdOrdQty}"
    Dim wProfBase As String
    wProfBase = IIF(ADC("UoYN4") = "Y", "OdCstVal", "OdSalVal")
    'GRP_REP.SetFormula Rep, "wProfitPer", "{@wProfit}/iif(Abs({@w" + wProfBase + "})=0, 1, Abs({@w" + wProfBase + ")})*100"
    GRP_REP.SetFormula Rep, "wProfitPer", "{@wProfit}/iif(Abs({@w" + wProfBase + "})=0, 1, Abs({@w" + wProfBase + "}))*100"
    GRP_REP.SetFormula Rep, "g_ProfPer", "{@g_Prof}/iif({@g_" + wProfBase + "}=0, 1, {@g1_" + wProfBase + "})*100"
    GRP_REP.SetFormula Rep, "g1_ProfPer", "{@g1_Prof}/iif({@g1_" + wProfBase + "}=0, 1, {@g1_" + wProfBase + "})*100"
    GRP_REP.SetFormula Rep, "g2_ProfPer", "{@g2_Prof}/iif({@g2_" + wProfBase + "}=0, 1, {@g2_" + wProfBase + "})*100"
    GRP_REP.SetFormula Rep, "g3_ProfPer", "{@g3_Prof}/iif({@g3_" + wProfBase + "}=0, 1, {@g3_" + wProfBase + "})*100"
    
'    GRP_REP.SetFormula Rep, "g_ProfPer", "0"
'    GRP_REP.SetFormula Rep, "g1_ProfPer", "0"
'    GRP_REP.SetFormula Rep, "g2_ProfPer", "0"
'    GRP_REP.SetFormula Rep, "g3_ProfPer", "0"
    '6.1b if Show profit is NO then cost price, cost value, profit, profit% will be invisible
    If ADC("UoYN5") = "N" Then
      Rep.hOdCstPrc.Suppress = True: Rep.hOdCstVal.Suppress = True: Rep.hProfit.Suppress = True:  Rep.hProfitPer.Suppress = True
      Rep.FldOdCstPrc.Suppress = True: Rep.wOdCstVal1.Suppress = True: Rep.wProfit1.Suppress = True: Rep.wProfitPer.Suppress = True
      Rep.ghOdCstVal.Suppress = True: Rep.ghProfit.Suppress = True: Rep.ghProfPer.Suppress = True
      Rep.g3OdCstVal1.Suppress = True: Rep.g3Prof1.Suppress = True: Rep.g3ProfPer1.Suppress = True
      Rep.g2OdCstVal1.Suppress = True: Rep.g2Prof1.Suppress = True: Rep.g2ProfPer1.Suppress = True
      Rep.g1OdCstVal1.Suppress = True: Rep.g1Prof1.Suppress = True: Rep.g1ProfPer1.Suppress = True
      Rep.gOdCstVal1.Suppress = True: Rep.gProf1.Suppress = True: Rep.gProfPer1.Suppress = True
    End If
    
    If ADC("UoYN3") = "N" Or wDetPos = 0 Then
      gb_ShowPic = False
      Rep.PicSec.Suppress = True
    ElseIf ADC("UoYN3") = "Y" And wDetPos <> 0 Then
      gb_ShowPic = True
      Rep.PicSec.Suppress = False
    End If
    
    Rep.TxtHead.SetText "Voucher Analysis - Designwise (Jwlry)"
  ElseIf UCase(ADC.MenuCd) = UCase("RepJtVchAnaRm") Then
    Rep.HOdDmId.Suppress = True: Rep.hOdDmSz.Suppress = True: Rep.hOdDmCd.Suppress = True
    Rep.hOdSfx.Suppress = True: Rep.hOdSalPrc.Suppress = True   'Rep.hDmCtg.Suppress = True:
    Rep.hOdCstPrc.Suppress = True: Rep.hProfit.Suppress = True: Rep.hProfitPer.Suppress = True
    Rep.hOdOrdQty.Suppress = True: Rep.hOdSalVal.Suppress = True: Rep.hOdCstVal.Suppress = True
    Rep.gHOdOrdQty.Suppress = True: Rep.ghOdSalVal.Suppress = True: Rep.ghOdCstVal.Suppress = True
    Rep.ghProfit.Suppress = True: Rep.ghProfPer.Suppress = True

    'Rep.hRmCtg.Suppress = False: Rep.hRmSCtg.Suppress = False: Rep.HRmCd.Suppress = False
    'Rep.hOrLn1.Suppress = False: Rep.hOrLn2.Suppress = False: Rep.hOrRmPtr.Suppress = False
    'Rep.hOrQty.Suppress = False: Rep.HOrWt.Suppress = False: Rep.hVal.Suppress = False
    'Rep.hWtInGms.Suppress = False: Rep.hPureWt.Suppress = False: Rep.HOrSalRt.Suppress = False
    'Rep.gHOrQty.Suppress = False: Rep.gHOrWt.Suppress = False: Rep.gHWtInGms.Suppress = False
    'Rep.gHPureWt.Suppress = False: Rep.gHVal.Suppress = False
    
    Rep.hLabMCd.Suppress = True: Rep.hLabSCd.Suppress = True    ' ***** Manali  3.10.0 - 14/07/12 - Lab Report

    'GRP_REP.SetFormula Rep, "wVal", "iif({rdo.RmQw}='Q', {rdo.OrQty}, {rdo.orWt})*{rdo.OrSalRt}"
    GRP_REP.SetFormula Rep, "wVal", "{rdo.OrSalVal}"
    GRP_REP.SetFormula Rep, "wProfitPer", "0"
    Rep.TxtHead.SetText "Voucher Analysis - RM (Jwlry)"
    
    gb_ShowPic = False
    Rep.PicSec.Suppress = True
  ElseIf UCase(ADC.MenuCd) = UCase("RepJtVchAnaLab") Then
    Rep.HOdDmId.Suppress = True: Rep.hOdDmSz.Suppress = True: Rep.hOdDmCd.Suppress = True
    Rep.hOdSfx.Suppress = True: Rep.hOdSalPrc.Suppress = True   'Rep.hDmCtg.Suppress = True:
    Rep.hOdCstPrc.Suppress = True: Rep.hProfit.Suppress = True: Rep.hProfitPer.Suppress = True
    Rep.hOdOrdQty.Suppress = True: Rep.hOdSalVal.Suppress = True: Rep.hOdCstVal.Suppress = True
    Rep.gHOdOrdQty.Suppress = True: Rep.ghOdSalVal.Suppress = True: Rep.ghOdCstVal.Suppress = True
    Rep.ghProfit.Suppress = True: Rep.ghProfPer.Suppress = True

    Rep.hRmCtg.Suppress = True: Rep.hRmSCtg.Suppress = True: Rep.hRmCd.Suppress = True
    Rep.hOrLn1.Suppress = True: Rep.hOrLn2.Suppress = True: Rep.hOrRmPtr.Suppress = True
    Rep.hOrQty.Suppress = True: Rep.HOrWt.Suppress = True: 'Rep.hVal.Suppress = True
    Rep.hWtInGms.Suppress = True: Rep.hPureWt.Suppress = True: 'Rep.HOrSalRt.Suppress = True
    Rep.gHOrQty.Suppress = True: Rep.gHOrWt.Suppress = True: Rep.gHWtInGms.Suppress = True
    Rep.gHPureWt.Suppress = True: 'Rep.gHVal.Suppress = True

    'GRP_REP.SetFormula Rep, "wVal", "iif({rdo.OlQw}='Q', {rdo.OlQty}, {rdo.OdLabAsWt})*{rdo.OlSalRt}"
    GRP_REP.SetFormula Rep, "wVal", "{rdo.OlSalVal}"
    GRP_REP.SetFormula Rep, "wProfitPer", "0"
    Rep.TxtHead.SetText "Voucher Analysis - Labour (Jwlry)"
    
    gb_ShowPic = False
    Rep.PicSec.Suppress = True
  End If
  
  'Rep.TxtCoCdFr.SetText adc("UoCoCdFr")
  'Rep.TxtCoCdTo.SetText adc("UoCoCdTo")
  'Rep.TxtCoCdSel.SetText adc("UoCoCdSel")
  
  If ws_MultiCoMod Then
    Rep.TxtCoCdFr.SetText ADC("UoCoCdFr"): Rep.TxtCoCdTo.SetText ADC("UoCoCdTo"): Rep.TxtCoCdSel.SetText ADC("UoCoCdSel")
  Else
    Rep.TxtCoCdToLbl.Suppress = True: Rep.TxtCoCdSelLbl.Suppress = True
    Rep.TxtCoCdFr.SetText ADC("UoCoCdFr")
  End If
  
  'sv.30 passing divide by value to dsr
  GRP_REP.SetFormula Rep, "wDivByVal", ADC("UoValDivBy")
  
  'sv.30 when divied by amount>1 then follwing text will display in header
  If ADC("UoValDivBy") > 1 Then Rep.TxtDivByVal.SetText "(Values in " + Format(ADC("UoValDivBy"), "##,##,##0") + ")"
  
  Rep.txtTcs.SetText "Tcs: " + IIF(ADC("UoOmTc") <> "", "(" + ADC("UoOmTc") + ")", "") + IIF(ADC("UoOmTc1") <> "", IIF(ADC("UoYn1") = "Y", " + ", " - ") + "(" + ADC("UoOmTc1") + ")", "") _
                      + IIF(ADC("UoOmTc2") <> "", IIF(ADC("UoYn2") = "Y", " + ", " - ") + "(" + ADC("UoOmTc2") + ")", "")
  
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.OrRmCd}", 8)
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(ADC)

End Sub


' Zubin 213
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  'uni.4  checking Include SalCtg2,SalCtg3 option
  Case Is = 0
    If .Value = Checked Then ADC("wAddAllSalCtg") = "Y" Else ADC("wAddAllSalCtg") = "N"
  Case Is = 1     'Tag = "AddVal2"
    If .Value = Checked Then ADC("UoYN1") = "Y" Else ADC("UoYN1") = "N"
  Case Is = 2     'Tag = "AddVal3"
    If .Value = Checked Then ADC("UoYN2") = "Y" Else ADC("UoYN2") = "N"
  Case Is = 3     'Tag = "ShowPic"
    If .Value = Checked Then ADC("UoYN3") = "Y" Else ADC("UoYN3") = "N"
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
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
'pg
  ADC.Child = "GrdDsgAna"
  Set GrdDsgAna.PrevCtl = ADC
  Set GrdDsgAna.NextCtl = ADC
  GrdDsgAna.HotKey = "A"
  
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = False
  gs_OmTcTyp = ctJtTcTyp
  gs_CmCtg = "C"
  gs_PTyp1 = "CmCtg"
  gb_CmCtgFor = True
  gs_CmCtg1 = "C"
  
  If ws_MultiCoMod = True Then
    Set ADC.FirNKeyCtl = ADC("UoCoCdFr")
  Else
    Set ADC.FirNKeyCtl = ADC("UoCmCtgFr")
  End If

  If UCase(ADC.MenuCd) = UCase("RepJtVchAnaDsg") Then
    FraRmScp.Enabled = False: FraRmScp.Visible = False
    FraLabScp.Enabled = False: FraLabScp.Visible = False
    gs_PTyp = "": gs_Tbl = ""
  ElseIf UCase(ADC.MenuCd) = UCase("RepJtVchAnaRm") Then
    FraRmScp.Enabled = True: FraRmScp.Visible = True
    ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False
    LblProfPer.Enabled = False: LblProfPer.Visible = False
    FraOptProfPer.Enabled = False: FraOptProfPer.Visible = False
    LblProfitYN.Enabled = False: LblProfitYN.Visible = False    '6.1b profit YN option will be enable for vchr ana dsg only
    FraShowProfYN.Enabled = False:  FraShowProfYN.Visible = False '6.1b profit YN option will be enable for vchr ana dsg only
    FraLabScp.Enabled = False: FraLabScp.Visible = False
    FraSc.Enabled = False: FraSc.Visible = False:    CmdSc.Visible = False
    'uni.4 Include SalCtg2,3 scope option not visible for Voucher analysis Rm
    ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False
    gs_PTyp = "": gs_Tbl = ""
  ElseIf UCase(ADC.MenuCd) = UCase("RepJtVchAnaLab") Then
    FraRmScp.Enabled = True: FraRmScp.Visible = True
    ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False
    LblProfPer.Enabled = False: LblProfPer.Visible = False
    FraOptProfPer.Enabled = False: FraOptProfPer.Visible = False
    LblProfitYN.Enabled = False: LblProfitYN.Visible = False    '6.1b profit YN option will be enable for vchr ana dsg only
    FraShowProfYN.Enabled = False:  FraShowProfYN.Visible = False '6.1b profit YN option will be enable for vchr ana dsg only
    FraRmScp.Enabled = False: FraRmScp.Visible = False
    FraSc.Enabled = False: FraSc.Visible = False:    CmdSc.Visible = False
    'uni.4 Include SalCtg2,3 scope option not visible for Voucher analysis Lab
    ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False
    gs_PTyp = "LABMCD"
    gs_Tbl = "Param"
  End If

  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  If ws_MultiCoMod = True Then
      ADC("UOCOCDFR").Enabled = True: ADC("UOCOCDTO").Visible = True: ADC("UOCOCDSEL").Visible = True
  Else
      ADC("UOCOCDFR").Enabled = False: ADC("UOCOCDTO").Visible = False: ADC("UOCOCDSEL").Visible = False
  End If

  Call SetGroupSort
  
End Sub

Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gb_CoCdFor = False
  gs_CmCtg = ADC("UoCmCtgFr")
  gs_OmTcTyp = ctJtTcTyp
  gs_PTyp1 = "CmCtg"
  gb_CmCtgFor = True
  gs_CmCtg1 = ADC("UoCmCtgFr")
  gs_PTyp = IIF(UCase(ADC.MenuCd) = UCase("RepJtVchAnaLab"), "LABMCD", "")
  gs_Tbl = IIF(UCase(ADC.MenuCd) = UCase("RepJtVchAnaLab"), "Param", "")
  'sv.30 divided by amount cannot be 0
  Select Case UCase(IdName)
  Case Is = UCase("UoValDivBy")
    If pv_NewValue = 0 Then Cancel = True: ErrMsg = "Divided By Amount Cannot be Zero": Exit Sub
  End Select
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)

  gb_CoCdFor = False
  gs_OmTcTyp = ctJtTcTyp
  gs_CmCtg = ADC("UoCmCtgFr")
  gs_PTyp1 = "CmCtg"
  gb_CmCtgFor = True
  gs_CmCtg1 = ADC("UoCmCtgFr")
  gs_PTyp = IIF(UCase(ADC.MenuCd) = UCase("RepJtVchAnaLab"), "LABMCD", "")
  gs_Tbl = IIF(UCase(ADC.MenuCd) = UCase("RepJtVchAnaLab"), "Param", "")
  SetProp Me, IdName, When
End Sub


Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub OptProfPer_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN4") = "Y"
Case Is = 1
  ADC("UoYN4") = "N"
End Select
End Sub

Private Sub OptProfPer_GotFocus(Index As Integer)
  DispMsg FraOptProfPer.ToolTipText, etInfo
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

  If ws_MultiCoMod = False Then _
      ADC("UoCoCdFr") = gs_CoCd: ADC("UoCoCdTo") = gs_CoCd: ADC("UoCoCdSel") = "": ADC("UoCoCdFr").Enabled = False
  
  gs_CmCtg = ADC("UoCmCtgFr")
  
  If ADC("UoYN1") = "" Then ADC("UoYN1") = "N"
  If ADC("UoYN2") = "" Then ADC("UoYN2") = "N"
  If ADC("UoYN3") = "" Then ADC("UoYN3") = "N"
  
  If ADC("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If ADC("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If ADC("UoYN3") = "Y" Then ChkBoxArr(3).Value = Checked Else ChkBoxArr(3).Value = Unchecked

  If ADC("UoYN4") = "Y" Then OptProfPer(0).Value = True
  If ADC("UoYN4") = "N" Then OptProfPer(1).Value = True
  'uni.4 default value for Include SalCtg2,SalCtg3 is Yes
  If UCase(ADC.MenuCd) = UCase("RepJtVchAnaDsg") Then
    ChkBoxArr(0).Value = Checked
    ADC("wAddAllSalCtg") = "Y"
  Else
    ChkBoxArr(0).Value = Unchecked
    ADC("wAddAllSalCtg") = "N"
  End If
  ADC("UoValDivBy") = 1     'sv.30 setting default value
  ADC("UoYN5") = "N"    '6.1b default value for SHow Profit is NO
  If ADC("UoYN5") = "Y" Then OptProfYN(0).Value = True   '6.1b
  If ADC("UoYN5") = "N" Then OptProfYN(1).Value = True   '6.1b
End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = True, ctSelfCoCd, ADC("UOCOCDFR")) + "' and " + _
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
  
  
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
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
    Cancel = atxt(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set Rep = Nothing
    '*** (09/08/05)
End Sub

'pg.8- Set visible /invisible state for analysis frame
Private Sub CmdSc_Click()
  If FraSc.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraSc.Enabled = True
    FraSc.Visible = True
    TAB_REP.Enabled = False
    ADC.AllowSave = False
    FraSc.ZOrder
  Else
    TAB_REP.Enabled = True
    ADC.AllowSave = True
    FraSc.Enabled = False
    FraSc.Visible = False
  End If
End Sub

'pg.8-checking Valid Analysis Sr & field
Private Sub GrdDsgAna_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  With GrdDsgAna
    Select Case UCase(ColName)
    Case Is = UCase("UdAnaSr")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                                "PTyp= 'DAANAFLD' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Analysis Sr.": Exit Sub
    Case Is = UCase("UdAnaCdFr"), UCase("UdAnaCdTo")
      Cancel = NewValue <> "" And Not moCn.RecSeek("Select PSCd From Param Where PTyp= 'DAANACD' And PMCd= '" + _
                                GrdDsgAna.Value(RowNum, "UdAnaSr") + "'And PSCd = '" + NewValue + "'")
      ErrMsg = "Invalid Analysis Field ": Exit Sub
    End Select
  End With
End Sub
'pg- calling hlp for GrdDsgAna
Private Sub GrdDsgAna_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Select Case ColName
  Case Is = UCase("UdAnaSr")
    Call HlpList.PMCd("DAANAFLD")
  Case Is = UCase("UdAnaCdFr"), UCase("UdAnaCdTo")
    If moCn.GetFldVal("Select PValue from Param where PTyp= 'DAANAFLD' and " + _
       "PMCd= '" + GrdDsgAna.Value(RowNum, "UdAnaSr") + "' ") = "N" Then _
       Cancel = True: ErrMsg = "Cannot Enter Code For This Analysis Field": Exit Sub
  
    Call HlpList.PSCd("DAANACD", GrdDsgAna.Value(RowNum, "UdAnaSr"))
  End Select
End Sub

Private Sub GrdDsgAna_InitKey(ByVal RowNum As Integer)
  With GrdDsgAna
    If .IsNew(RowNum) Then
      .Store "UdUsrCd", ADC.UsrCd
      .Store "UdMnuCd", ADC.MenuCd
    End If
  End With
End Sub

Private Sub GrdDsgAna_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdDsgAna
    Select Case UCase(ColName)
    Case Is = UCase("UdAnaCdFr")
      If .Mode = fgmnorm Then .Value(RowNum, "UdAnaCdFr") = ""
    Case Is = UCase("UdAnaCdTo")
      If .Mode = fgmnorm Then .Value(RowNum, "UdAnaCdTo") = ""
    End Select
  End With
End Sub

Private Sub GrdDsgAna_SetRecSource()
  GrdDsgAna.RecSource = "Select * from UsrOptDsgAna where UdUsrCd= '" + ADC.UsrCd + "' and " + _
                        "UdMnuCd= '" + ADC.MenuCd + "'"
                        
End Sub

Private Sub GrdDsgAna_Validate(Cancel As Boolean)
  Cancel = GrdDsgAna.Validate
End Sub
'6.1b
Private Sub OptProfYN_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN5") = "Y"
  FraOptProfPer.Enabled = True
Case Is = 1
  ADC("UoYN5") = "N"
  FraOptProfPer.Enabled = False
End Select
End Sub
'6.1b
Private Sub OptProfYN_GotFocus(Index As Integer)
  DispMsg FraShowProfYN.ToolTipText, etInfo
End Sub

