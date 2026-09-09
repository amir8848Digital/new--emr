VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpSalAnaLab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Analysis (Lab Level)"
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
      Left            =   2010
      TabIndex        =   142
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9720
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
      Left            =   12240
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
      TabIndex        =   87
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
      Left            =   6345
      TabIndex        =   97
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   98
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
         TabIndex        =   99
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
      Left            =   7560
      TabIndex        =   96
      Top             =   9645
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
      TabIndex        =   89
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
      TabIndex        =   92
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   91
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpSalAnaLab.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpSalAnaLab.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   95
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   90
               Top             =   360
               Width           =   9435
               _ExtentX        =   16642
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   94
            Top             =   360
            Width           =   15105
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
               TabIndex        =   40
               Tag             =   "DmHold"
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   4150
               Width           =   2250
            End
            Begin VB.Frame FraOptDt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   8115
               TabIndex        =   136
               ToolTipText     =   "Select if to show Invoice Date Or Invoice Printing Date"
               Top             =   7605
               Width           =   4995
               Begin VB.OptionButton OptInvDt 
                  Caption         =   "Print Date"
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
                  Left            =   3000
                  TabIndex        =   79
                  ToolTipText     =   "Select if to show Invoice Date Or Invoice Printing Date"
                  Top             =   0
                  Width           =   2085
               End
               Begin VB.OptionButton OptInvDt 
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
                  Height          =   285
                  Index           =   0
                  Left            =   0
                  TabIndex        =   78
                  ToolTipText     =   "Select if to show Invoice Date Or Invoice Printing Date"
                  Top             =   0
                  Width           =   1875
               End
            End
            Begin VB.Frame FraOptCust 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   8115
               TabIndex        =   135
               ToolTipText     =   "Select eihter Invoice Customer Or Order Customer"
               Top             =   7320
               Width           =   4665
               Begin VB.OptionButton OptCust 
                  Caption         =   "Order Cust"
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
                  Left            =   3000
                  TabIndex        =   76
                  ToolTipText     =   "Select eihter Invoice Customer Or Order Customer"
                  Top             =   0
                  Width           =   1425
               End
               Begin VB.OptionButton OptCust 
                  Caption         =   "Invoice Cust"
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
                  TabIndex        =   75
                  ToolTipText     =   "Select eihter Invoice Customer Or Order Customer"
                  Top             =   0
                  Width           =   1545
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5940
               TabIndex        =   7
               ToolTipText     =   "Enter To Invoice Year"
               Top             =   705
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYTO"
               IdName          =   "UOINYYTO"
               CmpStr          =   "iLYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2880
               TabIndex        =   4
               ToolTipText     =   "Enter From Invoice Year"
               Top             =   705
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "iLYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   2115
               TabIndex        =   3
               ToolTipText     =   "Enter From Invoice Tc"
               Top             =   705
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "iLTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3300
               TabIndex        =   5
               ToolTipText     =   "Enter From Invoice Character"
               Top             =   705
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "iLChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   5160
               TabIndex        =   6
               ToolTipText     =   "Enter To Invoice Tc"
               Top             =   705
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCTO"
               IdName          =   "UOINTCTO"
               CmpStr          =   "iLTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   6360
               TabIndex        =   8
               ToolTipText     =   "Enter To Invoice Character"
               Top             =   705
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRTO"
               IdName          =   "UOINCHRTO"
               CmpStr          =   "iLChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   2115
               TabIndex        =   9
               ToolTipText     =   "Enter From Invoice Number"
               Top             =   990
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "IlNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   5160
               TabIndex        =   11
               ToolTipText     =   "Enter To Invoice Number"
               Top             =   990
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOTO"
               IdName          =   "UOINNOTO"
               CmpStr          =   "IlNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   2115
               TabIndex        =   73
               ToolTipText     =   "Specify Whether Value/ Weight To Be Shown Is Actual, Invoice Or Customs "
               Top             =   7485
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOINVAL"
               ReCalcParent    =   "UOCURCDFR"
               IdName          =   "UOINVAL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   5160
               TabIndex        =   35
               ToolTipText     =   "Enter To Design Category"
               Top             =   3585
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
               Left            =   8140
               TabIndex        =   36
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   3585
               Width           =   6870
               _ExtentX        =   12118
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
               Left            =   2115
               TabIndex        =   34
               ToolTipText     =   "Enter From Design Category"
               Top             =   3585
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
               Left            =   5160
               TabIndex        =   42
               ToolTipText     =   "Enter To Design Code"
               Top             =   4635
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "IdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   8145
               TabIndex        =   43
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   4635
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "IdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   2115
               TabIndex        =   41
               ToolTipText     =   "Enter From Design Code"
               Top             =   4635
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "IdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   5160
               TabIndex        =   16
               ToolTipText     =   "Enter To Customer"
               Top             =   1560
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
               Index           =   41
               Left            =   8140
               TabIndex        =   17
               ToolTipText     =   "Enter Customer Selection"
               Top             =   1560
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   2115
               TabIndex        =   15
               ToolTipText     =   "Enter From Customer"
               Top             =   1560
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
               Index           =   2
               Left            =   2115
               TabIndex        =   48
               ToolTipText     =   "Enter From Karat"
               Top             =   5490
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               CmpStr          =   "IdKt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   5160
               TabIndex        =   38
               ToolTipText     =   "Enter To Design Sales Category"
               Top             =   3870
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
               Left            =   8140
               TabIndex        =   39
               ToolTipText     =   "Enter Design Sales Category Selection"
               Top             =   3870
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMSALCTGSEL"
               IdName          =   "UODMSALCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2115
               TabIndex        =   37
               ToolTipText     =   "Enter From Design Sales Category"
               Top             =   3870
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
               Index           =   7
               Left            =   5160
               TabIndex        =   25
               ToolTipText     =   "Enter To Export Number"
               Top             =   2445
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOTO"
               IdName          =   "UOINEXPNOTO"
               CmpStr          =   "InExpNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   8140
               TabIndex        =   26
               ToolTipText     =   "Enter Export Number Selection"
               Top             =   2445
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOINEXPNOSEL"
               IdName          =   "UOINEXPNOSEL"
               CmpStr          =   "InExpNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   2115
               TabIndex        =   24
               ToolTipText     =   "Enter From Export Number"
               Top             =   2445
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOFR"
               IdName          =   "UOINEXPNOFR"
               CmpStr          =   "InExpNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   5160
               TabIndex        =   45
               ToolTipText     =   "Enter To Design Size"
               Top             =   4920
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
               CmpStr          =   "IdDmSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   2115
               TabIndex        =   44
               ToolTipText     =   "Enter From Design Size"
               Top             =   4920
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "IdDmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   5160
               TabIndex        =   47
               ToolTipText     =   "Enter To Suffix"
               Top             =   5205
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOODSFXTO"
               IdName          =   "UOODSFXTO"
               CmpStr          =   "IdSfx <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   2115
               TabIndex        =   46
               ToolTipText     =   "Enter From Suffix"
               Top             =   5205
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOODSFXFR"
               IdName          =   "UOODSFXFR"
               CmpStr          =   "IdSfx >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   5925
               TabIndex        =   55
               ToolTipText     =   "Enter To Order Year"
               Top             =   5775
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "IdExpOdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   2895
               TabIndex        =   52
               ToolTipText     =   "Enter From Order Year"
               Top             =   5775
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "IdExpOdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   2115
               TabIndex        =   51
               ToolTipText     =   "Enter From Order Tc"
               Top             =   5775
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "IdExpOdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   3315
               TabIndex        =   53
               ToolTipText     =   "Enter From Order Character"
               Top             =   5775
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "IdExpOdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   5160
               TabIndex        =   54
               ToolTipText     =   "Enter To Order Tc"
               Top             =   5790
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "IdExpOdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   6345
               TabIndex        =   56
               ToolTipText     =   "Enter To Order Character"
               Top             =   5775
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "IdExpOdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   6045
               TabIndex        =   60
               ToolTipText     =   "Enter To Order Serial"
               Top             =   6060
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "IdExpOdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   3015
               TabIndex        =   58
               ToolTipText     =   "Enter From Order Serial"
               Top             =   6060
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "IdExpOdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   2115
               TabIndex        =   57
               ToolTipText     =   "Enter From Order Number"
               Top             =   6060
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "IdExpOdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   5160
               TabIndex        =   59
               ToolTipText     =   "Enter To Order Number"
               Top             =   6075
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "IdExpOdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   13515
               TabIndex        =   74
               ToolTipText     =   "Show Invoice Or Order Customer ('Y' for Invoice / 'N' For Order)"
               Top             =   7320
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
               Index           =   4
               Left            =   2115
               TabIndex        =   72
               ToolTipText     =   "Enter Level Of Detail To Be Shown"
               Top             =   7200
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOSALABGRP"
               IdName          =   "UOSALABGRP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   5160
               TabIndex        =   33
               ToolTipText     =   "Enter To Labour Rate"
               Top             =   3300
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOSALRTTO"
               IdName          =   "UOSALRTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   2115
               TabIndex        =   32
               ToolTipText     =   "Enter From Labour Rate"
               Top             =   3300
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOSALRTFR"
               IdName          =   "UOSALRTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   5160
               TabIndex        =   49
               ToolTipText     =   "Enter To Karat"
               Top             =   5490
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTTO"
               IdName          =   "UOKTTO"
               CmpStr          =   "IdKt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   8145
               TabIndex        =   50
               ToolTipText     =   "Enter Karat Selection"
               Top             =   5490
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "IdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   5160
               TabIndex        =   28
               ToolTipText     =   "Enter To Labour Main Code"
               Top             =   2730
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
               CmpStr          =   "iLMCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2115
               TabIndex        =   27
               ToolTipText     =   "Enter From Labour Main Code"
               Top             =   2730
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
               CmpStr          =   "iLMCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   8140
               TabIndex        =   29
               ToolTipText     =   "Enter Labour Main Code Selection"
               Top             =   2730
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPMCDSEL"
               IdName          =   "UOPMCDSEL"
               CmpStr          =   "iLMCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   5160
               TabIndex        =   31
               ToolTipText     =   "Enter To Labour Sub Code"
               Top             =   3015
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOLSCDTO"
               IdName          =   "UOLSCDTO"
               CmpStr          =   "iLSCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   2115
               TabIndex        =   30
               ToolTipText     =   "Enter From Labour Sub Code"
               Top             =   3015
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOLSCDFR"
               IdName          =   "UOLSCDFR"
               CmpStr          =   "iLSCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   5160
               TabIndex        =   14
               ToolTipText     =   "Enter To Invoice Voucher Date"
               Top             =   1275
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2115
               TabIndex        =   13
               ToolTipText     =   "Enter From Invoice Voucher Date"
               Top             =   1275
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   2115
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   420
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "IdCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   8145
               TabIndex        =   71
               ToolTipText     =   "Enter Sales Executive Selection"
               Top             =   6930
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSALPERSEL"
               IdName          =   "UOSALPERSEL"
               CmpStr          =   "OmCmSalPer In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   5160
               TabIndex        =   70
               ToolTipText     =   "Enter To Sales Executive"
               Top             =   6930
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
               Index           =   55
               Left            =   2115
               TabIndex        =   69
               ToolTipText     =   "Enter From Sales Executive"
               Top             =   6915
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
               Index           =   62
               Left            =   2115
               TabIndex        =   80
               ToolTipText     =   "Enter Currency Code Whose Invoice You Wish To View Or Leave It Blank If You Wish To View All Invoices In Base Currency"
               Top             =   7785
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
               Index           =   73
               Left            =   5160
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   420
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "IdCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   8140
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   450
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "IdCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   80
               Left            =   13515
               TabIndex        =   77
               ToolTipText     =   "Show invoice Date/ Printng Dat (Yes / No)"
               Top             =   7605
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
               Index           =   85
               Left            =   8140
               TabIndex        =   20
               ToolTipText     =   "Enter Customer Group Selection"
               Top             =   1845
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMGRPSEL"
               IdName          =   "UOCMGRPSEL"
               CmpStr          =   "CE.CmGrp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   86
               Left            =   5160
               TabIndex        =   19
               ToolTipText     =   "Enter To Customer Group"
               Top             =   1845
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMGRPTO"
               IdName          =   "UOCMGRPTO"
               CmpStr          =   "CE.CmGrp<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   87
               Left            =   2115
               TabIndex        =   18
               ToolTipText     =   "Enter From Customer Group"
               Top             =   1850
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMGRPFR"
               IdName          =   "UOCMGRPFR"
               CmpStr          =   "CE.CmGrp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   88
               Left            =   8140
               TabIndex        =   23
               ToolTipText     =   "Enter Customer Type Selection"
               Top             =   2145
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMTYPSEL"
               IdName          =   "UOCMTYPSEL"
               CmpStr          =   "CE.CmTyp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   89
               Left            =   5160
               TabIndex        =   22
               ToolTipText     =   "Enter To Customer Type"
               Top             =   2145
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMTYPTO"
               IdName          =   "UOCMTYPTO"
               CmpStr          =   "CE.CmTyp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   90
               Left            =   2115
               TabIndex        =   21
               ToolTipText     =   "Enter From Customer Type"
               Top             =   2150
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMTYPFR"
               IdName          =   "UOCMTYPFR"
               CmpStr          =   "CE.CmTyp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   91
               Left            =   3045
               TabIndex        =   10
               ToolTipText     =   "Enter From Invoice Serial"
               Top             =   990
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOIDSRFR"
               IdName          =   "UOIDSRFR"
               CmpStr          =   "IdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   92
               Left            =   6090
               TabIndex        =   12
               ToolTipText     =   "Enter From Invoice Serial"
               Top             =   990
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOIDSRTO"
               IdName          =   "UOIDSRTO"
               CmpStr          =   "IdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   5160
               TabIndex        =   86
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   8640
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   5160
               TabIndex        =   84
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   8355
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGTO"
               IdName          =   "UORMSCTGTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   2115
               TabIndex        =   81
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   8070
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   59
               Left            =   2115
               TabIndex        =   83
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   8355
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
               Data            =   "ABCDW"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   2115
               TabIndex        =   85
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   8640
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   10380
               TabIndex        =   82
               ToolTipText     =   "Enter To Raw Material Category"
               Top             =   7590
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
               CmpStr          =   "RmCtg <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   93
               Left            =   5160
               TabIndex        =   62
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   6360
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
               Index           =   94
               Left            =   8145
               TabIndex        =   63
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   6360
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   95
               Left            =   2115
               TabIndex        =   61
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   6345
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
               Index           =   96
               Left            =   3405
               TabIndex        =   65
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   6630
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
               Index           =   97
               Left            =   6435
               TabIndex        =   67
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   6645
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
               Index           =   98
               Left            =   8145
               TabIndex        =   68
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   6645
               Visible         =   0   'False
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "OdPoNo in"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   115
               Left            =   2115
               TabIndex        =   64
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   6645
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
               Index           =   116
               Left            =   5160
               TabIndex        =   66
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   6645
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
               Index           =   99
               Left            =   2775
               TabIndex        =   152
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   4260
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               IdName          =   "WADDALLSALCTG"
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
               Left            =   3345
               TabIndex        =   151
               Top             =   6645
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
               Index           =   39
               Left            =   6375
               TabIndex        =   150
               Top             =   6645
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
               Left            =   90
               TabIndex        =   149
               Top             =   6645
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
               Index           =   28
               Left            =   90
               TabIndex        =   148
               Top             =   6360
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
               Index           =   37
               Left            =   120
               TabIndex        =   147
               ToolTipText     =   "Location"
               Top             =   8070
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
               Index           =   36
               Left            =   120
               TabIndex        =   146
               Top             =   8355
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
               Index           =   29
               Left            =   120
               TabIndex        =   145
               Top             =   8640
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
               Left            =   5985
               TabIndex        =   141
               Top             =   990
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
               Left            =   2940
               TabIndex        =   140
               Top             =   990
               Width           =   105
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
               Index           =   32
               Left            =   90
               TabIndex        =   139
               Top             =   2150
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
               Index           =   33
               Left            =   90
               TabIndex        =   138
               Top             =   1850
               Width           =   1425
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Consider Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5200
               TabIndex        =   137
               Top             =   7605
               Width           =   2055
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
               TabIndex        =   134
               Top             =   7785
               Width           =   1005
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
               Index           =   23
               Left            =   90
               TabIndex        =   133
               Top             =   6915
               Width           =   1995
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
               Index           =   22
               Left            =   90
               TabIndex        =   132
               Top             =   420
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
               Index           =   21
               Left            =   8100
               TabIndex        =   131
               Top             =   60
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Date/Print Date"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   90
               TabIndex        =   130
               Top             =   1275
               Width           =   1995
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
               Index           =   19
               Left            =   90
               TabIndex        =   129
               Top             =   3015
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
               Index           =   18
               Left            =   90
               TabIndex        =   128
               Top             =   2730
               Width           =   1425
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
               Index           =   15
               Left            =   90
               TabIndex        =   127
               Top             =   3300
               Width           =   1425
            End
            Begin VB.Label LblRmGrp 
               BackStyle       =   0  'Transparent
               Caption         =   "Rep Det Level"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   126
               Top             =   7200
               Width           =   1845
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Inv/ Ord Cust"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5200
               TabIndex        =   125
               Top             =   7320
               Width           =   2055
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order Tc/Yy/Chr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   124
               Top             =   5775
               Width           =   1785
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
               Index           =   11
               Left            =   90
               TabIndex        =   123
               Top             =   6060
               Width           =   1485
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
               Index           =   8
               Left            =   6285
               TabIndex        =   122
               Top             =   5775
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
               Index           =   4
               Left            =   3255
               TabIndex        =   121
               Top             =   5775
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
               Index           =   3
               Left            =   5865
               TabIndex        =   120
               Top             =   5775
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
               Index           =   0
               Left            =   2835
               TabIndex        =   119
               Top             =   5775
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
               Left            =   5985
               TabIndex        =   118
               Top             =   6060
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
               Index           =   31
               Left            =   2955
               TabIndex        =   117
               Top             =   6060
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
               TabIndex        =   116
               Top             =   5205
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
               TabIndex        =   115
               Top             =   4920
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Exp No"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   114
               Top             =   2445
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
               Left            =   90
               TabIndex        =   113
               Top             =   5490
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
               TabIndex        =   112
               Top             =   1560
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
               TabIndex        =   111
               Top             =   4635
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
               TabIndex        =   110
               Top             =   3585
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
               TabIndex        =   109
               Top             =   3870
               Width           =   1365
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Value As"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   108
               Top             =   7485
               Width           =   1845
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
               Left            =   2115
               TabIndex        =   107
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
               Left            =   5160
               TabIndex        =   106
               Top             =   60
               Width           =   1275
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
               Left            =   2820
               TabIndex        =   105
               Top             =   705
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
               Left            =   5880
               TabIndex        =   104
               Top             =   705
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
               Left            =   3240
               TabIndex        =   103
               Top             =   705
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
               Index           =   24
               Left            =   6300
               TabIndex        =   102
               Top             =   705
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv No/Sr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   101
               Top             =   990
               Width           =   1605
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
               Index           =   9
               Left            =   90
               TabIndex        =   100
               Top             =   705
               Width           =   1725
            End
         End
      End
      Begin VB.Frame FraSc 
         Height          =   2925
         Left            =   1560
         TabIndex        =   143
         Top             =   6480
         Visible         =   0   'False
         Width           =   3945
         Begin MwfCtl.MWCTL_FLX GrdDsgAna 
            Height          =   2430
            Left            =   120
            TabIndex        =   144
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
Attribute VB_Name = "EmrFrpSalAnaLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepSalAnaLab
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "DmCtg", "hDmCtg", "Sales Ctg", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= max(DmCtg)) "
  'Geeta***Crm***212
  'uni.4 max(DmSalCtg) replaced by qDmSalCtgRepl,qDmSalCtgDescRepl to get either SalCtg,SalCTg2 or SalCTg3
  GRP_REP.Add "Sales Ctg", "||qDmSalCtgRepl||", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and PmCd=max(DmCtg) and PSCd= ||qDmSalCtgDescRepl|| ) "
  
  GRP_REP.Add "Inv Vch No", "IdCoCd+IdTc+IdYy+IdChr+str(IdNo)", "IdCoCd+'/'+IdTc+'/'+IdYy+'/'+IdChr+'/'+LTrim(Str(IdNo))", "", "", "Company Code,Invoice Cust,Inv Date,Inv Exp No,Customer Grp,Customer Typ", "", "", "max(InExpNo) "
  GRP_REP.Add "Inv Exp No", "IdCoCd+max(InExpNo)", "'['+IdCoCd+'] '+max(InExpNo)", "InExpNo", "hInExpNo", "Company Code,Inv Date,Inv Vch No,Customer Grp,Customer Typ", "", "", ""
  
  ' ***** Manali  3.10.0 - 10/03/12 - Sort Based on InDt/PrnDt
  '6.1
  GRP_REP.Add "Inv/Prn Date", "convert(varchar(10), max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End), 112)", "convert(varchar(8), max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End), " + CStr(DtFmtSlahVal()) + ")", "", "", "Quarter,Month,Week", "", "", ""
  GRP_REP.Add "Week", "LTrim(Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))))+' - '+DateName(WW,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))", "", "", "", "Quarter,Month", "", "", ""
  GRP_REP.Add "Month", "Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End)))+Str(DatePart(mm,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End)))", "LTrim(Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))))+' '+DateName(MM,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))", "", "", "Quarter", "", "", ""
  'MW.134 - group sort Year added.
  GRP_REP.Add "Year", "LTrim(Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))))", "", "", "", "", "", "", ""
  GRP_REP.Add "Quarter", "LTrim(Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))))+' - '+DateName(QQ,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))", "", "", "", "", "", "", ""
  ' ***** Manali  3.10.0 - 10/03/12 - Sort Based on InDt/PrnDt
''  Bef 3.10.0
''  GRP_REP.Add "Inv Date", "convert(varchar(10), max(InDt), 112)", "convert(varchar(8), max(InDt), 3)", "", "", "Quarter,Month,Week", "", "", ""
''  GRP_REP.Add "Prn Date", "convert(varchar(10), max(InPrnDt), 112)", "convert(varchar(8), max(InPrnDt), 3)", "", "", "Quarter,Month,Week", "", "", ""
''  GRP_REP.Add "Week", "LTrim(Str(DatePart(yyyy,max(InDt))))+' - '+DateName(WW,max(InDt))", "", "", "", "Quarter,Month", "", "", ""
''  GRP_REP.Add "Month", "Str(DatePart(yyyy,max(InDt)))+Str(DatePart(mm,max(InDt)))", "LTrim(Str(DatePart(yyyy,max(InDt))))+' '+DateName(MM,max(InDt))", "", "", "Quarter", "", "", ""
''  GRP_REP.Add "Quarter", "LTrim(Str(DatePart(yyyy,max(InDt))))+' - '+DateName(QQ,max(InDt))", "", "", "", "", "", "", ""
  
  GRP_REP.Add "Design Code", "max(IdDmCd)", "", "IdDmCd", "hIdDmCd", "Design Ctg,Sales Ctg", "", "", "max(DmDesc) "
  GRP_REP.Add "Design Size", "max(IdDmSz)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= max(IdDmSz))"
  GRP_REP.Add "Suffix", "max(IdSfx)", "", "", "", "", "", "", ""
  '******Geeta*****************Emr206****06/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "ExpOrd Cust", "max(IdExpCmCd)", "", "", "", "Sales Exec", "", "", "case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then max(CE.CmName) else max(IdExpCmCd) end " ' From CustMst Where CmCtg= 'C' and CmCd= max(IdExpCmCd))
  GRP_REP.Add "Invoice Cust", "max(InCmCd)", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then CmName else  max(InCmCd) end) From CustMst Where CmCtg= 'C' and CmCd= max(InCmCd)) "
  '**********
  
    'Sachin - 4.1.0.0   `
  GRP_REP.Add "Customer Grp", "max(CE.CmGrp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMGRP' and PmCd = Max(CE.CmGrp))"
  GRP_REP.Add "Customer Typ", "max(CE.CmTyp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMTYP' and PmCd = Max(CE.CmTyp))"
  
  GRP_REP.Add "Order", "IdCoCd+max(IdExpOdTc)+max(IdExpOdYy)+max(IdExpOdChr)+str(max(IdExpOdNo))", "IdCoCd+'/'+max(IdExpOdTc)+'/'+max(IdExpOdYy)+'/'+max(IdExpOdChr)+'/'+LTrim(Str(max(IdExpOdNo)))", "", "", "Company Code,ExpOrd Cust,Sales Exec", "", "", "max(IdExpCmCd) "
  
  GRP_REP.Add "Cust PO", "max(OmPoNo)", "max(OmPoNo)", "", "", "", "", "", "max(OmPoNo) "
  'vk.19 sub customer added
  GRP_REP.Add "Cust SUB PO", "max(OdSubCust)+max(OdPoNo)", "max(OdSubCust)+(case when max(OdPoNo)<> '' then ' : ' else '' end)+max(OdPoNo)", "", "", "", "", "", "max(OdSubCust)+(case when max(OdPoNo)<> '' then ' : ' else '' end)+max(OdPoNo) "
  
'  GRP_REP.Add "Cust PO", "(Select max(OmPoNo) From OrdMst Where OmCoCd=IdCoCd and OmTc=Max(IdExpOdTc) and OmYy=Max(IdExpOdYy) and OmChr=Max(IdExpOdChr)  and OmNo=Max(IdExpOdNo)) ", "(Select max(OmPoNo) From OrdMst Where OmCoCd=IdCoCd and OmTc=Max(IdExpOdTc) and OmYy=Max(IdExpOdYy) and OmChr=Max(IdExpOdChr)  and OmNo=Max(IdExpOdNo)) ", "", "", "", "", "", "(Select max(OmPoNo) From OrdMst Where OmCoCd=IdCoCd and OmTc=Max(IdExpOdTc) and OmYy=Max(IdExpOdYy) and OmChr=Max(IdExpOdChr)  and OmNo=Max(IdExpOdNo)) "
'  GRP_REP.Add "Cust Sub PO", "(Select max(OdPoNo) From OrdDsg Where OdCoCd=IdCoCd and OdTc=Max(IdExpOdTc) and OdYy=Max(IdExpOdYy) and OdChr=Max(IdExpOdChr)  and OdNo=Max(IdExpOdNo) and OdSr=Max(IdExpOdSr))", "(Select max(OdPoNo) From OrdDsg Where OdCoCd=IdCoCd and OdTc=Max(IdExpOdTc) and OdYy=Max(IdExpOdYy) and OdChr=Max(IdExpOdChr)  and OdNo=Max(IdExpOdNo) and OdSr=Max(IdExpOdSr))", "", "", "", "", "", "(Select max(OdPoNo) From OrdDsg Where OdCoCd=IdCoCd and OdTc=Max(IdExpOdTc) and OdYy=Max(IdExpOdYy) and OdChr=Max(IdExpOdChr)  and OdNo=Max(IdExpOdNo) and OdSr=Max(IdExpOdSr))"

  GRP_REP.Add "Karat", "max(IdKt)", "", "IdKt", "hIdKt", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = max(IdKt)) "
  GRP_REP.Add "PCS/PRS", "max(DmUom)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMUOM' and PmCd= max(DmUom)) "
  GRP_REP.Add "VaCtg", "max(IdVaCtg)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'VACTG' and PMCd = max(IdVaCtg)) "
  GRP_REP.Add "(DsgCd)Colour", "max(IdDmCd+OdDmCol)", "'('+max(IdDmCd)+')'+max(OdDmCol)", "IdDmCd", "hIdDmCd", "Design Ctg,Sales Ctg,Design Code,Design Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(OdDmCol)) "
  GRP_REP.Add "Design Colour", "max(OdDmCol)", "", "", "", "(DsgCd)Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(OdDmCol)) "
  
  GRP_REP.Add "Lab Main Cd", Array("max(iLMCd)", "'SET'"), "", "iLMCd", "hiLMCd", "", "", "", _
            Array("(Select PDesc From Param Where PTyp= 'LABMCD' and PMCd= max(iLMCd))", "(Select PDesc From Param Where PTyp= 'LABMCD' and PMCd= 'SET')")
  GRP_REP.Add "Lab Sub Cd", Array("max(iLMCd+iLSCd)", "'SET'+Ltrim(RTrim(max(IrSetSCd)))"), Array("'('+max(iLMCd)+') '+max(iLSCd)", "'(SET)'+max(IrSetSCd)"), "iLMCd,iLSCd", "hiLMCd,hiLSCd", "Lab Main Cd", "", "", Array("(Select PDesc From Param Where PTyp= 'LABSCD' and PMCd= max(iLMCd) and PSCd = max(iLSCd)) ", "(Select PDesc From Param Where PTyp= 'LABSCD' and PMCd= 'SET' and PSCd = max(IrSetSCd)) ")
  
  GRP_REP.Add "Rm Ctg", Array("''", "max(IrRmCtg)"), "", "IrRmCtg", "hIrRmCtg", "", "", "", _
    Array("''", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = max(IrRmCtg))")
  
  GRP_REP.Add "Rm SubCtg", Array("''", "max(IrRmCtg+IrRmSCtg)"), Array("''", "'('+ RTrim(max(IrRmCtg))+ ') '+ max(IrRmSCtg)", "''"), "IrRmCtg,IrRmSCtg", _
                    "hIrRmCtg,hIrRmSCtg", "", "", "", _
    Array("''", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = max(IrRmSCtg))")
    
  GRP_REP.Add "Rm Code", Array("''", "max(IrRmCtg+IrRmSCtg+IrRmCd)"), Array("''", "'('+ RTrim(max(IrRmCtg))+ '/'+ max(IrRmSCtg) +')'+max(IrRmCd)", "''"), "IrRmCtg,IrRmSCtg,IrRmCd", _
                    "hIrRmCtg,hIrRmSCtg,hIrRmCd", "", "", "", _
    Array("''", "(Select (SubString(RmDesc, 1, 30)) From RmMst where RmCd=max(IrRmCd))")
    
  
  GRP_REP.Add "Sales Exec", "max(OmCmSalPer)", "", "", "", "", "", "", "(Select PDesc from Param Where PTyp = 'SP' and PMCd= max(OmCmSalPer)) "
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ,Process,OrderNo,OrderSr,Customer PO,FgSubLoc ******
  If ws_MultiCoMod = True Then _
    GRP_REP.Add "Company Code", "IdCoCd", "IdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=IdCoCd) "
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******
  
End Sub

Private Sub adc_setreprecsource()
  '*** Report Sql ***
'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'IlCoCd, IlTc, IlYy, IlChr, IlNo, IlSr, space(4) as IlSrNo, InExpNo, InDt, InCmCd, DmCtg,
'IdDmCd, IdDmSz, IdSfx, IdQty, IdExpOdTc, IdExpOdYy, IdExpOdChr,
'IdExpOdNo, IdExpOdSr, IdVaCtg, iLMCd, iLSCd,
'iLiRt As qLabRt, iLiVal As qLabVal
'From InvLab, InvDsg, InvHd, DsgMst where 1= 2
  '*** Report Sql ***
  
  Dim grpflds0 As String, grpflds1 As String, wSqlStrg As String, wrepcnd As String, wCnd As String
  Dim wLabDet As String, wLabGrp As String, wCust As String, wOrdJoin As String
  Dim wOpt As String, wiLLabRtOpt As String, wiLLabValOpt As String, wSqlSetStrg As String
  Dim ws_DsgAnaCnd As String
  ' Zubin 213
  Dim wb_InBaseCur As Boolean
  wb_InBaseCur = IIF(ADC("UoCurCdFr") = "", True, False)
  
  Set Rep = Nothing
  Set ADC.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)
   
  '*** To check if 'Design Colour' group is selected ***
  Dim wGrp(6) As String, i As Integer, wDsgColGrpPos As Integer, wLabGrpOpt As String
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDsgColGrpPos = 0
  For i = 1 To 3
    If wGrp(i) = "(DsgCd)Colour" Or wGrp(i) = "Design Colour" Then wDsgColGrpPos = i
    If wGrp(i) = "Lab Main Cd" Or wGrp(i) = "Lab Sub Cd" Then wLabGrpOpt = wGrp(i)
  Next i
  
  '*** To 'Design Colour' group not selected then ***
  If wDsgColGrpPos = 0 Then
    wOrdJoin = ""
  Else
    wOrdJoin = " Join OrdDsg on " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey=IdPrtKey and ", "") + " OdIdNo=IdOdIdNo and OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
               "and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr "
  End If
    wOrdJoin = ""
  If ADC("UoYN") = "Y" Then
    wCust = "InCmCd"
    Rep.hInCmCd.SetText "Inv Cust"
  Else
    wCust = "IdExpCmCd"
    Rep.hInCmCd.SetText "Ord Cust"
  End If
  
  ADC("UoCmCdFr").CmpStr = wCust + ">= "
  ADC("UoCmCdTo").CmpStr = wCust + "<= "
  ADC("UoCmCdSel").CmpStr = wCust + " In "
  
'  adc("UoSalPerFr").CmpStr = IIF(adc("UoYN") = "Y", "c.", "ce.") + "CmSalPer >= "
'  adc("UoSalPerTo").CmpStr = IIF(adc("UoYN") = "Y", "c.", "ce.") + "CmSalPer <= "
'  adc("UoSalPerSel").CmpStr = IIF(adc("UoYN") = "Y", "c.", "ce.") + "CmSalPer In "
  
  ' ****** Manali 3.9.0 - Option for InDt / InPrnDt added
  Dim wInvDt As String
  If ADC("UoYN3") = "Y" Then
    wInvDt = "InDt "
  Else
    wInvDt = "InPrnDt "
  End If
    ADC("UoDtFr").CmpStr = wInvDt + ">= "
    ADC("UoDtTo").CmpStr = wInvDt + "<= "
  ' ****** Manali 3.9.0 - Option for InDt / InPrnDt added
  
  If UCase(ADC("UoInVal")) = "A" Then
    wOpt = "Act"     '''"Actual"
    wiLLabRtOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLARt/InCnvFct else iLARt*InCnvFct end) ", "iLARt")
    wiLLabValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLAVal/InCnvFct else iLAVal*InCnvFct end) ", "iLAVal")
  ElseIf UCase(ADC("UoInVal")) = "I" Then
    wOpt = "Inv"      '''"Invoice"
    wiLLabRtOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiRt/InCnvFct else iLiRt*InCnvFct end) ", "iLiRt")
    wiLLabValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiVal/InCnvFct else iLiVal*InCnvFct end) ", "iLiVal")
  ElseIf UCase(ADC("UoInVal")) = "Z" Then
    wOpt = "Custm"    ''''"Customs"
    wiLLabRtOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZRt/InCnvFct else iLZRt*InCnvFct end) ", "iLZRt")
    wiLLabValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZVal/InCnvFct else iLZVal*InCnvFct end) ", "iLZVal")
  End If
  
  ADC("UoSalRtFr").CmpStr = wiLLabRtOpt + ">= "
  ADC("UoSalRtTo").CmpStr = wiLLabRtOpt + "<= "
  ADC("UoCoCdFr").CmpStr = "iLCoCd = "
  ADC("UoInTcFr").CmpStr = "iLTc>= "
  ADC("UoInTcTo").CmpStr = "iLTc<= "
  ADC("UoInYyFr").CmpStr = "iLYy>= "
  ADC("UoInYyTo").CmpStr = "iLYy<= "
  ADC("UoInChrFr").CmpStr = "iLchr>= "
  ADC("UoInChrTo").CmpStr = "iLChr<= "
  ADC("UoInNoFr").CmpStr = "iLNo>= "
  ADC("UoInNoTo").CmpStr = "iLNo<= "
  ADC("UoPmCdFr").CmpStr = "iLMCd >= "
  ADC("UoPmCdTo").CmpStr = "iLMcd <= "
  ADC("UoPmCdSel").CmpStr = "iLMCd In "
  ADC("UoLsCdFr").CmpStr = "iLSCd >= "
  ADC("UoLsCdTo").CmpStr = "iLSCd <= "
    
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
   
  If ADC("UoRmCtgFr") <> "" Or ADC("UoRmCtgTo") <> "" Or ADC("UoRmSCtgFr") <> "" Or ADC("UoRmSCtgTo") <> "" Or _
         ADC("UoRmCdFr") <> "" Or ADC("UoRmCdTo") <> "" Then
        wCnd = IIF(wCnd <> "", " and ", " where ") + "iLMCd='SET'"
  End If
  
  ' Zubin 213
  wCnd = wCnd + IIF(wb_InBaseCur = False, " and CE.CmCurCd= '" + ADC("UoCurCdFr") + "'", "")
   
  Select Case UCase(ADC("UoSaLabGrp"))
  Case Is = "SR"
    wLabDet = " '' as IrRmCtg,'' as IrRmSCtg,'' as IrRmCd, max(iLMCd) as iLMCd, max(iLSCd) as iLSCd, max(" + wiLLabRtOpt + ") as qLabRt, "
    wLabGrp = ", iLSrNo "
    Rep.hiLSrNo.Suppress = False: Rep.FldiLSrNo.Suppress = False
    Rep.hiLRt.SetText wOpt + Chr(13) + "Rate"
  Case Is = "LABRT"
    wLabDet = " '' as IrRmCtg,'' as IrRmSCtg,'' as IrRmCd, max(iLMCd) as iLMCd, iLSCd, " + wiLLabRtOpt + " as qLabRt, "
    wLabGrp = ", iLSCd, " + wiLLabRtOpt + " "
    Rep.hiLSrNo.Suppress = False: Rep.FldiLSrNo.Suppress = False
    Rep.hiLRt.SetText wOpt + Chr(13) + "Rate"
  Case Is = "LABSCD"
    wLabDet = " '' as IrRmCtg,'' as IrRmSCtg,'' as IrRmCd, iLMCd, iLSCd, 0 as qLabRt, "
    wLabGrp = ", iLMCd, iLSCd "
    Rep.hiLSrNo.Suppress = True: Rep.FldiLSrNo.Suppress = True
    Rep.hiLRt.SetText " "
  Case Is = "LABMCD"
    If wLabGrpOpt = "Lab Sub Cd" Then _
      DispMsg "Group Option Has To Be A Higher Level Than Detail", etError: wCnd = IIF(wCnd <> "", " and 1= 2", " where 1= 2")
    wLabDet = " '' as IrRmCtg,'' as IrRmSCtg,'' as IrRmCd, iLMCd, '' as iLSCd, 0 as qLabRt, "
    wLabGrp = ", iLMCd "
    Rep.hiLSrNo.Suppress = True: Rep.FldiLSrNo.Suppress = True
    Rep.hiLRt.SetText " ": Rep.hiLSCd.SetText " "
  End Select
  
  
   ws_DsgAnaCnd = " And (Select Count(*) From UsrOptDsgAna " + _
                 "Join DsgAna on " + IIF(gs_Partition = ctCurrPrtn, " DaPrtKey='" + ctCurrPrtn + "' and ", "") + " DaAnaSr = UdAnaSr " + _
                 "where (DaAnaCd between UdAnaCdFr and UdAnaCdTo) " + _
                 "and DaTcTyp = DmTcTyp and DaCd= DmCd and DaSz= DmSz and " + _
                 "UdUsrCd = '" + ADC("UoUsrCd") + "' and " + _
                 "UdMnuCd= '" + ADC("UoMnuCd") + "')= " & (GrdDsgAna.Rows - 1) & " "
 
    
  'Uni.4 - SalCtg scope will check all of 3 Sal Ctg(SalCtg, SalCtg2 & SalCtg3).
  'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is
  'calling function to get where clause to join DsgMst
  Dim wDmSalCtgCnd As String, ws_DmSalCtg As String
  If ADC("UoDmSalCtgFr") <> "" Or ADC("UoDmSalCtgTo") <> "" Or ADC("UoDmSalCtgSel") <> "" Then
    wDmSalCtgCnd = GetDmSalCtgWhrCnd(ADC("UoDmSalCtgFr"), ADC("UoDmSalCtgTo"), ADC("UoDmSalCtgSel"), ADC("wAddAllSalCtg"))
  End If
    
  'uni.4 selecting Dmsalctg either from SalCtg,SalCtg2 or SalCtg3
  'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is
  'calling function to get SalCtg value to show in detail section and group sort
  ws_DmSalCtg = GetDmSalCtg(ADC("UoDmSalCtgFr"), ADC("UoDmSalCtgTo"), ADC("UoDmSalCtgSel"), ADC("wAddAllSalCtg"), "Y")
  
  
  '****** Sachin 3.02 - Id fields in Joins
  ' ***** Manali  3.10.0 - 10/03/12 - Sort Based on InDt/PrnDt
  'uni.4 wDmSalCtgCnd added to DsgMst join, to get SalCtg value from scope.
  'To avoid SalCtg scope condition in where clause
  wSqlStrg = "Select " + grpflds0 + ", " + _
             " IdCoCd as IlCoCd, IdTc as IlTc, IdYy as IlYy, IdChr as IlChr, IdNo as IlNo, IdSr as IlSr, 'L'+ Rtrim(Ltrim(str(max(iLSrNo)))) as IlSrNo, " + _
             " max(InExpNo) as InExpNo, max(InDt) as InDt, max(" + wCust + ") as InCmCd, max(DmCtg) as DmCtg, " + _
             " max(IdDmCd) as IdDmCd, max(IdDmSz) as IdDmSz, max(IdSfx) as IdSfx, max(IdQty) as IdQty, " + _
             " max(IdExpOdTc) as IdExpOdTc, max(IdExpOdYy) as IdExpOdYy, max(IdExpOdChr) as IdExpOdChr, " + _
             " max(IdExpOdNo) as IdExpOdNo, max(IdExpOdSr) as IdExpOdSr, max(IdVaCtg) as IdVaCtg, " + wLabDet + _
             " sum(" + wiLLabValOpt + ") as qLabVal " + _
             "From InvLab join InvDsg on IlIdIdNo=IdIdNo and IdCoCd= iLCoCd and IdTc= iLTc and IdYy= iLYy " + _
             " and IdChr= iLChr and IdNo= iLNo and IdSr= iLSr " + IIF(gs_Partition = ctCurrPrtn, " and IdPrtKey=IlPrtKey ", "") + _
             " join InvHd on InCoCd= iLCoCd and InTc= iLTc and InYy= iLYy " + _
             "  and InChr= iLChr and InNo= iLNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IlPrtKey ", "") + _
             " Join OrdDsg on OdCoCd= IdCoCd and OdTc=IdExpOdTc and OdYy=IdExpOdYy and OdChr =IdExpOdChr and OdNo=IdExpOdNo and OdSr = IdExpOdSr and OdIdNo=IdOdIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=IdPrtKey ", "") + _
             " Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=IdPrtKey ", "") + _
             " Join CustMst CE on CE.CmCtg= 'C' and CE.CmCd = IdExpCmCd " + _
             " Join CustMst C on C.CmCtg= 'C' and C.CmCd =InCmCd " + _
             " join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=IdPrtKey ", "") + wDmSalCtgCnd + wOrdJoin + _
             " Join (Select '" + Trim(wInvDt) + "' as qDtOpt) A on 1=1 " + _
             wCnd + ws_DsgAnaCnd + IIF(gs_Partition = ctCurrPrtn, " and IlPrtKey='" + ctCurrPrtn + "' ", "") + " Group by IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr " + wLabGrp
    
  '********************************* SetCd*******************************************
  
  If UCase(ADC("UoInVal")) = "A" Then
    wOpt = "Act"
    wiLLabRtOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetARt/InCnvFct else IrSetARt*InCnvFct end) ", "IrSetARt")
    wiLLabValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) ", "IrSetAVal")
  ElseIf UCase(ADC("UoInVal")) = "I" Then
    wOpt = "Inv"
    wiLLabRtOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIRt/InCnvFct else IrSetIRt*InCnvFct end) ", "IrSetIRt")
    wiLLabValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIVal/InCnvFct else IrSetIVal*InCnvFct end) ", "IrSetIVal")
  ElseIf UCase(ADC("UoInVal")) = "Z" Then
    wOpt = "Custm"
    wiLLabRtOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZRt/InCnvFct else IrSetZRt*InCnvFct end) ", "IrSetZRt")
    wiLLabValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZVal/InCnvFct else IrSetZVal*InCnvFct end) ", "IrSetZVal")
  End If
  
  ADC("UoSalRtFr").CmpStr = wiLLabRtOpt + ">= "
  ADC("UoSalRtTo").CmpStr = wiLLabRtOpt + "<= "
  ADC("UoCoCdFr").CmpStr = "IrCoCd = "
  ADC("UoInTcFr").CmpStr = "IrTc>= "
  ADC("UoInTcTo").CmpStr = "IrTc<= "
  ADC("UoInYyFr").CmpStr = "IrYy>= "
  ADC("UoInYyTo").CmpStr = "IrYy<= "
  ADC("UoInChrFr").CmpStr = "Irchr>= "
  ADC("UoInChrTo").CmpStr = "IrChr<= "
  ADC("UoInNoFr").CmpStr = "IrNo>= "
  ADC("UoInNoTo").CmpStr = "IrNo<= "
  ADC("UoPmCdFr").CmpStr = "'SET' >= "
  ADC("UoPmCdTo").CmpStr = "'SET' <= "
  ADC("UoPmCdSel").CmpStr = "'SET' In "
  ADC("UoLsCdFr").CmpStr = "IrSetSCd >= "
  ADC("UoLsCdTo").CmpStr = "IrSetSCd <= "
  ADC("UoRmCtgFr").CmpStr = "IrRmCtg>="
  ADC("UoRmCtgTo").CmpStr = "IrRmCtg<="
  ADC("UoRmSCtgFr").CmpStr = "IrRmSCtg>="
  ADC("UoRmSCtgTo").CmpStr = "IrRmSCtg<="
  ADC("UoRmCdFr").CmpStr = "IrRmCd>="
  ADC("UoRmCdTo").CmpStr = "IrRmCd<="
    
  wrepcnd = ADC.RepCond
  'wCnd = IIF(wrepcnd <> "", " where IrSetSCd<>'' and " + wrepcnd, "")      '*** (Bef 3.2.0)
  wCnd = " where IrSetSCd<> '' " + IIF(wrepcnd <> "", " and " + wrepcnd, "")  '*** (Jen 3.2.0)
  
  ' Zubin 213
  wCnd = wCnd + IIF(wb_InBaseCur = False, " and CE.CmCurCd= '" + ADC("UoCurCdFr") + "'", "")
  
  Select Case UCase(ADC("UoSaLabGrp"))
  Case Is = "SR"
    wLabDet = " max(IrRmCtg) as IrRmCtg,max(IrRmSCtg) as IrRmSCtg,max(IrRmCd) as IrRmCd,'SET' as iLMCd, max(IrSetSCd) as iLSCd, max(" + wiLLabRtOpt + ") as qLabRt, "
    wLabGrp = ", 'R' + Str(IrSrNo) "
    Rep.hiLSrNo.Suppress = False: Rep.FldiLSrNo.Suppress = False
    Rep.hiLRt.SetText wOpt + Chr(13) + "Rate"
  Case Is = "LABRT"
    wLabDet = " '' as IrRmCtg,'' as IrRmSCtg,'' as IrRmCd,'SET' as iLMCd, IrSetSCd as iLSCd, " + wiLLabRtOpt + " as qLabRt, "
    wLabGrp = ", IrSetSCd, " + wiLLabRtOpt + " "
    Rep.hiLSrNo.Suppress = False: Rep.FldiLSrNo.Suppress = False
    Rep.hiLRt.SetText wOpt + Chr(13) + "Rate"
  Case Is = "LABSCD"
    wLabDet = " '' as IrRmCtg,'' as IrRmSCtg,'' as IrRmCd,'SET' as iLMCd, IrSetSCd as iLSCd, 0 as qLabRt, "
    wLabGrp = ", IrSetSCd "
    Rep.hiLSrNo.Suppress = True: Rep.FldiLSrNo.Suppress = True
    Rep.hiLRt.SetText " "
  Case Is = "LABMCD"
    If wLabGrpOpt = "Lab Sub Cd" Then _
      DispMsg "Group Option Has To Be A Higher Level Than Detail", etError: wCnd = IIF(wCnd <> "", " and 1= 2", " where 1= 2")
    wLabDet = " '' as IrRmCtg,'' as IrRmSCtg,'' as IrRmCd,'SET' as iLMCd, '' as iLSCd, 0 as qLabRt, "
    wLabGrp = ""
    Rep.hiLSrNo.Suppress = True: Rep.FldiLSrNo.Suppress = True
    Rep.hiLRt.SetText " ": Rep.hiLSCd.SetText " "
  End Select
 
  '****** Sachin 3.02 - Id fields in Joins
  ' ***** Manali  3.10.0 - 10/03/12 - Sort Based on InDt/PrnDt
  'uni.4 wDmSalCtgCnd added to DsgMst join, to get SalCtg value from scope.
  'To avoid SalCtg scope condition in where clause
  wSqlSetStrg = "Select " + grpflds1 + ", " + _
             "IdCoCd as IlCoCd, IdTc as IlTc, IdYy as IlYy, IdChr as IlChr, IdNo as IlNo, IdSr as IlSr, " + _
             "'R'+ Rtrim(Ltrim(str(max(IrSrNo)))) as IlSrNo, max(InExpNo) as InExpNo, max(InDt) as InDt, " + _
             "max(" + wCust + ") as InCmCd, max(DmCtg) as DmCtg, max(IdDmCd) as IdDmCd, max(IdDmSz) as IdDmSz, " + _
             "max(IdSfx) as IdSfx, max(IdQty) as IdQty, max(IdExpOdTc) as IdExpOdTc, max(IdExpOdYy) as IdExpOdYy, " + _
             "max(IdExpOdChr) as IdExpOdChr, max(IdExpOdNo) as IdExpOdNo, " + _
             "max(IdExpOdSr) as IdExpOdSr, max(IdVaCtg) as IdVaCtg, " + wLabDet + _
             "sum(" + wiLLabValOpt + ") as qLabVal " + _
             "From InvRm join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy " + _
             "and IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + IIF(gs_Partition = ctCurrPrtn, " and IdPrtKey=IrPrtKey ", "") + _
             "join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy " + _
             "and InChr= IrChr and InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IrPrtKey ", "") + _
             " Join OrdDsg on OdCoCd= IdCoCd and OdTc=IdExpOdTc and OdYy=IdExpOdYy and OdChr =IdExpOdChr and OdNo=IdExpOdNo and OdSr = IdExpOdSr and OdIdNo=IdOdIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=IdPrtKey ", "") + _
             " Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=IdPrtKey ", "") + _
             "Join CustMst CE on CE.CmCtg= 'C' and CE.CmCd = IdExpCmCd " + _
             "Join CustMst C on C.CmCtg= 'C' and C.CmCd = InCmCd " + _
             " join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=IdPrtKey ", "") + wDmSalCtgCnd + wOrdJoin + _
            "Join (Select '" + Trim(wInvDt) + "' as qDtOpt) A on 1=1 " + _
             wCnd + ws_DsgAnaCnd + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey='" + ctCurrPrtn + "' ", "") + " Group by IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr " + wLabGrp
  
  'uni.4 replacing DmSalCtg value and its description in group sort of Sales Ctg
  wSqlStrg = Replace(wSqlStrg, "||qDmSalCtgRepl||", ws_DmSalCtg)
  wSqlStrg = Replace(wSqlStrg, "||qDmSalCtgDescRepl||", ws_DmSalCtg)
  wSqlSetStrg = Replace(wSqlSetStrg, "||qDmSalCtgRepl||", ws_DmSalCtg)
  wSqlSetStrg = Replace(wSqlSetStrg, "||qDmSalCtgDescRepl||", ws_DmSalCtg)
  
  ADC.RepRecSource = wSqlStrg + " Union All " + wSqlSetStrg + " Order By Grp1, Grp2, Grp3, IlCoCd, IlTc, IlYy, IlChr, IlNo, IlSr, IlSrNo " + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
 
 
  ' ***** Manali 3.10.0 - 09/02/12 - Dsg Qty added
  ShowRepSummFld Rep.Fldr1IdQty, Rep.Fldr2IdQty, Rep.Fldr3IdQty, Rep.FldrIdQty, GRP_REP.GrpOpts, "Lab Main Cd,Lab Sub Cd"

 
'  Rep.h1Val.SetText wOpt + Chr(13) + "Value": Rep.h2Val.SetText wOpt + Chr(13) + "Value"
  Rep.h1Val.SetText wOpt + " Value": Rep.h2Val.SetText wOpt + " Value"

  ' Zubin 213
  If wb_InBaseCur Then
    Rep.TxtValIn.SetText "Value In " + moCn.GetFldVal("Select HBaseCurCd from Head " + _
                                        "where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ADC("UoCoCdFr") + "'") + ""
  Else
    Rep.TxtValIn.SetText "Value In " + UCase(ADC("UoCurCdFr"))
  End If
  
  ' ***** Manali 3.9.0 - Fr and To Date Scope Report
  Rep.TxtFrDt.SetText ADC("UoDtFr")
  Rep.TxtToDt.SetText ADC("UoDtTo")
  ' ***** Manali 3.9.0 - Fr and To Date Scope Report
  
  ' ***** Manali 3.10.0 - Date Scope based on Date Selected
  If ADC("UoYN3") = "Y" Then
    Rep.LblFrDt.SetText "Fr Inv Dt :": Rep.LblToDt.SetText "To Inv Dt :"
  Else
    Rep.LblFrDt.SetText "Fr Prn Dt :": Rep.LblToDt.SetText "To Prn Dt :"
  End If
  ' ***** Manali 3.10.0 - Date Scope based on Date Selected
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.IrRmCd}", 7)
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******

End Sub

' Zubin 213
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
  Case Is = UCase("UoCurCdFr")
    If UCase(ADC("UoInVal")) = "S" Then ADC("UoCurCdFr") = "": Exit Sub
  End Select
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
  With ChkBoxArr(Index)
    Select Case Index
    'uni.4  checking Include SalCtg2,SalCtg3 option
    Case Is = 0
      If .Value = Checked Then ADC("wAddAllSalCtg") = "Y" Else ADC("wAddAllSalCtg") = "N"
    End Select
  End With
End Sub
'uni.4 Tooltip for check box
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub

Private Sub CmdSc_Click()
  If FraSC.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraSC.Enabled = True
    FraSC.Visible = True
    TAB_REP.Enabled = False
    ADC.AllowSave = False
    FraSC.ZOrder
  Else
    TAB_REP.Enabled = True
    ADC.AllowSave = True
    FraSC.Enabled = False
    FraSC.Visible = False
  End If
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
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = False
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    ADC.Child = "GrdDsgAna"
    Set GrdDsgAna.PrevCtl = ADC
    Set GrdDsgAna.NextCtl = ADC
    GrdDsgAna.HotKey = "A"
    
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
    Set ADC.FirNKeyCtl = ADC("UoCoCdFr")
  Else
    Set ADC.FirNKeyCtl = ADC("UOInTcFr")
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  'Set adc.FirNKeyCtl = adc("UoInTcFr") 'adc("UoCoCdFr")
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_PTyp = "LABMCD"
  gs_Tbl = "Param"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
      ADC("UOCOCDFR").Enabled = True: ADC("UOCOCDTO").Visible = True: ADC("UOCOCDSEL").Visible = True
  Else
      ADC("UOCOCDFR").Enabled = False: ADC("UOCOCDTO").Visible = False: ADC("UOCOCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)

  gb_CoCdFor = False  '****** Sachin 2.14.0 - [Multi-Company] ******

  Select Case UCase(IdName)
  Case Is = UCase("UoPmCdFr"), UCase("UoPmCdTo"), UCase("UoPmCdSel")
    gs_PTyp = "LABMCD"
    gs_Tbl = "Param"
  Case Is = UCase("UoLSCdFr")
    gs_LabMCd = ADC("UoPmCdFr")
  Case Is = UCase("UoLSCdTo")
    gs_LabMCd = ADC("UoPmCdTo")
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
    'gs_LabMCd = ""
    
  gb_CoCdFor = False  '****** Sachin 2.14.0 - [Multi-Company] ******
    
  Select Case UCase(IdName)
  Case Is = UCase("UoPmCdFr"), UCase("UoPmCdTo"), UCase("UoPmCdSel")
    gs_PTyp = "LABMCD"
    gs_Tbl = "Param"
    
  Case Is = UCase("UoLSCdFr")
    gs_LabMCd = ADC("UoPmCdFr")
    
  Case Is = UCase("UoLSCdTo")
    gs_LabMCd = ADC("UoPmCdTo")
    
  ' Zubin 213
  Case Is = UCase("UoCurCdFr")
    If UCase(ADC("UoInVal")) = "S" Then Cancel = True: Exit Sub
  End Select
  
  SetProp Me, IdName, When
End Sub

Private Sub GrdDsgAna_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  With GrdDsgAna
    Select Case UCase(ColName)
    Case Is = UCase("UdAnaSr")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                                "PTyp= 'DAANAFLD' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Analysis Sr.": Exit Sub
    End Select
  End With
End Sub

Private Sub GrdDsgAna_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
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
' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  '*** Initialise Key Fields
      'The fields of the Base table which are not existing in the grid are stored in the
      'table using Grid_InitKey
  
  With GrdDsgAna
    If .IsNew(RowNum) Then
      .Store "UdUsrCd", ADC("UoUsrCd")
      .Store "UdMnuCd", ADC("UoMnuCd")
    End If
  End With

End Sub

Private Sub GrdDsgAna_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
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
  '*** Set the Record Source of the Grid GrdDsgAna
  ' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  GrdDsgAna.RecSource = "Select * from UsrOptDsgAna where UdUsrCd= '" + ADC("UoUsrCd") + "' and " + _
                        "UdMnuCd= '" + ADC("UoMnuCd") + "'"
End Sub

Private Sub GrdDsgAna_Validate(Cancel As Boolean)
' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  Cancel = GrdDsgAna.Validate
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptCust_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN") = "Y"
Case Is = 1
  ADC("UoYN") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptCust_GotFocus(Index As Integer)
  DispMsg FraOptCust.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

Private Sub OptInvDt_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN3") = "Y"
Case Is = 1
  ADC("UoYN3") = "N"
End Select
End Sub

Private Sub OptInvDt_GotFocus(Index As Integer)
  DispMsg FraOptDt.ToolTipText, etInfo
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
  
  ' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
  If ADC("UoYN") = "Y" Then OptCust(0).Value = True
  If ADC("UoYN") = "N" Then OptCust(1).Value = True
  ' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
  ' ****** Manali 3.9.0 - Option for InDt / InPrnDt added
  If ADC("UoYN3") = "Y" Then OptInvDt(0).Value = True
  If ADC("UoYN3") = "N" Then OptInvDt(1).Value = True
  ' ****** Manali 3.9.0 - Option for InDt / InPrnDt added
  'uni.4 default value for Include SalCtg2,SalCtg3 is Yes
  ChkBoxArr(0).Value = Checked
  ADC("wAddAllSalCtg") = "Y"

End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = True, gs_CoCd, ADC("UOCOCDFR")) + "' and " + _
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
  If ws_MultiCoMod = True Then
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdFr.SetText ADC("UOCOCDFR"): Rep.TxtCoCdTo.SetText ADC("UOCOCDTO"): Rep.TxtCoCdSel.SetText ADC("UOCOCDSEL")
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
