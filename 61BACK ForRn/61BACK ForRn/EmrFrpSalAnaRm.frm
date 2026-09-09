VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpSalAnaRm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Analysis (Rm Level)"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdSc 
      Height          =   390
      Left            =   2040
      TabIndex        =   158
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9720
      WhatsThisHelpID =   9800
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
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   60
      TabIndex        =   102
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
      Left            =   10860
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   480
      Left            =   4455
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   847
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4965
      TabIndex        =   110
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   111
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
         TabIndex        =   112
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
      Height          =   465
      Left            =   6180
      TabIndex        =   109
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   106
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
      Height          =   9735
      Left            =   0
      TabIndex        =   105
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   104
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpSalAnaRm.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpSalAnaRm.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   108
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   103
               Top             =   360
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   107
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Include SalCtg2, SalCtg3   "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Index           =   1
               Left            =   30
               TabIndex        =   55
               Tag             =   "DmHold"
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   5400
               Width           =   2150
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1995
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   88
               Top             =   8430
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
               Index           =   0
               Left            =   5000
               TabIndex        =   89
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Stock Rate"
               Top             =   8430
               Width           =   2050
            End
            Begin VB.Frame FraOptDt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   9915
               TabIndex        =   151
               ToolTipText     =   "Select if to show Invoice Date Or Invoice Printing Date"
               Top             =   8460
               Width           =   4365
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
                  Left            =   2640
                  TabIndex        =   95
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
                  TabIndex        =   94
                  ToolTipText     =   "Select if to show Invoice Date Or Invoice Printing Date"
                  Top             =   0
                  Width           =   1875
               End
            End
            Begin VB.Frame FraOptCust 
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   9915
               TabIndex        =   149
               ToolTipText     =   "Select eihter Invoice Customer Or Order Customer"
               Top             =   8175
               Width           =   4305
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
                  Left            =   2640
                  TabIndex        =   93
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
                  TabIndex        =   92
                  Top             =   0
                  Width           =   1545
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5820
               TabIndex        =   7
               ToolTipText     =   "Enter To Invoice Year"
               Top             =   525
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYTO"
               IdName          =   "UOINYYTO"
               CmpStr          =   "IrYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2760
               TabIndex        =   4
               ToolTipText     =   "Enter From Invoice Year"
               Top             =   525
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "IrYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1995
               TabIndex        =   3
               ToolTipText     =   "Enter From Invoice Tc"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "IrTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3180
               TabIndex        =   5
               ToolTipText     =   "Enter From Invoice Character"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "IrChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   5040
               TabIndex        =   6
               ToolTipText     =   "Enter To Invoice Tc"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCTO"
               IdName          =   "UOINTCTO"
               CmpStr          =   "IrTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   6240
               TabIndex        =   8
               ToolTipText     =   "Enter To Invoice Character"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRTO"
               IdName          =   "UOINCHRTO"
               CmpStr          =   "IrChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1995
               TabIndex        =   9
               ToolTipText     =   "Enter From Invoice Number"
               Top             =   810
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "IrNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   5040
               TabIndex        =   11
               ToolTipText     =   "Enter To Invoice Number"
               Top             =   810
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOTO"
               IdName          =   "UOINNOTO"
               CmpStr          =   "IrNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   1995
               TabIndex        =   49
               ToolTipText     =   "Enter From Design Category"
               Top             =   4830
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
               Left            =   5040
               TabIndex        =   57
               ToolTipText     =   "Enter To Design Code"
               Top             =   5880
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
               Left            =   8100
               TabIndex        =   58
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   5880
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "IdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   5040
               TabIndex        =   16
               ToolTipText     =   "Enter To Customer"
               Top             =   1380
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
               Left            =   8100
               TabIndex        =   17
               ToolTipText     =   "Enter Customer Selection"
               Top             =   1380
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1995
               TabIndex        =   15
               ToolTipText     =   "Enter From Customer"
               Top             =   1380
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
               Left            =   1995
               TabIndex        =   63
               ToolTipText     =   "Enter From Karat"
               Top             =   6735
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
               Left            =   5040
               TabIndex        =   53
               ToolTipText     =   "Enter To Design Sales Category"
               Top             =   5115
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
               Left            =   8100
               TabIndex        =   54
               ToolTipText     =   "Enter Design Sales Category Selection"
               Top             =   5115
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMSALCTGSEL"
               IdName          =   "UODMSALCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   1995
               TabIndex        =   52
               ToolTipText     =   "Enter From Design Sales Category"
               Top             =   5115
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
               Left            =   5040
               TabIndex        =   25
               ToolTipText     =   "Enter To Export Number"
               Top             =   2265
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
               Left            =   8100
               TabIndex        =   26
               ToolTipText     =   "Enter Export Number Selection"
               Top             =   2265
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOINEXPNOSEL"
               IdName          =   "UOINEXPNOSEL"
               CmpStr          =   "InExpNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1995
               TabIndex        =   24
               ToolTipText     =   "Enter From Export Number"
               Top             =   2265
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
               Left            =   5040
               TabIndex        =   60
               ToolTipText     =   "Enter To Design Size"
               Top             =   6165
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
               Index           =   37
               Left            =   5040
               TabIndex        =   62
               ToolTipText     =   "Enter To Suffix"
               Top             =   6450
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
               Left            =   1995
               TabIndex        =   61
               ToolTipText     =   "Enter From Suffix"
               Top             =   6450
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
               Index           =   21
               Left            =   7550
               TabIndex        =   78
               ToolTipText     =   "Enter To Order Serial"
               Top             =   7305
               Width           =   500
               _ExtentX        =   873
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
               Left            =   4500
               TabIndex        =   73
               ToolTipText     =   "Enter From Order Serial"
               Top             =   7305
               Width           =   500
               _ExtentX        =   873
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
               Left            =   3630
               TabIndex        =   72
               ToolTipText     =   "Enter From Order Number"
               Top             =   7305
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
               Left            =   6680
               TabIndex        =   77
               ToolTipText     =   "Enter To Order Number"
               Top             =   7305
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
               Left            =   14670
               TabIndex        =   98
               ToolTipText     =   "Show Invoice Or Order Customer ('Y' for Invoice / 'N' For Order)"
               Top             =   8085
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
               Left            =   1995
               TabIndex        =   87
               ToolTipText     =   "Enter Level Of Detail To Be Shown"
               Top             =   8145
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOSARMGRP"
               IdName          =   "UOSARMGRP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   5040
               TabIndex        =   40
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   3690
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "IrRmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   8100
               TabIndex        =   41
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   3690
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "IrRmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   5040
               TabIndex        =   34
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   3120
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGTO"
               IdName          =   "UORMSCTGTO"
               CmpStr          =   "Rm.RmSCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   8100
               TabIndex        =   35
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   3120
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "Rm.RmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   1995
               TabIndex        =   27
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   2550
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "Rm.RmCtg >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1995
               TabIndex        =   33
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   3120
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
               Data            =   "ABCDW"
               CmpStr          =   "Rm.RmSCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   1995
               TabIndex        =   39
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   3690
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "IrRmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   5040
               TabIndex        =   43
               ToolTipText     =   "Enter To Sale Rate"
               Top             =   3975
               Width           =   1245
               _ExtentX        =   2196
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
               Left            =   1995
               TabIndex        =   42
               ToolTipText     =   "Enter From Sale Rate"
               Top             =   3975
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOSALRTFR"
               IdName          =   "UOSALRTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   5040
               TabIndex        =   28
               ToolTipText     =   "Enter To Raw Material Category"
               Top             =   2550
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
               CmpStr          =   "Rm.RmCtg <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   8100
               TabIndex        =   29
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   2550
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCTGSEL"
               IdName          =   "UORMCTGSEL"
               CmpStr          =   "Rm.RmCtg In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   5040
               TabIndex        =   64
               ToolTipText     =   "Enter To Karat"
               Top             =   6735
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
               Left            =   8100
               TabIndex        =   65
               ToolTipText     =   "Enter Karat Selection"
               Top             =   6735
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "IdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   5040
               TabIndex        =   14
               ToolTipText     =   "Enter To Invoice Voucher Date"
               Top             =   1095
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
               Index           =   55
               Left            =   1995
               TabIndex        =   13
               ToolTipText     =   "Enter From Invoice Voucher Date"
               Top             =   1095
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
               Index           =   56
               Left            =   1995
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   240
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "IrCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   8100
               TabIndex        =   68
               ToolTipText     =   "Enter Sales Executive Selection"
               Top             =   7020
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSALPERSEL"
               IdName          =   "UOSALPERSEL"
               CmpStr          =   "OmCmSalPer In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   5040
               TabIndex        =   67
               ToolTipText     =   "Enter To Sales Executive"
               Top             =   7020
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
               Index           =   59
               Left            =   1995
               TabIndex        =   66
               ToolTipText     =   "Enter From Sales Executive"
               Top             =   7020
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
               Index           =   1
               Left            =   5610
               TabIndex        =   75
               ToolTipText     =   "Enter To Order Year"
               Top             =   7305
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
               Index           =   13
               Left            =   5040
               TabIndex        =   74
               ToolTipText     =   "Enter To Order Tc"
               Top             =   7305
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "IdExpOdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   6020
               TabIndex        =   76
               ToolTipText     =   "Enter To Order Character"
               Top             =   7305
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "IdExpOdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   2570
               TabIndex        =   70
               ToolTipText     =   "Enter From Order Year"
               Top             =   7305
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
               Left            =   1995
               TabIndex        =   69
               ToolTipText     =   "Enter From Order Tc"
               Top             =   7305
               Width           =   530
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "IdExpOdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   2970
               TabIndex        =   71
               ToolTipText     =   "Enter From Order Character"
               Top             =   7305
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "IdExpOdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   5040
               TabIndex        =   45
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   4260
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1TO"
               IdName          =   "UOLN1TO"
               CmpStr          =   "IrRmSz<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   1995
               TabIndex        =   44
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   4260
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1FR"
               IdName          =   "UOLN1FR"
               CmpStr          =   "IrRmSz>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   5040
               TabIndex        =   50
               ToolTipText     =   "Enter To Design Category"
               Top             =   4830
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
               Left            =   8100
               TabIndex        =   51
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   4830
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
               CmpStr          =   "DmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   9915
               TabIndex        =   96
               Top             =   8730
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
               Left            =   5040
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   240
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "IrCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   8100
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   270
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "IrCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   1995
               TabIndex        =   46
               ToolTipText     =   "Enter From Lot No."
               Top             =   4545
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataType        =   4
               DataField       =   "UOLOTNOFR"
               IdName          =   "UOLOTNOFR"
               CmpStr          =   "IrLotNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   64
               Left            =   5040
               TabIndex        =   47
               ToolTipText     =   "Enter To Lot No."
               Top             =   4545
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOLOTNOTO"
               IdName          =   "UOLOTNOTO"
               CmpStr          =   "IrLotNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   65
               Left            =   8100
               TabIndex        =   48
               ToolTipText     =   "Enter Lot No. Selection"
               Top             =   4545
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOLOTNOSEL"
               IdName          =   "UOLOTNOSEL"
               CmpStr          =   "IrLotNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   80
               Left            =   14670
               TabIndex        =   99
               ToolTipText     =   "Show invoice Date/ Printng Dat (Yes / No)"
               Top             =   8370
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
               Index           =   66
               Left            =   5040
               TabIndex        =   37
               ToolTipText     =   "Enter To Parent Code"
               Top             =   3405
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOPARENTCDTO"
               IdName          =   "UOPARENTCDTO"
               CmpStr          =   "Rm.RmParentCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   67
               Left            =   1995
               TabIndex        =   36
               ToolTipText     =   "Enter From Parent Code"
               Top             =   3405
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOPARENTCDFR"
               IdName          =   "UOPARENTCDFR"
               CmpStr          =   "Rm.RmParentCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   68
               Left            =   8100
               TabIndex        =   38
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   3405
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
               CmpStr          =   "Rm.RmParentCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1995
               TabIndex        =   56
               ToolTipText     =   "Enter From Design Code"
               Top             =   5880
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
               Index           =   36
               Left            =   1995
               TabIndex        =   59
               ToolTipText     =   "Enter From Design Size"
               Top             =   6165
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
               Index           =   85
               Left            =   8100
               TabIndex        =   20
               ToolTipText     =   "Enter Customer Group Selection"
               Top             =   1680
               Width           =   6930
               _ExtentX        =   12224
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
               Left            =   5040
               TabIndex        =   19
               ToolTipText     =   "Enter To Customer Group"
               Top             =   1680
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMGRPTO"
               IdName          =   "UOCMGRPTO"
               CmpStr          =   "CE.CmGrp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   87
               Left            =   1995
               TabIndex        =   18
               ToolTipText     =   "Enter From Customer Group"
               Top             =   1680
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
               Left            =   8100
               TabIndex        =   23
               ToolTipText     =   "Enter Customer Type Selection"
               Top             =   1980
               Width           =   6930
               _ExtentX        =   12224
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
               Left            =   5040
               TabIndex        =   22
               ToolTipText     =   "Enter To Customer Type"
               Top             =   1980
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
               Left            =   1995
               TabIndex        =   21
               ToolTipText     =   "Enter From Customer Type"
               Top             =   1980
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
               Left            =   2925
               TabIndex        =   10
               ToolTipText     =   "Enter From Invoice Serial"
               Top             =   810
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
               Left            =   5985
               TabIndex        =   12
               ToolTipText     =   "Enter From Invoice Serial"
               Top             =   810
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
               Index           =   3
               Left            =   1995
               TabIndex        =   90
               ToolTipText     =   "Specify Whether Value/ Weight To Be Shown Is Actual, Invoice Or Customs "
               Top             =   8715
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
               Index           =   69
               Left            =   6855
               TabIndex        =   91
               ToolTipText     =   "Specify Whether Value/ Weight To Be Shown Is Actual, Invoice Or Customs "
               Top             =   8715
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOINVAL1"
               ReCalcParent    =   "UOCURCDFR"
               IdName          =   "UOINVAL1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   93
               Left            =   5040
               TabIndex        =   80
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   7590
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
               Left            =   8100
               TabIndex        =   81
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   7590
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   95
               Left            =   1995
               TabIndex        =   79
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   7590
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
               Left            =   3285
               TabIndex        =   83
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   7875
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
               Left            =   6345
               TabIndex        =   85
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   7875
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
               Left            =   8100
               TabIndex        =   86
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   7875
               Visible         =   0   'False
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "OdPoNo in"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   12580
               TabIndex        =   97
               Top             =   8730
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOYN1"
               ReCalcParent    =   "UOCURCDFR"
               IdName          =   "UOYN1"
               CmpStr          =   "InCls = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   240
               Index           =   71
               Left            =   7380
               TabIndex        =   165
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   8400
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   423
               Enabled         =   0   'False
               DataField       =   "UOYN4"
               IdName          =   "UOYN4"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   115
               Left            =   1995
               TabIndex        =   82
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   7875
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
               Left            =   5040
               TabIndex        =   84
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   7875
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
               TabIndex        =   169
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   5460
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
               Index           =   72
               Left            =   1995
               TabIndex        =   30
               ToolTipText     =   "Enter Raw Material Stone Class From Range"
               Top             =   2835
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSTNCLSFR"
               IdName          =   "UOSTNCLSFR"
               CmpStr          =   "ISNULL(SC.PSCd,'') >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   74
               Left            =   5040
               TabIndex        =   31
               ToolTipText     =   "Enter Raw Material Stone Class To Range"
               Top             =   2835
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSTNCLSTO"
               IdName          =   "UOSTNCLSTO"
               CmpStr          =   "ISNULL(SC.PSCd,'') <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   76
               Left            =   8100
               TabIndex        =   32
               ToolTipText     =   "Enter Raw Material Stone Class Selection"
               Top             =   2835
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSTNCLSSEL"
               IdName          =   "UOSTNCLSSEL"
               CmpStr          =   "ISNULL(SC.PSCd,'') In"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Stone Class"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   60
               TabIndex        =   172
               Top             =   2835
               Width           =   1305
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
               Left            =   6630
               TabIndex        =   171
               Top             =   7320
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
               Index           =   11
               Left            =   3590
               TabIndex        =   170
               Top             =   7320
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
               Index           =   41
               Left            =   6285
               TabIndex        =   168
               Top             =   7875
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
               Left            =   3225
               TabIndex        =   167
               Top             =   7875
               Width           =   105
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
               Index           =   39
               Left            =   60
               TabIndex        =   166
               Top             =   8430
               Width           =   1575
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Cls (Y/N)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   11280
               TabIndex        =   164
               Top             =   8730
               Width           =   1605
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
               TabIndex        =   163
               Top             =   7875
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
               Index           =   37
               Left            =   60
               TabIndex        =   162
               Top             =   7590
               Width           =   1425
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Other Value As"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5040
               TabIndex        =   161
               Top             =   8730
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
               Index           =   36
               Left            =   5880
               TabIndex        =   157
               Top             =   810
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
               Left            =   2820
               TabIndex        =   156
               Top             =   810
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
               Index           =   34
               Left            =   60
               TabIndex        =   155
               Top             =   1980
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
               Left            =   60
               TabIndex        =   154
               Top             =   1680
               Width           =   1425
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
               Index           =   32
               Left            =   60
               TabIndex        =   153
               Top             =   3405
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
               Left            =   8100
               TabIndex        =   152
               Top             =   8460
               Width           =   1965
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
               TabIndex        =   150
               Top             =   4545
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
               Left            =   8100
               TabIndex        =   148
               Top             =   8730
               Width           =   1005
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Size/Ptr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   147
               Top             =   4260
               Width           =   1425
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
               Index           =   28
               Left            =   60
               TabIndex        =   146
               Top             =   7020
               Width           =   1785
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
               Index           =   23
               Left            =   60
               TabIndex        =   145
               Top             =   240
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
               Index           =   22
               Left            =   8100
               TabIndex        =   144
               Top             =   -30
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
               Index           =   21
               Left            =   60
               TabIndex        =   143
               Top             =   1095
               Width           =   1965
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
               Index           =   20
               Left            =   60
               TabIndex        =   142
               Top             =   3690
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
               Index           =   19
               Left            =   60
               TabIndex        =   141
               Top             =   3120
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
               Index           =   18
               Left            =   60
               TabIndex        =   140
               ToolTipText     =   "Location"
               Top             =   2550
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
               Index           =   15
               Left            =   60
               TabIndex        =   139
               Top             =   3975
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
               Left            =   60
               TabIndex        =   138
               Top             =   8145
               Width           =   1845
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Customer"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   8100
               TabIndex        =   137
               Top             =   8175
               Width           =   1515
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Ord Tc/Yy/Chr/No/Sr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   136
               Top             =   7305
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
               Index           =   8
               Left            =   5960
               TabIndex        =   135
               Top             =   7305
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
               Left            =   2930
               TabIndex        =   134
               Top             =   7305
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
               Left            =   5570
               TabIndex        =   133
               Top             =   7305
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
               Index           =   0
               Left            =   2500
               TabIndex        =   132
               Top             =   7305
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
               Left            =   7500
               TabIndex        =   131
               Top             =   7305
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
               Left            =   4450
               TabIndex        =   130
               Top             =   7305
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
               Left            =   60
               TabIndex        =   129
               Top             =   6450
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
               Left            =   60
               TabIndex        =   128
               Top             =   6165
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
               Left            =   60
               TabIndex        =   127
               Top             =   2265
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
               TabIndex        =   126
               Top             =   6735
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
               TabIndex        =   125
               Top             =   1380
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
               TabIndex        =   124
               Top             =   5880
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
               TabIndex        =   123
               Top             =   4830
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
               Left            =   60
               TabIndex        =   122
               Top             =   5115
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
               Left            =   60
               TabIndex        =   121
               Top             =   8730
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
               Left            =   1995
               TabIndex        =   120
               Top             =   -30
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
               Left            =   5040
               TabIndex        =   119
               Top             =   -30
               Width           =   1275
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
               Left            =   2700
               TabIndex        =   118
               Top             =   525
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
               Left            =   5760
               TabIndex        =   117
               Top             =   525
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
               Left            =   3120
               TabIndex        =   116
               Top             =   525
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
               Left            =   6180
               TabIndex        =   115
               Top             =   525
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
               Left            =   60
               TabIndex        =   114
               Top             =   810
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
               Left            =   60
               TabIndex        =   113
               Top             =   525
               Width           =   1725
            End
         End
      End
      Begin VB.Frame FraSc 
         Height          =   2925
         Left            =   1800
         TabIndex        =   159
         Top             =   6480
         Visible         =   0   'False
         Width           =   3945
         Begin MwfCtl.MWCTL_FLX GrdDsgAna 
            Height          =   2430
            Left            =   120
            TabIndex        =   160
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
Attribute VB_Name = "EmrFrpSalAnaRm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepSalAnaRm
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim mb_LotMod As Boolean  ' ***** Manali 350Nxt

Private Sub SetGroupSort()

  GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "DmCtg", "hDmCtg", "Sales Ctg", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= max(DmCtg)) "
  'Geeta***Crm***212
  'uni.4 max(DmSalCtg) replaced by qDmSalCtgRepl,qDmSalCtgDescRepl to get either SalCtg,SalCTg2 or SalCTg3
  GRP_REP.Add "Sales Ctg", "||qDmSalCtgRepl||", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and PmCd=max(DmCtg) and PSCd= ||qDmSalCtgDescRepl|| ) "
  
  GRP_REP.Add "Inv Vch No", "IrCoCd+IrTc+IrYy+IrChr+str(IrNo)", "IrCoCd+'/'+IrTc+'/'+IrYy+'/'+IrChr+'/'+LTrim(Str(IrNo))", "", "", "Company Code,Invoice Cust,Inv Date,Inv Exp No,Customer Grp,Customer Typ", "", "", "max(InExpNo) "
  GRP_REP.Add "Inv Exp No", "IrCoCd+max(InExpNo)", "'['+IrCoCd+'] '+max(InExpNo)", "InExpNo", "hInExpNo", "Company Code,Inv Date,Inv Vch No,Customer Grp,Customer Typ", "", "", ""
  ' ***** Manali  3.10.0 - 10/03/12 - Sort Based on InDt/PrnDt
  '6.1
  GRP_REP.Add "Inv/Prn Date", "convert(varchar(10), max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End), 112)", "convert(varchar(8), max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End), " + CStr(DtFmtSlahVal()) + ")", "", "", "Quarter,Month,Week", "", "", ""
  GRP_REP.Add "Week", "LTrim(Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))))+' - '+DateName(WW,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))", "", "", "", "Quarter,Month", "", "", ""
  GRP_REP.Add "Month", "Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End)))+Str(DatePart(mm,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End)))", "LTrim(Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))))+' '+DateName(MM,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))", "", "", "Quarter", "", "", ""
  'MW.134 - group sort Year added.
  GRP_REP.Add "Year", "LTrim(Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))))", "", "", "", "", "", "", ""
  GRP_REP.Add "Quarter", "LTrim(Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))))+' - '+DateName(QQ,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))", "", "", "", "", "", "", ""
''  ' bef 3.10.0
''  GRP_REP.Add "Inv Date", "convert(varchar(10), max(InDt), 112)", "convert(varchar(8), max(InDt), 3)", "", "", "Quarter,Month,Week", "", "", ""
''  GRP_REP.Add "Prn Date", "convert(varchar(10), max(InPrnDt), 112)", "convert(varchar(8), max(InPrnDt), 3)", "", "", "Quarter,Month,Week", "", "", ""
''  GRP_REP.Add "Week", "LTrim(Str(DatePart(yyyy,max(InDt))))+' - '+DateName(WW,max(InDt))", "", "", "", "Quarter,Month", "", "", ""
''  GRP_REP.Add "Month", "Str(DatePart(yyyy,max(InDt)))+Str(DatePart(mm,max(InDt)))", "LTrim(Str(DatePart(yyyy,max(InDt))))+' '+DateName(MM,max(InDt))", "", "", "Quarter", "", "", ""
''  GRP_REP.Add "Quarter", "LTrim(Str(DatePart(yyyy,max(InDt))))+' - '+DateName(QQ,max(InDt))", "", "", "", "", "", "", ""
  GRP_REP.Add "Design Code", "max(IdDmCd)", "", "IdDmCd", "hIdDmCd", "Design Ctg,Sales Ctg", "", "", "max(DmDesc) "
  GRP_REP.Add "Design Size", "max(IdDmSz)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= max(IdDmSz))"
  GRP_REP.Add "Suffix", "max(IdSfx)", "", "", "", "", "", "", ""
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  
  'Sachin - 4.1.0.0
  GRP_REP.Add "ExpOrd Cust", "max(IdExpCmCd)", "", "", "", "Sales Exec", "", "", "case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  max(CE.CmName) else max(IdExpCmCd) end " ' From CustMst Where CmCtg= 'C' and CmCd= max(IdExpCmCd)) "
  GRP_REP.Add "Invoice Cust", "max(InCmCd)", "", "", "", "Customer Grp,Customer Typ", "", "", "(Select case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else  max(InCmCd) end From CustMst Where CmCtg= 'C' and CmCd= max(InCmCd)) "
  '***************
  
  'Sachin - 4.1.0.0
  GRP_REP.Add "Customer Grp", "max(CE.CmGrp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMGRP' and PmCd = Max(CE.CmGrp))"
  GRP_REP.Add "Customer Typ", "max(CE.CmTyp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMTYP' and PmCd = Max(CE.CmTyp))"
  
  GRP_REP.Add "Order", "max(IdExpOdTc)+max(IdExpOdYy)+max(IdExpOdChr)+str(max(IdExpOdNo))", "max(IdExpOdTc)+'/'+max(IdExpOdYy)+'/'+max(IdExpOdChr)+'/'+LTrim(Str(max(IdExpOdNo)))", "", "", "ExpOrd Cust,Sales Exec", "", "", "max(IdExpCmCd) "
  
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
  
  GRP_REP.Add "Rm Ctg", "max(Rm.RmCtg)", "", "RmCtg", "hRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = max(Rm.RmCtg)) "
  GRP_REP.Add "Rm SubCtg", "max(Rm.RmCtg+Rm.RmSCtg)", "'('+ RTrim(max(Rm.RmCtg))+ ') '+ max(Rm.RmSCtg)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = max(Rm.RmCtg) and PSCd= max(Rm.RmSCtg)) "
  GRP_REP.Add "Rm Code", "max(Rm.RmCtg+Rm.RmSCtg+IrRmCd)", "'('+ RTrim(max(Rm.RmCtg))+'/'+ max(Rm.RmSCtg)+ ') '+ max(IrRmCd)", "RmCtg,RmSCtg,IrRmCd", "hRmCtg,hRmSCtg,hIrRmCd", "Rm Ctg,Rm SubCtg,BaseCd + RmCd", "", "", "max(SubString(Rm.RmDesc, 1, 30)) "
  
  ' ***** Manali 3.9.0 - Parent Code
  GRP_REP.Add "Parent Cd", "Max(Rm.RmCtg+Rm.RmSCtg+(case when Rm.RmParentCd= '' then IrRmCd else Rm.RmParentCd end))", "'('+ RTrim(max(Rm.RmCtg))+'/'+ max(Rm.RmSCtg)+ ') '+ max(case when Rm.RmParentCd= '' then IrRmCd else Rm.RmParentCd end)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg,Rm SubCtg,BaseCd + RmCd", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= max(Rm.RmParentCd)) "
  
  '****** Sachin 2.13.0 - 16/06/2006 ******
  Dim wsPtrYN As String
  wsPtrYN = "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= max(Rm.RmCtg) and PSCd= max(Rm.RmSCtg))"
  'GRP_REP.Add "Rm Sv/Len", "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else str(IrRmSz,6,3) end", "", "qRmSz", "hRmSz", "Rm Ptr,(RmCd) Sv/Len,Range Size", "", "", ""
  'GRP_REP.Add "Rm Sv/Len", "case when " + wsPtrYN + "='Y' then   str(max(IrRmSz),7,4) else str(max(IrRmSz),6,3) end", "", "IrRmSz", "hIrRmSz", "Rm Ptr,(RmCd) Sv/Len,Range Size", "", "", ""
  ' ***** Manali 3.03 - 07/07/08 - RmSzDesc
'  GRP_REP.Add "Rm Sv/Sz", "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("max(IrRmSz)", "max(IrRmsz2)", "max(IrRmsz3)") + " else IsNull(max(RisRmSzDesc), " + GetRmSzDBFldNm("max(IrRmSz)", "max(IrRmsz2)", "max(IrRmsz3)") + ") end", "", "wIrRmSz", "hIrRmSz", "Rm Ptr,(RmCd) Sv/Len,Range Size", "", "", ""
  GRP_REP.Add "Rm Sv/Sz", "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("max(IrRmSz)", "max(IrRmsz2)", "max(IrRmsz3)") + " else (case when max(RisRmSzDesc) is null then " + GetRmSzDBFldNm("max(IrRmSz)", "max(IrRmsz2)", "max(IrRmsz3)") + " else max(RisRmSzDesc) end) end", "", "wRmSz", "hIrRmSz", "Rm Ptr,(RmCd) Sv/Len,Range Size", "", "", ""
  
  '****** Sachin 2.13.0 - 16/06/2006 ******
  
  GRP_REP.Add "BaseCd + RmCd", "max(Rm.RmBaseCd+IrRmCd)", "'('+ RTrim(max(Rm.RmBaseCd))+ ') '+ max(IrRmCd)", "IrRmCd", "hIrRmCd", "Rm Ctg,Rm SubCtg,Rm Code", "", "", "max(SubString(Rm.RmDesc, 1, 30)) "
  GRP_REP.Add "Sales Exec", "max(OmCmSalPer)", "", "", "", "", "", "", "(Select PDesc from Param Where PTyp = 'SP' and PMCd= max(OmCmSalPer)) "
  'sv.52 Stone Class group sort added
  GRP_REP.Add "Stone Class", "max(Rm.RmCtg+IsNull(SC.PSCd, ''))", "'('+ RTrim(max(Rm.RmCtg))+') '+max(IsNull(SC.PSCd, ''))", "RmCtg", "hRmCtg", "Rm Ctg", "", "", "max(SC.PDesc)"

  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ,Process,OrderNo,OrderSr,Customer PO,FgSubLoc ******
  If ws_MultiCoMod = True Then _
    GRP_REP.Add "Company Code", "IrCoCd", "IrCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=IrCoCd) "
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******

  If mb_LotMod Then
    GRP_REP.Add "RM Lot No.", "Max(IrLotNo)", "", "", "", "", "", "", "Max(IrLotNo)"
  End If


'  Select Case UCase(adc("UoSaRmGrp"))
'  Case Is = "SR", "RMRT", "RMCD"
'    GRP_REP.Add "Rm Code", "max(RmCtg+RmSCtg+IrRmCd)", "'('+ RTrim(max(RmCtg))+'/'+ max(RmSCtg)+ ') '+ max(IrRmCd)", "RmCtg,RmSCtg,IrRmCd", "hRmCtg,hRmSCtg,hIrRmCd", "Rm Ctg,Rm SubCtg", "", "", "max(RmDesc) "
'    GRP_REP.Add "Rm SubCtg", "max(RmCtg+RmSCtg)", "'('+ RTrim(max(RmCtg))+ ') '+ max(RmSCtg)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = max(RmCtg) and PSCd= max(RmSCtg)) "
'  Case Is = "RMSCTG"
'    GRP_REP.Add "Rm SubCtg", "max(RmCtg+RmSCtg)", "'('+ RTrim(max(RmCtg))+ ') '+ max(RmSCtg)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = max(RmCtg) and PSCd= max(RmSCtg)) "
'  End Select
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  ' **** Manali 3.03 - RisRmSzDesc as qRmSzDesc added
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(45) as DGrp1, space(45) as DGrp2, space(45) as DGrp3,
  'space(45) as G1Desc, space(45) as G2Desc, space(45) as G3Desc,
  'IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr, IrSrNo, InExpNo, InDt, InCmCd, DmCtg,
  'IdDmCd, IdDmSz, IdSfx, IdQty, IdExpOdTc, IdExpOdYy, IdExpOdChr,
  'IdExpOdNo, IdExpOdSr, IdVaCtg, RmCtg, RmSCtg, IrRmCd, RmQw, IrRmSz, RisRmSzDesc as qRmSzDesc,
  'IrRmIRt as qRmRt, IrRmIWt as qRmWt, IrRmQty as qRmQty, IrRmIWt as qGmsWt,
  'IrRmIWt As qPureWt, IrRmIVal As qRmVal
  'From InvRm, InvDsg, InvHd, DsgMst, RmMst, RmIdSz where 1= 2
  '*** Report Sql ***
  
  Dim grpflds As String, wSqlStrg As String, wrepcnd As String, wCnd As String
  Dim wRmDet As String, wRmGrp As String, wCust As String, wOrdJoin As String
  Dim wOpt As String, wIrRmWtOpt As String, wIrRmRtOpt As String, wIrRmValOpt As String
    
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  Dim wRmSzDesc  As String, wRmIdSzJoin As String
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  
  Dim wAddlCnd As String  '****** Sachin 3.02.0
  
  Dim wPureWt As String     ' ***** Manali 3.6.0
  Dim ws_DsgAnaCnd As String
    
  ' Zubin 213
  Dim wb_InBaseCur As Boolean
  wb_InBaseCur = IIF(ADC("UoCurCdFr") = "", True, False)
  
  
  Set Rep = Nothing
  Set ADC.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
   
  '*** To check if 'Design Colour' group is selected ***
  Dim wGrp(6) As String, i As Integer, wDsgColGrpPos As Integer, wRmGrpOpt As String
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDsgColGrpPos = 0
  For i = 1 To 3
    If wGrp(i) = "(DsgCd)Colour" Or wGrp(i) = "Design Colour" Then wDsgColGrpPos = i
    If wGrp(i) = "Rm Ctg" Or wGrp(i) = "Rm SubCtg" Or wGrp(i) = "Rm Code" Or wGrp(i) = "Rm Sv/Len" Then wRmGrpOpt = wGrp(i)
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
    wOpt = "Act"   '''"Actual"
    wIrRmWtOpt = "IrRmAWt"
    wIrRmRtOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmARt/InCnvFct else IrRmARt*InCnvFct end) ", "IrRmARt")
    wIrRmValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal")
  ElseIf UCase(ADC("UoInVal")) = "I" Then
    wOpt = "Inv"    ''''"Invoice"
    wIrRmWtOpt = "IrRmIWt"
    wIrRmRtOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIRt/InCnvFct else IrRmIRt*InCnvFct end) ", "IrRmIRt")
    wIrRmValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIVal/InCnvFct else IrRmIVal*InCnvFct end) ", "IrRmIVal")
  ElseIf UCase(ADC("UoInVal")) = "Z" Then
    wOpt = "Custm"    '''"Customs"
    wIrRmWtOpt = "IrRmZWt"
    wIrRmRtOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZRt/InCnvFct else IrRmZRt*InCnvFct end) ", "IrRmZRt")
    wIrRmValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal*InCnvFct end) ", "IrRmZVal")
    ' ***** Manali 3.6.0 - Customs purity for GPSL
    'KBS.21- formula for pure wt is changed
    wPureWt = "sum((case when Rm.RmCtg in ('D', 'C') then " + wIrRmWtOpt + "/ 5 else " + wIrRmWtOpt + " end) " + _
               "*(case when Rm.RmCtg in ('G', 'P', 'S', 'L', 'M') then Rm.RmPurityZ when Rm.RmCtg= 'X' then B.RmPurityZ else 0 end) )" + _
               "/(case when max(Rm.RmCtg) in ('G', 'P', 'S', 'L', 'X') then (Select (Case When RmPurityZ >0.0 Then RmPurityZ Else 1 End) " + _
               " from RmMst Z where Z.RmZ='Y' and Z.RmCd= max(Rm.RmZCd) " + IIF(gs_Partition = ctCurrPrtn, " and  Z.RmPrtKey='" + ctCurrPrtn + "' ", "") + " )" + _
               " else 1 end) as qPureWt, "
  ElseIf UCase(ADC("UoInVal")) = "S" Then
    wOpt = "Stk"    ''''"Stock"
    wIrRmWtOpt = "IrRmAWt"
    wIrRmRtOpt = "IrStkRt"
    'wIrRmValOpt = "(case when RmQW= 'Q' then IrRmQty else IrRmAWt end)* IrStkRt"
    '****** Sachin 2.13.0 - 05/06/2006 - Stock Value should always be Weight Wise ******
    wIrRmValOpt = "IrRmAWt*IrStkRt"
    '****** Sachin 2.13.0 - 05/06/2006 - Stock Value should always be Weight Wise ******
    
    '4.1.4
    If Trim(ADC("UoInVal1")) <> "" Then
        Dim wAIrRmRtOpt As String, wZIrRmRtOpt As String, wIIrRmRtOpt As String
        Dim wAIrRmValOpt As String, wZIrRmValOpt As String, wIIrRmValOpt As String
        
        wAIrRmRtOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal")
        wIIrRmRtOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIVal/InCnvFct else IrRmIVal* InCnvFct end) ", "IrRmIVal")
        wZIrRmRtOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal* InCnvFct end) ", "IrRmZVal")
        wIrRmRtOpt = " (Case When RSCtg.PValue1='N' Then " + IIF(UCase(ADC("UoInVal1")) = "A", wAIrRmRtOpt, _
                                                             IIF(UCase(ADC("UoInVal1")) = "I", wIIrRmRtOpt, _
                                                             IIF(UCase(ADC("UoInVal1")) = "Z", wZIrRmRtOpt, wIrRmRtOpt))) + " Else " + wIrRmRtOpt + " End)"
        
        wAIrRmValOpt = " (Case When RSCtg.PValue1='N' Then " + _
                        IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal") + " Else IrStkRt*IrRmAWt End) "
                        
        wZIrRmValOpt = " (Case When RSCtg.PValue1='N' Then " + _
                        IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal*InCnvFct end) ", "IrRmZVal") + " Else IrStkRt*IrRmZWt End) "
        
        wIIrRmValOpt = " (Case When RSCtg.PValue1='N' Then " + _
                        IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIVal/InCnvFct else IrRmIVal*InCnvFct end) ", "IrRmIVal") + " Else IrStkRt*IrRmIWt End) "
        
        wIrRmValOpt = IIF(UCase(ADC("UoInVal1")) = "A", wAIrRmValOpt, _
                      IIF(UCase(ADC("UoInVal1")) = "I", wIIrRmValOpt, _
                      IIF(UCase(ADC("UoInVal1")) = "Z", wZIrRmValOpt, wIrRmValOpt)))
    End If
  End If
  
  ' ***** Manali 3.6.0 - Customs purity for GPSL other than Z
  If UCase(ADC("UoInVal")) <> "Z" Then
    wPureWt = "sum((case when Rm.RmCtg in ('D', 'C') then " + wIrRmWtOpt + "/ 5 else " + wIrRmWtOpt + " end) " + _
               "*(case when Rm.RmCtg in ('G', 'P', 'S', 'L', 'M') then Rm.RmPurityWt when Rm.RmCtg= 'X' then Rm.RmBasePurityWt else 0 end) )" + _
               "/(case when max(Rm.RmCtg) in ('G', 'P', 'S', 'L', 'M') then " + _
                            "IsNull((Select RmPurityWt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where " + _
                            "PTyp= 'RMCTG' and PMCd= max(Rm.RmCtg)), 1) " + _
               "when max(Rm.RmCtg)= 'X' then " + _
                            "IsNull((Select RmPurityWt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where " + _
                            "PTyp= 'RMCTG' and PMCd= IsNull((Select RmCtg From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= max(Rm.RmBaseCd)), '')), 1) " + _
              "else 1 end) as qPureWt, "
  End If
  ' ***** Manali 3.6.0 - Prd purity for GPSL other than Z
  
  ADC("UoSalRtFr").CmpStr = wIrRmRtOpt + ">= "
  ADC("UoSalRtTo").CmpStr = wIrRmRtOpt + "<= "
    
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")

  ' Zubin 213
  wCnd = wCnd + IIF(wb_InBaseCur = False, " and CE.CmCurCd= '" + ADC("UoCurCdFr") + "'", "")
  wCnd = wCnd + IIF(ADC("UOYN1") <> "", " AND INCLS = '" + ADC("UOYN1") + "'", "")


  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  wRmSzDesc = "(case when (Select PDesc225 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(Max(RisRmCtg), '') and PSCd= max(Rm.RmSCtg))= 'Y' " + _
              "then IsNull(max(RisRmSzDesc), '') Else '' End )"
  wRmIdSzJoin = "Left Outer Join RmIdSz on RisRmCtg=Rm.RmCtg and RisRmId=Rm.RmId and RisSz1=IrRmSz "
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  
  '****** Sachin 2.13.0 [16-08-2006] - RmQw added below
  ' ****** Manali 3.03 - 08/07/08 - RisRmSzDesc added below
  Select Case UCase(ADC("UoSaRmGrp"))
  Case Is = "SR"
    'wRmDet = " max(RmCtg) as RmCtg, max(RmSCtg) as RmSCtg, max(IrRmCd) as IrRmCd, max(RmQw) As RmQw, max(IrRmSz) as IrRmSz, max(" + wIrRmRtOpt + ") as qRmRt, "
    wRmDet = " max(Rm.RmCtg) as RmCtg, max(Rm.RmSCtg) as RmSCtg, max(IrRmCd) as IrRmCd, max(Rm.RmQw) As RmQw, max(IrRmSz) as IrRmSz, max(IrRmSz2) as IrRmSz2, max(IrRmSz3) as IrRmSz3, " + wRmSzDesc + " as qRmSzDesc, max(" + wIrRmRtOpt + ") as qRmRt, "
    wRmGrp = ", IrSrNo "
    Rep.hIrSrNo.Suppress = False: Rep.FldIrSrNo.Suppress = False
    Rep.hIrRmRt.SetText wOpt + Chr(13) + "Rate"
  Case Is = "RMRT"
    wRmDet = " max(Rm.RmCtg) as RmCtg, max(Rm.RmSCtg) as RmSCtg, IrRmCd, max(Rm.RmQw) As RmQw, IrRmSz, IrRmSz2, IrRmSz3, " + wRmSzDesc + " as qRmSzDesc, " + wIrRmRtOpt + " as qRmRt, "
    wRmGrp = ", IrRmCd, IrRmSz, IrRmSz2, IrRmSz3, " + wIrRmRtOpt + " "
    Rep.hIrSrNo.Suppress = True: Rep.FldIrSrNo.Suppress = True
    Rep.hIrRmRt.SetText wOpt + Chr(13) + "Rate"
  Case Is = "RMCD"
    wRmDet = " max(Rm.RmCtg) as RmCtg, max(Rm.RmSCtg) as RmSCtg, IrRmCd, max(Rm.RmQw) As RmQw, IrRmSz, IrRmSz2, IrRmSz3, " + wRmSzDesc + " as qRmSzDesc, 0.0 as qRmWt, "
    wRmGrp = ", IrRmCd, IrRmSz, IrRmSz2, IrRmSz3 "
    Rep.hIrSrNo.Suppress = True: Rep.FldIrSrNo.Suppress = True
    Rep.hIrRmRt.SetText " "
  Case Is = "RMSCTG"
    If wRmGrpOpt = "Rm Code" Or wRmGrpOpt = "Rm Sv/Len" Then _
      DispMsg "Group Option Has To Be A Higher Level Than Detail", etError: wCnd = wCnd + IIF(wCnd <> "", " and 1= 2", " where 1= 2")
    wRmDet = " Rm.RmCtg, Rm.RmSCtg, '' as IrRmCd, '' as RmQw, 0.0 as IrRmSz, 0.0 as IrRmSz2, 0.0 as IrRmSz3, '' as qRmSzDesc, 0.0 as qRmWt, "
    wRmGrp = ", Rm.RmCtg, Rm.RmSCtg "
    Rep.hIrSrNo.Suppress = True: Rep.FldIrSrNo.Suppress = True
    Rep.hIrRmRt.SetText " ": Rep.hIrRmCd.SetText " "
  Case Is = "RMCTG"
    If wRmGrpOpt = "Rm SubCtg" Or wRmGrpOpt = "Rm Code" Or wRmGrpOpt = "Rm Sv/Len" Then _
      DispMsg "Group Option Has To Be A Higher Level Than Detail", etError: wCnd = wCnd + IIF(wCnd <> "", " and 1= 2", " where 1= 2")
    wRmDet = " Rm.RmCtg, '' as RmSCtg, '' as IrRmCd, '' As RmQw, 0.0 as IrRmSz, 0.0 as IrRmSz2, 0.0 as IrRmSz3, '' as qRmSzDesc, 0.0 as qRmRt, "
    wRmGrp = ", Rm.RmCtg "
    Rep.hIrSrNo.Suppress = True: Rep.FldIrSrNo.Suppress = True
    Rep.hIrRmRt.SetText " ": Rep.hIrRmCd.SetText " ": Rep.hRmSCtg.SetText " "
  End Select
    
     ws_DsgAnaCnd = " And (Select Count(*) From UsrOptDsgAna " + _
                 "Join DsgAna on " + IIF(gs_Partition = ctCurrPrtn, " DaPrtKey='" + ctCurrPrtn + "' and ", "") + " DaAnaSr = UdAnaSr " + _
                 "where (DaAnaCd between UdAnaCdFr and UdAnaCdTo) " + _
                 "and DaTcTyp = DmTcTyp and DaCd= DmCd and DaSz= DmSz and " + _
                 "UdUsrCd = '" + ADC("UoUsrCd") + "' and " + _
                 "UdMnuCd= '" + ADC("UoMnuCd") + "')= " & (GrdDsgAna.Rows - 1) & " "
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(ADC("UORMCTGFR")) <> "", " IrRmCtg >= '" + Trim(ADC("UORMCTGFR")) + "'", "")
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMCTGTO")) <> "", " And ", "") + IIF(Trim(ADC("UORMCTGTO")) <> "", " IrRmCtg <= '" + Trim(ADC("UORMCTGTO")) + "'", "")
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMCTGSEL")) <> "", " And ", "") + IIF(Trim(ADC("UORMCTGSEL")) <> "", " IrRmCtg In (" + Trim(ADC("UORMCTGSEL")) + ")", "")
      
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGFR")) <> "", " IrRmSCtg >= '" + Trim(ADC("UORMSCTGFR")) + "'", "")
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGTO")) <> "", " IrRmSCtg <= '" + Trim(ADC("UORMSCTGTO")) + "'", "")
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGSEL")) <> "", " IrRmSCtg In (" + Trim(ADC("UORMSCTGSEL")) + ")", "")
      
      ' ****** Manali 3.9.0 - Option for InDt / InPrnDt added
      If ADC("UoYN3") = "Y" Then
        wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(ADC("UODTFR")) <> "01/01/80", " IrDt >= '" + Trim(ADC("UODTFR")) + "'", "")
        wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(ADC("UODTTO")) <> "01/01/80", " IrDt <= '" + Trim(ADC("UODTTO")) + "'", "")
      End If
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
    
 
             '"/(case when max(Rm.RmCtg) in ('G', 'P', 'S') then " + _
             '"IsNull((Select RmPurityWt from Param join RmMst on RmCd= PValue where " + _
             '"PTyp= 'RMCTG' and PMCd= max(Rm.RmCtg)), 1) else 1 end) as qPureWt, "
    
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
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    ' ***** Manali 3.03 - 07/07/08 - Join with RmIdSz added
    ' ***** Manali  3.10.0 - 10/03/12 - Sort Based on InDt/PrnDt
    ' Sachin 4.1.0.0
    'uni.4 wDmSalCtgCnd added to DsgMst join, to get SalCtg value from scope.
    'To avoid SalCtg scope condition in where clause
    'sv.52 Left outer join Param SC on SC.PTyp= 'STNCLS' and SC.PMCD=Rm.RmCtg and SC.PSCd=Rm.RmStnCls added
    wSqlStrg = "Select " + grpflds + ", " + _
             "IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr, max(IrSrNo) as IrSrNo, max(InExpNo) as InExpNo, max(InDt) as InDt, " + _
             "max(" + wCust + ") as InCmCd, max(DmCtg) as DmCtg, max(IdDmCd) as IdDmCd, max(IdDmSz) as IdDmSz, max(IdSfx) as IdSfx, max(IdQty) as IdQty, " + _
             "max(IdExpOdTc) as IdExpOdTc, max(IdExpOdYy) as IdExpOdYy, max(IdExpOdChr) as IdExpOdChr, max(IdExpOdNo) as IdExpOdNo, " + _
             "max(IdExpOdSr) as IdExpOdSr, max(IdVaCtg) as IdVaCtg, " + wRmDet + _
             "sum(" + wIrRmWtOpt + ") as qRmWt, Sum(IrRmQty) as qRmQty, " + _
             "sum(case when Rm.RmCtg in ('D', 'C') then " + wIrRmWtOpt + "/ 5 else " + wIrRmWtOpt + " end) as qGmsWt, " + _
             wPureWt + "sum(" + wIrRmValOpt + ") as qRmVal " + _
             " From InvRm join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy " + _
             " and IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + IIF(gs_Partition = ctCurrPrtn, " and IdPrtKey=IrPrtKey ", "") + _
             " Join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy " + _
             " and InChr= IrChr and InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IrPrtKey ", "") + _
             " Join OrdDsg on OdCoCd= IdCoCd and OdTc=IdExpOdTc and OdYy=IdExpOdYy and OdChr =IdExpOdChr and OdNo=IdExpOdNo and OdSr = IdExpOdSr and OdIdNo=IdOdIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=IdPrtKey ", "") + _
             " Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=IdPrtKey ", "") + _
             " Join CustMst CE on CE.CmCtg= 'C' and CE.CmCd = IdExpCmCd " + _
             " Join CustMst C on C.CmCtg= 'C' and C.CmCd =InCmCd " + _
             " Join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=IdPrtKey ", "") + wDmSalCtgCnd + wOrdJoin + _
             " Join (Select '" + Trim(wInvDt) + "' as qDtOpt) A on 1=1 " + _
             " Join RmMst Rm on RmCd= IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=IrPrtKey ", "") + _
             " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg " + _
             " Left Outer Join RmMst B On B.RmCd=Rm.RmBaseCd  and B.RmPrtKey  =Rm.RmPrtKey " + wRmIdSzJoin + _
             " Left outer join Param SC on SC.PTyp= 'STNCLS' and SC.PMCD=Rm.RmCtg and SC.PSCd=Rm.RmStnCls " + _
             wCnd + ws_DsgAnaCnd + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey='" + ctCurrPrtn + "' ", "") + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + " group by IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr " + wRmGrp
    
  'uni.4 replacing DmSalCtg value and its description in group sort of Sales Ctg
  wSqlStrg = Replace(wSqlStrg, "||qDmSalCtgRepl||", ws_DmSalCtg)
  wSqlStrg = Replace(wSqlStrg, "||qDmSalCtgDescRepl||", ws_DmSalCtg)
  
  ADC.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  ' ***** Manali 3.6.0
  ' **** Manali 3.9.0 - Change Reverted
'  Rep.LblValAs.SetText moCn.GetFldVal("Select PDesc from Param Where PTyp='INVAL' and PMCd='" + adc("UoInVal") + "'")
  ' Bef 3.6.0
  
'  Rep.h1RmWt.SetText wOpt + Chr(13) + "Rm Wt": Rep.h2RmWt.SetText wOpt + Chr(13) + "Rm Wt"
'  Rep.h1RmQty.SetText wOpt + Chr(13) + "Rm Qty": Rep.h2RmQty.SetText wOpt + Chr(13) + "Rm Qty"
'  Rep.h1WtInGms.SetText wOpt + Chr(13) + "Wt In Gms": Rep.h2WtInGms.SetText wOpt + Chr(13) + "Wt In Gms"
'  Rep.h1PureWt.SetText wOpt + Chr(13) + "Pure Wt": Rep.h2PureWt.SetText wOpt + Chr(13) + "Pure Wt"
'  Rep.h1Val.SetText wOpt + Chr(13) + "Value": Rep.h2Val.SetText wOpt + Chr(13) + "Value"
  
  Rep.h1RmWt.SetText wOpt + " Rm Wt": Rep.h2RmWt.SetText wOpt + " Rm Wt"
  Rep.h1RmQty.SetText wOpt + " Rm Qty": Rep.h2RmQty.SetText wOpt + " Rm Qty"
  Rep.h1WtInGms.SetText wOpt + " Wt In Gms": Rep.h2WtInGms.SetText wOpt + " Wt In Gms"
  Rep.h1PureWt.SetText wOpt + " Pure Wt": Rep.h2PureWt.SetText wOpt + " Pure Wt"
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
  
  'KC.3- StkRt-when Show Stock Rate=YES then Stock Rate is visible in report
  'if it is NO then Stock Rate is invisible in report
 
  If UCase(ADC("UOINVAL")) = "S" Then
  If ADC("UoYN4") = "Y" Then
    'Rep.hIrRmRt.SetText "Stock Rt"
    Rep.FldIrStkRt.Suppress = False
    Rep.FldqRmVal.Suppress = False
    Rep.FldG3qRmVal.Suppress = False
    Rep.FldG2qRmVal.Suppress = False
    Rep.FldG1qRmVal.Suppress = False
    Rep.FldGqRmVal.Suppress = False
  Else
    Rep.hIrRmRt.SetText ""
    Rep.FldIrStkRt.Suppress = True
    Rep.h1Val.SetText ""
    Rep.h2Val.SetText ""
    Rep.FldqRmVal.Suppress = True
    Rep.FldG3qRmVal.Suppress = True
    Rep.FldG2qRmVal.Suppress = True
    Rep.FldG1qRmVal.Suppress = True
    Rep.FldGqRmVal.Suppress = True
  End If
  End If
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
  'StkRt-To select this option(Show Stock rate) either Seo user can give their password or
  ' can enter SEOPWD password otherwise this option can't not select
  With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0
    If .Value = Checked Then
       If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
            "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + ADC.MenuCd + "' and " + _
            "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
            "Isnull ((Select PValue1 from Param where " + _
            "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ), ''))") Then _
            .Value = Unchecked
    End If
    If .Value = Checked Then ADC("UoYN4") = "Y" Else ADC("UoYN4") = "N"
  'uni.4  checking Include SalCtg2,SalCtg3 option
  Case Is = 1
    If .Value = Checked Then ADC("wAddAllSalCtg") = "Y" Else ADC("wAddAllSalCtg") = "N"
    
  End Select
  End With
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
  Set moCn = ADC.Connection  '*** (Bef speed) bef 24/11/06
  Set ADC.FirNKeyCtl = ADC("UoInTcFr") 'adc("UoCoCdFr")
  Call InitProp(Me)
  
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
  
  If ws_MultiCoMod = True Then
      ADC("UOCOCDFR").Enabled = True: ADC("UOCOCDTO").Visible = True: ADC("UOCOCDSEL").Visible = True
  Else
      ADC("UOCOCDFR").Enabled = False: ADC("UOCOCDTO").Visible = False: ADC("UOCOCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gb_RmCtgFor = False
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  
  If ws_MultiCoMod = False Then
    mb_LotMod = IIF(moCn.GetFldVal("Select HLotNoYn From Head Where HCd='" + ctSelfCmCd + "' And " + _
                              "HCoCd ='" + gs_CoCd + "' ") = "Y", True, False)
  Else
    mb_LotMod = IIF(moCn.GetFldVal("Select Count('X') From Head Where HCd='" + ctSelfCmCd + "' And " + _
                              "HCoCd <> '" + ctSelfCoCd + "' and HLotNoYn='Y' ") > 0, True, False)
  End If
  
  If mb_LotMod = False Then
    LblLotNo.Visible = False
    ADC("UoLotNoFr").Enabled = False: ADC("UoLotNoFr").Visible = False: ADC("UoLotNoFr") = ""
    ADC("UoLotNoTo").Enabled = False: ADC("UoLotNoTo").Visible = False: ADC("UoLotNoTo") = ""
    ADC("UoLotNoSel").Enabled = False: ADC("UoLotNoSel").Visible = False: ADC("UoLotNoSel") = ""
  End If
    
  Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gb_CoCdFor = False  '****** Sachin 2.14.0 - [Multi-Company]
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
    gb_CoCdFor = False  '****** Sachin 2.14.0 - [Multi-Company]
  ' Zubin 213
  Select Case UCase(IdName)
  Case Is = UCase("UoCurCdFr")
    If UCase(ADC("UoInVal")) = "S" Then Cancel = True: Exit Sub
  Case Is = UCase("UoLotNoFr"), UCase("UoLotNoFr"), UCase("UoLotNoSel")
    gs_RmLocTyp = "XS"
  End Select
  ' Zubin 213
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

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptCust_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN") = "Y"
Case Is = 1
  ADC("UoYN") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

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
  
    ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  If ADC("UoYN") = "Y" Then OptCust(0).Value = True
  If ADC("UoYN") = "N" Then OptCust(1).Value = True
  ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  ' ****** Manali 3.9.0 - Option for InDt / InPrnDt added
  If ADC("UoYN3") = "Y" Then OptInvDt(0).Value = True
  If ADC("UoYN3") = "N" Then OptInvDt(1).Value = True
  ' ****** Manali 3.9.0 - Option for InDt / InPrnDt added
  
  'StkRt-Set Show StockRate option is No
  ADC("UoYN4") = "N"
  'uni.4 default value for Include SalCtg2,SalCtg3 is Yes
  ChkBoxArr(1).Value = Checked
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
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ")
  If Trim(TxtPwd.text) <> "" And Trim(TxtPwd.text) <> ws_VldPwd Then
      Cancel = True:  DispMsg "Invalid Password / User Does Not Have Seo Rights ", etError
  End If

End Sub

Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
