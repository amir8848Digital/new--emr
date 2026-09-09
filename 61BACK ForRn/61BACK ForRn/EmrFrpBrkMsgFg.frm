VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpBrkMsgFg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Breakage Missing at FG Level"
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
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   60
      TabIndex        =   74
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
      TabIndex        =   73
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
      TabIndex        =   72
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
      TabIndex        =   82
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   83
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
         TabIndex        =   84
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
      TabIndex        =   81
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   78
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
      TabIndex        =   77
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   76
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpBrkMsgFg.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpBrkMsgFg.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   80
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   75
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
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   14985
            Begin VB.Frame FraOptBrkMsg 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2115
               TabIndex        =   130
               ToolTipText     =   "Select If To Show Report For Broken Or Missing Or Both"
               Top             =   820
               Width           =   4035
               Begin VB.OptionButton OptBrkMsg 
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
                  TabIndex        =   5
                  ToolTipText     =   "Select If To Show Report For Broken Or Missing Or Both"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptBrkMsg 
                  Caption         =   "Missing"
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
                  TabIndex        =   4
                  ToolTipText     =   "Select If To Show Report For Broken Or Missing Or Both"
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton OptBrkMsg 
                  Caption         =   "Broken"
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
                  TabIndex        =   3
                  ToolTipText     =   "Select If To Show Report For Broken Or Missing Or Both"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.Frame FraPwd 
               Height          =   870
               Left            =   10440
               TabIndex        =   127
               Top             =   7200
               Width           =   4380
               Begin VB.TextBox TxtPwd 
                  Height          =   285
                  IMEMode         =   3  'DISABLE
                  Left            =   1920
                  MaxLength       =   30
                  PasswordChar    =   "*"
                  TabIndex        =   67
                  Top             =   480
                  Width           =   2340
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pwd To View Val"
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
                  Index           =   51
                  Left            =   120
                  TabIndex        =   129
                  Top             =   480
                  Width           =   1755
               End
               Begin VB.Label LblMain 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "  Password"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   0
                  Left            =   -30
                  TabIndex        =   128
                  Top             =   90
                  Width           =   4365
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Value  "
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
               Left            =   60
               TabIndex        =   64
               Tag             =   "ShowVal"
               ToolTipText     =   "Check To Show Value"
               Top             =   6915
               Width           =   2295
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Box Details"
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
               Left            =   60
               TabIndex        =   70
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show box Details"
               Top             =   7575
               Width           =   2355
            End
            Begin VB.Frame FraOptCust 
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   2115
               TabIndex        =   112
               ToolTipText     =   "Select eihter Invoice Customer Or Order Customer"
               Top             =   7230
               Width           =   4280
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
                  TabIndex        =   69
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
                  TabIndex        =   68
                  Top             =   0
                  Width           =   1545
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5580
               TabIndex        =   45
               ToolTipText     =   "Enter To Invoice Year"
               Top             =   4470
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOFGYYTO"
               IdName          =   "UOFGYYTO"
               CmpStr          =   "FdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2880
               TabIndex        =   42
               ToolTipText     =   "Enter From Invoice Year"
               Top             =   4470
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOFGYYFR"
               IdName          =   "UOFGYYFR"
               CmpStr          =   "FdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   2115
               TabIndex        =   41
               ToolTipText     =   "Enter From Invoice Tc"
               Top             =   4470
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFGTCFR"
               IdName          =   "UOFGTCFR"
               CmpStr          =   "FdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3300
               TabIndex        =   43
               ToolTipText     =   "Enter From Invoice Character"
               Top             =   4470
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
               Index           =   46
               Left            =   4800
               TabIndex        =   44
               ToolTipText     =   "Enter To Invoice Tc"
               Top             =   4470
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFGTCTO"
               IdName          =   "UOFGTCTO"
               CmpStr          =   "FdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   6000
               TabIndex        =   46
               ToolTipText     =   "Enter To Invoice Character"
               Top             =   4470
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
               Index           =   26
               Left            =   2115
               TabIndex        =   47
               ToolTipText     =   "Enter From Invoice Number"
               Top             =   4780
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
               Index           =   29
               Left            =   4800
               TabIndex        =   48
               ToolTipText     =   "Enter To Invoice Number"
               Top             =   4780
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
               Index           =   30
               Left            =   2115
               TabIndex        =   32
               ToolTipText     =   "Enter From Design Category"
               Top             =   3540
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
               Index           =   40
               Left            =   4800
               TabIndex        =   7
               ToolTipText     =   "Enter To Customer"
               Top             =   1140
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
               Left            =   7500
               TabIndex        =   8
               ToolTipText     =   "Enter Customer Selection"
               Top             =   1140
               Width           =   7290
               _ExtentX        =   12859
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
               TabIndex        =   6
               ToolTipText     =   "Enter From Customer"
               Top             =   1140
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
               Index           =   11
               Left            =   4800
               TabIndex        =   39
               ToolTipText     =   "Enter To Design Sales Category"
               Top             =   4160
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMPRDCTGTO"
               IdName          =   "UODMPRDCTGTO"
               CmpStr          =   "DmSalCtg<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   7500
               TabIndex        =   40
               ToolTipText     =   "Enter Design Sales Category Selection"
               Top             =   4160
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMPRDCTGSEL"
               IdName          =   "UODMPRDCTGSEL"
               CmpStr          =   "DmSalCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2115
               TabIndex        =   38
               ToolTipText     =   "Enter From Design Sales Category"
               Top             =   4160
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMPRDCTGFR"
               IdName          =   "UODMPRDCTGFR"
               CmpStr          =   "DmSalCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5685
               TabIndex        =   18
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1755
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
               Left            =   3015
               TabIndex        =   16
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1755
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
               Left            =   2115
               TabIndex        =   15
               ToolTipText     =   "Enter From Order Number"
               Top             =   1755
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
               Top             =   1755
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
               Left            =   6630
               TabIndex        =   71
               ToolTipText     =   "Show Invoice Or Order Customer ('Y' for Invoice / 'N' For Order)"
               Top             =   7260
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
               Index           =   25
               Left            =   4800
               TabIndex        =   28
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   2925
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   7500
               TabIndex        =   29
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   2925
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   4800
               TabIndex        =   25
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   2625
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGTO"
               IdName          =   "UORMSCTGTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   7500
               TabIndex        =   26
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   2625
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   2115
               TabIndex        =   21
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   2340
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   2115
               TabIndex        =   24
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   2625
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
               Index           =   49
               Left            =   2115
               TabIndex        =   27
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   2925
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   4800
               TabIndex        =   22
               ToolTipText     =   "Enter To Raw Material Category"
               Top             =   2340
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   7500
               TabIndex        =   23
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   2340
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCTGSEL"
               IdName          =   "UORMCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   4800
               TabIndex        =   50
               ToolTipText     =   "Enter To Invoice Voucher Date"
               Top             =   5085
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODOCDTTO"
               IdName          =   "UODOCDTTO"
               CmpStr          =   "FdDt<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   2115
               TabIndex        =   49
               ToolTipText     =   "Enter From Invoice Voucher Date"
               Top             =   5085
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODOCDTFR"
               IdName          =   "UODOCDTFR"
               CmpStr          =   "FdDt>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   2115
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   480
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
               Index           =   1
               Left            =   5565
               TabIndex        =   13
               ToolTipText     =   "Enter To Order Year"
               Top             =   1455
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
               Top             =   1455
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
               Index           =   17
               Left            =   5985
               TabIndex        =   14
               ToolTipText     =   "Enter To Order Character"
               Top             =   1455
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
               Index           =   5
               Left            =   2895
               TabIndex        =   10
               ToolTipText     =   "Enter From Order Year"
               Top             =   1455
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
               Index           =   6
               Left            =   2115
               TabIndex        =   9
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1455
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
               Index           =   10
               Left            =   3315
               TabIndex        =   11
               ToolTipText     =   "Enter From Order Character"
               Top             =   1455
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
               Index           =   60
               Left            =   4800
               TabIndex        =   31
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   3225
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1TO"
               IdName          =   "UOLN1TO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   2115
               TabIndex        =   30
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   3225
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1FR"
               IdName          =   "UOLN1FR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4800
               TabIndex        =   33
               ToolTipText     =   "Enter To Design Category"
               Top             =   3540
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
               Left            =   7500
               TabIndex        =   34
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   3540
               Width           =   7290
               _ExtentX        =   12859
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
               Left            =   4800
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   480
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
               Left            =   7500
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   510
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "IrCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   4800
               TabIndex        =   20
               ToolTipText     =   "Enter To Invoice Voucher Date"
               Top             =   2055
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODT1TO"
               IdName          =   "UODT1TO"
               CmpStr          =   "OdOmDt<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   2115
               TabIndex        =   19
               ToolTipText     =   "Enter From Invoice Voucher Date"
               Top             =   2055
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODT1FR"
               IdName          =   "UODT1FR"
               CmpStr          =   "OdOmDt>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   5610
               TabIndex        =   55
               ToolTipText     =   "Enter To Invoice Year"
               Top             =   5400
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYTO"
               IdName          =   "UOINYYTO"
               CmpStr          =   "InYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   2895
               TabIndex        =   52
               ToolTipText     =   "Enter From Invoice Year"
               Top             =   5400
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "InYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   2115
               TabIndex        =   51
               ToolTipText     =   "Enter From Invoice Tc"
               Top             =   5400
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "InTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   3315
               TabIndex        =   53
               ToolTipText     =   "Enter From Invoice Character"
               Top             =   5400
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "InChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   4800
               TabIndex        =   54
               ToolTipText     =   "Enter To Invoice Tc"
               Top             =   5400
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCTO"
               IdName          =   "UOINTCTO"
               CmpStr          =   "InTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   6060
               TabIndex        =   56
               ToolTipText     =   "Enter To Invoice Character"
               Top             =   5400
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRTO"
               IdName          =   "UOINCHRTO"
               CmpStr          =   "InChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   2115
               TabIndex        =   57
               ToolTipText     =   "Enter From Invoice Number"
               Top             =   5685
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "InNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   4800
               TabIndex        =   58
               ToolTipText     =   "Enter To Invoice Number"
               Top             =   5685
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOTO"
               IdName          =   "UOINNOTO"
               CmpStr          =   "InNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   4800
               TabIndex        =   60
               ToolTipText     =   "Enter To Invoice Voucher Date"
               Top             =   5985
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "InDt<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   2115
               TabIndex        =   59
               ToolTipText     =   "Enter From Invoice Voucher Date"
               Top             =   5985
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "InDt>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   2880
               TabIndex        =   121
               ToolTipText     =   "Show Box Detail (Yes / No)"
               Top             =   7560
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
               Index           =   8
               Left            =   2115
               TabIndex        =   61
               ToolTipText     =   "Enter Level Of Detail To Be Shown"
               Top             =   6285
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_BTN1 CmdPwd 
               Height          =   405
               Left            =   10440
               TabIndex        =   123
               ToolTipText     =   "Enter Password to Edit Design"
               Top             =   6840
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   714
               BackColor       =   -2147483633
               MaskColor       =   8454143
               Caption         =   "P&assword"
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
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3511
               Left            =   4800
               TabIndex        =   63
               ToolTipText     =   "Enter To Stock Rate"
               Top             =   6600
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOSTKRTTO"
               IdName          =   "UOSTKRTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3611
               Left            =   2115
               TabIndex        =   62
               ToolTipText     =   "Enter From Stock Rate"
               Top             =   6600
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOSTKRTFR"
               IdName          =   "UOSTKRTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   2985
               TabIndex        =   124
               ToolTipText     =   "Show Value(Y/N)"
               Top             =   6885
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               ReCalcParent    =   "UOPMCD,UOPMCD"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   4800
               TabIndex        =   65
               ToolTipText     =   "Enter Value Type"
               Top             =   6885
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCD"
               ReCalcOn        =   "UOYN2"
               IdName          =   "UOPMCD"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   6675
               TabIndex        =   131
               ToolTipText     =   "Enter 'B' For Broken, 'M' For Missing Or Blank For Both"
               Top             =   840
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOBRKMSG"
               IdName          =   "UOBRKMSG"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   4800
               TabIndex        =   36
               ToolTipText     =   "Enter To Design Code"
               Top             =   3840
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
               Index           =   58
               Left            =   7500
               TabIndex        =   37
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   3840
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "BOdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   59
               Left            =   2115
               TabIndex        =   35
               ToolTipText     =   "Enter From Design Code"
               Top             =   3840
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
               Index           =   84
               Left            =   8340
               TabIndex        =   66
               ToolTipText     =   "Specify If The Other Value To Be Shown Should Be Stock Value, Cost Value Or Sale Value Or No Value"
               Top             =   6885
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Other Value Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   6480
               TabIndex        =   134
               ToolTipText     =   "Location"
               Top             =   6885
               Width           =   1845
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
               Index           =   35
               Left            =   60
               TabIndex        =   133
               Top             =   3840
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Brk/Msg/Both"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   132
               ToolTipText     =   "Location"
               Top             =   840
               Width           =   1605
            End
            Begin VB.Label LblStkRt 
               BackStyle       =   0  'Transparent
               Caption         =   "Stock Rate"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Top             =   6600
               Width           =   1545
            End
            Begin VB.Label LblValTyp 
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
               Left            =   3480
               TabIndex        =   125
               ToolTipText     =   "Location"
               Top             =   6915
               Width           =   1245
            End
            Begin VB.Label LblRmGrp 
               BackStyle       =   0  'Transparent
               Caption         =   "Group Srt"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Top             =   6300
               Width           =   1845
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
               Index           =   34
               Left            =   60
               TabIndex        =   120
               Top             =   5400
               Width           =   1725
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
               Index           =   33
               Left            =   60
               TabIndex        =   119
               Top             =   5685
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
               Index           =   32
               Left            =   6000
               TabIndex        =   118
               Top             =   5400
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
               Index           =   28
               Left            =   3255
               TabIndex        =   117
               Top             =   5400
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
               Index           =   16
               Left            =   5535
               TabIndex        =   116
               Top             =   5400
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
               Left            =   2820
               TabIndex        =   115
               Top             =   5400
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Date"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   114
               Top             =   5985
               Width           =   1965
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order Date"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   113
               Top             =   2070
               Width           =   1965
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
               TabIndex        =   111
               Top             =   3240
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
               Index           =   23
               Left            =   60
               TabIndex        =   110
               Top             =   480
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
               Left            =   7500
               TabIndex        =   109
               Top             =   90
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Fg Date"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   108
               Top             =   5055
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
               TabIndex        =   107
               Top             =   2940
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
               TabIndex        =   106
               Top             =   2625
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
               TabIndex        =   105
               ToolTipText     =   "Location"
               Top             =   2340
               Width           =   1425
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
               Left            =   60
               TabIndex        =   104
               Top             =   7230
               Width           =   1515
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
               Left            =   60
               TabIndex        =   103
               Top             =   1455
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
               Left            =   60
               TabIndex        =   102
               Top             =   1770
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
               Index           =   8
               Left            =   5925
               TabIndex        =   101
               Top             =   1500
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
               Left            =   3255
               TabIndex        =   100
               Top             =   1500
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
               TabIndex        =   99
               Top             =   1500
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
               Left            =   2835
               TabIndex        =   98
               Top             =   1500
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
               Left            =   5625
               TabIndex        =   97
               Top             =   1785
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
               Left            =   2955
               TabIndex        =   96
               Top             =   1785
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
               TabIndex        =   95
               Top             =   1140
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
               TabIndex        =   94
               Top             =   3540
               Width           =   1425
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
               Index           =   6
               Left            =   60
               TabIndex        =   93
               Top             =   4160
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
               Left            =   2115
               TabIndex        =   92
               Top             =   90
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
               TabIndex        =   91
               Top             =   90
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
               Left            =   2820
               TabIndex        =   90
               Top             =   4485
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
               Left            =   5520
               TabIndex        =   89
               Top             =   4485
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
               Left            =   3240
               TabIndex        =   88
               Top             =   4485
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
               Left            =   5940
               TabIndex        =   87
               Top             =   4485
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Fg No"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   86
               Top             =   4775
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Fg Tc/Yy/Chr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   85
               Top             =   4470
               Width           =   1725
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpBrkMsgFg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepBrkMsgFg
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim mb_LotMod As Boolean  ' ***** Manali 350Nxt
Dim ms_PwdFlag As Boolean   '4.1.0.0
Private Sub SetGroupSort()

  GRP_REP.Add "Design Ctg", "DmCtg", "", "DmCtg", "", "", "", "", "DmCtg"
  GRP_REP.Add "Design Code", "DmCd", "", "DmCd", "", "", "", "", "DmCd"
  GRP_REP.Add "Ord KT", "OdKt", "", "OdKt", "", "", "", "", "OdKt"
  GRP_REP.Add "Sales Ctg", "DmSalCtg", "", "", "", "", "", "", "DmSalCtg "
  GRP_REP.Add "Prd Ctg", "DmPrdCtg", "", "", "", "", "", "", "DmPrdCtg "
  GRP_REP.Add "Inv No", "OdCoCd+IfTc+IfYy+IfChr+str(IfNo)", "OdCoCd+'/'+IfTc+'/'+IfYy+'/'+IfChr+'/'+LTrim(Str(IfNo))", "", "", "", "", "", "OdCoCd+'/'+IfTc+'/'+IfYy+'/'+IfChr+'/'+LTrim(Str(IfNo))"
  GRP_REP.Add "Order No", "OdCoCd+OdTc+OdYy+OdChr+str(OdNo)", "OdCoCd+'/'+OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))", "", "", "", "", "", "OdCoCd+'/'+OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))"
  GRP_REP.Add "Customer", "OdOmCmCd", "", "", "", "", "", "", " IsNull((Select CmName From CustMst Where CmCtg= 'C' and CmCd= OdOmCmCd),'-') "
  GRP_REP.Add "Inv Exp No", "OdCoCd+InExpNo", "'['+OdCoCd+'] '+InExpNo", "InExpNo", "hInExpNo", "Company Code,Inv Date,Inv Vch No", "", "", ""
  GRP_REP.Add "Rm Ctg", "RmCtg", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = RmCtg) "
  GRP_REP.Add "Rm SubCtg", "RmCtg+RmSCtg", "'('+ RTrim(RmCtg)+ ') '+ RmSCtg", "", "", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd=RmSCtg) "
  GRP_REP.Add "Rm Code", "RmCtg+RmSCtg+RmCd", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ RmCd", "", "", "Rm Ctg,Rm SubCtg,Rm Code", "", "", "SubString(RmDesc, 1, 30) "
  
'  GRP_REP.Add "Rm Ctg", "max(RmCtg)", "", "RmCtg", "hRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = max(RmCtg)) "
'  GRP_REP.Add "Rm SubCtg", "max(RmCtg+RmSCtg)", "'('+ RTrim(max(RmCtg))+ ') '+ max(RmSCtg)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = max(RmCtg) and PSCd= max(RmSCtg)) "
'  GRP_REP.Add "Rm Code", "max(RmCtg+RmSCtg+IrRmCd)", "'('+ RTrim(max(RmCtg))+'/'+ max(RmSCtg)+ ') '+ max(IrRmCd)", "RmCtg,RmSCtg,IrRmCd", "hRmCtg,hRmSCtg,hIrRmCd", "Rm Ctg,Rm SubCtg,BaseCd + RmCd", "", "", "max(SubString(RmDesc, 1, 30)) "
'
  If ws_MultiCoMod = True Then _
    GRP_REP.Add "Company Code", "OdCoCd", "OdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=OdCoCd) "
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******


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
  'Select  space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(45) as DGrp1, space(45) as DGrp2, space(45) as DGrp3,
  'space(45) as G1Desc, space(45) as G2Desc, space(45) as G3Desc,
  'OdOmCmCd,OdCoCd,OdTc,OdYy,OdChr,OdNo,OdSr ,IfTc as InTc,IfYy as InYy,
  'IfChr as InChr,IfNo as InNo,IfSr as InSr,
  ' Space(1) as qTag, TdCoCd, BYy, BChr, BNo, TdRmWt as qBMWt,0 as qMul,
  ' TdRmQty as qBMQty, 0.0 as qBRmWt, 0 as qBRmQty, FgDt as FgDt, TdRmCtg as qRmCtg, TdRmSCtg as qRmSCtg, TdRmCd as qRmCd,
  ' TdRmSz as qTdRmSz,RrRngSz as qRngSz,TdToRmDc as qRmDc,
  ' TdTc, TdYy, TdChr, TdNo, TdSr, DmCtg, DmPrdCtg,OdDmCd, OdKt, OmPoNo,TdFrRmLocTyp as qBM ,
  '  '' As qQw, 0.0 as qRmVal , 0.0 as qRmWt, 0 as qRmQty
  ' From Txnd, Fgd, RmRt, FgRm, OrdDsg, InvFgd, Bag, FG, DsgMst, OrdMst
  ' Where 1 = 2
  '*** Report Sql ***
  
  Dim grpflds As String, wSqlStrg As String, wrepcnd As String, wCnd As String, wBM As String
  Dim wRmDet As String, wRmGrp As String, wCust As String, wOrdJoin As String
  Dim wOpt As String, wIrRmWtOpt As String, wIrRmRtOpt As String, wIrRmValOpt As String
  Dim wAddlCnd As String
  Dim ws_TmpTblA As String
  Dim wDetPos As Integer
    
  ' Zubin 213
  Dim wb_InBaseCur As Boolean
 
  Set Rep = Nothing
  Set adc.RepSource = Rep
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
  wDetPos = 0
  For i = 1 To 3
    If wGrp(i) = "(DsgCd)Colour" Or wGrp(i) = "Design Colour" Then wDsgColGrpPos = i
    If wGrp(i) = "Rm Ctg" Or wGrp(i) = "Rm SubCtg" Or wGrp(i) = "Rm Code" Or wGrp(i) = "Rm Sv/Len" Then wRmGrpOpt = wGrp(i)
    If wGrp(i) = "(Detail)" Then wDetPos = i
  Next i
  If wDetPos = 0 Then Rep.DetSec2.Suppress = True
  
  Dim ws_BaseCurCd, ws_RmRt As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")

  If adc("UoYN") = "Y" Then
    wCust = "IsNull(InCmCd,'')"
    Rep.hOdOmCmCd.SetText "Inv Cust"
  Else
    wCust = "OmCmCd"
    Rep.hOdOmCmCd.SetText "Ord Cust"
  End If
  
  adc("UoCmCdFr").CmpStr = wCust + ">= "
  adc("UoCmCdTo").CmpStr = wCust + "<= "
  adc("UoCmCdSel").CmpStr = wCust + " In "
  
 If (adc("UOPMCD") = UCase("CstVal")) Then
    ws_RmRt = ",(Select rQw.RmQwCst From RmMst rQw Where " + IIF(gs_Partition = ctCurrPrtn, " rQw.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " rQw.RmCd=TdRmCd) As qQw, isnull((Select RrCstRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
    "RrCtg= TdRmCtg and RrSCtg= TdRmSCtg and RrCd= TdRmCd and RrCmCd= '" + ctValueCmCd + "' and " + _
    "RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn<= TdRmSz And RrToLn>= TdRmSz and RrDmCtg= '' and RrLabMCd= ''), 0) as qRmVal "
    
 ElseIf (adc("UOPMCD") = UCase("SalVal")) Then
    ws_RmRt = ",(Select rQw.RmQw From RmMst rQw Where " + IIF(gs_Partition = ctCurrPrtn, " rQw.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " rQw.RmCd=TdRmCd) As qQw, isnull((Select RrSalRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
    "RrCtg= TdRmCtg and RrSCtg= TdRmSCtg and RrCd= TdRmCd and RrCmCd= '" + ctValueCmCd + "' and " + _
    "RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn<= TdRmSz And RrToLn>= TdRmSz and RrDmCtg= '' and RrLabMCd= ''), 0) as qRmVal "
    
 ElseIf (adc("UOPMCD") = UCase("StkVal")) Then
    ws_RmRt = ",'W' As qQw, TdRmStkRt as qRmVal "
    
    '4.1.4
    Dim ws_CRmRt As String, ws_SRmRt As String
    If UCase(adc("UoPMCdTo")) = UCase("CstVal") Then
        ws_RmRt = " ,(Case When RSCtg.PValue1='N' Then (Select rQw.RmQwCst From RmMst rQw Where " + IIF(gs_Partition = ctCurrPrtn, " rQw.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " rQw.RmCd=TdRmCd) Else 'W' End) As qQw, " + _
                   " (Case When RSCtg.PValue1='N' Then Isnull((Select RrCstRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
                   " RrCtg= TdRmCtg and RrSCtg= TdRmSCtg and RrCd= TdRmCd and RrCmCd= '" + ctValueCmCd + "' and " + _
                   " RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn<= TdRmSz And RrToLn>= TdRmSz and RrDmCtg= '' and RrLabMCd= ''), 0) Else TdRmStkRt End) as qRmVal "
    
    ElseIf UCase((adc("UoPMCdTo")) = UCase("SalVal")) Then
        ws_RmRt = " ,(Case When RSCtg.PValue1='N' Then (Select rQw.RmQw From RmMst rQw Where " + IIF(gs_Partition = ctCurrPrtn, " rQw.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " rQw.RmCd=TdRmCd) Else 'W' End) As qQw, " + _
                  " (Case When RSCtg.PValue1='N' Then IsNull((Select RrSalRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
                  " RrCtg= TdRmCtg and RrSCtg= TdRmSCtg and RrCd= TdRmCd and RrCmCd= '" + ctValueCmCd + "' and " + _
                  " RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn<= TdRmSz And RrToLn>= TdRmSz and RrDmCtg= '' and RrLabMCd= ''), 0) Else TdRmStkRt End) as qRmVal "
    End If
    
 ElseIf (adc("UOPMCD") = "") Then
    ws_RmRt = ",' ' As qQw, 0.0 as qRmVal "
 End If
 
 '*** Brk/Msg or Both ***
  If adc("UoBrkMsg") = "B" Then
    wBM = "'B'"
    Rep.TxtHead.SetText "Breakage Report (Fg Level)"
  ElseIf adc("UoBrkMsg") = "M" Then
    wBM = "'M'"
    Rep.TxtHead.SetText "Missing Report (Fg Level)"
  ElseIf adc("UoBrkMsg") = "" Then
    wBM = "'B','M'"
    Rep.TxtHead.SetText "Breakage, Missing Report (Fg Level)"
  End If
  
  Dim wInvDt As String
'  If adc("UoYN3") = "Y" Then
'    wInvDt = "InDt "
'  Else
'    wInvDt = "InPrnDt "
'  End If
' adc("UoDtFr").CmpStr = wInvDt + ">= "
' adc("UoDtTo").CmpStr = wInvDt + "<= "
  
 Dim wFgBmSrt, wFormulaStr As String
 
wFormulaStr = " Trim (ToText ({rdo.IdIdNo}, '######0'))+'/ '+Trim (ToText ({rdo.OdIdNo}, '######0'))"
 Rep.FldqRmCtg.Suppress = True: Rep.FldqRmSCtg.Suppress = True: Rep.FldqRmCd.Suppress = True: Rep.FldqTdRmSz.Suppress = True: Rep.FldqRngSz.Suppress = True
 Rep.FldqRmCtg1.Suppress = True: Rep.FldqRmSCtg1.Suppress = True: Rep.FldqRmCd1.Suppress = True: Rep.FldqTdRmSz1.Suppress = True: Rep.FldqRngSz1.Suppress = True
 Rep.hqRmCtg.Suppress = True: Rep.hqRmsCtg.Suppress = True: Rep.hqRmCd.Suppress = True: Rep.hqRmSz.Suppress = True: Rep.hqRngSz.Suppress = True
 If adc("UoPMCdFr") <> "" Then
    wFgBmSrt = moCn.GetFldVal("Select PDesc225 From Param where PTyp='FGBMSRT' and PMCd='" + adc("UoPMCdFr") + "' and PScd='' ")
    wFgBmSrt = UCase(wFgBmSrt)
    If InStr(1, "," + wFgBmSrt + ",", ",RMCTG,") > 0 Then Rep.FldqRmCtg.Suppress = False: Rep.FldqRmCtg1.Suppress = False: Rep.hqRmCtg.Suppress = False: wFormulaStr = wFormulaStr + "+'/'+{rdo.qRmCtg}"
    If InStr(1, "," + wFgBmSrt + ",", ",RMSCTG,") > 0 Then Rep.FldqRmSCtg.Suppress = False: Rep.FldqRmSCtg1.Suppress = False: Rep.hqRmsCtg.Suppress = False: wFormulaStr = wFormulaStr + "+'/'+{rdo.qRmSCtg}"
    If InStr(1, "," + wFgBmSrt + ",", ",RMCD,") > 0 Then Rep.FldqRmCd.Suppress = False: Rep.FldqRmCd1.Suppress = False: Rep.hqRmCd.Suppress = False: wFormulaStr = wFormulaStr + "+'/'+{rdo.qRmCd}"
    If InStr(1, "," + wFgBmSrt + ",", ",RMSZ,") > 0 Then Rep.FldqTdRmSz.Suppress = False: Rep.FldqTdRmSz1.Suppress = False: Rep.hqRmSz.Suppress = False: wFormulaStr = wFormulaStr + "+'/'+Trim (ToText ({rdo.qTdRmSz}, '####.###0'))"
    If InStr(1, "," + wFgBmSrt + ",", ",RRRNGSZ,") > 0 Then Rep.FldqRngSz.Suppress = False: Rep.FldqRngSz1.Suppress = False: Rep.hqRngSz.Suppress = False: wFormulaStr = wFormulaStr + "+'/'+{rdo.qRngSz}"
 End If
 GRP_REP.SetFormula Rep, "wGrp4", wFormulaStr
 
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  '6.1
 ws_TmpTblA = "Set DateFormat " + DtFmtStr() + "; Select OdCoCd as OdCoCd," + wCust + " as OdOmCmCd,OdTc,OdYy,OdChr,OdNo,OdSr,OdIdNo ,IsNull(IdIdNo,0) as IdIdNo,BIdNo,BPrtKey,FdYy+Str(FdKey) as qFdYyKey, " + _
              "   IsNull(IfTc,'') as IfTc,IsNull(IfYy,'') as IfYy,IsNull(IfChr,'') as IfChr,IsNull(IfNo,'') as IfNo,IsNull(IfSr,0) as IfSr,IsNull(InExpNo,'') as InExpNo,BYy as BYy,BChr,BNo,FdDt as FgDt, " + _
              "   DmCtg ,DmCd,DmSalCtg, DmPrdCtg, OdDmCd, OdKt, OmPoNo Into #TmpA " + _
              "  From Fgd Fd " + _
              "  Join Bag On FdCoCd=BCoCd and FdBYy=BYy and FdBChr=BChr and FdBNo=BNo and FdBIdNo=BIdNo " + _
              "  Join Param On PTyp= 'BCHR' and PMCd= FdBChr and PSCd= ''  and PValue= 'P' " + _
              "  Join OrdDsg on OdCoCd= FdCoCd and OdTc= FdPrdOdTc and OdYy= FdPrdOdYy and OdChr= FdPrdOdChr and OdNo= FdPrdOdNo and OdSr= FdPrdOdSr  And OdPrtKey=FdPrtKey " + _
              "  Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo  And OmPrtKey=OdPrtKey " + _
              "  Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp and DmCd= OdDmCd and DmSz= ''  And DmPrtKey=OdPrtKey " + _
              "  Left Outer Join InvFgd On IfTc='IN' and  IfCoCd=FdCoCd and IfFdBYy=FdBYy and IfFdBChr=FdBChr and IfFdBNo=FdBNo and IfBIdNo=FdBIdNo and IfPrtKey=FdPrtKey " + _
              "  Left Outer Join InvDsg On IfCoCd=IdCoCd and IfTc=IdTc and  IfYy=IdYy and IfChr=IdChr and IfNo=IdNo and IfSr=IdSr and IfPrtKey=IdPrtKey " + _
              "  Left Outer Join InvHd On InCoCd=IdCoCd and InTc=IdTc and InYy=IdYy and InChr=IdChr and InNo=IdNo " + _
              "   where BLoc='PFG' and FdYy+ Str(FdKey)=(Select max(FdYy+ Str(FdKey)) From Fgd where FdCoCd=BCoCd and FdBYy=BYy and FdBChr=BChr and FdBNo=BNo and FdBIdNo=BIdNo) " + wCnd
              

     moCn.CreateTmpTable "#TMPA", ws_TmpTblA
    
    
     wCnd = IIF(adc("UORMCTGFR") <> "", " And TdRmCtg>='" + adc("UORMCTGFR") + "'", "")
     wCnd = wCnd + IIF(adc("UORMCTGTo") <> "", " And TdRmCtg<='" + adc("UORMCTGTo") + "'", "")
     wCnd = wCnd + IIF(adc("UORMCTGSEL") <> "", " And TdRmCtg in (" + adc("UORMCTGFR") + ")", "")
     
     wCnd = wCnd + IIF(adc("UORMSCTGFR") <> "", " And TdRmSCtg>='" + adc("UORMSCTGFR") + "'", "")
     wCnd = wCnd + IIF(adc("UORMSCTGTo") <> "", " And TdRmSCtg<='" + adc("UORMSCTGTo") + "'", "")
     wCnd = wCnd + IIF(adc("UORMSCTGSEL") <> "", " And TdRmSCtg in (" + adc("UORMSCTGSEL") + ")", "")
     
     wCnd = wCnd + IIF(adc("UORMCDFR") <> "", " And TdRmCd>='" + adc("UORMCDFR") + "'", "")
     wCnd = wCnd + IIF(adc("UORMCDTo") <> "", " And TdRmCd<='" + adc("UORMCDTo") + "'", "")
     wCnd = wCnd + IIF(adc("UORMCDSEL") <> "", " And TdRmCd in (" + adc("UORMCDSEL") + ")", "")
     
     wCnd = wCnd + IIF(adc("UOLN1FR") <> 0, " And TdRmSz<=" + CStr(adc("UOLN1FR")) + "", "")
     wCnd = wCnd + IIF(adc("UOLN1To") <> 0, " And TdRmSz>=" + CStr(adc("UOLN1To")) + "", "")
    ' wCnd = ""
     
     
    
    wSqlStrg = "Select " + grpflds + ", " + _
               " OdOmCmCd,OdCoCd,OdTc,OdYy,OdChr,OdNo,OdSr ,IfTc as InTc,IfYy as InYy, " + _
               "IfChr as InChr,IfNo as InNo,IfSr as InSr,OdIdNo,IdIdNo,'D' as qTag, TdCoCd, BYy, BChr, BNo, TdRmWt as qBMWt," + _
               "(Case when TdToRmLocTyp in ('B', 'M') and TdToRmDc= 'D' then 1 " + _
               "    when TdToRmLocTyp in ('B', 'M') and TdToRmDc= 'C' then -1 " + _
               "    when TdFrRmLocTyp in ('B', 'M') and TdFrRmDc= 'D' then 1 " + _
               "    when TdFrRmLocTyp in ('B', 'M') and TdFrRmDc= 'C' then -1 Else 0 end) as qMul, " + _
               "  TdRmQty as qBMQty, 0.0 as qBRmWt, 0 as qBRmQty, FgDt as FgDt, TdRmCtg as qRmCtg, TdRmSCtg as qRmSCtg, TdRmCd as qRmCd, " + _
               "  TdRmSz as qTdRmSz,IsNull(RrRngSz, '') as qRngSz, (Case when TdToRmLocTyp in ('B', 'M') then TdToRmDc " + _
               "           when TdFrRmLocTyp in ('B', 'M') then TdFrRmDc Else '' end) as qRmDc, " + _
               "  TdTc, TdYy, TdChr, TdNo, TdSr, DmCtg, DmPrdCtg,OdDmCd, OdKt, OmPoNo, " + _
               "  (Case when TdToRmLocTyp in ('B', 'M') then TdToRmLocTyp  when TdFrRmLocTyp in ('B', 'M') then TdFrRmLocTyp Else '' end) as qBM " + _
               ws_RmRt + ", 0.0 as qRmWt, 0 as qRmQty " + _
               "  From Txnd " + _
               "  Join #TMPA  On TdCoCd=OdCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo and TdBIdNo=BIdNo and TdPrtKey=BPrtKey " + _
               " Join RmMst On RmCd=TdRmCd " + _
               " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=RmCtg And RSCtg.PSCd=RmSCtg " + _
               "  Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= 'ZSELF' and " + _
               "  RrCtg= TdRmCtg and RrSCtg= TdRmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
               "  RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= TdRmSz and RrToLn>= TdRmSz " + _
               " where (TdToRmLocTyp in (" + wBM + ") or TdFrRmLocTyp in (" + wBM + "))  And (TdToRmLocTyp in (" + wBM + ") Or TdFrRmLocTyp in (" + wBM + ")) and not(TdFrRmLocTyp in (" + wBM + ") " + _
               "  and TdToRmLocTyp in (" + wBM + ")) and not(TdFrRmLocTyp in (" + wBM + ") and TdToRmLocTyp in (" + wBM + ")) and TdSrNo<> 0  And TdPrtKey='C' " + _
               "   And TdRmCtg in ('D','C') " + wCnd + _
               " Union All "



     wCnd = IIF(adc("UORMCTGFR") <> "", " And FrRmCtg>='" + adc("UORMCTGFR") + "'", "")
     wCnd = wCnd + IIF(adc("UORMCTGTo") <> "", " And FrRmCtg<='" + adc("UORMCTGTo") + "'", "")
     wCnd = wCnd + IIF(adc("UORMCTGSEL") <> "", " And FrRmCtg in (" + adc("UORMCTGSEL") + ")", "")
     
     wCnd = wCnd + IIF(adc("UORMSCTGFR") <> "", " And FrRmSCtg>='" + adc("UORMSCTGFR") + "'", "")
     wCnd = wCnd + IIF(adc("UORMSCTGTo") <> "", " And FrRmSCtg<='" + adc("UORMSCTGTo") + "'", "")
     wCnd = wCnd + IIF(adc("UORMSCTGSEL") <> "", " And FrRmSCtg in (" + adc("UORMSCTGSEL") + ")", "")
     
     wCnd = wCnd + IIF(adc("UORMCDFR") <> "", " And FrRmCd>='" + adc("UORMCDFR") + "'", "")
     wCnd = wCnd + IIF(adc("UORMCDTo") <> "", " And FrRmCd<='" + adc("UORMCDTo") + "'", "")
     wCnd = wCnd + IIF(adc("UORMCDSEL") <> "", " And FrRmCd in (" + adc("UORMCDSEL") + ")", "")
     
     wCnd = wCnd + IIF(adc("UOLN1FR") <> 0, " And FrRmSz<=" + CStr(adc("UOLN1FR")) + "", "")
     wCnd = wCnd + IIF(adc("UOLN1To") <> 0, " And FrRmSz>=" + CStr(adc("UOLN1To")) + "", "")
     'wCnd = ""


        wSqlStrg = wSqlStrg + _
                    "  Select " + grpflds + ",  OdOmCmCd,OdCoCd,OdTc,OdYy,OdChr,OdNo,OdSr ,IfTc,IfYy, " + _
                    "  IfChr,IfNo,IfSr,OdIdNo,IdIdNo,'P' as qTag, FdCoCd, FdBYy, FdBChr, FdBNo, 0 as qBMWt, " + _
                    "  0 as qMul,0 as qBMQty,FrRmWt*(Case when FrFrRmLocTyp= 'F' and FrFrRmDc= 'D' then 1 " + _
                    "        when FrFrRmLocTyp= 'F' and FrFrRmDc= 'C' then -1 " + _
                    "        when FrToRmLocTyp= 'F' and FrToRmDc= 'D' then 1 " + _
                    "        when FrToRmLocTyp= 'F' and FrToRmDc= 'C' then -1 " + _
                    "         Else 0 end) as qBRmWt, FrRmQty*(Case when FrFrRmLocTyp= 'F' and FrFrRmDc= 'D' then 1 " + _
                    "       when FrFrRmLocTyp= 'F' and FrFrRmDc= 'C' then -1 " + _
                    "       when FrToRmLocTyp= 'F' and FrToRmDc= 'D' then 1 " + _
                    "       when FrToRmLocTyp= 'F' and FrToRmDc= 'C' then -1 " + _
                    "        Else 0 end) as qBRmQty,FdDt as TDt, FrRmCtg as qRmCtg, FrRmSCtg as qRmSCtg, FrRmCd as qRmCd," + _
                    "        FrRmSz as qTdRmSz,IsNull(RrRngSz, '') as qRngSz, '' as qRmDc,FdTc, FdYy, FdChr, FdNo, FdSr, DmCtg, DmPrdCtg,OdDmCd, OdKt, OmPoNo, '' as qBM, " + _
                    "     '' As qQw, 0.0 as qRmVal , FrRmWt*(Case when FrFrRmLocTyp= 'F' and FrFrRmDc= 'D' then 1 " + _
                    "        when FrFrRmLocTyp= 'F' and FrFrRmDc= 'C' then -1 " + _
                    "        when FrToRmLocTyp= 'F' and FrToRmDc= 'D' then 1 " + _
                    "        when FrToRmLocTyp= 'F' and FrToRmDc= 'C' then -1 " + _
                    "         Else 0 end) as qRmWt, FrRmQty*(Case when FrFrRmLocTyp= 'F' and FrFrRmDc= 'D' then 1 " + _
                    "       when FrFrRmLocTyp= 'F' and FrFrRmDc= 'C' then -1 " + _
                    "       when FrToRmLocTyp= 'F' and FrToRmDc= 'D' then 1 " + _
                    "       when FrToRmLocTyp= 'F' and FrToRmDc= 'C' then -1 " + _
                    "        Else 0 end) as qRmQty"

          wSqlStrg = wSqlStrg + _
                    "   From FgRm " + _
                    "   Join Fgd On FdCoCd=FrCoCd and FdTc=FrTc and FdYy=FrYy and FdChr=FrChr and FdNo=FrNo and FdSr=FrSr " + _
                    "   Join #TmpA  On FdCoCd=OdCoCd and BYy=FdBYy and BChr=FdBChr and BNo=FdBNo and BIdNo=FdBIdNo and BPrtKey=FdPrtKey and  FdYy+Str(FdKey)=qFdYyKey " + _
                    " Join RmMst On RmCd=FrRmCd " + _
                    " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=RmCtg And RSCtg.PSCd=RmSCtg " + _
                    "   Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= 'ZSELF' and " + _
                    "    RrCtg= FrRmCtg and RrSCtg= FrRmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                    "     RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
                    "  where FrRmCtg in ('D','C') " + wCnd
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  
  If adc("UoYN1") = "N" Or wDetPos = 0 Then
    Rep.BoxHd.Suppress = True
    Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True
    Rep.DET.Suppress = True
  Else
    Rep.DET.Suppress = False
    Rep.FldOdOmCmCd.Suppress = True
    Rep.FldwInvNo.Suppress = True
    Rep.FldwOrdNo.Suppress = True
    Rep.FldqRmCtg.Suppress = True
    Rep.FldqRmSCtg.Suppress = True
    Rep.FldqRmCd.Suppress = True
    Rep.FldqRngSz.Suppress = True
    Rep.FldqTdRmSz.Suppress = True
  End If
    
  CRV_REP.DisplayGroupTree = False
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
End Sub
' Zubin 213
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
  Case Is = UCase("UoPmCd")
    If adc("UoYn2") <> "Y" Then adc("UoPmCd") = ""

  Case Is = UCase("UoCurCdFr")
    If UCase(adc("UoInVal")) = "S" Then adc("UoCurCdFr") = "": Exit Sub
  End Select
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
    
  '4.1.0.0
  Case Is = 1
    If .Value = Checked Then
      If ms_PwdFlag = False Then
        ChkBoxArr(1).Value = Unchecked: adc("UoYN2") = "N": DispMsg "Enter SEO Password To Show Value", etError
      Else
        adc("UoYN2") = "Y"
      End If
    Else
      adc("UoYN2") = "N"
    End If
  End Select
End With
End Sub
Private Sub CmdPwd_Click()
   '4.1.0.0
   If FraPwd.Visible = True Then
      FraPwd.Visible = False
      FraPwd.Enabled = False
      CmdPwd.SetFocus
   Else
      FraPwd.Visible = True
      FraPwd.Enabled = True
      FraPwd.ZOrder
      TxtPwd.SetFocus
   End If
End Sub
Private Sub OptBrkMsg_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoBrkMsg") = "B"
Case Is = 1
  adc("UoBrkMsg") = "M"
Case Is = 2
  adc("UoBrkMsg") = ""
End Select
End Sub
Private Sub OptBrkMsg_GotFocus(Index As Integer)
  DispMsg FraOptBrkMsg.ToolTipText, etInfo
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
  Set moCn = adc.Connection  '*** (Bef speed) bef 24/11/06
  Set adc.FirNKeyCtl = adc("UoInTcFr") 'adc("UoCoCdFr")
  Call InitProp(Me)
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = False
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******

  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
    Set adc.FirNKeyCtl = adc("UoCoCdFr")
  Else
    Set adc.FirNKeyCtl = adc("UOInTcFr")
  End If
  
  If ws_MultiCoMod = True Then
      adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
  Else
      adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  FraPwd.Visible = False
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gb_RmCtgFor = False
    
  Call SetGroupSort
  
  '*** Jay 2.14 *** (SEO Change)
  ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")

End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  
  gb_CoCdFor = False  '****** Sachin 2.14.0 - [Multi-Company]
  
  Select Case UCase(IdName)
   Case Is = UCase("UoPMCd"), UCase("UOPMCDTO")      '4.1.4
        If UCase(IdName) = UCase("UoPmCd") And pv_NewValue = "" Then adc("UoPMCdTo") = "" '4.1.4
        gs_PTyp1 = "WIPVAL": gs_PTyp = "WIPVAL"
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = False  '****** Sachin 2.14.0 - [Multi-Company]
  
  ' Zubin 213
  Select Case UCase(IdName)
  Case Is = UCase("UOPMCD"), UCase("UOPMCDTO")   '4.1.4
    If adc("UOYN2") <> "Y" Then Cancel = True: ErrMsg = "Can Not Edit When Show Value Option Is Not Checked": Exit Sub
    gs_PTyp1 = "WIPVAL"
    
    If UCase(IdName) = UCase("UOPMCDTO") Then    '4.1.4
      gs_PTyp = "WIPVAL"
      If adc.Mode = xNorm Then _
        If UCase(adc("UoPMCd")) <> UCase("STKVAL") Then Cancel = True: ErrMsg = "Can Enter Only if Value Type is 'STKVAL'": Exit Sub
    End If
    
  Case Is = UCase("UoCurCdFr")
    If UCase(adc("UoInVal")) = "S" Then Cancel = True: Exit Sub
    
  Case Is = UCase("UoLotNoFr"), UCase("UoLotNoFr"), UCase("UoLotNoSel")
    gs_RmLocTyp = "XS"
    
  Case Is = UCase("UoPMCdFr")
    gs_PTyp = "FGBMSRT"
    gs_Tbl = "Param"
    
  End Select
  ' Zubin 213
  SetProp Me, IdName, When
End Sub

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptCust_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN") = "Y"
Case Is = 1
  adc("UoYN") = "N"
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
  adc("UoYN3") = "Y"
Case Is = 1
  adc("UoYN3") = "N"
End Select
End Sub

Private Sub OptInvDt_GotFocus(Index As Integer)
  'DispMsg FraOptDt.ToolTipText, etInfo
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
  
    ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  If adc("UoYN") = "Y" Then OptCust(0).Value = True
  If adc("UoYN") = "N" Then OptCust(1).Value = True
  ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  ' ****** Manali 3.9.0 - Option for InDt / InPrnDt added
 
  adc("UoYn2") = "N"
  adc("UoPMCD") = ""
 
End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
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
  If ws_MultiCoMod = True Then
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdFr.SetText adc("UOCOCDFR"): Rep.TxtCoCdTo.SetText adc("UOCOCDTO"): Rep.TxtCoCdSel.SetText adc("UOCOCDSEL")
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
Private Sub HideAllFras()
  '*** This procedure hides all the frames
  FraPwd.Visible = False
  FraPwd.Enabled = False
End Sub
Private Sub TxtPwd_Validate(Cancel As Boolean)
    If Trim(TxtPwd.text) = "" Then
      adc("UoYn1") = ""
      ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
      adc("UoYn2") = "N": ChkBoxArr(1).Value = Unchecked
    Else
      If Trim(TxtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                     "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + adc.MenuCd + "' and " + _
                     "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                     "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
          ms_PwdFlag = True: Exit Sub
      Else
          ms_PwdFlag = False
          adc("UoYn2") = "N": ChkBoxArr(1).Value = Unchecked
          Cancel = True:  DispMsg "Invalid Password / User Does Not Have Seo Rights ", etError
      End If
    End If
End Sub
