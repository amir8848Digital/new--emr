VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpInvProfitMrg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Profit Margin"
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11430
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   9630
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4410
      TabIndex        =   47
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
      Left            =   5535
      TabIndex        =   57
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   58
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
         TabIndex        =   59
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
      Left            =   6750
      TabIndex        =   56
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   53
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
      TabIndex        =   49
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
      TabIndex        =   52
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   51
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpInvProfitMrg.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpInvProfitMrg.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   55
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   50
               Top             =   360
               Width           =   8475
               _ExtentX        =   14949
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   54
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
               Left            =   0
               TabIndex        =   46
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Stock Rate"
               Top             =   6150
               Width           =   2100
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1950
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   45
               Top             =   5790
               Width           =   2340
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   1950
               TabIndex        =   32
               ToolTipText     =   "Enter From Karat"
               Top             =   3855
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               Data            =   "ABCDW"
               CmpStr          =   "IdKt >="
            End
            Begin VB.Frame FraOptCstRt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1950
               TabIndex        =   82
               ToolTipText     =   "Specify Whether To Consider The Invoice LME Cost Rates Or Order Lme Cost Rates For Calculating Actual Value"
               Top             =   4995
               Width           =   4665
               Begin VB.OptionButton OptCstRt 
                  Caption         =   "Order"
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
                  TabIndex        =   43
                  ToolTipText     =   "Specify Whether To Consider The Invoice LME Cost Rates Or Order Lme Cost Rates For Calculating Actual Value"
                  Top             =   0
                  Width           =   1425
               End
               Begin VB.OptionButton OptCstRt 
                  Caption         =   "Invoice "
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
                  TabIndex        =   42
                  ToolTipText     =   "Specify Whether To Consider The Invoice LME Cost Rates Or Order Lme Cost Rates For Calculating Actual Value"
                  Top             =   0
                  Width           =   1545
               End
            End
            Begin VB.Frame FraOptValWt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1950
               TabIndex        =   81
               ToolTipText     =   "Select if to show Value or Weight"
               Top             =   4710
               Width           =   4665
               Begin VB.OptionButton OptValWt 
                  Caption         =   "Value"
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
                  TabIndex        =   39
                  ToolTipText     =   "Select if to show Value or Weight"
                  Top             =   0
                  Width           =   1425
               End
               Begin VB.OptionButton OptValWt 
                  Caption         =   "Weight"
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
                  Left            =   2610
                  TabIndex        =   40
                  ToolTipText     =   "Select if to show Value or Weight"
                  Top             =   0
                  Width           =   1425
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5430
               TabIndex        =   5
               ToolTipText     =   "Enter To Invoice Year"
               Top             =   975
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYTO"
               IdName          =   "UOINYYTO"
               CmpStr          =   "IdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2730
               TabIndex        =   2
               ToolTipText     =   "Enter From Invoice Year"
               Top             =   975
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "IdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1950
               TabIndex        =   1
               ToolTipText     =   "Enter From Invoice Tc"
               Top             =   975
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "IdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3150
               TabIndex        =   3
               ToolTipText     =   "Enter From Invoice Character"
               Top             =   975
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "IdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4650
               TabIndex        =   4
               ToolTipText     =   "Enter To Invoice Tc"
               Top             =   975
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCTO"
               IdName          =   "UOINTCTO"
               CmpStr          =   "IdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   5850
               TabIndex        =   6
               ToolTipText     =   "Enter To Invoice Character"
               Top             =   975
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRTO"
               IdName          =   "UOINCHRTO"
               CmpStr          =   "IdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1950
               TabIndex        =   7
               ToolTipText     =   "Enter From Invoice Number"
               Top             =   1260
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "IdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   4650
               TabIndex        =   8
               ToolTipText     =   "Enter To Invoice Number"
               Top             =   1260
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOTO"
               IdName          =   "UOINNOTO"
               CmpStr          =   "IdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1950
               TabIndex        =   35
               ToolTipText     =   "Specify Whether Base Value Is Actual, Invoice Or Customs "
               Top             =   4140
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
               Index           =   4
               Left            =   1950
               TabIndex        =   38
               ToolTipText     =   "Show Value Or Weight Figures ('Y' for Value / 'N' For Weight)"
               Top             =   4110
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
               Index           =   1
               Left            =   1950
               TabIndex        =   37
               ToolTipText     =   "Specify Whether Other Value Is Actual, Invoice Or Customs "
               Top             =   4425
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
               Index           =   28
               Left            =   4650
               TabIndex        =   24
               ToolTipText     =   "Enter To Design Category"
               Top             =   3000
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
               Left            =   7380
               TabIndex        =   25
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   3000
               Width           =   7470
               _ExtentX        =   13176
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
               Left            =   1950
               TabIndex        =   23
               ToolTipText     =   "Enter From Design Category"
               Top             =   3000
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
               Left            =   4650
               TabIndex        =   30
               ToolTipText     =   "Enter To Design Code"
               Top             =   3570
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
               Left            =   7380
               TabIndex        =   31
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   3570
               Width           =   7470
               _ExtentX        =   13176
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "IdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1950
               TabIndex        =   29
               ToolTipText     =   "Enter From Design Code"
               Top             =   3570
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
               Left            =   4650
               TabIndex        =   12
               ToolTipText     =   "Enter To Customer"
               Top             =   1830
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "InCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   7380
               TabIndex        =   13
               ToolTipText     =   "Enter Customer Selection"
               Top             =   1830
               Width           =   7470
               _ExtentX        =   13176
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "InCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1950
               TabIndex        =   11
               ToolTipText     =   "Enter From Customer"
               Top             =   1830
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "InCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   4650
               TabIndex        =   33
               ToolTipText     =   "Enter To Karat"
               Top             =   3855
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
               Left            =   7380
               TabIndex        =   34
               ToolTipText     =   "Enter Karat Selection"
               Top             =   3855
               Width           =   7470
               _ExtentX        =   13176
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "IdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4650
               TabIndex        =   27
               ToolTipText     =   "Enter To Design Sales Category"
               Top             =   3285
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGTO"
               IdName          =   "UODMSALCTGTO"
               CmpStr          =   "DmSalCtg<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   7380
               TabIndex        =   28
               ToolTipText     =   "Enter Design Sales Category Selection"
               Top             =   3285
               Width           =   7470
               _ExtentX        =   13176
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMSALCTGSEL"
               IdName          =   "UODMSALCTGSEL"
               CmpStr          =   "DmSalCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   1950
               TabIndex        =   26
               ToolTipText     =   "Enter From Design Sales Category"
               Top             =   3285
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGFR"
               IdName          =   "UODMSALCTGFR"
               CmpStr          =   "DmSalCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4650
               TabIndex        =   21
               ToolTipText     =   "Enter To Export Number"
               Top             =   2715
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
               Left            =   7380
               TabIndex        =   22
               ToolTipText     =   "Enter Export Number Selection"
               Top             =   2715
               Width           =   7470
               _ExtentX        =   13176
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOINEXPNOSEL"
               IdName          =   "UOINEXPNOSEL"
               CmpStr          =   "InExpNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1950
               TabIndex        =   20
               ToolTipText     =   "Enter From Export Number"
               Top             =   2715
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
               Index           =   17
               Left            =   1950
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   685
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "IdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   4650
               TabIndex        =   10
               ToolTipText     =   "Enter To Finished Goods Transaction Date"
               Top             =   1545
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "InDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1950
               TabIndex        =   9
               ToolTipText     =   "Enter From Finished Goods Transaction Date"
               Top             =   1545
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "InDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1950
               TabIndex        =   41
               ToolTipText     =   "Specify Whether To Consider The Invoice LME Cost Rates For Calculating Actual Value ('Y' for Invoice / 'N' For Order)"
               Top             =   4395
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
               Index           =   44
               Left            =   1950
               TabIndex        =   44
               ToolTipText     =   "Enter Currency Code Whose Invoices You Wish To View Or Leave It Blank If You Wish To View All Invoices In Base Currency"
               Top             =   5520
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               ReCalcOn        =   "UOINVAL,UOINVAL1"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   7380
               TabIndex        =   16
               ToolTipText     =   "Enter Customer Group Selection"
               Top             =   2130
               Width           =   7620
               _ExtentX        =   13441
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMGRPSEL"
               IdName          =   "UOCMGRPSEL"
               CmpStr          =   "CmGrp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   4650
               TabIndex        =   15
               ToolTipText     =   "Enter To Customer Group"
               Top             =   2130
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
               Index           =   21
               Left            =   1950
               TabIndex        =   14
               ToolTipText     =   "Enter From Customer Group"
               Top             =   2130
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
               Index           =   22
               Left            =   7380
               TabIndex        =   19
               ToolTipText     =   "Enter Customer Type Selection"
               Top             =   2430
               Width           =   7620
               _ExtentX        =   13441
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMTYPSEL"
               IdName          =   "UOCMTYPSEL"
               CmpStr          =   "CmTyp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   4650
               TabIndex        =   18
               ToolTipText     =   "Enter To Customer Type"
               Top             =   2430
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
               Index           =   27
               Left            =   1950
               TabIndex        =   17
               ToolTipText     =   "Enter From Customer Type"
               Top             =   2430
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
               Index           =   23
               Left            =   4650
               TabIndex        =   36
               ToolTipText     =   "Specify Whether Other Value Is Actual, Invoice Or Customs "
               Top             =   4140
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOINVAL2"
               ReCalcParent    =   "UOCURCDFR"
               IdName          =   "UOINVAL2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   240
               Index           =   71
               Left            =   2640
               TabIndex        =   86
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   6120
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   423
               Enabled         =   0   'False
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
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
               TabIndex        =   87
               Top             =   5790
               Width           =   1575
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
               Left            =   3000
               TabIndex        =   85
               Top             =   4140
               Width           =   1485
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
               Index           =   11
               Left            =   60
               TabIndex        =   84
               Top             =   2120
               Width           =   1425
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
               TabIndex        =   83
               Top             =   2400
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
               Left            =   60
               TabIndex        =   80
               Top             =   5520
               Width           =   1005
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Take LME Cost Rt From "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   570
               Left            =   60
               TabIndex        =   79
               Top             =   4995
               Width           =   1845
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
               TabIndex        =   78
               Top             =   1545
               Width           =   1605
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
               Index           =   10
               Left            =   60
               TabIndex        =   77
               Top             =   1260
               Width           =   1605
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
               TabIndex        =   76
               Top             =   685
               Width           =   1335
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
               TabIndex        =   75
               Top             =   2700
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
               TabIndex        =   74
               Top             =   3855
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
               TabIndex        =   73
               Top             =   1830
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
               TabIndex        =   72
               Top             =   3570
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
               TabIndex        =   71
               Top             =   3000
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
               TabIndex        =   70
               Top             =   3285
               Width           =   1365
            End
            Begin VB.Label Label2 
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
               Left            =   60
               TabIndex        =   69
               Top             =   4425
               Width           =   1845
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Base Value As"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   68
               Top             =   4140
               Width           =   1845
            End
            Begin VB.Label LblVal 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Value/ Wt"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   67
               Top             =   4710
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
               Left            =   1950
               TabIndex        =   66
               Top             =   180
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
               Left            =   4650
               TabIndex        =   65
               Top             =   180
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
               Left            =   2670
               TabIndex        =   64
               Top             =   975
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
               Left            =   5370
               TabIndex        =   63
               Top             =   975
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
               Left            =   3090
               TabIndex        =   62
               Top             =   975
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
               Left            =   5790
               TabIndex        =   61
               Top             =   975
               Width           =   105
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
               TabIndex        =   60
               Top             =   975
               Width           =   1725
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpInvProfitMrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepProfitMrg
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "", "", "Sales Ctg", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= max(DmCtg)) "
  'Geeta***Crm***212***
  GRP_REP.Add "Sales Ctg", "max(DmSalCtg)", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and PmCd=Max(DmCtg) and PSCd= max(DmSalCtg)) "
  GRP_REP.Add "Inv Vch No", "IdTc+IdYy+IdChr+str(IdNo)", "IdTc+'/'+IdYy+'/'+IdChr+'/'+LTrim(Str(IdNo))", "", "", "Invoice Cust,Customer Grp,Customer Typ", "", "", "max(InExpNo) "  '4.1.0.0 - Cust Grp and Typ
  GRP_REP.Add "Inv Exp No", "max(InExpNo)", "", "", "", "Customer Grp,Customer Typ", "", "", ""
  'GRP_REP.Add "Design Code", "max(IdDmCd)", "", "IdDmCd", "hIdDmCd", "Design Ctg,Sales Ctg", "", "", "max(DmDesc) "
  GRP_REP.Add "Design Code", "max(IdDmCd)", "", "", "", "Design Ctg,Sales Ctg", "", "", "max(DmDesc) "
  GRP_REP.Add "ExpOrd Cust", "max(IdExpCmCd)", "", "", "", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= max(IdExpCmCd)) "
  
  'Sachin 4.1.0.0
  GRP_REP.Add "Invoice Cust", "max(InCmCd)", "", "", "", "Customer Grp,Customer Typ", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= max(InCmCd)) "
  
  'Sachin 4.1.0.0
  GRP_REP.Add "Customer Grp", "max(CmGrp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMGRP' and PmCd = Max(CmGrp))"
  GRP_REP.Add "Customer Typ", "max(CmTyp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMTYP' and PmCd = Max(CmTyp))"
  
  GRP_REP.Add "Karat", "max(IdKt)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = max(IdKt)) "
  GRP_REP.Add "PCS/PRS", "max(DmUom)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMUOM' and PmCd= max(DmUom)) "
  GRP_REP.Add "VaCtg", "max(IdVaCtg)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'VACTG' and PMCd = max(IdVaCtg)) "
  GRP_REP.Add "(DsgCd)Colour", "max(IdDmCd+OdDmCol)", "'('+max(IdDmCd)+')'+max(OdDmCol)", "IdDmCd", "hIdDmCd", "Design Ctg,Sales Ctg,Design Code,Design Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(OdDmCol)) "
  GRP_REP.Add "Design Colour", "max(OdDmCol)", "", "", "", "(DsgCd)Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(OdDmCol)) "
End Sub
Private Sub adc_setreprecsource()
' **** Zubin - Previous Report SQL **** '
  '*** Report Sql ***
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'IdTc, IdYy, IdChr, IdNo, IdSr, IdDmCd, IdQty as IdQty,
  'IdiFOB as qPcVal1, IdiFOB as qPcVal2, IrRmiVal as qMetVal1, IrRmiVal as qMetVal2,
  'IrRmiVal as qDCVal1, IrRmiVal as qDCVal2, IrRmiVal as qXVal1, IrRmiVal as qXVal2,
  'IrSetiVal as qSetVal1, IrSetiVal as qSetVal2, iLiVal as qLabVal1, iLiVal as qLabVal2,
  'IrRmiWt as qPcWt1, IrRmiWt as qPcWt2, IrRmiWt as qMetWt1, IrRmiWt as qMetWt2,
  'IrRmiWt As qDCWt1, IrRmiWt As qDCWt2, IrRmiWt As qXWt1, IrRmiWt As qXWt2
  'From InvDsg, DsgMst, InvHd, InvRm, InvLab, RmMst where 1= 2
  '*** Report Sql ***
'===============================================================================================================

  ' **** Zubin - New Report SQL **** '
  '  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  '  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  '  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  '  OdTc, OdYy, OdChr, OdNo, OdSr, OdDmCd, OdOrdQty, OdSalPrc as qPcVal1, OdCstPrc as qPcVal2,
  '  0 as qMetVal1, 0 as qMetVal2, 0 as qDCVal1, 0 as qDCVal2,  0 as qXVal1, 0 as qXVal2, OrSetSalVal as qSetVal1,
  '  OrSetCstVal as qSetVal2, 0 as qLabVal1, 0 as qLabVal2, 0 as qPcWt1, 0 as qPcWt2, 0 as qMetWt1, 0 as qMetWt2,
  '  0 as qDCWt1, 0 as qDCWt2, 0 as qXWt1, 0 as qXWt2
  '  From OrdDsg, DsgMst, OrdRm, RmMst, OrdMst, CustMst
  '  where 1= 2
  
  Dim grpflds As String, wSqlStrg As String, wStr1 As String, wStr2 As String
  Dim wDetPos As Integer, wNonePos As Integer, wrepcnd As String, wCnd As String
  Dim wBaseIdFOB As String, wBaseIrRmWt As String, wBaseIrSetVal As String, wBaseIrRmVal As String, wBaseiLVal As String
  Dim wOthIdFOB As String, wOthIrRmWt As String, wOthIrSetVal As String, wOthIrRmVal As String, wOthiLVal As String
  Dim wBase As String, wOth As String
    
  Dim wBaseIdAFOB As String, wBaseIrASetVal As String, wBaseIrARmVal As String, wBaseAiLVal As String
  Dim wOthIdAFOB As String, wOthIrASetVal As String, wOthIrARmVal As String, wOthAiLVal As String
  
  Dim wABaseIrRmVal As String, wIBaseIrRmVal As String, wZBaseIrRmVal As String     '4.1.4
  Dim wABaseFob As String, wIBaseFob As String, wZBaseFob As String                 '4.1.4
  Dim wAOthIrRmVal As String, wIOthIrRmVal As String, wZOthIrRmVal As String        '4.1.4
  Dim wAOthFob As String, wIOthFob As String, wZOthFob As String                    '4.1.4

  
  '*** (Jen 2.13)
  Dim ws_CstAdj As String, ws_BaseAdj As String, ws_OthAdj As String
  '*** (Jen 2.13)
  
  ' Zubin 213
  Dim wb_InBaseCur As Boolean
  Dim ws_CurJoin As String, ws_InIlJoin As String, ws_BaseCurCd As String
  ' Zubin 213
  Dim wAddlCnd As String '****** Sachin 3.02.0 - New fields in where clause
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  '*** To check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0: wNonePos = 0
  For i = 1 To 3
    If wGrp(i) = "(Detail)" Then wDetPos = i
    If wGrp(i) = "(None)" Then wNonePos = i: Exit For
  Next i
  
  '*** If Detail not selected then ***
  If wDetPos = 0 Then
    If wNonePos = 2 Then Rep.G1FWt.Suppress = True: Rep.G2FWt.Suppress = True
    If wNonePos = 3 Then Rep.G1FWt.Suppress = True
  Else
    If wNonePos = 2 Then Rep.G1FWt.Suppress = True: Rep.G2FWt.Suppress = True: Rep.G3FWt.Suppress = True
    If wNonePos = 3 Then Rep.G1FWt.Suppress = True: Rep.G3FWt.Suppress = True
    If wNonePos = 0 Then Rep.G3FWt.Suppress = True
  End If
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  
  ' Zubin 213
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  If adc("UoCurCdFr") = "" Then
    wb_InBaseCur = True
    ws_InIlJoin = " Join InvHd on " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey=IlPrtKey and ", "") + " InCoCd= IlCoCd and InTc= IlTc and InYy= IlYy and InChr= IlChr and InNo= IlNo "
    Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
    
  Else
    wb_InBaseCur = False
    'ws_CurJoin = " Join CustMst Cm On CmCtg= 'C' and CmCd= InCmCd "    'Sachin 4.1.0.0
    wCnd = wCnd + " and CmCurCd= '" + adc("UoCurCdFr") + "' "
    Rep.TxtValIn.SetText "Value In " + adc("UoCurCdFr")
  End If
  ' Zubin 213
  
  ws_CurJoin = " Join CustMst Cm On CmCtg= 'C' and CmCd= InCmCd "       'Sachin 4.1.0.0

  '*** (Jen 2.13)
  If adc("UoYN1") = "Y" Then
    'ws_CstAdj = " * (case RmCtg when 'G' then InLmgCst/ OmLmgCst when 'P' then InLmpCst/ OmLmpCst when 'S' then InLmsCst/ OmLmsCst else 0 end) "
    ' ***** Manali 3.03 - 27/06/08 - Pd Mod - 'L' and 'LML' included
    ws_CstAdj = " * (case RmCtg when 'G' then " + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then InLmgCst/InCnvFct else InLmgCst*InCnvFct end) ", "InLmgCst") + "/ " + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OmLmgCst/OmCnvFct else OmLmgCst*OmCnvFct end) ", "OmLmgCst") + _
                " when 'P' then " + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then InLmpCst/InCnvFct else InLmpCst*InCnvFct end) ", "InLmpCst") + "/ " + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OmLmpCst/OmCnvFct else OmLmpCst*OmCnvFct end) ", "OmLmpCst") + _
                " when 'S' then " + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then InLmsCst/InCnvFct else InLmsCst*InCnvFct end) ", "InLmsCst") + "/ " + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OmLmsCst/OmCnvFct else OmLmsCst*OmCnvFct end) ", "OmLmsCst") + _
                " when 'L' then " + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then InLmlCst/InCnvFct else InLmlCst*InCnvFct end) ", "InLmlCst") + "/ " + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OmLmlCst/OmCnvFct else OmLmlCst*OmCnvFct end) ", "OmLmlCst") + " else 0 end) "
    ws_BaseAdj = " ": ws_OthAdj = " "
  Else
    ws_CstAdj = " "
    ws_BaseAdj = " ": ws_OthAdj = " "
  End If
  '*** (Jen 2.13)
  
  If UCase(adc("UoInVal")) = "A" Then
    wBase = "Actual"
    wBaseIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdAVal/InCnvFct else IdAVal*InCnvFct end) ", "IdAVal")
    wBaseIrRmWt = "IrRmAWt"
    wBaseIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) ", "IrSetAVal")
    wBaseIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal")
    wBaseiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLAVal/InCnvFct else iLAVal* InCnvFct end) ", "iLAVal")
    Rep.TxtBase.SetText "(Base Value = Actual)"
    ws_BaseAdj = ws_CstAdj    '*** (Jen 2.13)
    
  ElseIf UCase(adc("UoInVal")) = "S" Then
    wBase = "Stock"
    wBaseIdFOB = "sum((IrStkRt* IrRmAWt)+ (Case when InMulDiv= 'M' then IrSetAVal/ InCnvFct else IrSetAVal* InCnvFct end)) " + _
                 "+ IsNull((Select sum((Case when InMulDiv= 'M' then iLAVal/ InCnvFct else iLAVal* InCnvFct end)) from InvLab " + _
                 ws_InIlJoin + " where " + IIF(gs_Partition = ctCurrPrtn, " IlPrtKey=IdPrtKey and ", "") + " iLCoCd= IdCoCd and iLTc= IdTc " + _
                 "and iLYy= IdYy and iLChr= IdChr and iLNo= IdNo and iLSr= IdSr), 0) "
    wBaseIrRmWt = "IrRmAWt"
    wBaseIrSetVal = " (Case when InMulDiv= 'M' then IrSetAVal/ InCnvFct else IrSetAVal* InCnvFct end) "
    wBaseIrRmVal = "IrStkRt*IrRmAWt"
    wBaseiLVal = " (Case when InMulDiv= 'M' then iLAVal/ InCnvFct else iLAVal* InCnvFct end) "
    Rep.TxtBase.SetText "(Base Value = Stock)"
    
    '4.1.4
    If Trim(adc("UoInVal2")) <> "" Then
        wABaseIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal")
        wIBaseIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIVal/InCnvFct else IrRmIVal* InCnvFct end) ", "IrRmIVal")
        wZBaseIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal* InCnvFct end) ", "IrRmZVal")
        
        wBaseIrRmVal = " (Case When RSCtg.PValue1='N' Then " + IIF(UCase(adc("UoInVal2")) = "A", wABaseIrRmVal, _
                                                               IIF(UCase(adc("UoInVal2")) = "I", wIBaseIrRmVal, _
                                                               IIF(UCase(adc("UoInVal2")) = "Z", wZBaseIrRmVal, wBaseIrRmVal))) + " Else " + wBaseIrRmVal + " End)"
        
        wABaseFob = "sum((Case When RSCtg.PValue1='N' Then IrRmAVal Else IrStkRt* IrRmAWt End) + (Case when InMulDiv= 'M' then IrSetAVal/ InCnvFct else IrSetAVal* InCnvFct end)) " + _
                    "+ IsNull((Select sum((Case when InMulDiv= 'M' then iLAVal/ InCnvFct else iLAVal* InCnvFct end)) from InvLab " + _
                    ws_InIlJoin + " where " + IIF(gs_Partition = ctCurrPrtn, " IlPrtKey=IdPrtKey and ", "") + " iLCoCd= IdCoCd and iLTc= IdTc " + _
                    "and iLYy= IdYy and iLChr= IdChr and iLNo= IdNo and iLSr= IdSr), 0) "
        
        wIBaseFob = "sum((Case When RSCtg.PValue1='N' Then IrRmIVal Else IrStkRt*IrRmIWt End)+ (Case when InMulDiv= 'M' then IrSetIVal/ InCnvFct else IrSetIVal* InCnvFct end)) " + _
                    "+ IsNull((Select sum((Case when InMulDiv= 'M' then IlIVal/ InCnvFct else IlIVal* InCnvFct end)) from InvLab " + _
                    ws_InIlJoin + " where " + IIF(gs_Partition = ctCurrPrtn, " IlPrtKey=IdPrtKey and ", "") + " iLCoCd= IdCoCd and iLTc= IdTc " + _
                    "and iLYy= IdYy and iLChr= IdChr and iLNo= IdNo and iLSr= IdSr), 0) "
        
        wZBaseFob = "sum((Case When RSCtg.PValue1='N' Then IrRmZVal Else IrStkRt*IrRmZWt End)+ (Case when InMulDiv= 'M' then IrSetZVal/ InCnvFct else IrSetZVal* InCnvFct end)) " + _
                    "+ IsNull((Select sum((Case when InMulDiv= 'M' then IlZVal/ InCnvFct else IlZVal* InCnvFct end)) from InvLab " + _
                    ws_InIlJoin + " where " + IIF(gs_Partition = ctCurrPrtn, " IlPrtKey=IdPrtKey and ", "") + " iLCoCd= IdCoCd and iLTc= IdTc " + _
                    "and iLYy= IdYy and iLChr= IdChr and iLNo= IdNo and iLSr= IdSr), 0) "
        
        wBaseIdFOB = IIF(UCase(adc("UoInVal2")) = "A", wABaseFob, _
                                                             IIF(UCase(adc("UoInVal2")) = "I", wIBaseFob, _
                                                             IIF(UCase(adc("UoInVal2")) = "Z", wZBaseFob, wBaseIdFOB)))
                                                             
        Rep.TxtBase.SetText "(Base Value = Stk-Cst)"
    End If
    
    
    
  ElseIf UCase(adc("UoInVal")) = "I" Then
    wBase = "Invoice"
    wBaseIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdIFob/ InCnvFct else IdIFob* InCnvFct end) ", "IdIFob")
    wBaseIrRmWt = "IrRmIWt"
    wBaseIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIVal/ InCnvFct else IrSetIVal* InCnvFct end) ", "IrSetIVal")
    wBaseIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIVal/ InCnvFct else IrRmIVal* InCnvFct end) ", "IrRmIVal")
    wBaseiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiVal/ InCnvFct else iLiVal* InCnvFct end) ", "iLiVal")
    Rep.TxtBase.SetText "(Base Value = Invoice)"
    
  ElseIf UCase(adc("UoInVal")) = "Z" Then
    wBase = "Custom"
    wBaseIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdZFob/ InCnvFct else IdZFob* InCnvFct end) ", "IdZFob")
    wBaseIrRmWt = "IrRmZWt"
    wBaseIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZVal/ InCnvFct else IrSetZVal* InCnvFct end) ", "IrSetZVal")
    wBaseIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZVal/ InCnvFct else IrRmZVal* InCnvFct end) ", "IrRmZVal")
    wBaseiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZVal/ InCnvFct else iLZVal* InCnvFct end) ", "iLZVal")
    Rep.TxtBase.SetText "(Base Value = Customs)"
    
  End If
  
  If UCase(adc("UoInVal1")) = "A" Then
    wOth = "Actual"
    wOthIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdAVal/ InCnvFct else IdAVal* InCnvFct end) ", "IdAVal")
    wOthIrRmWt = "IrRmAWt"
    wOthIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetAVal/ InCnvFct else IrSetAVal* InCnvFct end) ", "IrSetAVal")
    wOthIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmAVal/ InCnvFct else IrRmAVal* InCnvFct end) ", "IrRmAVal")
    wOthiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLAVal/InCnvFct else iLAVal*InCnvFct end) ", "iLAVal")
    Rep.TxtOth.SetText "(Other Value = Actual)"
    ws_OthAdj = ws_CstAdj    '*** (Jen 2.13)
  ElseIf UCase(adc("UoInVal1")) = "S" Then
    wOth = "Stock"
    wOthIdFOB = "sum(IrStkRt* IrRmAWt) + sum((Case when InMulDiv= 'M' then IrSetAVal/ InCnvFct else IrSetAVal* InCnvFct end)) + " + _
                "IsNull((Select sum((Case when InMulDiv= 'M' then iLAVal/ InCnvFct else iLAVal* InCnvFct end)) from InvLab " + _
                ws_InIlJoin + " where " + IIF(gs_Partition = ctCurrPrtn, " IlPrtKey=IdPrtKey and ", "") + " iLCoCd= IdCoCd and iLTc= IdTc " + _
                "and iLYy= IdYy and iLChr= IdChr and iLNo= IdNo and iLSr= IdSr), 0) "
    wOthIrRmWt = "IrRmAWt"
    wOthIrSetVal = " (Case when InMulDiv= 'M' then IrSetAVal/ InCnvFct else IrSetAVal* InCnvFct end) "
    wOthIrRmVal = "IrStkRt* IrRmAWt"
    wOthiLVal = " (Case when InMulDiv= 'M' then iLAVal/ InCnvFct else iLAVal* InCnvFct end) "
    
    '4.1.4
    If Trim(adc("UoInVal2")) <> "" Then
        wAOthIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal")
        wIOthIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIVal/InCnvFct else IrRmIVal* InCnvFct end) ", "IrRmIVal")
        wZOthIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal* InCnvFct end) ", "IrRmZVal")
        
        wOthIrRmVal = " (Case When RSCtg.PValue1='N' Then " + IIF(UCase(adc("UoInVal2")) = "A", wABaseIrRmVal, _
                                                               IIF(UCase(adc("UoInVal2")) = "I", wIBaseIrRmVal, _
                                                               IIF(UCase(adc("UoInVal2")) = "Z", wZBaseIrRmVal, wOthIrRmVal))) + " Else " + wOthIrRmVal + " End)"
        
        wAOthFob = "sum((Case When RSCtg.PValue1='N' Then IrRmAVal Else IrStkRt* IrRmAWt End) + (Case when InMulDiv= 'M' then IrSetAVal/ InCnvFct else IrSetAVal* InCnvFct end)) " + _
                    "+ IsNull((Select sum((Case when InMulDiv= 'M' then iLAVal/ InCnvFct else iLAVal* InCnvFct end)) from InvLab " + _
                    ws_InIlJoin + " where " + IIF(gs_Partition = ctCurrPrtn, " IlPrtKey=IdPrtKey and ", "") + " iLCoCd= IdCoCd and iLTc= IdTc " + _
                    "and iLYy= IdYy and iLChr= IdChr and iLNo= IdNo and iLSr= IdSr), 0) "
        
        wIOthFob = "sum((Case When RSCtg.PValue1='N' Then IrRmIVal Else IrStkRt*IrRmIWt End)+ (Case when InMulDiv= 'M' then IrSetIVal/ InCnvFct else IrSetIVal* InCnvFct end)) " + _
                    "+ IsNull((Select sum((Case when InMulDiv= 'M' then IlIVal/ InCnvFct else IlIVal* InCnvFct end)) from InvLab " + _
                    ws_InIlJoin + " where " + IIF(gs_Partition = ctCurrPrtn, " IlPrtKey=IdPrtKey and ", "") + " iLCoCd= IdCoCd and iLTc= IdTc " + _
                    "and iLYy= IdYy and iLChr= IdChr and iLNo= IdNo and iLSr= IdSr), 0) "
        
        wZOthFob = "sum((Case When RSCtg.PValue1='N' Then IrRmZVal Else IrStkRt*IrRmZWt End)+ (Case when InMulDiv= 'M' then IrSetZVal/ InCnvFct else IrSetZVal* InCnvFct end)) " + _
                    "+ IsNull((Select sum((Case when InMulDiv= 'M' then IlZVal/ InCnvFct else IlZVal* InCnvFct end)) from InvLab " + _
                    ws_InIlJoin + " where " + IIF(gs_Partition = ctCurrPrtn, " IlPrtKey=IdPrtKey and ", "") + " iLCoCd= IdCoCd and iLTc= IdTc " + _
                    "and iLYy= IdYy and iLChr= IdChr and iLNo= IdNo and iLSr= IdSr), 0) "
        
        wOthIdFOB = IIF(UCase(adc("UoInVal2")) = "A", wABaseFob, _
                                                             IIF(UCase(adc("UoInVal2")) = "I", wIBaseFob, _
                                                             IIF(UCase(adc("UoInVal2")) = "Z", wZBaseFob, wOthIdFOB)))
    End If
    
    
    Rep.TxtOth.SetText "(Other Value = Stock)"
  ElseIf UCase(adc("UoInVal1")) = "I" Then
    wOth = "Invoice"
    wOthIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdIFob/ InCnvFct else IdIFob* InCnvFct end) ", "IdIFob")
    wOthIrRmWt = "IrRmIWt"
    wOthIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIVal/ InCnvFct else IrSetIVal* InCnvFct end) ", "IrSetIVal")
    wOthIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIVal/ InCnvFct else IrRmIVal* InCnvFct end) ", "IrRmIVal")
    wOthiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiVal/ InCnvFct else iLiVal* InCnvFct end) ", "iLiVal")
    Rep.TxtOth.SetText "(Other Value = Invoice)"
  ElseIf UCase(adc("UoInVal1")) = "Z" Then
    wOth = "Custom"
    wOthIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdZFob/ InCnvFct else IdZFob* InCnvFct end) ", "IdZFob")
    wOthIrRmWt = "IrRmZWt"
    wOthIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZVal/ InCnvFct else IrSetZVal* InCnvFct end) ", "IrSetZVal")
    wOthIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZVal/ InCnvFct else IrRmZVal* InCnvFct end) ", "IrRmZVal")
    wOthiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZVal/ InCnvFct else iLZVal* InCnvFct end) ", "iLZVal")
    Rep.TxtOth.SetText "(Other Value = Customs)"
  End If
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " IrDt >= '" + Trim(adc("UODTFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " IrDt <= '" + Trim(adc("UODTTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  '*** (Bef 2.13)
  'wStr1 = "Select " + grpflds + ", " + _
  '        "IdTc as OdTc, IdYy as OdYy, IdChr as OdChr, IdNo as OdNo, IdSr as OdSr, max(IdDmCd) as OdDmCd, " + _
  '        "max(IdQty) as OdOrdQty, " + _
  '        IIF(UCase(adc("UoInVal")) <> "S", "max(" + wBaseIdFOB + ")", wBaseIdFOB) + " as qPcVal1, " + _
  '        IIF(UCase(adc("UoInVal1")) <> "S", "max(" + wOthIdFOB + ") ", wOthIdFOB) + " as qPcVal2, " + _
  '        "sum(case when RmCtg in('G', 'P', 'S') then " + wBaseIrRmVal + " else 0 end) as qMetVal1, " + _
  '        "sum(case when RmCtg in('G', 'P', 'S') then " + wOthIrRmVal + " else 0 end) as qMetVal2, " + _
  '        "sum(case when RmCtg in('D', 'C') then " + wBaseIrRmVal + " else 0 end) as qDCVal1, " + _
  '        "sum(case when RmCtg in('D', 'C') then " + wOthIrRmVal + " else 0 end) as qDCVal2, " + _
  '        "sum(case when RmCtg in('X') then " + wBaseIrRmVal + " else 0 end) as qXVal1, " + _
  '        "sum(case when RmCtg in('X') then " + wOthIrRmVal + " else 0 end) as qXVal2, " + _
  '        "sum(" + wBaseIrSetVal + ") as qSetVal1, sum(" + wOthIrSetVal + ") as qSetVal2, " + _
  '        "(Select IsNull(sum(" + wBaseiLVal + "),0) from InvLab where iLCoCd= IdCoCd and iLTc= IdTc " + _
  '        "and iLYy= IdYy and iLChr= IdChr and iLNo= IdNo and iLSr= IdSr) as qLabVal1, " + _
  '        "(Select IsNull(sum(" + wOthiLVal + "),0) from InvLab where iLCoCd= IdCoCd and iLTc= IdTc " + _
  '        "and iLYy= IdYy and iLChr= IdChr and iLNo= IdNo and iLSr= IdSr) as qLabVal2, "
  '*** (Bef 2.13)
  
  ' ***** Manali 3.03 -28/06/08 - Pd Mod - 'L' added
  wStr1 = " Select " + grpflds + ", " + _
          " IdTc as OdTc, IdYy as OdYy, IdChr as OdChr, IdNo as OdNo, IdSr as OdSr, max(IdDmCd) as OdDmCd, " + _
          " max(IdQty) as OdOrdQty, " + _
          IIF(UCase(adc("UoInVal")) <> "S", "max(" + wBaseIdFOB + ")", wBaseIdFOB) + IIF(Trim(ws_BaseAdj) <> "", " + " + _
          "   sum((case when RmCtg in('G', 'P', 'S', 'L') then " + wBaseIrRmVal + " else 0 end)" + ws_BaseAdj + ") - " + _
          "   sum((case when RmCtg in('G', 'P', 'S', 'L') then " + wBaseIrRmVal + " else 0 end)) ", " ") + " as qPcVal1, " + _
          IIF(UCase(adc("UoInVal1")) <> "S", "max(" + wOthIdFOB + ") ", wOthIdFOB) + IIF(Trim(ws_OthAdj) <> "", " + " + _
          "   sum((case when RmCtg in('G', 'P', 'S', 'L') then " + wOthIrRmVal + " else 0 end)" + ws_OthAdj + ") - " + _
          "   sum((case when RmCtg in('G', 'P', 'S', 'L') then " + wOthIrRmVal + " else 0 end)) ", " ") + " as qPcVal2, " + _
          " sum((case when RmCtg in('G', 'P', 'S', 'L') then " + wBaseIrRmVal + " else 0 end)" + ws_BaseAdj + ") as qMetVal1, " + _
          " sum((case when RmCtg in('G', 'P', 'S', 'L') then " + wOthIrRmVal + " else 0 end)" + ws_OthAdj + ") as qMetVal2, " + _
          " sum(case when RmCtg in('D', 'C') then " + wBaseIrRmVal + " else 0 end) as qDCVal1, " + _
          " sum(case when RmCtg in('D', 'C') then " + wOthIrRmVal + " else 0 end) as qDCVal2, " + _
          " sum(case when RmCtg in('X') then " + wBaseIrRmVal + " else 0 end) as qXVal1, " + _
          " sum(case when RmCtg in('X') then " + wOthIrRmVal + " else 0 end) as qXVal2, " + _
          " sum(" + wBaseIrSetVal + ") as qSetVal1, sum(" + wOthIrSetVal + ") as qSetVal2, " + _
          " IsNull((Select sum(" + wBaseiLVal + ") from InvLab " + ws_InIlJoin + " " + _
          " where iLCoCd= Id1.IdCoCd and iLTc= Id1.IdTc and iLYy= Id1.IdYy and iLChr= Id1.IdChr " + _
          " and iLNo= Id1.IdNo and iLSr= Id1.IdSr " + IIF(gs_Partition = ctCurrPrtn, " And iLPrtKey=Id1.IdPrtKey ", "") + " ), 0) as qLabVal1, " + _
          " IsNull((Select sum(" + wOthiLVal + ") from InvLab " + ws_InIlJoin + " " + _
          " where iLCoCd= Id1.IdCoCd and iLTc= Id1.IdTc and iLYy= Id1.IdYy and iLChr= Id1.IdChr " + _
          " and iLNo= Id1.IdNo and iLSr= Id1.IdSr " + IIF(gs_Partition = ctCurrPrtn, " And iLPrtKey=Id1.IdPrtKey ", "") + " ), 0) as qLabVal2, "
  
  ' Zubin 213 (ws_CurJoin added)
  '****** (Jen 2.13) Added the OrdMst join ******'
  ' ***** Manali 3.03 -28/06/08 - Pd Mod - 'L' added
  wStr2 = " sum(" + wBaseIrRmWt + "/(case when RmCtg in('C', 'D') then 5 else 1 end)) as qPcWt1, " + _
          " sum(" + wOthIrRmWt + "/(case when RmCtg in('C', 'D') then 5 else 1 end)) as qPcWt2, " + _
          " sum(case when RmCtg in('G', 'P', 'S', 'L') then " + wBaseIrRmWt + " else 0 end) as qMetWt1, " + _
          " sum(case when RmCtg in('G', 'P', 'S', 'L') then " + wOthIrRmWt + " else 0 end) as qMetWt2, " + _
          " sum(case when RmCtg in('D', 'C') then " + wBaseIrRmWt + " else 0 end) as qDCWt1, " + _
          " sum(case when RmCtg in('D', 'C') then " + wOthIrRmWt + " else 0 end) as qDCWt2, " + _
          " sum(case when RmCtg in('X') then " + wBaseIrRmWt + " else 0 end) as qXWt1, " + _
          " sum(case when RmCtg in('X') then " + wOthIrRmWt + " else 0 end) as qXWt2 " + _
          " From InvDsg Id1 join DsgMst on DmIdno=IdDmIdNo And DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + _
          "  Join InvRm on IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy and IrChr= IdChr " + _
          "   and IrNo= IdNo and IrSr= IdSr " + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey=IdPrtKey ", "") + _
          "  Join RmMst on RmCd= IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + _
          "  Join Param RSCtg On PTyp='RMSCTG' And PMCd=RmCtg And PSCd=RmSCtg " + _
          "  Join InvHd Ih1 on InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr " + _
          "   and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IdPrtKey ", "") + _
          "  Join OrdMst on  " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey=IdPrtKey and ", "") + " OmCoCd= IdCoCd and OmTc= IdExpOdTc and OmYy= IdExpOdYy and OmChr= IdExpOdChr " + _
          "  and OmNo= IdExpOdNo " + ws_CurJoin + wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey='" + ctCurrPrtn + "' ", "") + _
          " Group By IdPrtKey, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr "
 
  wSqlStrg = wStr1 + wStr2
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt
  Rep.hVchSr.SetText "Invoice Vch"
  
  If adc("UoYN") = "Y" Then
    'Piece
    Rep.TxtHead.SetText "Invoice Profit Margin (Value)"
    'REP.h1Pc1.SetText wBase + Chr(13) + "Pc Val": REP.h2Pc1.SetText wBase + Chr(13) + "Pc Val"
    Rep.h1Pc1.SetText " " + wBase + " Pieces Val": Rep.h2Pc1.SetText " " + wBase + " Pieces Val"
    Rep.h1Pc2.SetText " " + wOth + " Pieces Val": Rep.h2Pc2.SetText " " + wOth + " Pieces Val"
    Rep.h1Met1.SetText " " + wBase + " Metal Val": Rep.h2Met1.SetText " " + wBase + " Metal Val"
    Rep.h1Met2.SetText " " + wOth + " Metal Val": Rep.h2Met2.SetText " " + wOth + " Metal Val"
    Rep.h1DC1.SetText wBase + " Dia,Cs Val": Rep.h2DC1.SetText wBase + " Dia,Cs Val"
    Rep.h1DC2.SetText wOth + " Dia,Cs Val": Rep.h2DC2.SetText wOth + " Dia,Cs Val"
    Rep.h1X1.SetText "  " + wBase + " Acc Val": Rep.h2X1.SetText "  " + wBase + " Acc Val"
    Rep.h1X2.SetText "  " + wOth + " Acc Val": Rep.h2X2.SetText "  " + wOth + " Acc Val"
    Rep.h1Lab1.SetText "  " + wBase + " Lab Val": Rep.h2Lab1.SetText "  " + wBase + " Lab Val"
    Rep.h1Lab2.SetText "  " + wOth + " Lab Val": Rep.h2Lab2.SetText "  " + wOth + " Lab Val"
    Rep.FldPcWt1.Suppress = True: Rep.FldPcWt2.Suppress = True: Rep.FldPcWtPrc.Suppress = True
    Rep.FldMetWt1.Suppress = True: Rep.FldMetWt2.Suppress = True: Rep.FldMetWtPrc.Suppress = True
    Rep.FldDCWt1.Suppress = True: Rep.FldDCWt2.Suppress = True: Rep.FldDCWtPrc.Suppress = True
    Rep.FldXWt1.Suppress = True: Rep.FldXWt2.Suppress = True: Rep.FldXWtPrc.Suppress = True
    Rep.G1FWt.Suppress = True: Rep.G2FWt.Suppress = True: Rep.G3FWt.Suppress = True
    Rep.RepFtWt.Suppress = True
  ElseIf adc("UoYN") = "N" Then
    Rep.TxtHead.SetText "Invoice Differential Weights"
    Rep.h1Pc1.SetText " " + wBase + " Pieces Wt": Rep.h2Pc1.SetText " " + wBase + " Pieces Wt"
    Rep.h1Pc2.SetText " " + wOth + " Pieces Wt": Rep.h2Pc2.SetText " " + wOth + " Pieces Wt"
    Rep.h1Met1.SetText " " + wBase + " Metal Wt": Rep.h2Met1.SetText " " + wBase + " Metal Wt"
    Rep.h1Met2.SetText " " + wOth + " Metal Wt": Rep.h2Met2.SetText " " + wOth + " Metal Wt"
    Rep.h1DC1.SetText wBase + " Dia,Cs Wt": Rep.h2DC1.SetText wBase + " Dia,Cs Wt"
    Rep.h1DC2.SetText wOth + " Dia,Cs Wt": Rep.h2DC2.SetText wOth + " Dia,Cs Wt"
    Rep.h1X1.SetText "  " + wBase + " Acc Wt": Rep.h2X1.SetText "  " + wBase + " Acc Wt"
    Rep.h1X2.SetText "  " + wOth + " Acc Wt": Rep.h2X2.SetText "  " + wOth + " Acc Wt"
    Rep.h1Lab1.SetText " ": Rep.h2Lab1.SetText " "
    Rep.h1Lab2.SetText " ": Rep.h2Lab2.SetText " "
    Rep.h1LabPrc.SetText " ": Rep.h2LabPrc.SetText " "
    Rep.FldPcVal1.Suppress = True: Rep.FldPcVal2.Suppress = True: Rep.FldPcValPrc.Suppress = True
    Rep.FldMetVal1.Suppress = True: Rep.FldMetVal2.Suppress = True: Rep.FldMetValPrc.Suppress = True
    Rep.FldDCVal1.Suppress = True: Rep.FldDCVal2.Suppress = True: Rep.FldDCValPrc.Suppress = True
    Rep.FldXVal1.Suppress = True: Rep.FldXVal2.Suppress = True: Rep.FldXValPrc.Suppress = True
    Rep.FldLabVal1.Suppress = True: Rep.FldLabVal2.Suppress = True: Rep.FldLabValPrc.Suppress = True
    Rep.G1F.Suppress = True: Rep.G2F.Suppress = True: Rep.G3F.Suppress = True
    Rep.RepFt.Suppress = True
  End If
 
  'KC.3- StkRt-when Show Stock Rate=YES then Stock Rate is visible in report
  'if it is NO then Stock Rate is invisible in report
 
  If UCase(adc("UOINVAL")) = "S" And UCase(adc("UoYN")) = "Y" Then
  If adc("UoYN2") = "Y" Then
    Rep.FldPcVal1.Suppress = False
    Rep.FldPcVal13a.Suppress = False:    Rep.FldPcVal12a.Suppress = False
    Rep.FldPcVal11a.Suppress = False:    Rep.FldPcVal1A.Suppress = False
    Rep.FldMetVal1.Suppress = False
    Rep.FldMetVal13a.Suppress = False:   Rep.FldMetVal12a.Suppress = False
    Rep.FldMetVal11a.Suppress = False:   Rep.FldMetVal1A.Suppress = False
    Rep.FldDCVal1.Suppress = False
    Rep.FldDCVal13a.Suppress = False:    Rep.FldDCVal12a.Suppress = False
    Rep.FldDCVal11a.Suppress = False:    Rep.FldDCVal1A.Suppress = False
    Rep.FldXVal1.Suppress = False
    Rep.FldXVal13a.Suppress = False:     Rep.FldXVal12a.Suppress = False
    Rep.FldXVal11a.Suppress = False:     Rep.FldXVal1A.Suppress = False
    Rep.FldLabVal1.Suppress = False
    Rep.FldLabVal3a.Suppress = False:    Rep.FldLabVal2a.Suppress = False
    Rep.FldLabVal1a.Suppress = False:    Rep.FldLabValA.Suppress = False
  Else
    Rep.h1Pc1.SetText "":   Rep.h2Pc1.SetText "":   Rep.FldPcVal1.Suppress = True
    Rep.FldPcVal13a.Suppress = True:    Rep.FldPcVal12a.Suppress = True
    Rep.FldPcVal11a.Suppress = True:    Rep.FldPcVal1A.Suppress = True
    Rep.h1Met1.SetText "":  Rep.h2Met1.SetText "":  Rep.FldMetVal1.Suppress = True
    Rep.FldMetVal13a.Suppress = True:   Rep.FldMetVal12a.Suppress = True
    Rep.FldMetVal11a.Suppress = True:   Rep.FldMetVal1A.Suppress = True
    Rep.h1DC1.SetText "":    Rep.h2DC1.SetText "":    Rep.FldDCVal1.Suppress = True
    Rep.FldDCVal13a.Suppress = True:    Rep.FldDCVal12a.Suppress = True
    Rep.FldDCVal11a.Suppress = True:    Rep.FldDCVal1A.Suppress = True
    Rep.h1X1.SetText "":    Rep.h2X1.SetText "":    Rep.FldXVal1.Suppress = True
    Rep.FldXVal13a.Suppress = True:     Rep.FldXVal12a.Suppress = True
    Rep.FldXVal11a.Suppress = True:     Rep.FldXVal1A.Suppress = True
    Rep.h1Lab1.SetText "":  Rep.h2Lab1.SetText "":    Rep.FldLabVal1.Suppress = True
    Rep.FldLabVal3a.Suppress = True:    Rep.FldLabVal2a.Suppress = True
    Rep.FldLabVal1a.Suppress = True:    Rep.FldLabValA.Suppress = True
  End If
  End If
  'If the Other value and Show value =yes then the above logic applies
  If UCase(adc("UOINVAL1")) = "S" And UCase(adc("UoYN")) = "Y" Then
  If adc("UoYN2") = "Y" Then
    Rep.FldPcVal2.Suppress = False
    Rep.FldPcVal23a.Suppress = False:    Rep.FldPcVal22a.Suppress = False
    Rep.FldPcVal21a.Suppress = False:    Rep.FldPcVal2A.Suppress = False
    Rep.FldMetVal2.Suppress = False
    Rep.FldMetVal23a.Suppress = False:   Rep.FldMetVal22a.Suppress = False
    Rep.FldMetVal21a.Suppress = False:   Rep.FldMetVal2A.Suppress = False
    Rep.FldDCVal2.Suppress = False
    Rep.FldDCVal23a.Suppress = False:    Rep.FldDCVal22a.Suppress = False
    Rep.FldDCVal21a.Suppress = False:    Rep.FldDCVal2A.Suppress = False
    Rep.FldXVal2.Suppress = False
    Rep.FldXVal23a.Suppress = False:     Rep.FldXVal22a.Suppress = False
    Rep.FldXVal21a.Suppress = False:     Rep.FldXVal2A.Suppress = False
    Rep.FldLabVal2.Suppress = False
    Rep.FldLabVal23a.Suppress = False:    Rep.FldLabVal22a.Suppress = False
    Rep.FldLabVal21a.Suppress = False:    Rep.FldLabValA2.Suppress = False
  Else
    Rep.h1Pc2.SetText "":   Rep.h2Pc2.SetText "":   Rep.FldPcVal2.Suppress = True
    Rep.FldPcVal23a.Suppress = True:    Rep.FldPcVal22a.Suppress = True
    Rep.FldPcVal21a.Suppress = True:    Rep.FldPcVal2A.Suppress = True
    Rep.h1Met2.SetText "":  Rep.h2Met2.SetText "":  Rep.FldMetVal2.Suppress = True
    Rep.FldMetVal23a.Suppress = True:   Rep.FldMetVal22a.Suppress = True
    Rep.FldMetVal21a.Suppress = True:   Rep.FldMetVal2A.Suppress = True
    Rep.h1DC2.SetText "":    Rep.h2DC2.SetText "":  Rep.FldDCVal2.Suppress = True
    Rep.FldDCVal23a.Suppress = True:    Rep.FldDCVal22a.Suppress = True
    Rep.FldDCVal21a.Suppress = True:    Rep.FldDCVal2A.Suppress = True
    Rep.h1X2.SetText "":    Rep.h2X2.SetText "":    Rep.FldXVal2.Suppress = True
    Rep.FldXVal23a.Suppress = True:     Rep.FldXVal22a.Suppress = True
    Rep.FldXVal21a.Suppress = True:     Rep.FldXVal2A.Suppress = True
    Rep.h1Lab2.SetText "":  Rep.h2Lab2.SetText "":  Rep.FldLabVal2.Suppress = True
    Rep.FldLabVal23a.Suppress = True:    Rep.FldLabVal22a.Suppress = True
    Rep.FldLabVal21a.Suppress = True:    Rep.FldLabValA2.Suppress = True
  End If
  End If

  CRV_REP.DisplayGroupTree = False

Call DispReport(adc) '****** Sachin 3.02 - Display Report ******

End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  ' Zubin 213
  Select Case UCase(IdName)
  Case Is = UCase("UoCurCdFr")
    If UCase(adc("UoInVal")) = "S" Or UCase(adc("UoInVal1")) = "S" Then adc("UoCurCdFr") = "": Exit Sub
  End Select
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
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  Set adc.FirNKeyCtl = adc("UoInTcFr") 'adc("UoCoCdFr")
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_InTcTyp = "IN"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_InTcTyp = "IN"
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  
  gb_CoCdFor = True         '*** (Jen 2.14 Next Patch 1)
  gs_InTcTyp = "IN"
  
  ' Zubin 213
  Select Case UCase(IdName)
  Case Is = UCase("UoCurCdFr")
    If UCase(adc("UoInVal")) = "S" Or UCase(adc("UoInVal1")) = "S" Then Cancel = True: Exit Sub
    
  Case Is = UCase("UoInVal2")     '4.1.4
      If adc.Mode = xNorm Then _
        If UCase(adc("UoInVal")) <> UCase("S") Then Cancel = True: ErrMsg = "Can Enter Only if Value Type is 'STKVAL'": Exit Sub
        
  End Select
  ' Zubin 213
  SetProp Me, IdName, When
End Sub

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptCstRt_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN1") = "Y"
Case Is = 1
  adc("UoYN1") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptCstRt_GotFocus(Index As Integer)
  DispMsg FraOptCstRt.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptValWt_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN") = "Y"
Case Is = 1
  adc("UoYN") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptValWt_GotFocus(Index As Integer)
  DispMsg FraOptValWt.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    adc("UOUSRCD") = adc.UsrCd
    adc("UOMNUCD") = adc.MenuCd
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
 adc("UoCoCdFr") = gs_CoCd
 
 ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
 If adc("UoYN") = "Y" Then OptValWt(0).Value = True
 If adc("UoYN") = "N" Then OptValWt(1).Value = True
 If adc("UoYN1") = "Y" Then OptCstRt(0).Value = True
 If adc("UoYN1") = "N" Then OptCstRt(1).Value = True
 ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  'StkRt-Set Show StockRate option is No
  adc("UoYN2") = "N"
  
End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                                                    "HCd = '" + ctSelfCmCd + "' ")
  Rep.wCoCd.SetText ws_HName

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
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  End Select
  End With
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

Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub


