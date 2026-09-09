VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpBrkMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Breakage, Missing Report"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   12810
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   9690
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4492
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4665
      TabIndex        =   87
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   88
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
         TabIndex        =   89
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
      Left            =   7920
      TabIndex        =   86
      Top             =   9660
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
      TabIndex        =   83
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
      Height          =   9915
      Left            =   0
      TabIndex        =   84
      Top             =   60
      Width           =   15315
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   90
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpBrkMsg.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpBrkMsg.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   93
            Top             =   360
            Width           =   15105
            Begin VB.OptionButton OptLoc 
               Caption         =   "Process loc"
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
               Left            =   5340
               TabIndex        =   138
               ToolTipText     =   "Select Production or Process"
               Top             =   2280
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
               Left            =   2640
               TabIndex        =   137
               ToolTipText     =   "Select Production or Process"
               Top             =   2280
               Value           =   -1  'True
               Width           =   2085
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show All Prd Lines          "
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
               Left            =   60
               TabIndex        =   80
               Tag             =   "ShowPrdLine"
               ToolTipText     =   "Check To Show All Production Lines"
               Top             =   7845
               Width           =   2790
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Box Details            "
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
               TabIndex        =   78
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show box Details"
               Top             =   7560
               Width           =   2790
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Value                     "
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
               TabIndex        =   70
               Tag             =   "ShowVal"
               ToolTipText     =   "Check To Show Value"
               Top             =   6990
               Width           =   2805
            End
            Begin VB.Frame FraOptGrpBy 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2670
               TabIndex        =   134
               ToolTipText     =   "Select If To Group By Shape Or RmCd"
               Top             =   1140
               Width           =   2505
               Begin VB.OptionButton OptGrpBy 
                  Caption         =   "Shape"
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
                  TabIndex        =   9
                  ToolTipText     =   " Select If To Group By Shape Or RmCd"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptGrpBy 
                  Caption         =   "RmCd"
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
                  TabIndex        =   10
                  ToolTipText     =   " Select If To Group By Shape Or RmCd"
                  Top             =   0
                  Width           =   945
               End
            End
            Begin VB.Frame FraOptBrkMsg 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2670
               TabIndex        =   133
               ToolTipText     =   "Select If To Show Report For Broken Or Missing Or Both"
               Top             =   840
               Width           =   4035
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
                  TabIndex        =   6
                  ToolTipText     =   "Select If To Show Report For Broken Or Missing Or Both"
                  Top             =   0
                  Width           =   1095
               End
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
                  TabIndex        =   7
                  ToolTipText     =   "Select If To Show Report For Broken Or Missing Or Both"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.Frame FraPwd 
               Height          =   870
               Left            =   9630
               TabIndex        =   103
               Top             =   7320
               Width           =   5220
               Begin VB.TextBox TxtPwd 
                  Height          =   285
                  IMEMode         =   3  'DISABLE
                  Left            =   2730
                  MaxLength       =   30
                  PasswordChar    =   "*"
                  TabIndex        =   74
                  Top             =   405
                  Width           =   2340
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
                  Width           =   5220
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Password To View Value"
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
                  Left            =   60
                  TabIndex        =   104
                  Top             =   405
                  Width           =   2715
               End
            End
            Begin MwfCtl.MWCTL_BTN1 CmdPwd 
               Height          =   405
               Left            =   10560
               TabIndex        =   73
               ToolTipText     =   "Enter Password to Edit Design"
               Top             =   6960
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
               Index           =   47
               Left            =   6765
               TabIndex        =   4
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
               Index           =   48
               Left            =   6765
               TabIndex        =   8
               ToolTipText     =   "Group On Shape Or Raw Material Code?"
               Top             =   1125
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOSHPCD"
               IdName          =   "UOSHPCD"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2940
               TabIndex        =   77
               ToolTipText     =   "Show Box Detail (Yes / No)"
               Top             =   7560
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
               Index           =   3511
               Left            =   5325
               TabIndex        =   68
               ToolTipText     =   "Enter To Stock Rate"
               Top             =   6705
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
               Left            =   2655
               TabIndex        =   67
               ToolTipText     =   "Enter From Stock Rate"
               Top             =   6705
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
               Left            =   2925
               TabIndex        =   69
               ToolTipText     =   "Show Value(Y/N)"
               Top             =   6990
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               ReCalcParent    =   "UOPMCD"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   5325
               TabIndex        =   71
               ToolTipText     =   "Enter Value Type"
               Top             =   6990
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCD"
               ReCalcOn        =   "UOYN1"
               IdName          =   "UOPMCD"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   2655
               TabIndex        =   75
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   7275
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
               Index           =   53
               Left            =   5325
               TabIndex        =   76
               ToolTipText     =   "Enter Bag Character"
               Top             =   7275
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
               Index           =   3
               Left            =   2655
               TabIndex        =   3
               ToolTipText     =   "Enter Raw Material Category"
               Top             =   555
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "RmCtg ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   2655
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   255
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "TCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   5325
               TabIndex        =   12
               ToolTipText     =   "Enter To Transaction Date"
               Top             =   1410
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
               Index           =   1
               Left            =   2655
               TabIndex        =   11
               ToolTipText     =   "Enter From Transaction Date"
               Top             =   1410
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
               Index           =   5
               Left            =   5325
               TabIndex        =   65
               ToolTipText     =   "Enter To Raw Material Code "
               Top             =   6390
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "RmCd<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   7980
               TabIndex        =   66
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   6390
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "RmCd In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   2655
               TabIndex        =   64
               ToolTipText     =   "Enter From Raw Material Code "
               Top             =   6390
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "RmCd>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   5325
               TabIndex        =   59
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   5775
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
               Index           =   9
               Left            =   7980
               TabIndex        =   60
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   5775
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "RmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   2655
               TabIndex        =   58
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   5775
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
               Data            =   "ABCDW"
               CmpStr          =   "RmSCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   5325
               TabIndex        =   20
               ToolTipText     =   "Enter To By Location"
               Top             =   2640
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   7980
               TabIndex        =   21
               ToolTipText     =   "Enter By Location Selection"
               Top             =   2640
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   2655
               TabIndex        =   19
               ToolTipText     =   "Enter From By Location"
               Top             =   2640
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   5325
               TabIndex        =   23
               ToolTipText     =   "Enter To By Worker"
               Top             =   2925
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
               Index           =   15
               Left            =   7980
               TabIndex        =   24
               ToolTipText     =   "Enter By Worker Selection"
               Top             =   2925
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOWRKSEL"
               IdName          =   "UOWRKSEL"
               CmpStr          =   "Td.TdByWrk In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   2655
               TabIndex        =   22
               ToolTipText     =   "Enter From By Worker"
               Top             =   2925
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
               Index           =   20
               Left            =   6105
               TabIndex        =   41
               ToolTipText     =   "Enter To Order Year"
               Top             =   4350
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
               Index           =   17
               Left            =   3435
               TabIndex        =   38
               ToolTipText     =   "Enter From Order Year"
               Top             =   4350
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
               Index           =   16
               Left            =   2655
               TabIndex        =   37
               ToolTipText     =   "Enter From Order Tc"
               Top             =   4350
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
               Index           =   18
               Left            =   3855
               TabIndex        =   39
               ToolTipText     =   "Enter From Order Character"
               Top             =   4350
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
               Index           =   19
               Left            =   5325
               TabIndex        =   40
               ToolTipText     =   "Enter To Order Tc"
               Top             =   4350
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
               Index           =   21
               Left            =   6525
               TabIndex        =   42
               ToolTipText     =   "Enter To Order Character"
               Top             =   4350
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
               Index           =   22
               Left            =   2655
               TabIndex        =   43
               ToolTipText     =   "Enter From Order Number"
               Top             =   4635
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
               Index           =   24
               Left            =   5325
               TabIndex        =   45
               ToolTipText     =   "Enter To Order Number"
               Top             =   4635
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
               Index           =   25
               Left            =   6210
               TabIndex        =   46
               ToolTipText     =   "Enter To Order Serial"
               Top             =   4635
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
               Left            =   3540
               TabIndex        =   44
               ToolTipText     =   "Enter From Order Serial"
               Top             =   4635
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
               Index           =   33
               Left            =   5325
               TabIndex        =   26
               ToolTipText     =   "Enter To Design Category"
               Top             =   3210
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
               Index           =   34
               Left            =   7980
               TabIndex        =   27
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   3210
               Width           =   6750
               _ExtentX        =   11906
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
               Left            =   2655
               TabIndex        =   25
               ToolTipText     =   "Enter From Design Category"
               Top             =   3210
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
               Index           =   27
               Left            =   5325
               TabIndex        =   32
               ToolTipText     =   "Enter To Design Code"
               Top             =   3780
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
               Index           =   28
               Left            =   7980
               TabIndex        =   33
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   3780
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "BOdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   2655
               TabIndex        =   31
               ToolTipText     =   "Enter From Design Code"
               Top             =   3780
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
               Index           =   30
               Left            =   5325
               TabIndex        =   29
               ToolTipText     =   "Enter To Design Production Category"
               Top             =   3495
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
               Index           =   31
               Left            =   7980
               TabIndex        =   30
               ToolTipText     =   "Enter Design Production Category Selection"
               Top             =   3495
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMPRDCTGSEL"
               IdName          =   "UODMPRDCTGSEL"
               CmpStr          =   "DmPrdCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   2655
               TabIndex        =   28
               ToolTipText     =   "Enter From Design Production Category"
               Top             =   3495
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
               Index           =   36
               Left            =   5325
               TabIndex        =   35
               ToolTipText     =   "Enter To Karat"
               Top             =   4065
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
               Index           =   35
               Left            =   2655
               TabIndex        =   34
               ToolTipText     =   "Enter From Karat"
               Top             =   4065
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
               Index           =   37
               Left            =   7980
               TabIndex        =   36
               ToolTipText     =   "Enter Karat Selection"
               Top             =   4065
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "OdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   5325
               TabIndex        =   14
               ToolTipText     =   "Enter To Production Sequence"
               Top             =   1695
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
               CmpStr          =   "OdPrdSeq <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   7980
               TabIndex        =   15
               ToolTipText     =   "Enter Production Sequence Selection"
               Top             =   1695
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPMCDSEL"
               IdName          =   "UOPMCDSEL"
               CmpStr          =   "OdPrdSeq In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   2655
               TabIndex        =   13
               ToolTipText     =   "Enter From Productiopn Sequence"
               Top             =   1695
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
               CmpStr          =   "OdPrdSeq >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   5325
               TabIndex        =   48
               ToolTipText     =   "Enter To Customer"
               Top             =   4920
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
               Index           =   43
               Left            =   7980
               TabIndex        =   49
               ToolTipText     =   "Enter Customer Selection"
               Top             =   4920
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   2655
               TabIndex        =   47
               ToolTipText     =   "Enter From Customer"
               Top             =   4920
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
               Index           =   45
               Left            =   5325
               TabIndex        =   51
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   5205
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
               Index           =   46
               Left            =   7980
               TabIndex        =   52
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   5205
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   2655
               TabIndex        =   50
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   5205
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
               Index           =   54
               Left            =   2940
               TabIndex        =   79
               ToolTipText     =   "Show All Production Lines? (Enter 'Y' for All Or 'N' For Only Production Lines With Brk/ Msg)"
               Top             =   7845
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               ReCalcParent    =   "UOBCHRFR"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   73
               Left            =   5325
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
               CmpStr          =   "TCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   7980
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   240
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "TCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   5325
               TabIndex        =   17
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1980
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDTO"
               IdName          =   "UOPRNTCDTO"
               CmpStr          =   "Lc.LocPrntCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   7980
               TabIndex        =   18
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1980
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
               CmpStr          =   "Lc.LocPrntCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   2655
               TabIndex        =   16
               ToolTipText     =   "Enter From Parent Code"
               Top             =   1980
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDFR"
               IdName          =   "UOPRNTCDFR"
               CmpStr          =   "Lc.LocPrntCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   69
               Left            =   3945
               TabIndex        =   54
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   5490
               Width           =   1315
               _ExtentX        =   2328
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   6615
               TabIndex        =   56
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   5490
               Width           =   1315
               _ExtentX        =   2328
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   7980
               TabIndex        =   57
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   5490
               Visible         =   0   'False
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "OdPoNo in"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   5325
               TabIndex        =   62
               ToolTipText     =   "Enter To Parent Code"
               Top             =   6080
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
               Index           =   59
               Left            =   2655
               TabIndex        =   61
               ToolTipText     =   "Enter From Parent Code"
               Top             =   6080
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
               Index           =   72
               Left            =   7980
               TabIndex        =   63
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   6080
               Width           =   6750
               _ExtentX        =   11906
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
               CmpStr          =   "RmParentCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   8955
               TabIndex        =   72
               ToolTipText     =   "Enter Other Value Type"
               Top             =   6990
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCD1"
               IdName          =   "UOPMCD1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   90
               Left            =   8040
               TabIndex        =   139
               ToolTipText     =   "Show Production or Process Location"
               Top             =   2400
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
               Index           =   61
               Left            =   2655
               TabIndex        =   53
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   5490
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
               Index           =   62
               Left            =   5325
               TabIndex        =   55
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   5490
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
               Index           =   20
               Left            =   6555
               TabIndex        =   141
               Top             =   5490
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
               Left            =   3885
               TabIndex        =   140
               Top             =   5490
               Width           =   105
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Oth Value Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   7320
               TabIndex        =   136
               ToolTipText     =   "Location"
               Top             =   6990
               Width           =   1605
            End
            Begin VB.Label LblParentCd 
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
               Left            =   90
               TabIndex        =   135
               Top             =   6080
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
               Index           =   38
               Left            =   90
               TabIndex        =   132
               Top             =   5490
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
               Left            =   90
               TabIndex        =   131
               Top             =   2010
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
               Left            =   90
               TabIndex        =   130
               ToolTipText     =   "Location"
               Top             =   7275
               Width           =   1035
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Show All Prd Lines"
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
               Left            =   90
               TabIndex        =   129
               Top             =   7845
               Visible         =   0   'False
               Width           =   1965
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
               Left            =   90
               TabIndex        =   127
               Top             =   5205
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
               TabIndex        =   126
               Top             =   4920
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
               Index           =   18
               Left            =   90
               TabIndex        =   125
               Top             =   1695
               Width           =   1485
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
               TabIndex        =   124
               Top             =   4065
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
               TabIndex        =   123
               Top             =   3780
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
               TabIndex        =   122
               Top             =   3210
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
               Index           =   4
               Left            =   90
               TabIndex        =   121
               Top             =   3495
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
               Left            =   90
               TabIndex        =   120
               Top             =   4350
               Width           =   1725
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
               Left            =   90
               TabIndex        =   119
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
               Index           =   24
               Left            =   6465
               TabIndex        =   118
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
               Index           =   25
               Left            =   3795
               TabIndex        =   117
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
               Index           =   26
               Left            =   6045
               TabIndex        =   116
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
               Index           =   27
               Left            =   3375
               TabIndex        =   115
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
               Index           =   31
               Left            =   3480
               TabIndex        =   114
               Top             =   4635
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
               Left            =   6150
               TabIndex        =   113
               Top             =   4635
               Width           =   105
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
               Left            =   90
               TabIndex        =   112
               Top             =   2925
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               Index           =   15
               Left            =   90
               TabIndex        =   111
               Top             =   2640
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
               Left            =   90
               TabIndex        =   110
               Top             =   6390
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
               Left            =   90
               TabIndex        =   109
               Top             =   5775
               Width           =   1425
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
               Index           =   8
               Left            =   90
               TabIndex        =   108
               Top             =   1410
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
               Index           =   12
               Left            =   90
               TabIndex        =   107
               Top             =   240
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
               Left            =   90
               TabIndex        =   106
               ToolTipText     =   "Location"
               Top             =   525
               Width           =   1605
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
               Left            =   4170
               TabIndex        =   105
               Top             =   7275
               Width           =   885
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
               Left            =   4170
               TabIndex        =   102
               ToolTipText     =   "Location"
               Top             =   6990
               Width           =   1245
            End
            Begin VB.Label LblShowVal 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Value"
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
               Left            =   90
               TabIndex        =   101
               Top             =   6990
               Visible         =   0   'False
               Width           =   1455
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
               Left            =   90
               TabIndex        =   100
               Top             =   6705
               Width           =   1545
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
               Index           =   16
               Left            =   90
               TabIndex        =   99
               ToolTipText     =   "Location"
               Top             =   7560
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.Label LblUoShpCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Group By"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   98
               Top             =   1125
               Width           =   1605
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
               Index           =   10
               Left            =   90
               TabIndex        =   97
               ToolTipText     =   "Location"
               Top             =   840
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
               Left            =   2685
               TabIndex        =   96
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
               Left            =   5355
               TabIndex        =   95
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
               Left            =   8010
               TabIndex        =   94
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
            TabIndex        =   91
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   92
               Top             =   360
               Width           =   8535
               _ExtentX        =   15055
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpBrkMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepBrkMsg
Dim Rep1 As New EmrRepBrkMsgPrd
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
'*** Jay 2.14 *** (SEO Change)
Dim ms_PwdFlag As Boolean
'*** Jay 2.14 *** (SEO Change)

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  'GRP_REP.Add "Process", Array("Td.TdByLoc", "max(Td.TdByLoc)"), "", "TdByLoc", "hTdByLoc", "", "", "", Array("(Select LocDesc From Loc Where LocTyp= 'P' and LocCd= Td.TdByLoc) ", "(Select LocDesc From Loc Where LocTyp= 'P' and LocCd= max(Td.TdByLoc)) ")
  'GRP_REP.Add "(Process)Worker", Array("Td.TdByLoc+Td.TdByWrk", "max(Td.TdByLoc)+max(Td.TdByWrk)"), Array("'('+ RTrim(Td.TdByLoc)+ ') '+ Td.TdByWrk", "'('+ RTrim(max(Td.TdByLoc))+ ') '+ max(Td.TdByWrk)"), "TdByLoc,TdByWrk", "hTdByLoc,hTdByWrk", "Process", "", "", Array("(Select PDesc From Param Where PTyp= 'WORK' and PMCd= Td.TdByWrk) ", "(Select PDesc From Param Where PTyp= 'WORK' and PMCd= max(Td.TdByWrk)) ")
  
  '*** (Bef 2.13)
  'GRP_REP.Add "Process", Array("Td.TdByLoc", "max(Td.TdByLoc)"), "", "", "", "", "", "", Array("(Select LocDesc From Loc Where LocCoCd= Td.TdCoCd and LocTyp= 'P' and LocCd= Td.TdByLoc) ", "(Select LocDesc From Loc Where LocCoCd= max(Td.TdCoCd) and LocTyp= 'P' and LocCd= max(Td.TdByLoc)) ")
  'GRP_REP.Add "(Process)Worker", Array("Td.TdByLoc+Td.TdByWrk", "max(Td.TdByLoc)+max(Td.TdByWrk)"), Array("'('+ RTrim(Td.TdByLoc)+ ') '+ Td.TdByWrk", "'('+ RTrim(max(Td.TdByLoc))+ ') '+ max(Td.TdByWrk)"), "", "", "Process", "", "", Array("(Select vPDesc From vParam Where vPCoCd= Td.TdCoCd and vPTyp= 'WORK' and vPMCd= Td.TdByWrk) ", "(Select vPDesc From vParam Where vPCoCd= max(Td.TdCoCd) and vPTyp= 'WORK' and vPMCd= max(Td.TdByWrk)) ")
  'GRP_REP.Add "Design Ctg", Array("DmCtg", "max(DmCtg)"), "", "DmCtg", "hDmCtg", "", "", "", Array("(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) ", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) ")
  'GRP_REP.Add "Prd Ctg", Array("DmPrdCtg", "max(DmPrdCtg)"), "", "DmPrdCtg", "hDmPrdCtg", "", "", "", Array("(Select PDesc from Param where PTyp= 'PRDCTG' and PMCd= DmPrdCtg) ", "(Select PDesc from Param where PTyp= 'PRDCTG' and PMCd= max(DmPrdCtg)) ")
  'GRP_REP.Add "Design Code", Array("BOdDmCd", "max(BOdDmCd)"), "", "BOdDmCd", "hBOdDmCd", "Design Ctg,Prd Ctg,(DsgCd)Colour", "", "", Array("Substring(DmDesc,1,120) ", "Substring(max(DmDesc),1,120) ")
  'GRP_REP.Add "(DsgCd)Colour", Array("OdDmCd+OdDmCol", "max(OdDmCd)+max(OdDmCol)"), Array("'('+OdDmCd+') '+OdDmCol", "'('+max(OdDmCd)+') '+max(OdDmCol)"), "BOdDmCd", "hBOdDmCd", "Design Ctg,Prd Ctg,Design Code", "", "", Array("(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) ", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(OdDmCol)) ")
  'GRP_REP.Add "Ord Kt", Array("OdKt", "max(OdKt)"), "", "OdKt", "hOdKt", "", "", "", Array("(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt) ", "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(OdKt)) ")
  'GRP_REP.Add "OrderNo", Array("BOdTc+BOdYy+BOdChr+str(BOdNo)", "max(BOdTc)+max(BOdYy)+max(BOdChr)+str(max(BOdNo))"), Array("BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "max(BOdTc)+'/'+max(BOdYy)+'/'+max(BOdChr)+'/'+LTrim(Str(max(BOdNo)))"), "wOrdNo,OmCmCd,OmPoNo", "hOrdNo,hOmCmCd,hOmPoNo", "Customer Code,Customer PO", "", "", Array("'('+OmCmCd+ '/'+ OmPONo+ ')' ", "'('+max(OmCmCd)+ '/'+ max(OmPONo)+ ')' ")
  ' '******Geeta*****************Emr206****06/04/04*******
  ''**********If UserLevel>6 then don't display the Customer Name
  'GRP_REP.Add "Customer Code", Array("OmCmCd", "max(OmCmCd)"), "", "OmCmCd", "hOmCmCd", "", "", "", Array("(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end ) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) ", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= max(OmCmCd)) ")
  'GRP_REP.Add "Customer PO", Array("OmCmCd+OmPONo", "max(OmCmCd)+max(OmPONo)"), Array("'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "'('+ max(OmCmCd)+ ') '+ LTrim(max(OmPONo))"), "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "", "", "", Array("BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "max(BOdTc)+'/'+max(BOdYy)+'/'+max(BOdChr)+'/'+LTrim(Str(max(BOdNo)))")
  '
  ''/** urmi 2.06 PrdSeq*//
  '  GRP_REP.Add "Ord PrdSeq", Array("OdPrdSeq", "max(OdPrdSeq)"), "", "", "", "", "", "", Array("(Select PDesc from Param where PTyp= 'PrdSeq' and PMCd= OdPrdSeq) ", "(Select PDesc from Param where PTyp= 'PrdSeq' and PMCd= max(OdPrdSeq)) ")
  '
  ' **** Zubin 211 **** '
  'GRP_REP.Add "Bag Type", Array("(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)", "(case max(PValue) when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)"), _
  '                        Array("(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)", "(case max(PValue) when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)"), "", "", "Bag Typ+Chr", "", "", _
  '                        Array("(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)", "(case max(PValue) when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)")
  'GRP_REP.Add "Bag Character", Array("BChr", "max(BChr)"), Array("BChr", "max(BChr)"), "", "", "Bag Type,Bag Typ+Chr", "", "", Array("(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')", "(Select PDesc from param where PTyp ='BCHR' and PMCd= max(BChr) and PSCd= '')")
  'GRP_REP.Add "Bag Typ+Chr", Array("(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+BChr", "(case max(PValue) when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+max(BChr)"), Array("(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+ '/ ' + BChr", "(case max(PValue) when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+ '/ ' + max(BChr)"), "", "", "Bag Type,Bag Character", "", "", Array("(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')", "(Select PDesc from param where PTyp ='BCHR' and PMCd= max(BChr) and PSCd= '')")
  ' **** Zubin 211 **** '
  '*** (Bef 2.13)
  'Lc.LocPrcs
  '*** (Jen 2.13)
  GRP_REP.Add "Location", "TCoCd+Td.TdByLoc", "'['+TCoCd+'] '+Td.TdByLoc", "", "", "Company Code,Loc Prnt Code", "", "", "(Select LocDesc From Loc Where LocCoCd= Td.TdCoCd and LocTyp= 'P' and LocCd= Td.TdByLoc) "
  GRP_REP.Add "Process", "TCoCd+Lc.LocPrcs", "'['+TCoCd+'] '+Lc.LocPrcs", "", "", "Company Code,Loc Prnt Code", "", "", "(Select LocDesc From Loc Where LocCoCd= Td.TdCoCd and LocTyp= 'R' and LocCd= Lc.LocPrcs) "
  GRP_REP.Add "(Location)Worker", "TCoCd+Td.TdByLoc+Td.TdByWrk", "'('+TCoCd +'/'+RTrim(Td.TdByLoc)+ ') '+ Td.TdByWrk", "", "", "Company Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd= Td.TdCoCd and vPTyp= 'WORK' and vPMCd= Td.TdByWrk) "
  GRP_REP.Add "(Process)Worker", "TCoCd+Lc.LocPrcs+Td.TdByWrk", "'('+TCoCd +'/'+RTrim(Lc.LocPrcs)+ ') '+ Td.TdByWrk", "", "", "Company Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd= Td.TdCoCd and vPTyp= 'WORK' and vPMCd= Td.TdByWrk) "
  
  GRP_REP.Add "Design Ctg", "DmCtg", "", "DmCtg", "hDmCtg", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
  GRP_REP.Add "Prd Ctg", "DmPrdCtg", "", "DmPrdCtg", "hDmPrdCtg", "", "", "", "(Select PDesc from Param where PTyp= 'PRDCTG' and PMCd= DmPrdCtg) "
  GRP_REP.Add "Design Code", "BOdDmCd", "", "BOdDmCd", "hBOdDmCd", "Design Ctg,Prd Ctg,(DsgCd)Colour", "", "", "Substring(DmDesc,1,120) "
  GRP_REP.Add "(DsgCd)Colour", "OdDmCd+OdDmCol", "'('+OdDmCd+') '+OdDmCol", "BOdDmCd", "hBOdDmCd", "Design Ctg,Prd Ctg,Design Code", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
  GRP_REP.Add "Ord Kt", "OdKt", "", "OdKt", "hOdKt", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt) "
  GRP_REP.Add "OrderNo", "TCoCd+BOdTc+BOdYy+BOdChr+str(BOdNo)", "TCoCd+'/'+BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "wOrdNo,OmCmCd,OmPoNo", "hOrdNo,hOmCmCd,hOmPoNo", "Company Code,Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end ) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
  GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))"
  '*** Jay 3.2.0 [OdPoNo]
  'vk.19 sub customer added
  GRP_REP.Add "Cust SUB PO", "OmCmCd+OmPONo+OdSubCust+OdPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)+(case when LTrim(OdPONo)<> '' or LTrim(OdSubCust)<>'' then '/'+ LTrim(OdSubCust) + ' : ' + LTrim(OdPONo) else '' end )", "OmCmCd,OmPoNo,OmPoNo", "hOmCmCd,hOmPONo,hOdPONo", "Customer Code,Customer PO", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))+(case when OdPONo<> '' or OdSubCust<>'' then ' - '+ LTrim(OdSubCust) + ' : ' + LTrim(OdPONo) else '' end )"
  '*** Jay 3.2.0 [OdPoNo]
  GRP_REP.Add "Ord PrdSeq", "OdPrdSeq", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'PrdSeq' and PMCd= OdPrdSeq) "

'*** Jay 2.13(CT) ***
  GRP_REP.Add "Bag Type", "(case B.PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) ", _
                          "(case B.PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) ", "", "", "Bag Typ+Chr", "", "", _
                          "(case B.PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) "
  GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
  GRP_REP.Add "Bag Typ+Chr", "(case B.PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) +BChr", "(case B.PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) + '/ ' + BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'  GRP_REP.Add "Bag Type", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)", _
'                          "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)", "", "", "Bag Typ+Chr", "", "", _
'                          "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)"
'  GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'  GRP_REP.Add "Bag Typ+Chr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+BChr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+ '/ ' + BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'*** (Jen 2.13)
'*** Jay 2.13(CT) ***
   
  '*** These Groupings are below the Bag Level ###
  '"qRmCd,qRmCtg", "hqRmCd,hqRmCtg" ### Check why the group field in the header & detail are not getting removed from the rep when a particular group is selected.
  If UCase(adc.MenuCd) = UCase("RepBrkMsgRm") Then
    '*** (Bef 2.13)
    'GRP_REP.Add "Rm Code", "IsNull(RmCtg+ RmSCtg+ RmCd, '')", "'('+ RTrim(IsNull(RmCtg, ''))+'/'+ IsNull(RmSCtg, '')+ ') '+ IsNull(RmCd, '')", "", "", "(RmCd) Sv/Ln,(RmCd) StkRt, Rm Sub Ctg", "", "", Array("IsNull(RmDesc, '')", "IsNull(max(RmDesc), '')")
    ''IsNull(max(RmDesc), '')
    ''**** Shilpa Emr 206 *****
    'GRP_REP.Add "Rm Sub Ctg", "IsNull(RmSCtg, '')", "IsNull(RmSCtg, '')", "", "", "", "", "", "IsNull((Select PDesc from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RmCtg, '') and PSCd= IsNull(RmSCtg, '')), '')"
    ''IsNull((Select PDesc from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RmSCtg, '')), '')
    ''**** Shilpa Emr 206 *****
    'GRP_REP.Add "Rm Sv/Ln", Array("str(TdRmSz,6,3)", "str(max(Tr.TdRmSz),6,3)"), "", "", "", "Rm Code,Rm Ptr,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
    'GRP_REP.Add "Rm StkRt", Array("str(TdRmStkRt,9,2)", "str(max(Tr.TdRmStkRt),9,2)"), "", "", "", "Rm Code,Rm Sv/Ln,Rm Ptr,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
    'If moCn.GetFldVal("Select HPtrYN from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then _
    '   GRP_REP.Add "Rm Ptr", Array("str(TdRmSz,6,3)", "str(max(Tr.TdRmSz),6,3)"), "", "", "", "Rm Code,Rm Sv/Ln,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
    'GRP_REP.Add "(RmCd) Sv/Ln", Array("RmCd+str(TdRmSz,6,3)", "RmCd+str(max(Tr.TdRmSz),6,3)"), Array("'('+RmCd+') '+LTrim(str(TdRmSz,6,3))", "'('+RmCd+') '+LTrim(str(max(Tr.TdRmSz),6,3))"), "", "", "Rm Code,(RmCd) StkRt,Rm Sv/Ln,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
    'GRP_REP.Add "(RmCd) StkRt", Array("RmCd+str(TdRmStkRt,9,2)", "RmCd+str(max(Tr.TdRmStkRt),9,2)"), Array("'('+RmCd+') '+LTrim(str(TdRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(max(Tr.TdRmStkRt),9,2))"), "", "", "Rm Code,(RmCd) Sv/Ln,Rm Sv/Ln,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
    '*** (Bef 2.13)
        
    '*** (Jen 2.13)
    GRP_REP.Add "Rm Code", "IsNull(RmCtg+ RmSCtg+ RmCd, '')", "'('+ RTrim(IsNull(RmCtg, ''))+'/'+ IsNull(RmSCtg, '')+ ') '+ IsNull(RmCd, '')", "", "", "(RmCd) Sv/Ln,(RmCd) StkRt, Rm Sub Ctg", "", "", "IsNull(RmDesc, '')"
    GRP_REP.Add "Rm Sub Ctg", "IsNull(RmSCtg, '')", "IsNull(RmSCtg, '')", "", "", "", "", "", "IsNull((Select PDesc from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RmCtg, '') and PSCd= IsNull(RmSCtg, '')), '')"
'    GRP_REP.Add "Rm Sv/Ln", Array("str(TdRmSz, 6, 3)", "str(Tr.TdRmSz, 6, 3)"), "", "", "", "Rm Code,Rm Ptr,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
    GRP_REP.Add "Rm Sv/Sz", Array(GetRmSzDBFldNm("TdRmSz", "TdRmsz2", "TdRmsz3"), GetRmSzDBFldNm("Tr.TdRmSz", "Tr.TdRmsz2", "Tr.TdRmsz3")), "", "", "", "Rm Code,Rm Ptr,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
    GRP_REP.Add "Rm StkRt", Array("str(TdRmStkRt, 9, 2)", "str(Tr.TdRmStkRt, 9, 2)"), "", "", "", "Rm Code,Rm Sv/Ln,Rm Ptr,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
    If moCn.GetFldVal("Select HPtrYN from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then _
    GRP_REP.Add "Rm Ptr", Array("str(TdRmSz, 6, 3)", "str(Tr.TdRmSz, 6, 3)"), "", "", "", "Rm Code,Rm Sv/Ln,Rm StkRt,(RmCd) Sv/Ln,(RmCd) StkRt", "", "", ""
    GRP_REP.Add "(RmCd) Sv/Sz", Array("RmCd+ " + GetRmSzDBFldNm("TdRmSz", "TdRmsz2", "TdRmsz3"), "RmCd+ " + GetRmSzDBFldNm("Tr.TdRmSz", "Tr.TdRmsz2", "Tr.TdRmsz3")), Array("'('+RmCd+') '+LTrim(str(TdRmSz,6,3))", "'('+RmCd+') '+LTrim(str(Tr.TdRmSz, 6, 3))"), "", "", "Rm Code,(RmCd) StkRt,Rm Sv/Ln,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
    GRP_REP.Add "(RmCd) StkRt", Array("RmCd+ str(TdRmStkRt, 9, 2)", "RmCd+ str(Tr.TdRmStkRt, 9, 2)"), Array("'('+RmCd+') '+LTrim(str(TdRmStkRt, 9, 2))", "'('+RmCd+') '+LTrim(str(Tr.TdRmStkRt, 9, 2))"), "", "", "Rm Code,(RmCd) Sv/Ln,Rm Sv/Ln,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
    '*** (Jen 2.13)
    
    '4.1.2.0
    If adc("UoShpCd") = "S" Then
       'GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "FldDetRmCtg,FldDet2RmCtg,FldDet2RmCd", "hqRmCtg,hqRmCd", "", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
       GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
    Else
       'GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "FldDetRmCtg,FldDet2RmCtg", "hqRmCtg", "", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
       GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
    End If
    
  End If
  '*** These Groupings are below the Bag Level ###
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ,Process,OrderNo,OrderSr,Customer PO,FgSubLoc ******
  If ws_MultiCoMod = True Then
    GRP_REP.Add "Company Code", "TCoCd", "TCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=TCoCd) "
    GRP_REP.Add "Loc Prnt Code", "Lc.LocPrntCd", "Lc.LocPrntCd", "", "", "", "", "", "(Select LocDesc From Loc Where LocCoCd= '" + ctSelfCoCd + "' and LocCd=Lc.LocPrntCd) "
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******
  
'  GRP_REP.Add "Process", "Td.TdByLoc", "", "TdByLoc", "hTdByLoc", "", "", "", "(Select LocDesc From Loc Where LocTyp= 'P' and LocCd= Td.TdByLoc) "
'  GRP_REP.Add "(Process)Worker", "Td.TdByLoc+Td.TdByWrk", "'('+ RTrim(Td.TdByLoc)+ ') '+ Td.TdByWrk", "TdByLoc,TdByWrk", "hTdByLoc,hTdByWrk", "Process", "", "", "(Select PDesc From Param Where PTyp= 'WORK' and PMCd= Td.TdByWrk) "
'  GRP_REP.Add "Design Ctg", "DmCtg", "", "DmCtg", "hDmCtg", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
'  GRP_REP.Add "Prd Ctg", "DmPrdCtg", "", "DmPrdCtg", "hDmPrdCtg", "", "", "", "(Select PDesc from Param where PTyp= 'PRDCTG' and PMCd= DmPrdCtg) "
'  GRP_REP.Add "Design Code", "BOdDmCd", "", "BOdDmCd", "hBOdDmCd", "Design Ctg,Prd Ctg", "", "", "Substring(DmDesc,1,120) "
'  GRP_REP.Add "Ord Kt", "OdKt", "", "OdKt", "hOdKt", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt) "
'  GRP_REP.Add "OrderNo", "BOdTc+BOdYy+BOdChr+str(BOdNo)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "wOrdNo,OmCmCd,OmPoNo", "hOrdNo,hOmCmCd,hOmPoNo", "Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
'  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
'  GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))"

End Sub

Private Sub adc_setreprecsource()
On Error GoTo RepErr

  '****** (Jen 2.13) Added:  TdRmWt as qRmWt, TdRmQty as qRmQty ******'
  '*** Report Sql For BrkMsg Rm ***
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  space(1) as qTag, TdCoCd, TdBYy, TdBChr, TdBNo, TdRmWt as qBMWt,
'  0 as qMul, TdRmQty as qBMQty, TdRmWt as qBRmWt, TdRmQty as qBRmQty,
'  TDt, RmCtg as qRmCtg, RmSCtg as qRmSCtg, TdRmCd as qRmCd,
'  TdFrRmDc as qRmDc, TdByWrk, TdByLoc, TdFrRmLoc as qBMLoc,
'  TdTc, TdYy, TdChr, TdNo, TdSr, DmCtg, DmPrdCtg, BOdDmCd, OdKt, BOdTc,
'  BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo, space(1) as qBM, TdRmStkRt as qRmVal,
'  TdRmWt as qRmWt, TdRmQty as qRmQty
'  from Txnd, Txn, RmMst, Bag, OrdDsg, OrdMst, DsgMst Where 1 = 2
  '*** Report Sql For BrkMsg Rm ***
  
  '*** Report Sql For BrkMsg Prd ***
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  space(1) as qTag, TdCoCd, TdBYy, TdBChr, TdBNo, TdRmWt as qBMWt,
'  0 as qMul, TdRmQty as qBMQty, TdRmWt as qBRmWt, TdRmQty as qBRmQty,
'  TDt, RmCtg as qRmCtg, RmSCtg as qRmSCtg, TdRmCd as qRmCd,
'  TdFrRmDc as qRmDc, TdByWrk, TdByLoc, TdFrRmLoc as qBMLoc,
'  TdTc, TdYy, TdChr, TdNo, TdSr, DmCtg, DmPrdCtg, BOdDmCd, OdKt, BOdTc,
'  BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo, space(1) as qBM, TdRmStkRt as qRmVal
'  from Txnd, Txn, RmMst, Bag, OrdDsg, OrdMst, DsgMst Where 1 = 2
  '*** Report Sql For BrkMsg Prd ***
  
  
  Dim grpflds0 As String, grpflds1 As String, wBM As String, wrepcnd As String
  Dim wCnd As String, wSqlStrg1 As String, wSqlStrg2 As String, wSqlStrg3 As String
  Dim wSqlStrg4 As String, wSqlStrg5 As String, wSqlStrg As String
  
  Dim ws_TmpTblA As String, ws_TmpTblB As String, ws_ShowAllPrd As String  '*** (Jen 2.13)
  
  'Dim ws_qBM As String
  Dim ws_RmRt As String
' **** Zubin 211 **** '
  Dim ws_BagJoin As String
' **** Zubin 211 **** '

  Dim wAddlCndTd As String, wAddlCndTr As String, wAddlCndOth As String '****** Sachin 3.02.0 - New Fields In Where Clause
  
  ' Zubin 213
  Dim ws_BaseCurCd As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")

  Set Rep = Nothing
  Set Rep1 = Nothing
  
  Select Case UCase(adc.MenuCd)
  Case UCase("RepBrkMsgRm")
    Set adc.RepSource = Rep
    Set mRep = Rep
  Case UCase("RepBrkMsgPrd")
    Set adc.RepSource = Rep1
    Set mRep = Rep1
  End Select
  
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)
  
  '*** To check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer
    wGrp(1) = GRP_REP.Value(gltGroup1)
    wGrp(2) = GRP_REP.Value(gltGroup2)
    wGrp(3) = GRP_REP.Value(gltGroup3)
    wDetPos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
  Next i
  
  '### Note TdByLoc,TdByWrk,qRmSCtg, qRmCd fields place in G4-DET, will come from the last record of the Bag. OK.
    
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCndOth = wAddlCndOth + IIF(wAddlCndOth <> "" And Trim(adc("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(adc("UOCMCDFR")) + "'", "")
  wAddlCndOth = wAddlCndOth + IIF(wAddlCndOth <> "" And Trim(adc("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(adc("UOCMCDTO")) + "'", "")
  wAddlCndOth = wAddlCndOth + IIF(wAddlCndOth <> "" And Trim(adc("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(adc("UOCMCDSEL")) + ")", "")
    
  wAddlCndTd = IIF(Trim(adc("UORMCTGFR")) <> "", " And TdRmCtg = '" + Trim(adc("UORMCTGFR")) + "'", "")
  wAddlCndTd = wAddlCndTd + IIF(wAddlCndTd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " TdRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCndTd = wAddlCndTd + IIF(wAddlCndTd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " TdRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCndTd = wAddlCndTd + IIF(wAddlCndTd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " TdRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  wAddlCndTd = wAddlCndTd + IIF(wAddlCndTd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " TdDt >= '" + Trim(adc("UODTFR")) + "'", "")
  wAddlCndTd = wAddlCndTd + IIF(wAddlCndTd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " TdDt <= '" + Trim(adc("UODTTO")) + "'", "")
    
  wAddlCndTd = wAddlCndTd + IIF(wAddlCndTd <> "" And wAddlCndOth <> "", " And ", "") + wAddlCndOth
    
  wAddlCndTr = IIF(Trim(adc("UORMCTGFR")) <> "", " And Tr.TdRmCtg = '" + Trim(adc("UORMCTGFR")) + "'", "")
  wAddlCndTr = wAddlCndTr + IIF(wAddlCndTr <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " Tr.TdRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCndTr = wAddlCndTr + IIF(wAddlCndTr <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " Tr.TdRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCndTr = wAddlCndTr + IIF(wAddlCndTr <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " Tr.TdRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  wAddlCndTr = wAddlCndTr + IIF(wAddlCndTr <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " Td.TdDt >= '" + Trim(adc("UODTFR")) + "'", "")
  wAddlCndTr = wAddlCndTr + IIF(wAddlCndTr <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " Td.TdDt <= '" + Trim(adc("UODTTO")) + "'", "")
  
  wAddlCndTr = wAddlCndTr + IIF(wAddlCndTr <> "" And wAddlCndOth <> "", " And ", "") + wAddlCndOth
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
    
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then
    mRep.BoxHd.Suppress = True: mRep.BoxDet.Suppress = True: mRep.BoxFt.Suppress = True
    mRep.DetSec2.Suppress = True
  End If
  
  If UCase(adc.MenuCd) = UCase("RepBrkMsgRm") Then
    adc("UoRmCtgFr").CmpStr = "RmCtg= "
    adc("UoRmSCtgFr").CmpStr = "RmSCtg>= "
    adc("UoRmSCtgTo").CmpStr = "RmSCtg<= "
    adc("UoRmSCtgSel").CmpStr = "RmSCtg In "
    adc("UoRmCdFr").CmpStr = "RmCd>= "
    adc("UoRmCdTo").CmpStr = "RmCd<= "
    adc("UoRmCdSel").CmpStr = "RmCd In "
    adc("UoStkRtFr").CmpStr = "TdRmStkRt>= "
    adc("UoStkRtTo").CmpStr = "TdRmStkRt<= "
    
    'Sachin 4.1.2
    adc("UoParentCdFr").CmpStr = "RmParentCd>="
    adc("UoParentCdTo").CmpStr = "RmParentCd<="
    adc("UoParentCdSel").CmpStr = "RmParentCd In "
    
  ElseIf UCase(adc.MenuCd) = UCase("RepBrkMsgPrd") Then
    adc("UoRmCtgFr").CmpStr = "RmCtg= "
    adc("UoRmSCtgFr").CmpStr = "RmSCtg>= "
    adc("UoRmSCtgTo").CmpStr = "RmSCtg<= "
    adc("UoRmSCtgSel").CmpStr = "RmSCtg In "
    adc("UoRmCdFr").CmpStr = ""
    adc("UoRmCdTo").CmpStr = ""
    adc("UoRmCdSel").CmpStr = ""
    adc("UoStkRtFr").CmpStr = ""
    adc("UoStkRtTo").CmpStr = ""
    
    'Sachin 4.1.2
    adc("UoParentCdFr").CmpStr = ""
    adc("UoParentCdTo").CmpStr = ""
    adc("UoParentCdSel").CmpStr = ""
  End If
  
  '*** Whether grouping should be done on RmCtg, RmSCtg or RmCd ***
  If adc("UoShpCd") = "S" And UCase(adc.MenuCd) = UCase("RepBrkMsgRm") Then
    GRP_REP.SetFormula mRep, "wGrp4", "{rdo.TdCoCd}+{rdo.TdBYy}+{rdo.TdBChr}+ToText ({rdo.TdBNo},'######0')+{rdo.qRmCtg}+{rdo.qRmSCtg}"
    mRep.hqRmCd.SetText "Rm SubCtg"
    mRep.hTdByWrk.Suppress = True
    mRep.hTdByLoc.Suppress = True
    
    mRep.FldqRmCd.Suppress = True
    mRep.FldqRmSCtg.Suppress = False
    mRep.FldTdByWrk.Suppress = True
    mRep.FldTdByLoc.Suppress = True
  
    'mRep.FldDet2RmCd.Suppress = True
    'mRep.FldDet2RmSCtg.Suppress = False
    'mRep.FldDet2ByWrk.Suppress = True
    'mRep.FldDet2ByLoc.Suppress = True
  End If
  
  If adc("UoShpCd") = "R" And UCase(adc.MenuCd) = UCase("RepBrkMsgRm") Then
    GRP_REP.SetFormula mRep, "wGrp4", "{rdo.TdCoCd}+{rdo.TdBYy}+{rdo.TdBChr}+ToText ({rdo.TdBNo},'######0')+{rdo.qRmCd}"
    mRep.hqRmCd.SetText "Rm Code"
    mRep.hTdByWrk.Suppress = False
    mRep.hTdByLoc.Suppress = False
    
    mRep.FldqRmCd.Suppress = False
    mRep.FldqRmSCtg.Suppress = True
    mRep.FldTdByWrk.Suppress = False
    mRep.FldTdByLoc.Suppress = False
  
    'mRep.FldDet2RmCd.Suppress = False
    'mRep.FldDet2RmSCtg.Suppress = True
    'mRep.FldDet2ByWrk.Suppress = False
    'mRep.FldDet2ByLoc.Suppress = False
  End If
  
  '*** Brk/Msg or Both ***
  If adc("UoBrkMsg") = "B" Then
    wBM = "'B'"
    mRep.TxtHead.SetText "Breakage Report"
  ElseIf adc("UoBrkMsg") = "M" Then
    wBM = "'M'"
    mRep.TxtHead.SetText "Missing Report"
  ElseIf adc("UoBrkMsg") = "" Then
    wBM = "'B','M'"
    mRep.TxtHead.SetText "Breakage, Missing Report"
  End If
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  'wCnd = wCnd + " and BYy= '05' and BChr= 'REG' and BNo in (828, 851) " '*** (Jen 2.13 temporary code)
  
  'wDCXARt = " isnull((Select " + wRt + " from RmRt where RrTcTyp='RM' and RrCmCtg= 'C' and " + _
  "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd= '" + ctValueCmCd + "' and " + _
  "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrDmCtg= '' and RrLabMCd= '' and " + _
   wRt + " <> 0), 0) "
  
  ' Zubin 213 (ws_BaseCurCd cnd added)
  
   If OptLoc(0).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and Td.TdByLoc >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Td.TdByLoc <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Td.TdByLoc In (" + adc("UoIssLocSel") + ")"
   ElseIf OptLoc(1).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and Lc.LocPrcs >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Lc.LocPrcs <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Lc.LocPrcs In (" + adc("UoIssLocSel") + ")"
   End If
  
 
 
  If (adc("UOPMCD") = UCase("CstVal")) Then
    ws_RmRt = ",(Select rQw.RmQwCst From RmMst rQw Where " + IIF(gs_Partition = ctCurrPrtn, " rQw.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " rQw.RmCd=Rm.RmCd) As qQw, isnull((Select RrCstRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
    "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd= '" + ctValueCmCd + "' and " + _
    "RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn<= TdRmSz And RrToLn>= TdRmSz and RrDmCtg= '' and RrLabMCd= ''), 0) as qRmVal "
  ElseIf (adc("UOPMCD") = UCase("SalVal")) Then
    ws_RmRt = ",(Select rQw.RmQw From RmMst rQw Where " + IIF(gs_Partition = ctCurrPrtn, " rQw.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " rQw.RmCd=Rm.RmCd) As qQw, isnull((Select RrSalRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
    "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd= '" + ctValueCmCd + "' and " + _
    "RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn<= TdRmSz And RrToLn>= TdRmSz and RrDmCtg= '' and RrLabMCd= ''), 0) as qRmVal "
  ElseIf (adc("UOPMCD") = UCase("StkVal")) Then
    ws_RmRt = ",'W' As qQw, TdRmStkRt as qRmVal "
    '4.1.4
    If UCase(adc("UoPMCd1")) = UCase("CstVal") Then
          ws_RmRt = " , (Case When RSCtg.PValue1='N' Then ((Select rQw.RmQwCst From RmMst rQw Where " + IIF(gs_Partition = ctCurrPrtn, " rQw.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " rQw.RmCd=Rm.RmCd)) Else 'W' End) As qQw, " + _
              " (Case When RSCtg.PValue1='N' Then (Isnull((Select RrCstRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
              "   RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd= '" + ctValueCmCd + "' and " + _
              "   RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn<= TdRmSz And RrToLn>= TdRmSz and RrDmCtg= '' and RrLabMCd= ''), 0)) Else TdRmStkRt End) as qRmVal "
    
    ElseIf (adc("UoPMCd1") = UCase("SalVal")) Then
          ws_RmRt = ", (Case When RSCtg.PValue1='N' Then ((Select rQw.RmQw From RmMst rQw Where " + IIF(gs_Partition = ctCurrPrtn, " rQw.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " rQw.RmCd=Rm.RmCd)) Else 'W' End) As qQw, " + _
             "  (Case When RSCtg.PValue1='N' Then (isnull((Select RrSalRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
             "    RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd= '" + ctValueCmCd + "' and " + _
             "    RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn<= TdRmSz And RrToLn>= TdRmSz and RrDmCtg= '' and RrLabMCd= ''), 0)) Else TdRmStkRt End) as qRmVal "
      
    End If
  ElseIf (adc("UOPMCD") = "") Then
    ws_RmRt = ",' ' As qQw, 0.0 as qRmVal "
  End If
  '/*** urmi 2.06 */
    
'*** Jay 2.13(CT) ***
  ws_BagJoin = " Join Param B On B.PTyp= 'BCHR' and B.PMCd= BChr and B.PSCd= '' " + _
                    IIF(adc("UoBagTyp") <> "", " and B.PValue= '" + adc("UoBagTyp") + "' ", " ") + _
                    IIF(adc("UOBCHRFR") <> "", " and B.PMCd= '" + adc("UOBCHRFR") + "' ", " ")
' **** Zubin 211 **** '
'  If UCase(adc("UoBagTyp")) = "N" Then
'    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'N' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'  ElseIf UCase(adc("UoBagTyp")) = "Y" Then
'    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'Y' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'  Else
'    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'  End If
' **** Zubin 211 **** '
'*** Jay 2.13(CT) ***

  '****** (Jen 2.13) Added:  qRmWt, qRmQty & introduced the temporary table #TMPA ******'
  '***(Jen 2.11) (05/07/05)
  'Dust weight and BM Reason Code added
  '6.1
  ws_TmpTblA = " Set dateformat " + DtFmtStr() + " " + _
              "Select " + grpflds0 + ", " + _
              "'D' as qTag, TdCoCd, TdBYy, TdBChr, TdBNo, TdRmWt as qBMWt, " + _
              "(Case when L2.LocTyp in ('B', 'M') and TdToRmDc= 'D' then 1 " + _
              "      when L2.LocTyp in ('B', 'M') and TdToRmDc= 'C' then -1 " + _
              "      when L1.LocTyp in ('B', 'M') and TdFrRmDc= 'D' then 1 " + _
              "      when L1.LocTyp in ('B', 'M') and TdFrRmDc= 'C' then -1 " + _
              "Else 0 end) as qMul, " + _
              "TdRmQty as qBMQty, 0.0 as qBRmWt, 0 as qBRmQty, TDt, RmCtg as qRmCtg, " + _
              "RmSCtg as qRmSCtg, TdRmCd as qRmCd, " + _
              "(Case when L2.LocTyp in ('B', 'M') then TdToRmDc " + _
              "      when L1.LocTyp in ('B', 'M') then TdFrRmDc Else '' end) as qRmDc, " + _
              "TdByWrk,TdDustWt,TdBmRsnCd, Lc.LocPrntCd as LocPrntCd," + IIF(OptLoc(0).Value = True, "TdByLoc", "Lc.LocPrcs") + " as TdByLoc, " + _
              "(Case when L2.LocTyp in ('B', 'M') then TdToRmLoc " + _
              "      when L1.LocTyp in ('B', 'M') then TdFrRmLoc Else '' end) as qBMLoc, " + _
              "TdTc, TdYy, TdChr, TdNo, TdSr, DmCtg, DmPrdCtg, " + _
              "BOdDmCd, OdKt, BOdTc, BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo, " + _
              "(Case when L2.LocTyp in ('B', 'M') then L2.LocTyp " + _
              "      when L1.LocTyp in ('B', 'M') then L1.LocTyp Else '' end) as qBM " + _
              ws_RmRt + IIF(UCase(adc.MenuCd) = UCase("RepBrkMsgRm"), ", 0.0 as qRmWt, 0 as qRmQty ", " ")
              
  '****** Sachin 3.02 - Id fields in Joins
  ws_TmpTblA = ws_TmpTblA + " into #TMPA from Txnd Td " + _
              "Join Txn on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr=TdChr and TNo=TdNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey=TdPrtKey ", "") + _
              "Join RmMst Rm on RmCd= TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=TdPrtKey ", "") + _
              "Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg " + _
              "Join Bag on BIdNo=TdBIdNo and BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=TdPrtKey ", "") + _
              ws_BagJoin + _
              "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and OdChr= BOdChr " + _
              "and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "") + _
              "Join OrdMst on OmIdNo=BOmIdNo and OmCoCd=BCoCd and OmTc= BOdTc and OmYy= BOdYy and " + _
              "OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=BPrtKey ", "") + _
              "Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=OdPrtKey ", "") + _
              "join Loc L1 on L1.LocCoCd= TdCoCd and L1.LocCd= TdFrRmLoc " + _
              "join Loc L2 on L2.LocCoCd= TdCoCd and L2.LocCd= TdToRmLoc " + _
              "join Loc Lc on Lc.LocCoCd= TdCoCd and Lc.LocCd= TdByLoc " + _
              "where (L2.LocTyp in (" + wBM + ") or L1.LocTyp in (" + wBM + "))  " + _
              "And (TdToRmLocTyp in (" + wBM + ") Or TdFrRmLocTyp in (" + wBM + ")) " + _
              "and not(L1.LocTyp in (" + wBM + ") and L2.LocTyp in (" + wBM + ")) " + _
              "and not(TdFrRmLocTyp in (" + wBM + ") and TdToRmLocTyp in (" + wBM + ")) " + _
              "and TdSrNo<> 0 " + _
              IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey='" + ctCurrPrtn + "'", "") + _
              wCnd + IIF(wAddlCndTd <> "", IIF(wCnd <> "", "", " Where ") + wAddlCndTd, "")
  
  '*** (Jen 2.13)
  moCn.CreateTmpTable "#TMPA", ws_TmpTblA
  wSqlStrg1 = "Set dateformat " + DtFmtStr() + " Select * from #TMPA "    '6.1
  '*** (Jen 2.13)
  
  '***(Jen 2.11) (05/07/05)
  
  '***(Bef 05/07/05)
'  '/** selecting Los where TdRmFrLoc in wBM ('B','M') **/
'  wSqlStrg1 = "Set dateformat DMY " + _
'              "Select " + grpflds0 + ", " + _
'              "'D' as qTag, TdBYy, TdBChr, TdBNo, TdRmWt as qBMWt, " + _
'              "(Case when TdFrRmDc='D' then 1 when TdFrRmDc='C' then -1 Else 0 end) as qMul, " + _
'              "TdRmQty as qBMQty, 0.0 as qBRmWt, 0 as qBRmQty, TDt, RmCtg as qRmCtg, " + _
'              "RmSCtg as qRmSCtg, TdRmCd as qRmCd, TdFrRmDc as qRmDc, TdByWrk, TdByLoc, " + _
'              "TdFrRmLoc as qBMLoc, TdTc, TdYy, TdChr, TdNo, TdSr, DmCtg, DmPrdCtg, " + _
'              "BOdDmCd, OdKt, BOdTc, BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo, LocTyp as qBM " + _
'              ws_RmRt + " from Txnd Td " + _
'              "Join Txn on TCoCd= TdCoCd and TTc=TdTc and TYy=TdYy and TChr=TdChr and TNo=TdNo " + _
'              "Join RmMst on RmCd= TdRmCd " + _
'              "Join Bag on BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + _
'              ws_BagJoin + _
'              "Join OrdDsg on OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and OdChr= BOdChr " + _
'              "and OdNo= BOdNo and OdSr= BOdSr " + _
'              "Join OrdMst on OmCoCd=BCoCd and OmTc= BOdTc and OmYy= BOdYy and " + _
'              "OmChr= BOdChr and OmNo= BOdNo " + _
'              "Join DsgMst on DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + _
'              "join Loc on LocCoCd= TdCoCd and LocCd= TdFrRmLoc " + _
'              "where LocTyp in (" + wBM + ") and TdSrNo<> 0 " + wCnd
'
'  '/** selecting Loss where TdRmToLoc in wBM ('B','M') **/
'  'ws_qBM = "(select LocTyp from Loc where LocCoCd= TdCoCd and LocCd= TdToRmLoc) as qBM "
'
'  wSqlStrg2 = "Select " + grpflds0 + ", " + _
'              "'D' as qTag, TdBYy, TdBChr, TdBNo, TdRmWt as qBMWt, " + _
'              "(Case when  TdToRmDc='D' then 1 when  TdToRmDc='C' then -1 Else 0 end) as qMul, " + _
'              "TdRmQty as qBMQty, 0.0 as qBRmWt, 0 as qBRmQty, TDt, RmCtg as qRmCtg, " + _
'              "RmSCtg as qRmSCtg, TdRmCd as qRmCd, TdToRmDc as qRmDc, TdByWrk, TdByLoc, " + _
'              "TdToRmLoc as qBMLoc, TdTc, TdYy, TdChr, TdNo, TdSr, DmCtg, DmPrdCtg, " + _
'              "BOdDmCd , OdKt, BOdTc, BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo, LocTyp as qBM " + _
'              ws_RmRt + _
'              "from Txnd Td " + _
'              "Join Txn on TCoCd= TdCoCd and TTc=TdTc and TYy=TdYy and TChr=TdChr and TNo=TdNo " + _
'              "Join RmMst on RmCd=TdRmCd " + _
'              "Join Bag on BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + _
'              ws_BagJoin + _
'              "Join OrdDsg on OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
'              "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + _
'              "Join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy " + _
'              "and OmChr= BOdChr and OmNo= BOdNo " + _
'              "Join DsgMst on DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + _
'              "join Loc on LocCoCd= TdCoCd and LocCd= TdToRmLoc " + _
'              "where LocTyp in (" + wBM + ") and TdSrNo<> 0 " + wCnd
  '***(Bef 05/07/05)
  
  
  '/*** selecting Prd Records ***/
  '/* (qBRmWt) Dia/Cs weight in the Bag as of TdYyKey of PrdYN='Y' */
    
  If UCase(adc.MenuCd) = UCase("RepBrkMsgRm") Then
    adc("UoRmCtgFr").CmpStr = "RmCtg= "
    adc("UoRmSCtgFr").CmpStr = "RmSCtg>= "
    adc("UoRmSCtgTo").CmpStr = "RmSCtg<= "
    adc("UoRmSCtgSel").CmpStr = "RmSCtg In "
    adc("UoRmCdFr").CmpStr = "RmCd>= "
    adc("UoRmCdTo").CmpStr = "RmCd<= "
    adc("UoRmCdSel").CmpStr = "RmCd In "
    adc("UoStkRtFr").CmpStr = "Tr.TdRmStkRt>= "
    adc("UoStkRtTo").CmpStr = "Tr.TdRmStkRt<= "
    
    'Sachin 4.1.2
    adc("UoParentCdFr").CmpStr = "RmParentCd>="
    adc("UoParentCdTo").CmpStr = "RmParentCd<="
    adc("UoParentCdSel").CmpStr = "RmParentCd In "
    
  ElseIf UCase(adc.MenuCd) = UCase("RepBrkMsgPrd") Then
    adc("UoRmCtgFr").CmpStr = "TpRmCtg= "
    adc("UoRmSCtgFr").CmpStr = "TpRmSCtg>= "
    adc("UoRmSCtgTo").CmpStr = "TpRmSCtg<= "
    adc("UoRmSCtgSel").CmpStr = "TpRmSCtg In "
    adc("UoRmCdFr").CmpStr = ""
    adc("UoRmCdTo").CmpStr = ""
    adc("UoRmCdSel").CmpStr = ""
    adc("UoStkRtFr").CmpStr = ""
    adc("UoStkRtTo").CmpStr = ""
    
   'Sachin 4.1.2
    adc("UoParentCdFr").CmpStr = ""
    adc("UoParentCdTo").CmpStr = ""
    adc("UoParentCdSel").CmpStr = ""
  End If
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  If OptLoc(0).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and Td.TdByLoc >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Td.TdByLoc <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Td.TdByLoc In (" + adc("UoIssLocSel") + ")"
   ElseIf OptLoc(1).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and Lc.LocPrcs >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Lc.LocPrcs <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Lc.LocPrcs In (" + adc("UoIssLocSel") + ")"
   End If
   
  'wCnd = wCnd + " and BYy= '05' and BChr= 'REG' and BNo in (828, 851) " '*** (Jen 2.13 temporary code)
  
  '*** (Jen 2.13)
  If adc("UoYN3") = "N" Then
    ws_ShowAllPrd = " where exists (Select 'x' from #TMPA A where A.DGrp1= B.DGrp1 and A.DGrp2= B.DGrp2 and A.DGrp3= B.DGrp3 " + _
                    "and A.TdBYy= B.TdBYy and A.TdBChr= B.TdBChr and A.TdBNo= B.TdBNo and A.qRmCtg= B.qRmCtg " + _
                    "and A.qRmSCtg= B.qRmSCtg " + IIF(UCase(adc("UoShpCd")) = "S" Or UCase(adc.MenuCd) = UCase("RepBrkMsgPrd"), " ", " and A.qRmCd= B.qRmCd ") + ") "
  Else
    ws_ShowAllPrd = " "
  End If
  '*** (Jen 2.13)
  
  If UCase(adc.MenuCd) = UCase("RepBrkMsgRm") Then
    
    '*** (Bef 2.13)
    'wSqlStrg3 = "Select " + grpflds1 + ", " + _
    '            "'P' as qTag, Td.TdBYy as TdBYy, Td.TdBChr as TdBChr, Td.TdBNo as TdBNo, " + _
    '            "0.0 as qBMWt, 0 as qMul, 0 as qBMQty, Sum(isnull(Tr.TdRmWt* " + _
    '            "(Case when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='D' then 1 " + _
    '            "when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='C' then -1 " + _
    '            "when Tr.TdToRmLoc='W' and Tr.TdToRmDc='D' then 1 " + _
    '            "when Tr.TdToRmLoc='W' and Tr.TdToRmDc='C' then -1 " + _
    '            "Else 0 end),0.0)) as qBRmWt, Sum(isnull(Tr.TdRmQty* " + _
    '            "(Case when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='D' then 1 " + _
    '            "when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='C' then -1 " + _
    '            "when Tr.TdToRmLoc='W' and Tr.TdToRmDc='D' then 1 " + _
    '            "when Tr.TdToRmLoc='W' and Tr.TdToRmDc='C' then -1 " + _
    '            "Else 0 end),0)) as qBRmQty, "
    '
    'wSqlStrg4 = " max(TDt) as TDt, RmCtg as qRmCtg, RmSCtg as qRmSCtg, RmCd as qRmCd, '' as qRmDc, " + _
    '            "max(Td.TdByWrk) as TdByWrk, max(Td.TdByLoc) as TdByLoc, '' as qBMLoc, Td.TdTc as TdTc, " + _
    '            "Td.TdYy as TdYy, Td.TdChr as TdChr, Td.TdNo as TdNo, Td.TdSr as TdSr, " + _
    '            "max(DmCtg) as DmCtg, max(DmPrdCtg) as DmPrdCtg, max(BOdDmCd) as BOdDmCd, max(OdKt) as OdKt, " + _
    '            "max(BOdTc) as BOdTc, max(BOdYy) as BOdYy, max(BOdChr) as BOdChr, max(BOdNo) as BOdNo, " + _
    '            "max(OmCmCd) as OmCmCd, max(OmPoNo) as OmPoNo, '' as qBM, 0.0 as qRmVal " + _
    '            "from Txnd Td " + _
    '            "Join Txn on TCoCd=Td.TdCoCd and TTc=Td.TdTc and TYy=Td.TdYy and " + _
    '            "TChr=Td.TdChr and TNo=Td.TdNo " + _
    '            "Join Txnd Tr on Tr.TdCoCd=Td.TdCoCd and Tr.TdBYy=Td.TdByy and Tr.TdBChr = Td.TdBChr " + _
    '            "and Tr.TdBNo = Td.TdBNo and Tr.TdYyKey <= Td.TdYyKey  and Tr.TdSrno<> 0 " + _
    '            "Join RmMst on  RmCd=Tr.TdRmCd " + _
    '            "Join Bag on BCoCd= Td.TdCoCd and BYy= Td.TdBYy and BChr= Td.TdBChr and BNo= Td.TdBNo " + _
    '            ws_BagJoin + _
    '            " Join OrdDsg on OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
    '            "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + _
    '            "Join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy " + _
    '            "and OmChr= BOdChr and OmNo= BOdNo " + _
    '            "Join DsgMst on DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + _
    '            "where Td.TdSrNo= 0 and Td.TdPrdYN= 'Y' " + _
    '            wCnd + " group by Td.TdByy, Td.TdBChr, Td.TdBNo, RmCtg, RMSCtg, RmCd, Td.TdTc, " + _
    '            "Td.TdYy, Td.TdChr, Td.TdNo, Td.TdSr "
    '*** (Bef 2.13)
  
    '*** (Jen 2.13)
    '****** Sachin 2.14.0 - [Multi-Company] - Loc Join added below
    '6.1
    ws_TmpTblB = "Set dateformat " + DtFmtStr() + " Select " + grpflds1 + ", " + _
                 "Td.TdCoCd as TdCoCd, Td.TdBYy as TdBYy, Td.TdBChr as TdBChr, Td.TdBNo as TdBNo, " + _
                 "isnull(Tr.TdRmWt* " + _
                 "(Case when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='D' then 1 " + _
                 "when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='C' then -1 " + _
                 "when Tr.TdToRmLoc='W' and Tr.TdToRmDc='D' then 1 " + _
                 "when Tr.TdToRmLoc='W' and Tr.TdToRmDc='C' then -1 " + _
                 "Else 0 end), 0.0) as qBRmWt, isnull(Tr.TdRmQty* " + _
                 "(Case when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='D' then 1 " + _
                 "when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='C' then -1 " + _
                 "when Tr.TdToRmLoc='W' and Tr.TdToRmDc='D' then 1 " + _
                 "when Tr.TdToRmLoc='W' and Tr.TdToRmDc='C' then -1 " + _
                 "Else 0 end), 0) as qBRmQty, "
  
    '****** Sachin 3.02 - Id fields in Joins
    'Dust weight and BM Reason Code added
    ws_TmpTblB = ws_TmpTblB + " TDt, RmCtg as qRmCtg, RmSCtg as qRmSCtg, RmCd as qRmCd, " + _
                 "Td.TdByWrk as TdByWrk,Td.TdDustWt as TdDustWt,Td.TdBmRsnCd as TdBmRsnCd, Lc.LocPrntCd as LocPrntCd," + IIF(OptLoc(0).Value = True, "Td.TdByLoc", "Lc.LocPrcs") + " as TdByLoc, Td.TdTc as TdTc, " + _
                 "Td.TdYy as TdYy, Td.TdChr as TdChr, Td.TdNo as TdNo, Td.TdSr as TdSr, " + _
                 "DmCtg, DmPrdCtg, BOdDmCd, OdKt, BOdTc, BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo, Td.TdYyKey as qTdYyKey " + _
                 "into #TMPB from Txnd Td " + _
                 "Join Txn on TIdNo=Td.TdTIdNo and TCoCd= Td.TdCoCd and TTc= Td.TdTc and TYy= Td.TdYy and " + _
                 "TChr= Td.TdChr and TNo= Td.TdNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey=Td.TdPrtKey ", "") + _
                 "Join Loc Lc On LocCoCd=Td.TdCoCd And LocCd=Td.TdByLoc " + _
                 "Join Txnd Tr on Tr.TdBIdNo=Td.TdBIdNo and Tr.TdCoCd= Td.TdCoCd and Tr.TdBYy= Td.TdByy and Tr.TdBChr= Td.TdBChr " + _
                 "and Tr.TdBNo= Td.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " And Tr.TdPrtKey=Td.TdPrtKey ", "") + " and Tr.TdYyKey <= Td.TdYyKey  and Tr.TdSrno<> 0 " + _
                 "Join RmMst Rm on RmCd= Tr.TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=Tr.TdPrtKey ", "") + _
                 "Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=RmSCtg " + _
                 "Join Bag on BIdNo=Td.TdBIdNo and BCoCd= Td.TdCoCd and BYy= Td.TdBYy and BChr= Td.TdBChr and BNo= Td.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=Td.TdPrtKey ", "") + _
                 ws_BagJoin + _
                 " Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
                 "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "") + _
                 "Join OrdMst on OmIdNo=BOmIdNo and OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy " + _
                 "and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=BPrtKey ", "") + _
                 "Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=OdPrtKey ", "") + _
                 "where " + IIF(gs_Partition = ctCurrPrtn, " Td.TdPrtKey='" + ctCurrPrtn + "' And ", "") + " Td.TdSrNo= 0 and Td.TdPrdYN= 'Y' " + _
                 wCnd + IIF(wAddlCndTr <> "", IIF(wCnd <> "", "", " Where ") + wAddlCndTr, "")
    
    moCn.CreateTmpTable "#TMPB", ws_TmpTblB
    
    wSqlStrg2 = "Select max(Grp1) as Grp1, max(Grp2) as Grp2, max(Grp3) as Grp3, " + _
                "DGrp1, DGrp2, DGrp3, max(G1Desc) as G1Desc, max(G2Desc) as G2Desc, max(G3Desc) as G3Desc, " + _
                "'P' as qTag, TdCoCd, TdBYy, TdBChr, TdBNo, 0.0 as qBMWt, 0 as qMul, 0 as qBMQty, " + _
                "Sum(qBRmWt) as qBRmWt, Sum(qBRmQty) as qBRmQty, "
    
    '(2.13) Added:  qRmWt, qRmQty
    'Dust weight and BM Reason Code added
    wSqlStrg3 = " max(TDt) as TDt, qRmCtg, qRmSCtg, max(qRmCd) as qRmCd, '' as qRmDc, " + _
                "max(TdByWrk) as TdByWrk,max(TdDustWt) as TdDustWt,max(TdBmRsnCd) as TdBmRsnCd, max(LocPrntCd) as LocPrntCd, max(TdByLoc) as TdByLoc, '' as qBMLoc, TdTc, " + _
                "TdYy, TdChr, TdNo, TdSr, max(DmCtg) as DmCtg, max(DmPrdCtg) as DmPrdCtg, " + _
                "max(BOdDmCd) as BOdDmCd, max(OdKt) as OdKt, max(BOdTc) as BOdTc, max(BOdYy) as BOdYy, " + _
                "max(BOdChr) as BOdChr, max(BOdNo) as BOdNo, max(OmCmCd) as OmCmCd, max(OmPoNo) as OmPoNo, " + _
                "'' as qBM, '' As qQw, 0.0 as qRmVal, 0.0 as qRmWt, 0 as qRmQty " + _
                "from #TMPB B " + ws_ShowAllPrd + _
                " group by DGrp1, DGrp2, DGrp3, TdCoCd, TdBYy, TdBChr, TdBNo, qRmCtg, qRmSCtg " + IIF(UCase(adc("UoShpCd")) = "S", " ", ", qRmCd ") + _
                ", TdTc, TdYy, TdChr, TdNo, TdSr "
    'Dust weight and BM Reason Code added
    wSqlStrg5 = " Select max(Grp1) as Grp1, max(Grp2) as Grp2, max(Grp3) as Grp3, " + _
                "DGrp1, DGrp2, DGrp3, max(G1Desc) as G1Desc, max(G2Desc) as G2Desc, max(G3Desc) as G3Desc, " + _
                "'T' as qTag, TdCoCd, TdBYy, TdBChr, TdBNo, 0.0 as qBMWt, 0 as qMul, 0 as qBMQty, 0.0 as qBRmWt, " + _
                "0.0 as qBRmQty, '01/01/80' as TDt, qRmCtg, qRmSCtg, max(qRmCd) as qRmCd, '' as qRmDc, '' as TdByWrk, '' as TdDustWt,'' as TdBmRsnCd," + _
                "'' as LocPrntCd, '' as TdByLoc, '' as qBMLoc, '' as TdTc, '' as TdYy, '' as TdChr, 0 as TdNo, 0 as TdSr, " + _
                "'' as DmCtg, '' as DmPrdCtg, '' as BOdDmCd, '' as OdKt, '' as BOdTc, '' as BOdYy, '' as BOdChr, " + _
                "'' as BOdNo, '' as OmCmCd, '' as OmPoNo, '' as qBM, '' As qQw, 0.0 as qRmVal, " + _
                "max(qRmWt) as qRmWt, max(qRmQty) as qRmQty " + _
                " From (Select max(Grp1) as Grp1, max(Grp2) as Grp2, max(Grp3) as Grp3, " + _
                "DGrp1, DGrp2, DGrp3, max(G1Desc) as G1Desc, max(G2Desc) as G2Desc, max(G3Desc) as G3Desc, " + _
                "TdCoCd, TdBYy, TdBChr, TdBNo, qRmCtg, qRmSCtg, max(qRmCd) as qRmCd, " + _
                "sum(qBRmWt) as qRmWt, sum(qBRmQty) as qRmQty, max(qTdYyKey) as qTdYyKey " + _
                "from #TMPB B " + ws_ShowAllPrd + _
                " group by DGrp1, DGrp2, DGrp3, TdCoCd, TdBYy, TdBChr, TdBNo, qRmCtg, qRmSCtg " + IIF(UCase(adc("UoShpCd")) = "S", " ", ", qRmCd ") + " " + _
                ", TdTc, TdYy, TdChr, TdNo, TdSr ) Tbl " + _
                "where qTdYyKey= (Select max(qTdYyKey) from #TMPB TB where TB.DGrp1= Tbl.DGrp1 and TB.DGrp2= Tbl.DGrp2 " + _
                "and TB.DGrp3= Tbl.DGrp3 and TB.TdBYy= Tbl.TdBYy and TB.TdBChr= Tbl.TdBChr and TB.TdBNo= Tbl.TdBNo " + _
                "and TB.qRmCtg= Tbl.qRmCtg and TB.qRMSCtg= Tbl.qRMSCtg " + IIF(UCase(adc("UoShpCd")) = "S", " ", " and TB.qRmCd= Tbl.qRmCd ") + ") " + _
                "group by DGrp1, DGrp2, DGrp3, TdCoCd, TdByy, TdBChr, TdBNo, " + _
                "qRmCtg, qRMSCtg " + IIF(UCase(adc("UoShpCd")) = "S", " ", ", qRmCd ")
    
    '*** (Jen 2.13) ***
  
  ElseIf UCase(adc.MenuCd) = UCase("RepBrkMsgPrd") Then
    
    '*** (Jen 2.13)
    wSqlStrg2 = "Select max(Grp1) as Grp1, max(Grp2) as Grp2, max(Grp3) as Grp3, " + _
                "DGrp1, DGrp2, DGrp3, max(G1Desc) as G1Desc, max(G2Desc) as G2Desc, max(G3Desc) as G3Desc, " + _
                "'P' as qTag, TdCoCd, TdBYy, TdBChr, TdBNo, 0.0 as qBMWt, 0 as qMul, 0 as qBMQty, 0.0 as qBRmWt, " + _
                "sum(qBRmQty) as qBRmQty, max(TDt) as TDt, qRmCtg, qRmSCtg, '' as qRmCd, '' as qRmDc, " + _
                "max(TdByWrk) as TdByWrk, max(TdDustWt) as TdDustWt, max(TdBmRsnCd) as TdBmRsnCd, max(LocPrntCd) as LocPrntCd, max(TdByLoc) as TdByLoc , '' as qBMLoc, TdTc, " + _
                "TdYy, TdChr, TdNo, TdSr, max(DmCtg) as DmCtg, max(DmPrdCtg) as DmPrdCtg, max(BOdDmCd) as BOdDmCd, " + _
                "max(OdKt) as OdKt, max(BOdTc) as BOdTc, max(BOdYy) as BOdYy, max(BOdChr) as BOdChr, " + _
                "max(BOdNo) as BOdNo, max(OmCmCd) as OmCmCd, max(OmPoNo) as OmPoNo, '' as qBM,'' As qQw, 0.0 as qRmVal "
    
    wSqlStrg3 = " from ( Select " + grpflds1 + ", " + _
                "Td.TdCoCd as TdCoCd, Td.TdBYy as TdBYy, Td.TdBChr as TdBChr, Td.TdBNo as TdBNo, " + _
                "TpRmQty as qBRmQty, TDt, TpRmCtg as qRmCtg, TpRmSCtg as qRmSCtg, " + _
                "Td.TdByWrk as TdByWrk, Td.TdDustWt as TdDustWt, Td.TdBmRsnCd as TdBmRsnCd, LocPrntCd," + IIF(OptLoc(0).Value = True, "Td.TdByLoc", "Lc.LocPrcs") + " as TdByLoc , Td.TdTc as TdTc, " + _
                "Td.TdYy as TdYy, Td.TdChr as TdChr, Td.TdNo as TdNo, Td.TdSr as TdSr, " + _
                "DmCtg, DmPrdCtg, BOdDmCd, OdKt, BOdTc, BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo " + _
                "from Txnd Td " + _
                "Join Txn on TIdNo=Td.TdTIdNo and TCoCd= Td.TdCoCd and TTc= Td.TdTc and TYy= Td.TdYy and TChr= Td.TdChr and TNo= Td.TdNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey=Td.TdPrtKey ", "") + _
                "Join TxndPrd on TpTdIdNo=Td.TdIdNo and TpCoCd= Td.TdCoCd and TpTc= Td.TdTc and TpYy= Td.TdYy " + _
                "and TpChr= Td.TdChr and TpNo= Td.TdNo and TpSr= Td.TdSr " + IIF(gs_Partition = ctCurrPrtn, " And TpPrtKey=Td.TdPrtKey ", "") + _
                "Join Loc Lc on LocCoCd=Td.TdCoCd And LocCd= Td.TdByLoc " + _
                "Join Bag on BIdNo=Td.TdBIdNo and BCoCd= Td.TdCoCd and BYy= Td.TdBYy and BChr= Td.TdBChr and BNo= Td.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=Td.TdPrtKey ", "") + _
                ws_BagJoin + _
                "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
                "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "") + _
                "Join OrdMst on OmIdNo=BOmIdNo and OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy and " + _
                "OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=BPrtKey ", "") + _
                "Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=OdPrtKey ", "") + _
                "where " + IIF(gs_Partition = ctCurrPrtn, " Td.TdPrtKey='" + ctCurrPrtn + "' And ", "") + " Td.TdSrNo= 0 and Td.TdPrdYN= 'Y' " + wCnd + _
                IIF(wAddlCndOth <> "", IIF(wCnd <> "", "", " Where ") + wAddlCndOth, "") + ") B " + ws_ShowAllPrd + _
                " group by  DGrp1, DGrp2, DGrp3, TdCoCd, TdBYy, TdBChr, TdBNo, qRmCtg, qRmSCtg, TdTc, TdYy, TdChr, TdNo, TdSr "
    
    'wAddlCndTd replaced by wAddlCndOth
    '*** (Jen 2.13)
    
    '*** (Bef 2.13)
    'wSqlStrg2 = "Select " + grpflds1 + ", " + _
    '            "'P' as qTag, Td.TdBYy as TdBYy, Td.TdBChr as TdBChr, Td.TdBNo as TdBNo, " + _
    '            "0.0 as qBMWt, 0 as qMul, 0 as qBMQty, 0.0 as qBRmWt, sum(TpRmQty) as qBRmQty, "
    '
    'wSqlStrg3 = " max(TDt) as TDt, TpRmCtg as qRmCtg, TpRmSCtg as qRmSCtg, '' as qRmCd, '' as qRmDc, " + _
    '            "max(Td.TdByWrk) as TdByWrk, max(Td.TdByLoc) as TdByLoc , '' as qBMLoc, Td.TdTc as TdTc, " + _
    '            "Td.TdYy as TdYy, Td.TdChr as TdChr, Td.TdNo as TdNo, Td.TdSr as TdSr, " + _
    '            "max(DmCtg) as DmCtg, max(DmPrdCtg) as DmPrdCtg, max(BOdDmCd) as BOdDmCd, max(OdKt) as OdKt, " + _
    '            "max(BOdTc) as BOdTc, max(BOdYy) as BOdYy, max(BOdChr) as BOdChr, max(BOdNo) as BOdNo, " + _
    '            "max(OmCmCd) as OmCmCd, max(OmPoNo) as OmPoNo, '' as qBM, 0.0 as qRmVal " + _
    '            "from Txnd Td " + _
    '            "Join Txn on TCoCd= Td.TdCoCd and TTc= Td.TdTc and TYy= Td.TdYy " + _
    '            "and TChr= Td.TdChr and TNo= Td.TdNo " + _
    '            "Join TxndPrd on TpCoCd= Td.TdCoCd and TpTc= Td.TdTc and TpYy= Td.TdYy " + _
    '            "and TpChr= Td.TdChr and TpNo= Td.TdNo and TpSr= Td.TdSr " + _
    '            "Join Bag on BCoCd= Td.TdCoCd and BYy= Td.TdBYy and BChr= Td.TdBChr and BNo= Td.TdBNo " + _
    '            ws_BagJoin + _
    '            "Join OrdDsg on OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
    '            "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + _
    '            "Join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy and " + _
    '            "OmChr= BOdChr and OmNo= BOdNo " + _
    '            "Join DsgMst on DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + _
    '            "where Td.TdSrNo= 0 and Td.TdPrdYN= 'Y' " + _
    '            wCnd + " group by Td.TdBYy, Td.TdBChr, Td.TdBNo, TpRmCtg, TpRmSCtg, Td.TdTc, " + _
    '            "Td.TdYy, Td.TdChr, Td.TdNo, Td.TdSr "
    '*** (Bef 2.13)
  End If
  
  '*** (Bef 2.13) wSqlStrg5 = " Order By TdByy, TdBChr, TdBNo, qRmCtg, qRMSCtg, qRmCd, TdTc, TdYy, TdChr, TdNo, TdSr "
  
  '*** (Jen 2.13)
  wSqlStrg4 = " Order By TdCoCd, TdByy, TdBChr, TdBNo, qRmCtg, qRMSCtg, qRmCd, TdTc, TdYy, TdChr, TdNo, TdSr "
  '*** (Jen 2.13)
  
  '*** (Bef 05/07/05)
  'wSqlStrg = wSqlStrg1 + " Union All " + wSqlStrg2 + " Union All " + wSqlStrg3 + wSqlStrg4 + wSqlStrg5
  '*** (Bef 05/07/05)
  
  
  '***(Jen 2.11) (05/07/05)
  'wSqlStrg = wSqlStrg1 + wSqlStrg2 + " Union All " + wSqlStrg3 + wSqlStrg4 + wSqlStrg5    '*** (Bef 2.13)
  '***(Jen 2.11) (05/07/05)
  
  '*** (Jen 2.13)
  If UCase(adc.MenuCd) = UCase("RepBrkMsgRm") Then
    wSqlStrg = wSqlStrg1 + " Union All " + wSqlStrg2 + wSqlStrg3 + " Union All " + wSqlStrg5 + wSqlStrg4
  ElseIf UCase(adc.MenuCd) = UCase("RepBrkMsgPrd") Then
    wSqlStrg = wSqlStrg1 + " Union All " + wSqlStrg2 + wSqlStrg3 + wSqlStrg4
  End If
  '*** (Jen 2.13)
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added
  mRep.TxtFrDt.SetText (adc("UoDtFr"))
  mRep.TxtToDt.SetText (adc("UoDtTo"))
  
  If adc("UoYN") = "N" Then
    mRep.BoxHd.Suppress = True
    mRep.BoxDet.Suppress = True
    mRep.BoxFt.Suppress = True
    mRep.DET.Suppress = True
    mRep.FldDet2BagNo.Suppress = False
    
    'mRep.FldDet2RmCtg.Suppress = False
    
    'mRep.FldDet2ByLoc.Suppress = False
    'mRep.FldDet2ByWrk.Suppress = False
    'mRep.FldDet2RmCd.Suppress = False
    'mRep.FldDet2RmSCtg.Suppress = False
    'inuuu
    If UCase(adc("UoShpCd")) = "S" Then
      mRep.FldDet2RmCd.Suppress = True
      mRep.FldDet2RmSCtg.Suppress = False
      'mRep.FldDet2ByWrk.Suppress = True
      'mRep.FldDet2ByLoc.Suppress = True
    ElseIf UCase(adc("UoShpCd")) = "R" Then
      mRep.FldDet2RmCd.Suppress = False
      mRep.FldDet2RmSCtg.Suppress = True
      'mRep.FldDet2ByWrk.Suppress = False
      'mRep.FldDet2ByLoc.Suppress = False
    End If
  Else
    mRep.FldDet2BagNo.Suppress = True
    mRep.FldDet2RmCtg.Suppress = True
    'mRep.FldDet2ByLoc.Suppress = True
    'mRep.FldDet2ByWrk.Suppress = True
    If UCase(adc.MenuCd) = UCase("RepBrkMsgRm") Then mRep.FldDet2RmCd.Suppress = True
    mRep.FldDet2RmSCtg.Suppress = True
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
'  If adc("UoBagTyp") <> "" Then mRep.TxtBagTypAndChr.SetText IIF(adc("UoBagTyp") = "Y", "(Flute Bags", "(Primary Bags") & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
'  '***(Jen 2.11)
'*** Jay 2.13(CT) ***
  
  CRV_REP.DisplayGroupTree = False

  '*** (Jen 2.12) (This works. The actual groups for which the totals have to be suppressed have yet to be decided)
  Dim wo_Grp As New Collection
  If UCase(adc.MenuCd) = UCase("RepBrkMsgRm") Then
    ''AddItem wo_Grp, "PrcWt", "Rm Code"
    'AddItem wo_Grp, "", "Rm Code"
    'AddItem wo_Grp, "", "Rm Sub Ctg"
    'AddItem wo_Grp, "", "Rm StkRt"
    
    AddItem wo_Grp, "", "Process"
    AddItem wo_Grp, "", "(Process)Worker"
    Call SuppressInvalidGrpTotals(Rep, GRP_REP, wo_Grp, "PrcWt,PrcQty")
  End If
  '*** (Jen 2.12) (This works. The actual groups for which the totals have to be suppressed have yet to be decided)

  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******

'/** selecting Los where TdRmFrLoc in wBM ('B','M') **/
'select 'D' as qTag, TdBYy, TdBChr, TdBNo,
'  TdRmWt as qBMWt,
'     (Case when TdFrRmDc='D' then 1
'    when TdFrRmDc='C' then -1
'          Else 0 end) as qMul,
'  TdRmQty as qBMQty, 0.0 as qBRmWt, 0 as qBRmQty,
'  TDt, RmCtg as qRmCtg, RmSCtg as qRmSCtg,  TdRmCd as qRmCd, TdFrRmDc as qRmDc, TdByWrk, TdByLoc , TdFrRmLoc as qBMLoc,
'  TdTc, TdYy, TdChr, TdNo, TdSr, DmCtg, DmPrdCtg,
'        BOdDmCd , OdKt, BOdTc, BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo
'from Txnd Td
'Join Txn on TTc=TdTc and TYy=TdYy and TChr=TdChr and TNo=TdNo
'Join RmMst on RmCd=TdRmCd
'Join Bag on BYy= TdBYy and BChr= TdBChr and BNo= TdBNo
'Join OrdDsg on OdTc= BOdTc and OdYy= BOdYy and OdChr= BOdChr
'and OdNo= BOdNo and OdSr= BOdSr
'Join OrdMst on OmTc= BOdTc and OmYy= BOdYy and OmChr= BOdChr and OmNo= BOdNo
'Join DsgMst on DmTcTyp= 'DM' and DmCd= BOdDmCd and DmSz= ''
'where (select LocTyp from Loc where LocCd=TdFrRmLoc) in ('B','M') and RmCtg = 'D'
'and TdSrNo<>0   and Tdt >= '01/11/00' and Tdt <='01/11/00' and TdByLoc='PFIL'
'
'Union All
'
'/** selecting Loss where TdRmToLoc in wBM ('B','M') **/
'select 'D' as qTag, TdBYy, TdBChr, TdBNo,
'  TdRmWt as qBMWt,
'     (Case when  TdToRmDc='D' then 1
'    when  TdToRmDc='C' then -1
'          Else 0 end) as qMul,
'  TdRmQty as qBMQty, 0.0 as qBRmWt, 0 as qBRmQty,
'  TDt, RmCtg as qRmCtg, RmSCtg as qRmSCtg, TdRmCd as qRmCd, TdToRmDc as qRmDc, TdByWrk, TdByLoc , TdToRmLoc as qBMLoc,
'  TdTc , TdYy, TdChr, TdNo, TdSr
'from Txnd Td
'Join Txn on TTc=TdTc and TYy=TdYy and TChr=TdChr and TNo=TdNo
'Join RmMst on RmCd=TdRmCd
'Join Bag on BYy= TdBYy and BChr= TdBChr and BNo= TdBNo
'Join OrdDsg on OdTc= BOdTc and OdYy= BOdYy and OdChr= BOdChr
'and OdNo= BOdNo and OdSr= BOdSr
'Join OrdMst on OmTc= BOdTc and OmYy= BOdYy and OmChr= BOdChr and OmNo= BOdNo
'Join DsgMst on DmTcTyp= 'DM' and DmCd= BOdDmCd and DmSz= ''
'where (select LocTyp from Loc where LocCd=TdToRmLoc) in ('B','M') and  RmCtg= 'D'
'  and TdSrNo<>0 and Tdt >= '01/11/00' and Tdt <='01/11/00'  and TdByLoc='PFIL'
'
'Union All
'
'select 'P' as qTag, Td.TdBYy as TdBYy, Td.TdBChr as TdBChr, Td.TdBNo as TdBNo,
'  0.0 as qBMWt, 0 as qMul, 0 as qBMQty,
'  Sum(isnull(Tr.TdRmWt*
'    (Case when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='D' then 1
'      when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='C' then -1
'      when Tr.TdToRmLoc='W' and Tr.TdToRmDc='D' then 1
'      when Tr.TdToRmLoc='W' and Tr.TdToRmDc='C' then -1
'          Else 0 end),0.0))  as qBRmWt,
'  Sum(isnull(Tr.TdRmQty*
'    (Case when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='D' then 1
'      when Tr.TdFrRmLoc='W' and Tr.TdFrRmDc='C' then -1
'      when Tr.TdToRmLoc='W' and Tr.TdToRmDc='D' then 1
'      when Tr.TdToRmLoc='W' and Tr.TdToRmDc='C' then -1
'          Else 0 end),0))  as qBRmQty,
'  max(TDt), RmCtg as qRmCtg, RmSCtg as qRmSCtg, RmCd as qRmCd, '' as qRmDc, max(Td.TdByWrk), max(Td.TdByLoc) , '' as qBMLoc,
'  Td.TdTc as TdTc, Td.TdYy as TdYy, Td.TdChr as TdChr, Td.TdNo as TdNo, Td.TdSr as TdSr
'from Txnd Td
'Join Txn on TTc=Td.TdTc and TYy=Td.TdYy and TChr=Td.TdChr and TNo=Td.TdNo
'Join Txnd Tr on Tr.TdBYy=Td.TdByy and Tr.TdBChr = Td.TdBChr and Tr.TdBNo = Td.TdBNo
'       and Tr.TDYyKey <= Td.TdYyKey  and Tr.TdSrno<> 0
'Join RmMst on  RmCd=Tr.TdRmCd
'Join Bag on BYy= Td.TdBYy and BChr= Td.TdBChr and BNo= Td.TdBNo
'Join OrdDsg on OdTc= BOdTc and OdYy= BOdYy and OdChr= BOdChr
'and OdNo= BOdNo and OdSr= BOdSr
'Join OrdMst on OmTc= BOdTc and OmYy= BOdYy and OmChr= BOdChr and OmNo= BOdNo
'Join DsgMst on DmTcTyp= 'DM' and DmCd= BOdDmCd and DmSz= ''
'where Td.TdSrNo=0 and Td.TdPrdYN='Y' and Tdt >= '01/11/00' and Tdt <='30/11/00'
'and Td.TdFrBLoc = 'PFIL' and RmCtg ='D'
'group by  Td.TdByy, Td.TdBChr, Td.TdBNo,RmCtg, RMSCtg, RmCd, Td.TdTc, Td.TdYy, Td.TdChr, Td.TdNo, Td.TdSr
'order By  TdByy, TdBChr, TdBNo, qRmCtg, qRMSCtg, qRmCd, TdTc, TdYy, TdChr, TdNo, TdSr
Exit Sub

RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1
    If .Value = Checked Then
      If ms_PwdFlag = False Then
        ChkBoxArr(1).Value = Unchecked: adc("UoYN") = "N": DispMsg "Enter SEO Password To Show Value", etError
      Else
        adc("UoYN1") = "Y"
      End If
    Else
      adc("UoYN1") = "N"
    End If
  Case Is = 3
    If .Value = Checked Then adc("UoYN3") = "Y" Else adc("UoYN3") = "N"
  End Select
End With
End Sub
' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added

' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added

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
    Set adc.FirNKeyCtl = adc("UORMCTGFR")
  End If
  
  LblPrntCd.Visible = ws_MultiCoMod: adc("UoPrntCdFr").Visible = ws_MultiCoMod
  adc("UoPrntCdTo").Visible = ws_MultiCoMod: adc("UoPrntCdSel").Visible = ws_MultiCoMod
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
 
  Call InitProp(Me)
  
  If UCase(adc.MenuCd) = UCase("RepBrkMsgRm") Then
    LblUoShpCd.Visible = True     'adc("UoShpCd").Visible = True: adc("UoShpCd").Enabled = True
    FraOptGrpBy.Enabled = True: FraOptGrpBy.Visible = True  ' ******** Manali 3.5.0 - 17/11/08 - Radio button & CheckBoxes added
    adc("UoRmCdFr").Visible = True: adc("UoRmCdFr").Enabled = True
    adc("UoRmCdTo").Visible = True: adc("UoRmCdTo").Enabled = True
    adc("UoRmCdSel").Visible = True: adc("UoRmCdSel").Enabled = True
    LblRmCd.Visible = True
    adc("UoStkRtFr").Visible = True: adc("UoStkRtFr").Enabled = True
    adc("UoStkRtTo").Visible = True: adc("UoStkRtTo").Enabled = True
    LblStkRt.Visible = True
    
    'Sachin 4.1.2
    adc("UoParentCdFr").Visible = True: adc("UoParentCdFr").Enabled = True
    adc("UoParentCdTo").Visible = True: adc("UoParentCdTo").Enabled = True
    adc("UoParentCdSel").Visible = True: adc("UoParentCdSel").Enabled = True
    LblParentCd.Visible = True
    
  ElseIf UCase(adc.MenuCd) = UCase("RepBrkMsgPrd") Then
    LblUoShpCd.Visible = False      'adc("UoShpCd").Visible = False: adc("UoShpCd").Enabled = False:
    FraOptGrpBy.Enabled = False: FraOptGrpBy.Visible = False:    ' ******** Manali 3.5.0 - 17/11/08 - Radio button & CheckBoxes added
    adc("UoRmCdFr").Visible = True: adc("UoRmCdFr").Enabled = True
    adc("UoRmCdFr").Visible = False: adc("UoRmCdFr").Enabled = False
    adc("UoRmCdTo").Visible = False: adc("UoRmCdTo").Enabled = False
    adc("UoRmCdSel").Visible = False: adc("UoRmCdSel").Enabled = False
    LblRmCd.Visible = False
    adc("UoStkRtFr").Visible = False: adc("UoStkRtFr").Enabled = False
    adc("UoStkRtTo").Visible = False: adc("UoStkRtTo").Enabled = False
    LblStkRt.Visible = False
    ' Zubin 213 (Err correction, Value is not shown in BrkMsgPrd)
    LblShowVal.Visible = False: adc("UoYn1").Visible = False
    LblValTyp.Visible = False: adc("UoPmCd").Visible = False
    CmdPwd.Visible = False: FraPwd.Visible = False
    ChkBoxArr(1).Enabled = False: ChkBoxArr(1).Visible = False    ' ******** Manali 3.5.0 - 17/11/08 - Radio button & CheckBoxes added
    
    'Sachin 4.1.2
    adc("UoParentCdFr").Visible = False: adc("UoParentCdFr").Enabled = False
    adc("UoParentCdTo").Visible = False: adc("UoParentCdTo").Enabled = False
    adc("UoParentCdSel").Visible = False: adc("UoParentCdSel").Enabled = False
    LblParentCd.Visible = False
    
    ' Zubin 213
  End If
  
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
  Call SetGroupSort
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
      adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
  Else
      adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  '*** Jay 2.14 *** (SEO Change)
  ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
  '*** Jay 2.14 *** (SEO Change)
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
'/* urmi 2.06 */

  gb_CoCdFor = False '****** Sachin 2.14.0 - [Multi-Company]

  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr"), UCase("UoPMCdTo"), UCase("UoPMCdSel")
    gs_Tbl = "vParam"
    gs_PTyp = "PRDSEQ"
    
  Case Is = UCase("UoPMCd")
    gs_PTyp1 = "WIPVAL"
    If UCase(IdName) = UCase("UoPmCd") And pv_NewValue <> "STKVAL" Then adc("UoPMCd1") = "" '4.1.4

  Case Is = UCase("UoYn1")
    '*** Jay 2.14 *** (SEO Change)
    'If UCase(pv_NewValue) = "Y" And Trim(TxtPwd.text) <> moCn.GetFldVal("Select PValue1 from Param where " + _
                                   "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
    If UCase(pv_NewValue) = "Y" And ms_PwdFlag = False Then _
      Cancel = True: DispMsg "Enter SEO Password to View Values", etError
    '*** Jay 2.14 *** (SEO Change)
    
  End Select
'/* urmi 2.06 */

  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_LocTyp = "'P'"
  gb_CoCdFor = False '****** Sachin 2.14.0 - [Multi-Company]

 Select Case UCase(IdName)
  Case Is = UCase("UOPMCD")
    '*** Jay 2.14 *** (SEO Change)
    ' ****** Manali 3.5.0 - ChkBox added - ErrMsg Changed
    If adc("UoYn1") <> "Y" Then Cancel = True: ErrMsg = "Cannot Edit When Show Value Option Is Not Checked": Exit Sub
    'If adc("UoYn1") <> "Y" Then Cancel = True: ErrMsg = "Can Not Edit When Show Value Option Is Not Yes": Exit Sub
    'If adc("UoYn1") = "N" Then Cancel = True: ErrMsg = "Cannot Edit When Show Value is Set To NO"
    '*** Jay 2.14 *** (SEO Change)
    gs_PTyp1 = "WIPVAL"
    
  Case Is = UCase("UOPMCD1")    '4.1.4
    If adc("UoPMCd") <> "STKVAL" Then Cancel = True: ErrMsg = "Cannot Edit When Value Type is not STKVAL": Exit Sub
    gs_PTyp = "WIPVAL"
    
  '*** (Jen 2.11)
  Case Is = UCase("UoBChrFr")
    If adc.Mode = xNorm Then
      If adc("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Flute Bag Y/N.": Exit Sub
      HlpList.PMCd "BCHR", "'" + adc("UoBagTyp") + "'": Exit Sub
    End If
  '*** (Jen 2.11)
  Case Is = UCase("UOISSLOCFR"), UCase("UOISSLOCTO"), UCase("UOISSLOCSEL")
        If adc("UOLOCTYPFR") = "R" Then gs_LocTyp = "'R'"
        If adc("UOLOCTYPFR") = "P" Then gs_LocTyp = "'P'"
  
 End Select
  
 SetProp Me, IdName, When
  
'/* urmi 2.06 */
 Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr"), UCase("UoPMCdSel")
    gs_Tbl = "vParam"
    gs_PTyp = "PRDSEQ"
    HlpList.vPMCd adc("UoCoCdFr"), gs_PTyp
  Case Is = UCase("UoPMCdTo")
    gs_Tbl = "vParam"
    gs_PTyp = "PRDSEQ"
    HlpList.vPMCd adc("UoCoCdTo"), gs_PTyp
 End Select
'/* urmi 2.06 */
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case IdName
  Case Is = UCase("UoPmCd")
    If adc("UoYn1") <> "Y" Then adc("UoPmCd") = ""
    'dsffs
  Case Is = UCase("UOBCHRFR")
    'If ADC("UoBagTyp") = "N" Then
    adc("UOBCHRFR") = ""
  End Select
End Sub
'/* urmi 2.06 */
Private Sub CmdPwd_Click()
     If FraPwd.Visible = True Then
      FraPwd.Visible = False
      FraPwd.Enabled = False
      'CmdPwd.enable = True
      '*** Jenny (Old Code Bef Color)
'      CmdPwd.BackColor = &H8000000F
'      CmdPwd.FontBold = False
      '*** Jenny (Old Code Bef Color)
      CmdPwd.SetFocus
    Else
      FraPwd.Visible = True
      FraPwd.Enabled = True
      FraPwd.ZOrder
      '*** Jenny (Old Code Bef Color)
'      CmdPwd.BackColor = vbWhite
'      CmdPwd.FontBold = True
      '*** Jenny (Old Code Bef Color)
      TxtPwd.SetFocus
    End If
End Sub
'/* urmi 2.06 */
Private Sub HideAllFras()
  '*** This procedure hides all the frames
  FraPwd.Visible = False
  FraPwd.Enabled = False
End Sub

' ***** Manali 3.5.0 - 17/11/08 - Radio Buttons Added
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
' ***** Manali 3.5.0 - 17/11/08 - Radio Buttons Added

' ***** Manali 3.5.0 - 17/11/08 - Radio Buttons Added
Private Sub OptBrkMsg_GotFocus(Index As Integer)
  DispMsg FraOptBrkMsg.ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 17/11/08 - Radio Buttons Added

' ***** Manali 3.5.0 - 17/11/08 - Radio Buttons Added
Private Sub OptGrpBy_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoShpCd") = "R"
Case Is = 1
  adc("UoShpCd") = "S"
End Select
End Sub
' ***** Manali 3.5.0 - 17/11/08 - Radio Buttons Added

' ***** Manali 3.5.0 - 17/11/08 - Radio Buttons Added
Private Sub OptGrpBy_GotFocus(Index As Integer)
DispMsg FraOptGrpBy.ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 17/11/08 - Radio Buttons Added

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

'/* urmi 2.06 */
Private Sub TxtPwd_Validate(Cancel As Boolean)
'*** Jay 2.14 *** (SEO Change)
    If Trim(TxtPwd.text) = "" Then
      adc("UoYn1") = ""
      ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
    Else
      If Trim(TxtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                     "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + adc.MenuCd + "' and " + _
                     "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                     "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
          ms_PwdFlag = True: Exit Sub
      Else
          ms_PwdFlag = False
          Cancel = True:  DispMsg "Invalid Password / User Does Not Have Seo Rights ", etError
      End If
    End If
'*** Jay 2.14 *** (SEO Change)
'*** bef 2.14
'    If Trim(TxtPwd.text) = "" Then adc("UoYn1") = ""
'    If Trim(TxtPwd.text) = "" Or Trim(TxtPwd.text) = moCn.GetFldVal("Select PValue1 from Param where " + _
'                           "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'      Exit Sub
'    Else
'      Cancel = True
'      DispMsg "Invalid Password", etError
'    End If
'*** bef 2.14
End Sub
Private Sub tab_REP_Click(PreviousTab As Integer)
  fra_tabrep(TAB_REP.Tab).Enabled = True
  fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  adc("UOUSRCD") = adc.UsrCd
  adc("UOMNUCD") = adc.MenuCd
  Call HideAllFras
  
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)

  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = False Then _
      adc("UoCoCdFr") = gs_CoCd: adc("UoCoCdTo") = gs_CoCd: adc("UoCoCdSel") = "": adc("UoCoCdFr").Enabled = False
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
  adc("UoYn1") = ""
  adc("UoPMCD") = ""
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
'    adc("UoBagTyp") = "N": adc("UoBChrFr") = ""
'  End If
'  '*** (Jen 2.11)
'*** Jay 2.13(CT) ***

  ' ****** Manali 3.5.0 - 17/11/08 - Check Boxes and Radio Buttons Added
  If adc("UoBrkMsg") = "B" Then
    OptBrkMsg(0).Value = True
  ElseIf adc("UoBrkMsg") = "M" Then
    OptBrkMsg(1).Value = True
  ElseIf adc("UoBrkMsg") = "" Then
    OptBrkMsg(2).Value = True
  End If
  
  If adc("UoShpCd") = "R" Then
    OptGrpBy(0).Value = True
  ElseIf adc("UoShpCd") = "S" Then
    OptGrpBy(1).Value = True
  End If
  
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  If adc("UoYN3") = "" Then adc("UoYN3") = "N"

  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If adc("UoYN3") = "Y" Then ChkBoxArr(3).Value = Checked Else ChkBoxArr(3).Value = Unchecked
  ' ****** Manali 3.5.0 - 17/11/08 - Check Boxes And Radio Buttons Added
  If adc("UOLOCTYPFR") = "P" Then
    OptLoc(0).Value = True
  ElseIf adc("UOLOCTYPFR") = "R" Then
    OptLoc(1).Value = True
  End If

End Sub
Private Sub DispCoNm()
  'mRep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = True, gs_CoCd, adc("UOCOCDFR")) + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
  mRep.wCoCd.SetText ws_HName
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  mRep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    mRep.RHMain.UnderlaySection = True: mRep.wCoCd.Suppress = True
  Else
    mRep.RHLogo.Suppress = True
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
    
  '****  Manali - before 3.2.2
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  mRep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  mRep.wSrvrTime.SetText ws_DtTm
  
  '***** Manali 3.2.2 user date and time in report
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
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set mRep = Nothing
    Set Rep = Nothing
    Set Rep1 = Nothing
    '*** (09/08/05)
End Sub

Private Sub SuppressInvalidGrpTotals(ByVal po_Rep As CRAXDRT.Report, ByVal po_GrpRep As MWCTL_GRP, ByRef po_Grp As Collection, ByVal ps_FldNames As String)
  '*** (Jen 2.12)
  Dim ws_Grp(3) As String, wi_Lvl(3) As Integer, ws_FldArr() As String
  Dim wi_i As Integer, wi_j As Integer, wi_k As Integer

  On Error GoTo ErrH

  ws_Grp(1) = po_GrpRep.Value(gltGroup1)
  ws_Grp(2) = po_GrpRep.Value(gltGroup2)
  ws_Grp(3) = po_GrpRep.Value(gltGroup3)
  
  ws_FldArr = Split(ps_FldNames, ",")
  If Not IsArray(ws_FldArr) Then Exit Sub
  
  Call GetLevel(ws_Grp, wi_Lvl)
  
  For wi_i = 1 To 3
    '*** (Note That The Names Of The Header/ Footer Sections Of The 3 Groups Were Already Hard-Coded)
    If isExists(po_Grp, ws_Grp(wi_i)) = True And wi_Lvl(wi_i) > 0 Then
      '*** The Totals Of The Corresponding Field Will Be Suppressed For All Groups Above This Group
      For wi_j = wi_Lvl(wi_i) - 1 To 1 Step -1
        '*** (02/01/06)
        For wi_k = 0 To UBound(ws_FldArr)
            If Trim(ws_FldArr(wi_k)) <> "" Then po_Rep.Sections("G" + CStr(wi_j) + "F").ReportObjects("G" + CStr(wi_j) + Trim(ws_FldArr(wi_k))).Suppress = True
        Next wi_k
        '*** (02/01/06)
      
        '*** (Bef 02/01/06)
        'po_Rep.Sections("G" + CStr(wi_j) + "F").ReportObjects("G" + CStr(wi_j) + po_Grp.Item(ws_Grp(wi_i))).Suppress = True
        '*** (Bef 02/01/06)
      Next wi_j
      '*** The Report Footer Section For Displaying Totals Has Been Hard-Coded As 'RF'
      '    The Grand Total Will Always Be Suppressed
      '*** (Bef 02/01/06)
      'po_Rep.Sections("RF").ReportObjects("G" + po_Grp.Item(ws_Grp(wi_i))).Suppress = True
      '*** (Bef 02/01/06)
      '*** (02/01/06)
      For wi_k = 0 To UBound(ws_FldArr)
          If Trim(ws_FldArr(wi_k)) <> "" Then po_Rep.Sections("RF").ReportObjects("G" + Trim(ws_FldArr(wi_k))).Suppress = True
      Next wi_k
      '*** (02/01/06)
    
    End If
  Next wi_i
  Exit Sub

ErrH:
  DispMsg "Error In Group Field Suppression: " & Err.Description, etError

  '*** (Jen 2.12)
End Sub
Private Sub GetLevel(ByRef ps_Grp() As String, ByRef pi_Lvl() As Integer)
  '*** (Jen 2.12)
  Const ctNone = "(None)"
  Const ctDetail = "(Detail)"
  Const ctGrand = "(Grand)"
  
  Dim wi_i As Integer, wi_j As Integer, ws_LstLvl As String, wi_Levels As Integer
  Dim wf_OneLvl As Boolean, wf_TwoLvl As Boolean, wf_ThreeLvl As Boolean

  ws_LstLvl = ""
  For wi_i = 1 To 3
    If ps_Grp(wi_i) = ctGrand Then
      ws_LstLvl = ctGrand: Exit For
    End If
    
    If ps_Grp(wi_i) = ctDetail Then
      ws_LstLvl = ctDetail: Exit For
    End If
  Next wi_i
  
  wf_OneLvl = (ps_Grp(1) <> ctNone And ps_Grp(2) = ctNone And ps_Grp(3) = ctNone)
  wf_TwoLvl = (ps_Grp(1) <> ctNone And ps_Grp(2) <> ctNone And ps_Grp(3) = ctNone)
  wf_ThreeLvl = (ps_Grp(1) <> ctNone And ps_Grp(2) <> ctNone And ps_Grp(3) <> ctNone)
  
  For wi_i = 1 To 3
    pi_Lvl(wi_i) = -1
  Next wi_i
  
  wi_Levels = IIF(wf_OneLvl, 1, IIF(wf_TwoLvl, 2, IIF(wf_ThreeLvl, 3, 0)))
  If ws_LstLvl = ctDetail Or ws_LstLvl = ctGrand Then wi_Levels = wi_Levels - 1

  If wi_Levels > 0 Then
    If ws_LstLvl = ctDetail Then
      For wi_i = 1 To wi_Levels
        pi_Lvl(wi_i) = wi_i
      Next wi_i
    Else
      wi_j = wi_Levels
      For wi_i = 3 To (3 - wi_Levels) + 1 Step -1
        pi_Lvl(wi_j) = wi_i
        wi_j = wi_j - 1
      Next wi_i
    End If
  End If
  '*** (Jen 2.12)
End Sub
