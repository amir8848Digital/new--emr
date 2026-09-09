VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpDsgLst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Design Listing"
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
      TabIndex        =   54
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9660
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
      Left            =   11550
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   9540
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4455
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5415
      TabIndex        =   68
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   69
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
         TabIndex        =   70
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
      Left            =   6630
      TabIndex        =   67
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   64
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
      TabIndex        =   58
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
      Height          =   9795
      Left            =   0
      TabIndex        =   63
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   62
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpDsgLst.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpDsgLst.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   66
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   61
               Top             =   360
               Width           =   8745
               _ExtentX        =   15425
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   65
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
               Index           =   2
               Left            =   30
               TabIndex        =   10
               Tag             =   "DmHold"
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   1560
               Width           =   1950
            End
            Begin VB.Frame FraOptDmValid 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1785
               TabIndex        =   94
               ToolTipText     =   "Show Valid Or Invalid Or All Designs"
               Top             =   6120
               Width           =   3465
               Begin VB.OptionButton OptValid 
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
                  TabIndex        =   48
                  ToolTipText     =   "Show Valid Or Invalid Or All Designs"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptValid 
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
                  Left            =   1380
                  TabIndex        =   47
                  ToolTipText     =   "Show Valid Or Invalid Or All Designs"
                  Top             =   0
                  Width           =   945
               End
               Begin VB.OptionButton OptValid 
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
                  TabIndex        =   46
                  ToolTipText     =   "Show Valid Or Invalid Or All Designs"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show 2nd Line  "
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
               Index           =   0
               Left            =   30
               TabIndex        =   50
               Tag             =   "ShowLine"
               ToolTipText     =   "Check To Show 2nd Line"
               Top             =   6405
               Width           =   1950
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Dsg On Hold   "
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
               TabIndex        =   45
               Tag             =   "DmHold"
               ToolTipText     =   "Check To Show Design on Hold"
               Top             =   5835
               Width           =   1950
            End
            Begin VB.Frame FraOptRm 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1785
               TabIndex        =   88
               ToolTipText     =   "Select Production Rm Qty, Wt or Sales Rm Qty, Wt"
               Top             =   6915
               Width           =   3975
               Begin VB.OptionButton OptRmPrd 
                  Caption         =   "Prd Rm"
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
                  TabIndex        =   52
                  ToolTipText     =   "Select Production Rm Qty, Wt or Sales Rm Qty, Wt"
                  Top             =   0
                  Width           =   1245
               End
               Begin VB.OptionButton OptRmPrd 
                  Caption         =   "Sales Rm"
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
                  Left            =   1380
                  TabIndex        =   53
                  ToolTipText     =   "Select Production Rm Qty, Wt or Sales Rm Qty, Wt"
                  Top             =   0
                  Width           =   1455
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4245
               TabIndex        =   23
               ToolTipText     =   "Enter To Customer"
               Top             =   3240
               Visible         =   0   'False
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
               Left            =   6750
               TabIndex        =   24
               ToolTipText     =   "Enter Customer Selection"
               Top             =   3240
               Visible         =   0   'False
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
               Index           =   42
               Left            =   1785
               TabIndex        =   22
               ToolTipText     =   "Enter From Customer"
               Top             =   3240
               Visible         =   0   'False
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
               Index           =   28
               Left            =   4245
               TabIndex        =   2
               ToolTipText     =   "Enter To Design Category"
               Top             =   645
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
               Left            =   6750
               TabIndex        =   3
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   645
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
               Left            =   1785
               TabIndex        =   1
               ToolTipText     =   "Enter From Design Category"
               Top             =   645
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
               Left            =   4245
               TabIndex        =   12
               ToolTipText     =   "Enter To Design Code"
               Top             =   2100
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "DmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   6750
               TabIndex        =   13
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2100
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "DmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1785
               TabIndex        =   11
               ToolTipText     =   "Enter From Design Code"
               Top             =   2100
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "DmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4245
               TabIndex        =   5
               ToolTipText     =   "Enter To Design Production Category"
               Top             =   930
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
               Left            =   6750
               TabIndex        =   6
               ToolTipText     =   "Enter Design Production Category Selection"
               Top             =   930
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
               Left            =   1785
               TabIndex        =   4
               ToolTipText     =   "Enter From Design Production Category"
               Top             =   930
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
               Index           =   15
               Left            =   4245
               TabIndex        =   20
               ToolTipText     =   "Enter To Karat"
               Top             =   2955
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTTO"
               IdName          =   "UOKTTO"
               CmpStr          =   "DmKt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   6750
               TabIndex        =   21
               ToolTipText     =   "Enter Karat Selection"
               Top             =   2955
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "DmKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1785
               TabIndex        =   19
               ToolTipText     =   "Enter From Karat"
               Top             =   2955
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               Data            =   "ABCDW"
               CmpStr          =   "DmKt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   4245
               TabIndex        =   15
               ToolTipText     =   "Enter To Design Size"
               Top             =   2385
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
               CmpStr          =   "DmSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1785
               TabIndex        =   14
               ToolTipText     =   "Enter From Design Size"
               Top             =   2385
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "DmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   4245
               TabIndex        =   17
               ToolTipText     =   "Enter To Design Colour"
               Top             =   2670
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UODMCOLTO"
               IdName          =   "UODMCOLTO"
               CmpStr          =   "DmCol <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   6750
               TabIndex        =   18
               ToolTipText     =   "Enter Design Colour Selection"
               Top             =   2670
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCOLSEL"
               IdName          =   "UODMCOLSEL"
               CmpStr          =   "DmCol In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   1785
               TabIndex        =   16
               ToolTipText     =   "Enter From Design Colour"
               Top             =   2670
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UODMCOLFR"
               IdName          =   "UODMCOLFR"
               CmpStr          =   "DmCol >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   2415
               TabIndex        =   49
               ToolTipText     =   "Show Second Line Of Detail? ('N', 'Y', Blank)"
               Top             =   6450
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
               Left            =   4245
               TabIndex        =   26
               ToolTipText     =   "Enter To Set Family Code"
               Top             =   3525
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOSETCDTO"
               IdName          =   "UOSETCDTO"
               CmpStr          =   "DmSetCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   6750
               TabIndex        =   27
               ToolTipText     =   "Enter Set Family Code Selection"
               Top             =   3525
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSETCDSEL"
               IdName          =   "UOSETCDSEL"
               CmpStr          =   "DmSetCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1785
               TabIndex        =   25
               ToolTipText     =   "Enter From Set Family Code"
               Top             =   3525
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOSETCDFR"
               IdName          =   "UOSETCDFR"
               CmpStr          =   "DmSetCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4245
               TabIndex        =   29
               ToolTipText     =   "Enter To Loss Category "
               Top             =   3810
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOLSCTGTO"
               IdName          =   "UOLSCTGTO"
               CmpStr          =   "DmLsCtg <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   6750
               TabIndex        =   30
               ToolTipText     =   "Enter Loss Category Selection"
               Top             =   3810
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOLSCTGSEL"
               IdName          =   "UOLSCTGSEL"
               CmpStr          =   "DmLsCtg In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   1785
               TabIndex        =   28
               ToolTipText     =   "Enter From Loss Category "
               Top             =   3810
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOLSCTGFR"
               IdName          =   "UOLSCTGFR"
               CmpStr          =   "DmLsCtg >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4245
               TabIndex        =   8
               ToolTipText     =   "Enter To Design Sales Category"
               Top             =   1215
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
               Index           =   9
               Left            =   6750
               TabIndex        =   9
               ToolTipText     =   "Enter Design Sales Category Selection"
               Top             =   1215
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMSALCTGSEL"
               IdName          =   "UODMSALCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1785
               TabIndex        =   7
               ToolTipText     =   "Enter From Design Sales Category"
               Top             =   1215
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
               Index           =   19
               Left            =   2415
               TabIndex        =   44
               ToolTipText     =   "Show Designs On Hold? (Yes / No)"
               Top             =   5835
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
               CmpStr          =   "DmHld = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   5955
               TabIndex        =   51
               ToolTipText     =   "Show Production Rm Qty, Wt ('Y' For Prd/ 'N' For Ord)"
               Top             =   6915
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
               Index           =   17
               Left            =   1785
               TabIndex        =   0
               ToolTipText     =   "Enter Design Type"
               Top             =   360
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCD"
               IdName          =   "UOPMCD"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   1785
               TabIndex        =   35
               ToolTipText     =   "Enter From Design Date"
               Top             =   4665
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "DmDsgDt>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   4245
               TabIndex        =   36
               ToolTipText     =   "Enter To Design Date"
               Top             =   4665
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "DmDsgDt<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   4245
               TabIndex        =   32
               ToolTipText     =   "Enter To Model Maker"
               Top             =   4095
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UODMMODMKRTO"
               IdName          =   "UODMMODMKRTO"
               CmpStr          =   "DmModMkr <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   1785
               TabIndex        =   31
               ToolTipText     =   "Enter From Model Maker"
               Top             =   4095
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UODMMODMKRFR"
               IdName          =   "UODMMODMKRFR"
               CmpStr          =   "DmModMkr >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   4245
               TabIndex        =   34
               ToolTipText     =   "Enter To Design By"
               Top             =   4380
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UODMDSGBYTO"
               IdName          =   "UODMDSGBYTO"
               CmpStr          =   "DmDsgBy <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   1785
               TabIndex        =   33
               ToolTipText     =   "Enter From Design By"
               Top             =   4380
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UODMDSGBYFR"
               IdName          =   "UODMDSGBYFR"
               CmpStr          =   "DmDsgBy >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   4245
               TabIndex        =   38
               ToolTipText     =   "Enter To Parent Code"
               Top             =   4950
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOPARENTCDTO"
               IdName          =   "UOPARENTCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   1785
               TabIndex        =   37
               ToolTipText     =   "Enter From Parent Code"
               Top             =   4950
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOPARENTCDFR"
               IdName          =   "UOPARENTCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   6750
               TabIndex        =   39
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   4950
               Width           =   7980
               _ExtentX        =   14076
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   5595
               TabIndex        =   95
               ToolTipText     =   "Show Valid Designs (Yes / No / All)"
               Top             =   6120
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
               Index           =   110
               Left            =   1785
               TabIndex        =   40
               ToolTipText     =   "Enter From Diamond Weight"
               Top             =   5235
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   503
               Mask            =   "#00.00"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOGRSWTFR"
               IdName          =   "UOGRSWTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   4245
               TabIndex        =   41
               ToolTipText     =   "Enter To Diamond Weight"
               Top             =   5235
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   503
               Mask            =   "#00.00"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOGRSWTTO"
               IdName          =   "UOGRSWTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   128
               Left            =   1785
               TabIndex        =   42
               ToolTipText     =   "Enter From Colour Stone Weight"
               Top             =   5520
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#00.00"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UORMWTFR"
               IdName          =   "UORMWTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   129
               Left            =   4245
               TabIndex        =   43
               ToolTipText     =   "Enter to Colour Stone Weight"
               Top             =   5520
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#00.00"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UORMWTTO"
               IdName          =   "UORMWTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   77
               Left            =   2490
               TabIndex        =   99
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   1620
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               IdName          =   "WADDALLSALCTG"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "CS Wt"
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
               Index           =   115
               Left            =   60
               TabIndex        =   98
               Top             =   5520
               Width           =   1230
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Diamond Wt."
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   97
               Top             =   5235
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Dsg Valid"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   96
               Top             =   6135
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
               Index           =   19
               Left            =   60
               TabIndex        =   93
               Top             =   4950
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design By"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   91
               Top             =   4380
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Model Maker"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   90
               Top             =   4095
               Width           =   1305
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Date"
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
               Index           =   15
               Left            =   60
               TabIndex        =   89
               Top             =   4665
               Width           =   1410
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Typ"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   85
               Top             =   360
               Width           =   1425
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Rm Qty && Wt"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   495
               Left            =   60
               TabIndex        =   84
               Top             =   6885
               Width           =   1455
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
               Index           =   8
               Left            =   60
               TabIndex        =   83
               Top             =   1215
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Loss Ctg"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   82
               Top             =   3810
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Set Family Cd"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   81
               Top             =   3525
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Dsg Colour"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   80
               Top             =   2670
               Width           =   1275
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
               TabIndex        =   79
               Top             =   2385
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
               Left            =   60
               TabIndex        =   78
               Top             =   2955
               Width           =   1365
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
               TabIndex        =   77
               Top             =   930
               Width           =   1365
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
               TabIndex        =   76
               Top             =   645
               Width           =   1365
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
               TabIndex        =   75
               Top             =   2100
               Width           =   1365
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
               TabIndex        =   74
               Top             =   3240
               Visible         =   0   'False
               Width           =   1365
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
               Left            =   6750
               TabIndex        =   73
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
               Left            =   4215
               TabIndex        =   72
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
               Left            =   1785
               TabIndex        =   71
               Top             =   0
               Width           =   1335
            End
         End
      End
      Begin VB.Frame FraSc 
         Height          =   3045
         Left            =   3750
         TabIndex        =   86
         Top             =   6360
         Visible         =   0   'False
         Width           =   8505
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Dsgs With Below 'D','C' SCtg Only"
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
            Index           =   4
            Left            =   3840
            TabIndex        =   56
            ToolTipText     =   "Check To Show Show Designs Which Have Diamonds Given In Scope Only"
            Top             =   150
            Width           =   4185
         End
         Begin MwfCtl.MWCTL_FLX GrdDsgSz 
            Height          =   2430
            Left            =   3870
            TabIndex        =   57
            Top             =   510
            Width           =   4320
            _ExtentX        =   7620
            _ExtentY        =   4286
            Cols            =   6
            colname1        =   "UODRMCTG"
            heading1        =   "Ctg"
            datafld1        =   "UodRmCtg"
            datatype1       =   4
            recalcparent1   =   "GrdDsgSz(UODPTRYN),GrdDsgSz(UODRMSCTG)"
            maxlength1      =   1
            colname2        =   "UODRMSCTG"
            heading2        =   "SubCtg "
            datafld2        =   "UodRmSCtg"
            datatype2       =   4
            recalcon2       =   "GrdDsgSz(UODRMCTG)"
            recalcparent2   =   "GrdDsgSz(UODPTRYN)"
            maxlength2      =   5
            tooltiptext2    =   "Enter Rm Sub Category"
            colname3        =   "UODPTRYN"
            heading3        =   "Ptr "
            datafld3        =   "UodPtrYN"
            datatype3       =   4
            recalcon3       =   "GrdDsgSz(UODRMSCTG),GrdDsgSz(UODRMCTG)"
            maxlength3      =   1
            style3          =   2
            colname4        =   "UODLNSVPTRFR"
            heading4        =   "Fr Ln/Sv/Ptr "
            datafld4        =   "UodLnSvPtrFr"
            datatype4       =   2
            mask4           =   "##0.0000"
            maxlength4      =   8
            tooltiptext4    =   "Enter From Length/ Sieve/ Pointer"
            colname5        =   "UODLNSVPTRTO"
            heading5        =   "To Ln/Sv/Ptr "
            datafld5        =   "UodLnSvPtrTo"
            datatype5       =   2
            mask5           =   "##0.0000"
            maxlength5      =   8
            tooltiptext5    =   "Enter To Length/ Sieve/ Pointer"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   22
            Left            =   8010
            TabIndex        =   55
            ToolTipText     =   "Show Designs Which Have Diamonds Given In Scope Only (Y/N)"
            Top             =   150
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN4"
            IdName          =   "UOYN4"
         End
         Begin MwfCtl.MWCTL_FLX GrdDsgAna 
            Height          =   2430
            Left            =   180
            TabIndex        =   92
            Top             =   510
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
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Dsgs With Below 'D' SCtg Only"
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
            Index           =   12
            Left            =   3840
            TabIndex        =   87
            ToolTipText     =   "Location"
            Top             =   150
            Visible         =   0   'False
            Width           =   4155
         End
      End
   End
End
Attribute VB_Name = "EmrFrpDsgLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepDsgLst
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) "
  GRP_REP.Add "Prd Ctg", "max(DmPrdCtg)", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'PRDCTG' and PMCd= max(DmPrdCtg)) "
  '****Geeta***Crm***212
  'uni.4 max(DmSalCtg) replaced by qDmSalCtgRepl,qDmSalCtgDescRepl to get either SalCtg,SalCTg2 or SalCTg3
  GRP_REP.Add "Sales Ctg", "||qDmSalCtgRepl||", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and PmCd=max(DmCtg) and PSCd= ||qDmSalCtgDescRepl|| ) "
  
  GRP_REP.Add "Loss Ctg", "max(DmLsCtg)", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'LSCTG' and PMCd= max(DmLsCtg)) "
  'If UCase(adc.MenuCd) = UCase("RepDM") Then
    GRP_REP.Add "Design Size", "str(DmSz)", "", "", "", "Design Code", "", "", "(Select PDesc from Param where PTyp= 'DMSZ' and PMCd= DmSz) "
    GRP_REP.Add "Design Code", "max(DmCtg)+DmCd", "'('+max(DmCtg)+')'+DmCd", "", "", "Design Ctg,Prd Ctg,Sales Ctg,Design Size,(Detail)", "", "", "max(DmDesc)"
  'Else
  '  GRP_REP.Add "Design Code", "max(DmCtg)+DmCd", "'('+max(DmCtg)+')'+DmCd", "", "", "Design Ctg,Prd Ctg,Sales Ctg,(Detail)", "", "", "max(DmDesc)"
  'End If
  GRP_REP.Add "Dsg Colour", "max(DmCol)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCOL' and PMCd = max(DmCol))"
   '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  
  '****** Sachin 3.3.0 (28-06-08) - [37. Design Exclusive for Multiple Customers]******
  '****** Temporarily removed as will give an error for Multiple customers
  'GRP_REP.Add "Customer", "max(DmCmCd)", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end ) from CustMst where CmCtg= 'C' and ','+ Max(DmCmCd) + ',' Like '%,' + CmCd + ',%') "
  '****** Sachin 3.3.0 (28-06-08) - [37. Design Exclusive for Multiple Customers]******
  
  GRP_REP.Add "Karat", "max(DmKt)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = max(DmKt))"
  GRP_REP.Add "Set Cd", "max(DmSetCd)", "", "", "", "", "", "", ""
End Sub

Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
' Zubin 212
  'If adc("UoYN4") <> "" And GrdDsgSz.Rows <= 1 Then
  '  pr_ErrMsg = "Show Dsgs With Below 'D' SCtg Only Option Should Be Blank When Scope For 'D' Is Not Entered"
  '  pr_Cancel = True: Exit Sub
  'End If
  'If adc("UoYN4") = "" And GrdDsgSz.Rows > 1 Then
  '  pr_ErrMsg = "Show Dsgs With Below 'D' SCtg Only Option Cannot Be Blank When Scope For 'D' Is Not Entered. (Enter 'Y'/'N')"
  '  pr_Cancel = True: Exit Sub
  'End If
End Sub

Private Sub adc_setreprecsource()
On Error GoTo RepErr
  '*** Report Sql ***
  ' Zubin 212
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'DmTcTyp, DmCd, DmSz, DmDesc, DmCtg, DmPrdCtg, DmSalCtg, DmBagPcs, DmCmCd,
  'DmHld, DmHldDesc, DmParts, DmPartDesc, DmOldCd, DmCol, DmPrdSeq, DmPrdInst,
  'DmDsgBy, DmDsgDt, DmModMkr, DmLstMdf, DmUom, DmVaCtg, DmKt, DmLsCtg,
  'DmRegnExc, DmSetCd, DmWaxWt, DmModRunWt, DmSilModWt, DmCasPcWt, DrWt as qGldWt, DrWt as qPlWt, DrWt as qSlWt, DrWt as qPdWt, DrWt as qOthMtWt,
  'DrWt as qDiaWt, DrWt As qCSWt, DrWt As qAccWt, DrWt As qGrsWt
  'From DsgMst, DsgRm where 1= 2
  '*** Report Sql ***
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  If Not (moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y") Then
    Rep.HdValidCoCd.Suppress = True
  End If
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String
  Dim ws_RmWt As String
  
  '*** 2.11 (Jen)
  Dim ws_DsgLnSvPtr As String
  Dim ws_HDTPrdPtr As String
  '*** 2.11 (Jen)
  ' Zubin 212
  Dim ws_DsgLnSvPtrOnly As String
  Dim ws_DsgAnaCnd As String        ' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  
  ' ***** Manali 3.6.0 - 29/01/10
  Dim wRmMstCnd As String, wPrntCdCnd  As String

  wPrntCdCnd = IIF(adc("UoParentCdFr") <> "", " and RmParentCd >= '" + adc("UoParentCdFr") + "' ", "")
  wPrntCdCnd = wPrntCdCnd + IIF(adc("UoParentCdTo") <> "", " and RmParentCd <= '" + adc("UoParentCdTo") + "' ", "")
  wPrntCdCnd = wPrntCdCnd + IIF(adc("UoParentCdSel") <> "", " and RmParentCd In (" + adc("UoParentCdSel") + ") ", "")
  wRmMstCnd = IIF(wPrntCdCnd <> "", _
                  " and (Select count(*) from DsgRm Join RmMst On RmCd=DrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=DrPrtKey ", "") + _
                  "      Where DrDmIdNo = DmIdNo And DrTcTyp = DmTcTyp And DrCd = DmCd And DrSz = DmSz " + IIF(gs_Partition = ctCurrPrtn, " and DrPrtKey=DmPrtKey ", "") + _
                   wPrntCdCnd + ") > 0 ", _
                  "")
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  If adc("UoYN3") = "Y" Then
    ws_RmWt = "DrPrdWt"
  Else
    ws_RmWt = "DrWt"
  End If
    
  '*** To check if Detail is selected ***
  Dim wGrp(3) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 3
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
  Next i
  
  '*** If Detail not selected then suppress Section 'DET1' Of Detail ***
  If wDetPos = 0 Then Rep.DET1.Suppress = True
    
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  ' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  ws_DsgAnaCnd = " And (Select Count('*') From UsrOptDsgAna " + _
                 "Join DsgAna on " + IIF(gs_Partition = ctCurrPrtn, " DaPrtKey='" + ctCurrPrtn + "' and ", "") + " DaAnaSr = UdAnaSr " + _
                 "where (DaAnaCd between UdAnaCdFr and UdAnaCdTo) " + _
                 "and DaTcTyp = DmTcTyp and DaCd= DmCd and DaSz= DmSz and " + _
                 "UdUsrCd = '" + adc("UoUsrCd") + "' and " + _
                 "UdMnuCd= '" + adc("UoMnuCd") + "')= " & (GrdDsgAna.Rows - 1) & " "
  ' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  
  '*** 2.11 (Jen)
  ws_HDTPrdPtr = UCase(moCn.GetFldVal("Select HDTPrdPtr from Head where HCoCd= '" + gs_CoCd + "'"))
  ' ***** Manali 3.6.0 - 31/10/09 - Detail Scope applied for Colour Stones - RmCtg Condition added in Join (UodRmCtg)
  ws_DsgLnSvPtr = " And (Select Count('*') From UsrOptDet where " + _
                  "not exists(Select 'x' From DsgRm " + _
                  "     join Param on PTyp= 'RMSCTG' and PMCd= DrRmCtg and PSCd= DrRmSCtg " + _
                  "     where DrDmIdNo=DmIdNo and DrTcTyp= DmTcTyp and DrCd= DmCd and DrSz= DmSz " + IIF(gs_Partition = ctCurrPrtn, " and DrPrtKey=DmPrtKey ", "") + _
                  "     and /*DrRmCtg= 'D'*/ DrRmCtg=UodRmCtg and DrRmSCtg= UodRmSCtg " + _
                  "     and (  (PValue3= 'Y' and (case when DrPrdQty> 0 and '" + ws_HDTPrdPtr + "'= 'Y' and PValue2= 'N' " + _
                  "                              then DrPrdWt/DrPrdQty else DrRmPtr end) between UodLnSvPtrFr and UodLnSvPtrTo) " + _
                  "            Or (PValue3= 'N' and DrLn1 between UodLnSvPtrFr and UodLnSvPtrTo) " + _
                  "         ) " + _
                  "     ) " + _
                  "and UodUsrCd = '" + adc("UoUsrCd") + "' and UodMnuCd= '" + adc("UoMnuCd") + "')= 0 "
  '*** 2.11 (Jen)
  
  ' Zubin 212
  If adc("UOYN4") = "Y" Then
    '****** Sachin 3.02 - Id fields in Join
    ' ***** Manali 3.6.0 - 31/10/09 - Detail Scope applied for Colour Stones - RmCtg Condition added in Join (UodRmCtg)
    ws_DsgLnSvPtrOnly = " and (Select count('x') From DsgRm " + _
                    " join Param on PTyp= 'RMSCTG' and PMCd= DrRmCtg and PSCd= DrRmSCtg " + _
                    " where DrDmIdNo=DmIdNo and DrTcTyp= DmTcTyp and DrCd= DmCd and DrSz= DmSz and DrRmCtg In  ('D','C') " + IIF(gs_Partition = ctCurrPrtn, " and DrPrtKey=DmPrtKey ", "") + _
                    " and not Exists(Select 'x' From UsrOptDet where DrRmCtg= UodRmCtg and DrRmSCtg= UodRmSCtg and " + _
                    " UodUsrCd = '" + adc("UoUsrCd") + "' and UodMnuCd= '" + adc("UoMnuCd") + "' and " + _
                    " ( (PValue3= 'Y' and (case when DrPrdQty> 0 and '" + ws_HDTPrdPtr + "'= 'Y' and PValue2= 'N' " + _
                    "     then DrPrdWt/DrPrdQty else DrRmPtr end) between UodLnSvPtrFr and UodLnSvPtrTo) Or " + _
                    "   (PValue3= 'N' and DrLn1 between UodLnSvPtrFr and UodLnSvPtrTo)) )) = 0 "
  End If
  
  ' Zubin 212
  '****** Sachin 3.02 - Id fields in Joins
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  ' ***** Manali 3.6.0 - 29/01/10 - RmMst Cnd added
  
  '4.1.3.0
'  If adc("UOYN5") = "Y" Then
'    wCnd = wCnd + IIF(wCnd <> "", " and ", "") + " DmValidYN='Y' "
'  ElseIf adc("UOYN5") = "N" Then
'    wCnd = wCnd + IIF(wCnd <> "", " and ", "") + " DmValidYN='N' "
'  End If

  Dim wHaveCnd   As String
  If UCase(adc("UoGrsWtFr")) > 0 Then _
    wHaveCnd = wHaveCnd + " Having Sum(Case when IsNull(DrRmCtg, '') ='D' then isnull(" + ws_RmWt + ", 0) Else 0 End)>= " & adc("UoGrsWtFr")
  If UCase(adc("UoGrsWtTo")) > 0 Then _
    wHaveCnd = wHaveCnd + IIF(wHaveCnd = "", " Having ", " and ") + " Sum(Case when IsNull(DrRmCtg, '') ='D' then isnull(" + ws_RmWt + ", 0) Else 0 End)<= " & adc("UoGrsWtTo")
  
  If UCase(adc("UoRmWtFr")) > 0 Then _
    wHaveCnd = wHaveCnd + " Having Sum(Case when IsNull(DrRmCtg, '') ='C' then isnull(" + ws_RmWt + ", 0) Else 0 End)>= " & adc("UoRmWtFr")
  If UCase(adc("UoRmWtTo")) > 0 Then _
    wHaveCnd = wHaveCnd + IIF(wHaveCnd = "", " Having ", " and ") + " Sum(Case when IsNull(DrRmCtg, '') ='C' then isnull(" + ws_RmWt + ", 0) Else 0 End)<= " & adc("UoRmWtTo")

  'Uni.4 - SalCtg scope will check all of 3 Sal Ctg(SalCtg, SalCtg2 & SalCtg3).
  'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is
  'calling function to get where clause to join DsgMst
  Dim wDmSalCtgCnd As String, ws_DmSalCtg As String
  If adc("UoDmSalCtgFr") <> "" Or adc("UoDmSalCtgTo") <> "" Or adc("UoDmSalCtgSel") <> "" Then
    wDmSalCtgCnd = GetDmSalCtgWhrCnd(adc("UoDmSalCtgFr"), adc("UoDmSalCtgTo"), adc("UoDmSalCtgSel"), adc("wAddAllSalCtg"))
  End If
    
  'uni.4 selecting Dmsalctg either from SalCtg,SalCtg2 or SalCtg3
  'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is
  'calling function to get SalCtg value to show in detail section and group sort
  ws_DmSalCtg = GetDmSalCtg(adc("UoDmSalCtgFr"), adc("UoDmSalCtgTo"), adc("UoDmSalCtgSel"), adc("wAddAllSalCtg"), "Y")
  
  '578-12 condition for customer code in DmAllow table added
'  Dim wCustCdCnd As String, wDmAllowCnd As String
'  wCustCdCnd = IIF(ADC("UoCmCdFr") <> "", " and DmwCmCd >= '" + ADC("UoCmCdFr") + "' ", "")
'  wCustCdCnd = wCustCdCnd + IIF(ADC("UoCmCdTo") <> "", " and DmwCmCd <= '" + ADC("UoCmCdTo") + "' ", "")
'  wCustCdCnd = wCustCdCnd + IIF(ADC("UoCmCdSel") <> "", " and DmwCmCd In(" + ADC("UoCmCdSel") + ") ", "")
'
'  wDmAllowCnd = IIF(wCustCdCnd <> "", " and (Select isnull((Select top 1 DmwYN from DmAllow " + _
'                  " Where DmwDmTyp = DmTcTyp and DmwDmCd= DmCd and DmwDmSz = DmSz and DmwTyp='CUST' " + _
                  wCustCdCnd + " and DmwYN ='Y'),isnull((Select top 1 DmwYN from DmAllow " + _
                  " Where DmwDmTyp = DmTcTyp and DmwDmCd= DmCd and DmwDmSz = DmSz and DmwTyp='CUST' " + _
                  wCustCdCnd + " and DmwYN ='N'), (Select DmwYN from DmAllow " + _
                  " Where DmwDmTyp = DmTcTyp and DmwDmCd= DmCd and DmwDmSz = DmSz and DmwTyp='CUST' " + _
                  " and DmwCmCd='ZSELF' and DmwYN ='Y') ))) = 'Y' ", "")
                 
  '4.1.3.0 - DmValidYn added below
  'uni.4 max(DmSalCtg) replaced by ws_DmSalCtg
  'wDmSalCtgCnd added to DsgMst join, to get SalCtg value from scope.
  'To avoid SalCtg scope condition in where clause
  wSqlStrg = "Select " + grpflds + ", " + _
             "DmTcTyp, DmCd, DmSz, max(DmDesc) as DmDesc, max(DmCtg) as DmCtg, " + _
             "max(DmPrdCtg) as DmPrdCtg, " + ws_DmSalCtg + " as DmSalCtg, " + _
             "max(DmBagPcs) as DmBagPcs, max(DmCmCd) as DmCmCd, max(DmHld) as DmHld, max(DmValidYn) as DmValidYn, " + _
             "max(DmHldDesc) as DmHldDesc, max(DmParts) as DmParts, max(DmValidCoCd) as DmValidCoCd, " + _
             "max(DmPartDesc) as DmPartDesc, max(DmOldCd) as DmOldCd, " + _
             "max(DmCol) as DmCol, max(DmPrdSeq) as DmPrdSeq, max(DmPrdInst) as DmPrdInst, " + _
             "max(DmDsgBy) as DmDsgBy, max(DmDsgDt) as DmDsgDt, max(DmModMkr) as DmModMkr, " + _
             "max(DmLstMdf) as DmLstMdf, max(DmUom) as DmUom, max(DmVaCtg) as DmVaCtg, " + _
             "max(DmKt) as DmKt, max(DmLsCtg) as DmLsCtg, max(DmRegnExc) as DmRegnExc, " + _
             "max(DmSetCd) as DmSetCd, max(DmWaxWt) as DmWaxWt, max(DmModRunWt) as DmModRunWt, " + _
             "max(DmSilModWt) as DmSilModWt, max(DmCasPcWt) as DmCasPcWt, " + _
             "sum(case when DrRmCtg= 'G' then " + ws_RmWt + " else 0 end) as qGldWt, " + _
             "sum(case when DrRmCtg= 'P' then " + ws_RmWt + " else 0 end) as qPlWt, " + _
             "sum(case when DrRmCtg= 'S' then " + ws_RmWt + " else 0 end) as qSlWt, " + _
             "sum(case when DrRmCtg= 'L' then " + ws_RmWt + " else 0 end) as qPdWt, " + _
             "sum(case when DrRmCtg= 'M' then " + ws_RmWt + " else 0 end) as qOthMtWt, " + _
             "sum(case when DrRmCtg= 'D' then " + ws_RmWt + " else 0 end) as qDiaWt, " + _
             "sum(case when DrRmCtg= 'C' then " + ws_RmWt + " else 0 end) as qCSWt, " + _
             "sum(case when DrRmCtg= 'X' then " + ws_RmWt + " else 0 end) as qAccWt, " + _
             "sum(" + ws_RmWt + "/ (case when DrRmCtg= 'D' or DrRmCtg= 'C' then 5 else 1 end)) as qGrsWt " + _
             "From DsgRm join DsgMst on DmIdNo=DrDmIdNo and DmTcTyp= DrTcTyp and DmCd= DrCd and DmSz= DrSz " + IIF(gs_Partition = ctCurrPrtn, " and DrPrtKey=DmPrtKey ", "") + wDmSalCtgCnd + _
             " where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' And ", "") + " DmTcTyp= '" + gs_DmTcTyp + "' " + wCnd + ws_DsgLnSvPtr + ws_DsgLnSvPtrOnly + ws_DsgAnaCnd + _
             wRmMstCnd + " group By DmTcTyp, DmCd, DmSz " + wHaveCnd
  
  'uni.4 replacing DmSalCtg value and its description in group sort of Sales Ctg
  wSqlStrg = Replace(wSqlStrg, "||qDmSalCtgRepl||", ws_DmSalCtg)
  wSqlStrg = Replace(wSqlStrg, "||qDmSalCtgDescRepl||", ws_DmSalCtg)
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added


  If adc("UoPMCd") = "DM" Then
    Rep.TxtHead.SetText ("Design Master Listing")
  ElseIf adc("UoPMCd") = "SM" Then
    Rep.TxtHead.SetText ("Design Sketch Master Listing")
  ElseIf adc("UoPMCd") = "PM" Then
    Rep.TxtHead.SetText ("Design Part Master Listing")
  End If

  If adc("UoYN") = "N" Then
    Rep.DET1.Suppress = True
    Rep.hDsgPrts.Suppress = True
    Rep.hDesc.Suppress = True
    'Rep.hValid.Suppress = True      '4.1.3
    Rep.hDmLsCtg.Suppress = True
    Rep.hDmSetCd.Suppress = True
    ' Zubin 212
    Rep.hWaxWt.Suppress = True
    Rep.hCasPcWt.Suppress = True
    Rep.hModRunWt.Suppress = True
    Rep.hSilModWt.Suppress = True
    Rep.hValid.Suppress = True      '578-12 heading still display even show 2nd line not selected
    Rep.HdValidCoCd.Suppress = True '578-12 heading still display even show 2nd line not selected
    ' Zubin 212
  End If
  CRV_REP.DisplayGroupTree = False
  CmdSC.Top = Me.ScaleHeight - CmdSC.Height - 20
  
 Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  'uni.4  checking Include SalCtg2,SalCtg3 option
  Case Is = 2
    If .Value = Checked Then adc("wAddAllSalCtg") = "Y" Else adc("wAddAllSalCtg") = "N"
  Case Is = 1
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  Case Is = 4
    If .Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"
  End Select
End With
End Sub
' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added

' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added

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
      .Store "UdUsrCd", adc("UoUsrCd")
      .Store "UdMnuCd", adc("UoMnuCd")
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
  GrdDsgAna.RecSource = "Select * from UsrOptDsgAna where UdUsrCd= '" + adc("UoUsrCd") + "' and " + _
                        "UdMnuCd= '" + adc("UoMnuCd") + "'"
End Sub

Private Sub GrdDsgAna_Validate(Cancel As Boolean)
' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  Cancel = GrdDsgAna.Validate
End Sub

'*** 2.11 (Jen)
Private Sub GrdDsgSz_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Select Case ColName
  ' ***** Manali 3.6.0 - 31/10/09 - Detail Scope applied for Colour Stones
  Case Is = UCase("UodRmCtg")
    Call HlpList.PMCd("RMCTG")
  
  Case Is = UCase("UodRmSCtg")
    Call HlpList.PSCd("RMSCTG", GrdDsgSz.Value(RowNum, "UodRmCtg"))
    'Call HlpList.PSCd("RMSCTG", "D")
  End Select
End Sub
Private Sub GrdDsgSz_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  With GrdDsgSz
    Select Case UCase(ColName)
    ' ***** Manali 3.6.0 - 31/10/09 - Detail Scope applied for Colour Stones
    Case Is = UCase("UodRmCtg")
      Cancel = Not (NewValue = "D" Or NewValue = "C")
      ErrMsg = "Rm Ctg 'D' and 'C' allowed": Exit Sub

    Case Is = UCase("UodRmSCtg")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'RMSCTG' and PMCd = '" + GrdDsgSz.Value(RowNum, "UodRmCtg") + "' and PSCd= '" + NewValue + "'")
      'Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'RMSCTG' and PMCd = 'D' and PSCd= '" + NewValue + "'")
      ErrMsg = "Invalid Rm Sub Ctg": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdDsgSz_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdDsgSz
    Select Case UCase(ColName)
    ' ***** Manali 3.6.0 - 31/10/09 - Detail Scope applied for Colour Stones
    Case Is = UCase("UodRmSCtg")
      If .Mode = fgmnorm Then .Value(RowNum, "UodRmSCtg") = ""
      
    Case Is = UCase("UodPtrYN")
      If .Mode = fgmnorm Or .Mode = fgminit Then
        .Value(RowNum, "UodPtrYN") = IIF(moCn.GetFldVal("Select PValue3 from Param where PTyp= 'RMSCTG'  " + _
                                     "and PMCd= '" + GrdDsgSz.Value(RowNum, "UodRmCtg") + "' and PSCd= '" + .Value(RowNum, "UodRmSCtg") + "'") = "N", "N", "Y")
        
'        .Value(RowNum, "UodPtrYN") = IIF(moCn.GetFldVal("Select PValue3 from Param where PTyp= 'RMSCTG'  " + _
'                                     "and PMCd= 'D' and PSCd= '" + .Value(RowNum, "UodRmSCtg") + "'") = "N", "N", "Y")
      End If
    End Select
  End With
End Sub
Private Sub GrdDsgSz_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
      'The fields of the Base table which are not existing in the grid are stored in the
      'table using Grid_InitKey
  
  With GrdDsgSz
    If .IsNew(RowNum) Then
      .Store "UodUsrCd", adc("UoUsrCd")
      .Store "UodMnuCd", adc("UoMnuCd")
      .Store "UodTbl", "RmSz"
    End If
  End With
End Sub
Private Sub GrdDsgSz_SetRecSource()
  '*** Set the Record Source of the Grid GrdDsgAna
  GrdDsgSz.RecSource = "Select * from UsrOptDet where UodUsrCd= '" + adc("UoUsrCd") + "' and " + _
                       "UodMnuCd= '" + adc("UoMnuCd") + "' and UodTbl= 'RmSz'"
End Sub
Private Sub CmdSc_Click()
  If FraSC.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraSC.Enabled = True
    FraSC.Visible = True
    TAB_REP.Enabled = False
    adc.AllowSave = False
    FraSC.ZOrder
  Else
    TAB_REP.Enabled = True
    adc.AllowSave = True
    FraSC.Enabled = False
    FraSC.Visible = False
  End If
End Sub
'*** 2.11 (Jen)

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
    Set adc.FirNKeyCtl = adc("UoPMCd")
    
    adc.Child = "GrdDsgAna,GrdDsgSz"        ' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
    Set GrdDsgSz.PrevCtl = adc
    Set GrdDsgSz.NextCtl = adc
    GrdDsgSz.HotKey = "B"
    
     ' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
    Set GrdDsgAna.PrevCtl = adc
    Set GrdDsgAna.NextCtl = adc
    GrdDsgAna.HotKey = "A"
     ' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
     
    Call InitProp(Me)
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_PTyp1 = "DMTCTYP"
    gs_CmCtg = "C"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  'If UCase(adc.MenuCd) = UCase("RepDM") Then
  If adc("UoPMCd") = "DM" Then
    gs_DmTcTyp = "DM"
    LblDmSz.Caption = "Design Size"
  ElseIf adc("UoPMCd") = "SM" Then
    gs_DmTcTyp = "SM"
    LblDmSz.Caption = "Design Size"
  ElseIf adc("UoPMCd") = "PM" Then
    gs_DmTcTyp = "PM"
    LblDmSz.Caption = "Part No"
  End If
  
  If adc("UoPMCd") = "" Then adc("UoPMCd") = "DM"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  ' Zubin 212
  If adc("UoYN4") = "" Then adc("UoYN4") = "N"
  
  ' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
  If adc("UoYN3") = "Y" Then OptRmPrd(0).Value = True
  If adc("UoYN3") = "N" Then OptRmPrd(1).Value = True
  
  '4.1.3.0
  If adc("UoYN5") = "Y" Then
    OptValid(0).Value = True
  ElseIf adc("UoYN5") = "N" Then
    OptValid(1).Value = True
  ElseIf adc("UoYN5") = "" Then
    OptValid(2).Value = True
  End If
  
  ' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

 ' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
 If adc("UoYN") = "" Then adc("UoYN") = "N"
 If adc("UoYN1") = "" Then adc("UoYN1") = "N"
 If adc("UoYN4") = "" Then adc("UoYN4") = "N"
 
 If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
 If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
 If adc("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
 ' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
 'uni.4 default value for Include SalCtg2,SalCtg3 is Yes
  ChkBoxArr(2).Value = Checked
  adc("wAddAllSalCtg") = "Y"

End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_PTyp1 = "DMTCTYP"
  
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCd")
    If pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' from Param where PTyp= 'DMTCTYP' " + _
           "and PMCd= '" + pv_NewValue + "' and PSCd= ''") Then Cancel = True: ErrMsg = "Enter Valid Design Type": Exit Sub
           
    If pv_NewValue = "DM" Then gs_DmTcTyp = "DM"
    If pv_NewValue = "SM" Then gs_DmTcTyp = "SM"
    If pv_NewValue = "PM" Then gs_DmTcTyp = "PM"
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_PTyp1 = "DMTCTYP"
  SetProp Me, IdName, When
End Sub

Private Sub GrdDsgSz_Validate(Cancel As Boolean)
    Cancel = GrdDsgSz.Validate
End Sub

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptRmPrd_Click(Index As Integer)
Select Case Index
  Case Is = 0
    adc("UoYN3") = "Y"
  Case Is = 1
    adc("UoYN3") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptRmPrd_GotFocus(Index As Integer)
  DispMsg FraOptRm.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

Private Sub OptValid_Click(Index As Integer)
Select Case Index
  Case Is = 0
    adc("UoYN5") = "Y"
  Case Is = 1
    adc("UoYN5") = "N"
  Case Is = 2
    adc("UoYN5") = ""
End Select
End Sub

Private Sub OptValid_GotFocus(Index As Integer)
    DispMsg FraOptRm.ToolTipText, etInfo
End Sub

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  adc("UOUSRCD") = adc.UsrCd
  adc("UOMNUCD") = adc.MenuCd
  
  '4.1.3.0
  If adc("UOYN5") = "Y" Then
    OptValid(0).Value = True
  ElseIf adc("UOYN5") = "N" Then
    OptValid(1).Value = True
  Else
    OptValid(2).Value = True
  End If
  
End Sub
Private Sub DispCoNm()
  Rep.wCoCd.SetText gs_CoNm
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  Rep.wCoCdLogo.SetText gs_CoNm
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    Rep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
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
' Zubin 212
Private Sub adc_Validate(Cancel As Boolean)
  Cancel = GrdDsgSz.Validate
End Sub

