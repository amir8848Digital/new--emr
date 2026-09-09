VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpDsgCat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Design Catalogue"
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
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   7440
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MwfCtl.MWCTL_BTN1 CmdChgImg 
      Height          =   390
      Left            =   10920
      TabIndex        =   148
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9675
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   688
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Chg &Image"
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
   Begin MwfCtl.MWCTL_BTN1 CmdDispSc 
      Height          =   390
      Left            =   3240
      TabIndex        =   76
      ToolTipText     =   "Display Scope Options"
      Top             =   9680
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   688
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Extra Scope"
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
   Begin MwfCtl.MWCTL_BTN1 CmdSc 
      Height          =   390
      Left            =   2010
      TabIndex        =   75
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9680
      Width           =   1245
      _ExtentX        =   2196
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
      Left            =   13290
      TabIndex        =   82
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
      Left            =   7155
      TabIndex        =   90
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   91
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
         TabIndex        =   92
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
      Left            =   8370
      TabIndex        =   89
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   87
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
      Height          =   9765
      Left            =   0
      TabIndex        =   86
      Top             =   60
      Width           =   15285
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   85
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpDsgCat.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpDsgCat.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   88
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   84
               Top             =   360
               Width           =   9405
               _ExtentX        =   16589
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9075
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   14985
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Copy Picture"
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
               Index           =   16
               Left            =   2840
               TabIndex        =   73
               Tag             =   "SortOrdCust"
               ToolTipText     =   "Check To Copy Pictures to the Given folder"
               Top             =   8760
               Width           =   1605
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
               Index           =   2
               Left            =   30
               TabIndex        =   11
               Tag             =   "DmHold"
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   1620
               Width           =   1950
            End
            Begin VB.Frame FraOptDmValid 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1740
               TabIndex        =   155
               ToolTipText     =   "Show Valid Or Invalid Or All Designs"
               Top             =   5780
               Width           =   3465
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
                  Left            =   25
                  TabIndex        =   43
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
                  TabIndex        =   44
                  ToolTipText     =   "Show Valid Or Invalid Or All Designs"
                  Top             =   0
                  Width           =   945
               End
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
                  Left            =   2550
                  TabIndex        =   45
                  ToolTipText     =   "Show Valid Or Invalid Or All Designs"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Print Catalog "
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
               Index           =   13
               Left            =   30
               TabIndex        =   72
               Tag             =   "ShowDsgCat"
               ToolTipText     =   "Check To Print Catalogue"
               Top             =   8760
               Width           =   1950
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "On Hold     "
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
               TabIndex        =   42
               Tag             =   "DmHold"
               ToolTipText     =   "Check To Show Design on Hold"
               Top             =   5520
               Width           =   1950
            End
            Begin VB.Frame FraOptDsgExc 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1770
               TabIndex        =   114
               ToolTipText     =   "Select if to Show Free Design or Region Exclusive Or Not Exclusive Designs"
               Top             =   6570
               Visible         =   0   'False
               Width           =   9345
               Begin VB.OptionButton OptDsgExc 
                  Caption         =   "Free Design"
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
                  Left            =   5520
                  TabIndex        =   144
                  ToolTipText     =   "Select if to Show Free Design or Region Exclusive Or Not Exclusive Designs"
                  Top             =   0
                  Width           =   1545
               End
               Begin VB.OptionButton OptDsgExc 
                  Caption         =   "All Designs"
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
                  Left            =   7950
                  TabIndex        =   143
                  ToolTipText     =   "Select if to Show Free Design or Region Exclusive Or Not Exclusive Designs"
                  Top             =   0
                  Width           =   1665
               End
               Begin VB.OptionButton OptDsgExc 
                  Caption         =   "Not Exclusive"
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
                  Left            =   2550
                  TabIndex        =   52
                  ToolTipText     =   "Select if to Show Free Design or Region Exclusive Or Not Exclusive Designs"
                  Top             =   0
                  Width           =   1665
               End
               Begin VB.OptionButton OptDsgExc 
                  Caption         =   "Region  Exclusive"
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
                  TabIndex        =   51
                  ToolTipText     =   "Select if to Show Free Design or Region Exclusive Or Not Exclusive Designs"
                  Top             =   0
                  Width           =   2055
               End
            End
            Begin VB.Frame FraOptRm 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1770
               TabIndex        =   113
               ToolTipText     =   "Select Production Rm Qty, Wt or Sales Rm Qty, Wt"
               Top             =   6060
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
                  TabIndex        =   47
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
                  Left            =   2520
                  TabIndex        =   48
                  ToolTipText     =   "Select Production Rm Qty, Wt or Sales Rm Qty, Wt"
                  Top             =   0
                  Width           =   1455
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4245
               TabIndex        =   3
               ToolTipText     =   "Enter To Design Category"
               Top             =   765
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
               Left            =   6930
               TabIndex        =   4
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   765
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
               TabIndex        =   2
               ToolTipText     =   "Enter From Design Category"
               Top             =   765
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
               TabIndex        =   15
               ToolTipText     =   "Enter To Design Code"
               Top             =   2385
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
               Left            =   6930
               TabIndex        =   16
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2385
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
               TabIndex        =   14
               ToolTipText     =   "Enter From Design Code"
               Top             =   2385
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
               TabIndex        =   6
               ToolTipText     =   "Enter To Design Production Category"
               Top             =   1050
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
               Left            =   6930
               TabIndex        =   7
               ToolTipText     =   "Enter Design Production Category Selection"
               Top             =   1050
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
               TabIndex        =   5
               ToolTipText     =   "Enter From Design Production Category"
               Top             =   1050
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
               Index           =   34
               Left            =   4245
               TabIndex        =   18
               ToolTipText     =   "Enter To Design Size"
               Top             =   2670
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
               TabIndex        =   17
               ToolTipText     =   "Enter From Design Size"
               Top             =   2670
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
               Index           =   8
               Left            =   4245
               TabIndex        =   9
               ToolTipText     =   "Enter To Design Sales Category"
               Top             =   1335
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
               Left            =   6930
               TabIndex        =   10
               ToolTipText     =   "Enter Design Sales Category Selection"
               Top             =   1335
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
               TabIndex        =   8
               ToolTipText     =   "Enter From Design Sales Category"
               Top             =   1335
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
               Index           =   17
               Left            =   1785
               TabIndex        =   1
               ToolTipText     =   "Enter Design Type"
               Top             =   510
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCD"
               IdName          =   "UOPMCD"
               CmpStr          =   "DmTcTyp = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   8625
               TabIndex        =   46
               ToolTipText     =   "Show Production Rm Qty, Wt ('Y' For Prd/ 'N' For Ord)"
               Top             =   6060
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
               Index           =   110
               Left            =   1785
               TabIndex        =   12
               ToolTipText     =   "Enter From Diamond+ Colour Stone Weight"
               Top             =   2100
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
               Index           =   0
               Left            =   4245
               TabIndex        =   13
               ToolTipText     =   "Enter From Diamond+ Colour Stone Weight"
               Top             =   2100
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
               Index           =   1
               Left            =   1785
               TabIndex        =   19
               ToolTipText     =   "Enter From Design Set Family Code"
               Top             =   2955
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   15
               DataType        =   4
               DataField       =   "UOSETCDFR"
               IdName          =   "UOSETCDFR"
               CmpStr          =   "DmSetCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   4245
               TabIndex        =   20
               ToolTipText     =   "Enter  To Design Set Family Code"
               Top             =   2955
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   15
               DataType        =   4
               DataField       =   "UOSETCDTO"
               IdName          =   "UOSETCDTO"
               CmpStr          =   "DmSetCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1785
               TabIndex        =   25
               ToolTipText     =   "Enter From Design Set Code"
               Top             =   3525
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCOLFR"
               IdName          =   "UODMCOLFR"
               CmpStr          =   "DmCol >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   4245
               TabIndex        =   26
               ToolTipText     =   "Enter  To Design Set Code"
               Top             =   3525
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCOLTO"
               IdName          =   "UODMCOLTO"
               CmpStr          =   "DmCol <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   2025
               TabIndex        =   41
               ToolTipText     =   "Show Design On Hold (Y/N/' ')"
               Top             =   5535
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
               CmpStr          =   "DmHld="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1785
               TabIndex        =   28
               ToolTipText     =   "Enter From Main metal  Weight"
               Top             =   3810
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0.000"
               MaxLength       =   10
               DataType        =   2
               DataField       =   "UOMAINMETWTFR"
               IdName          =   "UOMAINMETWTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   114
               Left            =   4245
               TabIndex        =   29
               ToolTipText     =   "EnterTo Main metal  Weight"
               Top             =   3810
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0.000"
               MaxLength       =   10
               DataType        =   2
               DataField       =   "UOMAINMETWTTO"
               IdName          =   "UOMAINMETWTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   128
               Left            =   1785
               TabIndex        =   30
               ToolTipText     =   "Enter From Gross Weight"
               Top             =   4095
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0.000"
               MaxLength       =   10
               DataType        =   2
               DataField       =   "UORMWTFR"
               IdName          =   "UORMWTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   129
               Left            =   4245
               TabIndex        =   31
               ToolTipText     =   "Enter To Gross Weight"
               Top             =   4095
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0.000"
               MaxLength       =   10
               DataType        =   2
               DataField       =   "UORMWTTO"
               IdName          =   "UORMWTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1785
               TabIndex        =   36
               ToolTipText     =   "Enter From Design Date"
               Top             =   4950
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "DmDsgDt>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   4245
               TabIndex        =   37
               ToolTipText     =   "Enter To Design Date"
               Top             =   4950
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "DmDsgDt<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   11235
               TabIndex        =   49
               Top             =   6570
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
               Index           =   18
               Left            =   4245
               TabIndex        =   54
               ToolTipText     =   "Enter To Customer"
               Top             =   7050
               Visible         =   0   'False
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   6930
               TabIndex        =   55
               ToolTipText     =   "Enter Customer Selection"
               Top             =   7050
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
               Index           =   20
               Left            =   1785
               TabIndex        =   53
               ToolTipText     =   "Enter From Customer "
               Top             =   7050
               Visible         =   0   'False
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCMCDFR"
               ReCalcOn        =   "UODSGEXC"
               IdName          =   "UOCMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   4245
               TabIndex        =   23
               ToolTipText     =   "Enter To Design Production Category"
               Top             =   3240
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UODMPRDSEQTO"
               IdName          =   "UODMPRDSEQTO"
               CmpStr          =   "DmPrdSeq<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   6930
               TabIndex        =   24
               ToolTipText     =   "Enter Design Production Category Selection"
               Top             =   3270
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMPRDSEQSEL"
               IdName          =   "UODMPRDSEQSEL"
               CmpStr          =   "DmPrdSeq In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   1785
               TabIndex        =   22
               ToolTipText     =   "Enter From Design Production Category"
               Top             =   3240
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UODMPRDSEQFR"
               IdName          =   "UODMPRDSEQFR"
               CmpStr          =   "DmPrdSeq >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   6930
               TabIndex        =   21
               ToolTipText     =   "Enter Set Family Code Selection"
               Top             =   2985
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOSETCDSEL"
               IdName          =   "UOSETCDSEL"
               CmpStr          =   "DmSetCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   6930
               TabIndex        =   27
               ToolTipText     =   "Enter Design Production Category Selection"
               Top             =   3570
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCOLSEL"
               IdName          =   "UODMCOLSEL"
               CmpStr          =   "DmCol In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   4245
               TabIndex        =   33
               ToolTipText     =   "Enter To Model Maker"
               Top             =   4380
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
               TabIndex        =   32
               ToolTipText     =   "Enter From Model Maker"
               Top             =   4380
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
               TabIndex        =   35
               ToolTipText     =   "Enter To Design By"
               Top             =   4665
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
               TabIndex        =   34
               ToolTipText     =   "Enter From Design By"
               Top             =   4665
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
               Index           =   86
               Left            =   4245
               TabIndex        =   60
               ToolTipText     =   "Enter To Catalog"
               Top             =   7620
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataType        =   4
               DataField       =   "UODPCOLLTO"
               IdName          =   "UODPCOLLTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   88
               Left            =   6930
               TabIndex        =   61
               ToolTipText     =   "Enter Catalog Selection"
               Top             =   7620
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODPCOLLSEL"
               IdName          =   "UODPCOLLSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   89
               Left            =   4245
               TabIndex        =   66
               ToolTipText     =   "Enter To Set"
               Top             =   8190
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0"
               MaxLength       =   5
               DataType        =   1
               DataField       =   "UODPSETTO"
               IdName          =   "UODPSETTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   92
               Left            =   4245
               TabIndex        =   63
               ToolTipText     =   "Enter To Series"
               Top             =   7905
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0"
               MaxLength       =   5
               DataType        =   1
               DataField       =   "UODPSERTO"
               IdName          =   "UODPSERTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   94
               Left            =   4245
               TabIndex        =   69
               ToolTipText     =   "Enter To Version"
               Top             =   8475
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0"
               MaxLength       =   5
               DataType        =   1
               DataField       =   "UODPVERTO"
               IdName          =   "UODPVERTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   95
               Left            =   6930
               TabIndex        =   70
               ToolTipText     =   "Enter Version Selection"
               Top             =   8475
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODPVERSEL"
               IdName          =   "UODPVERSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   96
               Left            =   6930
               TabIndex        =   67
               ToolTipText     =   "Enter Set Selection"
               Top             =   8190
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODPSETSEL"
               IdName          =   "UODPSETSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   97
               Left            =   6930
               TabIndex        =   64
               ToolTipText     =   "Enter Series Selection"
               Top             =   7905
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODPSERSEL"
               IdName          =   "UODPSERSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   87
               Left            =   1785
               TabIndex        =   68
               ToolTipText     =   "Enter From Version"
               Top             =   8475
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0"
               MaxLength       =   5
               DataType        =   1
               DataField       =   "UODPVERFR"
               IdName          =   "UODPVERFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   90
               Left            =   1785
               TabIndex        =   65
               ToolTipText     =   "Enter From Set"
               Top             =   8190
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0"
               MaxLength       =   5
               DataType        =   1
               DataField       =   "UODPSETFR"
               IdName          =   "UODPSETFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   91
               Left            =   1785
               TabIndex        =   62
               ToolTipText     =   "Enter From Series"
               Top             =   7905
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0"
               MaxLength       =   5
               DataType        =   1
               DataField       =   "UODPSERFR"
               IdName          =   "UODPSERFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   93
               Left            =   1785
               TabIndex        =   59
               ToolTipText     =   "Enter From Catalog"
               Top             =   7620
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataType        =   4
               DataField       =   "UODPCOLLFR"
               IdName          =   "UODPCOLLFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   4245
               TabIndex        =   57
               ToolTipText     =   "Enter To Catalog"
               Top             =   7335
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataType        =   4
               DataField       =   "UODPCATTO"
               IdName          =   "UODPCATTO"
               CmpStr          =   "DpCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   6930
               TabIndex        =   58
               ToolTipText     =   "Enter Catalog Selection"
               Top             =   7335
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODPCATSEL"
               IdName          =   "UODPCATSEL"
               CmpStr          =   "DpCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   1785
               TabIndex        =   56
               ToolTipText     =   "Enter From Catalog"
               Top             =   7335
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataType        =   4
               DataField       =   "UODPCATFR"
               IdName          =   "UODPCATFR"
               CmpStr          =   "DpCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   2355
               TabIndex        =   71
               ToolTipText     =   "Show Catalogue"
               Top             =   8775
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN13"
               IdName          =   "UOYN13"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   11730
               TabIndex        =   50
               ToolTipText     =   "Show Production Rm Qty, Wt ('Y' For Prd/ 'N' For Ord)"
               Top             =   6570
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UODSGEXC"
               ReCalcParent    =   "UOCMCDFR"
               IdName          =   "UODSGEXC"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   4245
               TabIndex        =   39
               ToolTipText     =   "Enter To Parent Code"
               Top             =   5235
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
               Index           =   51
               Left            =   1785
               TabIndex        =   38
               ToolTipText     =   "Enter From Parent Code"
               Top             =   5235
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
               Index           =   53
               Left            =   6930
               TabIndex        =   40
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   5265
               Width           =   9060
               _ExtentX        =   15981
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   5535
               TabIndex        =   156
               ToolTipText     =   "Show Valid Designs (Yes / No / All)"
               Top             =   5760
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN15"
               IdName          =   "UOYN15"
               CmpStr          =   "DmValidYN="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   77
               Left            =   2520
               TabIndex        =   158
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   1680
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
               Index           =   71
               Left            =   6930
               TabIndex        =   74
               ToolTipText     =   "Enter Picture Path Where To Be Copied"
               Top             =   8760
               Width           =   4740
               _ExtentX        =   8361
               _ExtentY        =   503
               MaxLength       =   100
               DataType        =   4
               IdName          =   "WTARPICPATH"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   4680
               TabIndex        =   160
               Top             =   8760
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN16"
               IdName          =   "UOYN16"
            End
            Begin VB.Label LblPicPath 
               BackStyle       =   0  'Transparent
               Caption         =   "Picture Path"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5640
               TabIndex        =   159
               Top             =   8760
               Width           =   1185
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
               Index           =   12
               Left            =   30
               TabIndex        =   157
               Top             =   5790
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
               TabIndex        =   145
               Top             =   5250
               Width           =   1335
            End
            Begin VB.Label lblCat 
               BackStyle       =   0  'Transparent
               Caption         =   "Catalogue"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   30
               TabIndex        =   142
               Top             =   7335
               Width           =   1365
            End
            Begin VB.Label LblVer 
               BackStyle       =   0  'Transparent
               Caption         =   "Version"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   30
               TabIndex        =   141
               Top             =   8475
               Width           =   1365
            End
            Begin VB.Label LblSet 
               BackStyle       =   0  'Transparent
               Caption         =   "Set"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   30
               TabIndex        =   140
               Top             =   8190
               Width           =   1365
            End
            Begin VB.Label LblSer 
               BackStyle       =   0  'Transparent
               Caption         =   "Series"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   30
               TabIndex        =   139
               Top             =   7905
               Width           =   1365
            End
            Begin VB.Label LblColl 
               BackStyle       =   0  'Transparent
               Caption         =   "Collection"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   30
               TabIndex        =   138
               Top             =   7620
               Width           =   1365
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
               TabIndex        =   116
               Top             =   4380
               Width           =   1305
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
               TabIndex        =   115
               Top             =   4665
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Seq"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   112
               Top             =   3240
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
               Index           =   9
               Left            =   30
               TabIndex        =   111
               Top             =   7050
               Visible         =   0   'False
               Width           =   1365
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Exclusivity"
               BeginProperty Font 
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
               TabIndex        =   110
               Top             =   6570
               Visible         =   0   'False
               Width           =   1185
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
               Index           =   7
               Left            =   60
               TabIndex        =   109
               Top             =   4950
               Width           =   1410
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Main Met Wt"
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
               Index           =   99
               Left            =   60
               TabIndex        =   108
               Top             =   3810
               Width           =   1380
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Grs Wt"
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
               TabIndex        =   107
               Top             =   4095
               Width           =   1230
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "On Hold"
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
               Left            =   60
               TabIndex        =   106
               Top             =   5535
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Color"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   105
               Top             =   3525
               Width           =   1365
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
               Index           =   4
               Left            =   60
               TabIndex        =   104
               Top             =   2955
               Width           =   1335
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
               TabIndex        =   103
               Top             =   6060
               Width           =   1455
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
               Index           =   0
               Left            =   60
               TabIndex        =   102
               Top             =   2100
               Width           =   1365
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
               TabIndex        =   101
               Top             =   480
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
               Index           =   8
               Left            =   60
               TabIndex        =   100
               Top             =   1335
               Width           =   1365
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
               TabIndex        =   99
               Top             =   2670
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
               TabIndex        =   98
               Top             =   1050
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
               TabIndex        =   97
               Top             =   765
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
               TabIndex        =   96
               Top             =   2385
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
               Left            =   6960
               TabIndex        =   95
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
               Left            =   4215
               TabIndex        =   94
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
               Left            =   1785
               TabIndex        =   93
               Top             =   120
               Width           =   1335
            End
         End
      End
      Begin VB.Frame FraSc 
         Height          =   3105
         Left            =   3300
         TabIndex        =   77
         Top             =   6375
         Visible         =   0   'False
         Width           =   8385
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Dsgs With Below 'D','C' SCtg Only        "
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
            Left            =   3720
            TabIndex        =   79
            ToolTipText     =   "Check To Show Show Designs Which Have Diamonds Given In Scope Only"
            Top             =   360
            Width           =   4515
         End
         Begin MwfCtl.MWCTL_FLX GrdDsgAna 
            Height          =   2130
            Left            =   105
            TabIndex        =   78
            Top             =   750
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   3757
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
         Begin MwfCtl.MWCTL_FLX GrdDsgSz 
            Height          =   2130
            Left            =   3720
            TabIndex        =   80
            Top             =   750
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   3757
            Cols            =   6
            colname1        =   "UODRMCTG"
            heading1        =   "Ctg"
            datafld1        =   "UodRmCtg"
            datatype1       =   4
            recalcparent1   =   "GrdDsgSz(UODPTRYN),GrdDsgSz(UODRMSCTG)"
            maxlength1      =   1
            tooltiptext1    =   "Enter 'D' or 'C'"
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
            Index           =   54
            Left            =   3030
            TabIndex        =   146
            ToolTipText     =   "Show Designs Which Have Diamonds Given In Scope Only (Y/N)"
            Top             =   360
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN"
            IdName          =   "UOYN"
         End
      End
      Begin VB.Frame FraDispSc 
         Height          =   3675
         Left            =   6390
         TabIndex        =   117
         Top             =   5790
         Visible         =   0   'False
         Width           =   2895
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Wt On RmCtg "
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
            Index           =   14
            Left            =   90
            TabIndex        =   136
            Tag             =   "DmHold"
            ToolTipText     =   "Check To Show RmCtg wise Details instead of RmStg"
            Top             =   3320
            Width           =   2400
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Code Details        "
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
            Index           =   12
            Left            =   90
            TabIndex        =   135
            Tag             =   "ShowRmCd"
            ToolTipText     =   "Check To Show Rm Codes"
            Top             =   3030
            Width           =   2400
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Quantity                "
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
            Index           =   11
            Left            =   90
            TabIndex        =   133
            Tag             =   "ShowRmQty"
            ToolTipText     =   "Check To Show Rm Quantity"
            Top             =   2745
            Width           =   2400
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Wt Details             "
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
            Index           =   10
            Left            =   90
            TabIndex        =   131
            Tag             =   "ShowRmWt"
            ToolTipText     =   "Check To Show Rm Weight Details"
            Top             =   2460
            Width           =   2400
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm SubCtg                  "
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
            Index           =   9
            Left            =   90
            TabIndex        =   129
            Tag             =   "ShowRmSCtg"
            ToolTipText     =   "Check To Show Rm Sub Category "
            Top             =   2175
            Width           =   2400
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Prd/Sal Ctg      "
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
            Index           =   8
            Left            =   90
            TabIndex        =   127
            Tag             =   "ShowPrdSalCtg"
            ToolTipText     =   "Check To Show Production Rm and Sales Category ]"
            Top             =   1890
            Width           =   2400
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Grs Wt             "
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
            Index           =   7
            Left            =   90
            TabIndex        =   125
            Tag             =   "ShowGrsWt"
            ToolTipText     =   "Check To Show Design Category"
            Top             =   1605
            Width           =   2400
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Col                   "
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
            Index           =   6
            Left            =   90
            TabIndex        =   123
            Tag             =   "ShowDsgCol"
            ToolTipText     =   "Check To Show Design Colour"
            Top             =   1320
            Width           =   2400
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Sz                    "
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
            Index           =   5
            Left            =   90
            TabIndex        =   121
            Tag             =   "ShowDsgSz"
            ToolTipText     =   "Check To Show Design Size"
            Top             =   1035
            Width           =   2400
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Category          "
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
            Left            =   90
            TabIndex        =   119
            Tag             =   "ShowDsgCtg"
            ToolTipText     =   "Check To Show Design Category"
            Top             =   750
            Width           =   2400
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   37
            Left            =   2520
            TabIndex        =   118
            ToolTipText     =   "Show Design Category? (Yes/No)"
            Top             =   750
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
            Index           =   40
            Left            =   2520
            TabIndex        =   120
            ToolTipText     =   "Show Design Size? (Yes / No)"
            Top             =   1035
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
            Index           =   41
            Left            =   2520
            TabIndex        =   126
            ToolTipText     =   "Show Production Rm and Sales Category (Yes / No) ?"
            Top             =   1890
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN8"
            ReCalcParent    =   "UOYN10"
            IdName          =   "UOYN8"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   42
            Left            =   2520
            TabIndex        =   124
            ToolTipText     =   "Show Design Gross Weight  (Yes / No)?"
            Top             =   1605
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN7"
            IdName          =   "UOYN7"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   43
            Left            =   2520
            TabIndex        =   128
            ToolTipText     =   "Show Rm Sub Category (Yes/No) ?"
            Top             =   2175
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN9"
            IdName          =   "UOYN9"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   98
            Left            =   2520
            TabIndex        =   130
            ToolTipText     =   "Show Rm Weight Details (Yes/No)?"
            Top             =   2460
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN10"
            IdName          =   "UOYN10"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   44
            Left            =   2520
            TabIndex        =   132
            ToolTipText     =   "Show Rm Quantity (Yes/No) ?"
            Top             =   2745
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN11"
            IdName          =   "UOYN11"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   45
            Left            =   2520
            TabIndex        =   134
            ToolTipText     =   "Show Rm Codes (Yes/No) ?"
            Top             =   3030
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN12"
            IdName          =   "UOYN12"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   52
            Left            =   2520
            TabIndex        =   122
            ToolTipText     =   "Show Rm Value (Yes / No)"
            Top             =   1320
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN6"
            IdName          =   "UOYN6"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   55
            Left            =   2520
            TabIndex        =   147
            ToolTipText     =   "Show RmCtg wise Details (Y/N/' ')"
            Top             =   3320
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN14"
            IdName          =   "UOYN14"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Show Design Details"
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
            Index           =   4
            Left            =   0
            TabIndex        =   137
            Top             =   240
            Width           =   3660
         End
      End
   End
   Begin VB.Frame FraChgImg 
      Height          =   1275
      Left            =   6120
      TabIndex        =   149
      Top             =   8295
      Visible         =   0   'False
      Width           =   6135
      Begin MwfCtl.MWCTL_BTN1 CmdChgImgGo 
         Height          =   390
         Left            =   5400
         TabIndex        =   154
         ToolTipText     =   "Analysis Scope Options"
         Top             =   675
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   688
         BackColor       =   -2147483633
         MaskColor       =   8454143
         Caption         =   "Go"
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
      Begin MwfCtl.MWCTL_BTN1 CmdSelFile 
         Height          =   285
         Left            =   4920
         TabIndex        =   153
         ToolTipText     =   "Analysis Scope Options"
         Top             =   720
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         BackColor       =   -2147483633
         MaskColor       =   8454143
         Caption         =   "..."
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
         Index           =   57
         Left            =   1200
         TabIndex        =   152
         ToolTipText     =   "Enter Sort Option"
         Top             =   720
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   503
         MaxLength       =   0
         IdName          =   "WXLFILEPATH"
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Report File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   151
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Select Report Excel File"
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
         Left            =   0
         TabIndex        =   150
         Top             =   240
         Width           =   6180
      End
   End
End
Attribute VB_Name = "EmrFrpDsgCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepDsgCat
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim wo_rsPic As MDORowSet
Dim mb_DsgCatMod As Boolean   ' **** Manali 3.5.0 - 08/12/08
Private Sub SetGroupSort()
  '******Geeta*****************Emr208****
  GRP_REP.Add "Catogery", "max(DmCtg)", "", "", "", "", "", "", ""
  GRP_REP.Add "Size", "DmSz", "", "", "", "", "", "", ""
  GRP_REP.Add "Prd Ctg", "max(DmPrdCtg)", "", "", "", "", "", "", ""
  GRP_REP.Add "Sales Ctg", "max(DmSalCtg)", "", "", "", "", "", "", ""
  GRP_REP.Add "Color", "max(DmCol)", "", "", "", "", "", "", ""
  GRP_REP.Add "Prd Seq", "max(DmPrdSeq)", "", "", "", "", "", "", ""
  GRP_REP.Add "Dsg Date", "max(DmDsgDt)", "", "", "", "", "", "", ""
  End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
'*** (Jen 2.12)

'***Geeta**Emr208**
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  DmTcTyp, DmCd, DmSz, DmDesc, DmCtg, DmPrdCtg, DmSalCtg, DmCol, DmKt,
'  DrWt As qGrsWt, Space(30) as qColDesc
'  From DsgMst, DsgRm where 1= 2

'**Prev**
  '*** Report Sql ***
'  Select DmTcTyp, DmCd, DmSz, DmDesc, DmCtg, DmPrdCtg, DmSalCtg, DmCol, DmKt,
'  DrWt As qGrsWt, Space(30) as qColDesc  From DsgMst, DsgRm where 1= 2
  
  
'  *** Report Sql for Catalogue ***
'    Select DmTcTyp, DmCd, DmSz, DrRmCtg, 0 as qTag, DrRmSCtg, DrRmCd, DrWt, DrQty
'    From DsgRm, DsgMst Where 1 = 2
  '*** Report Sql for Catalogue ***
  '*** Report Sql ***
  
  Dim wrepcnd As String, wCnd As String, wSqlStrg As String, grpflds As String
  Dim ws_RmWt As String, wRmDet As String, wHaveCnd As String
  Dim ws_DsgAnaCnd As String
  Dim ws_RmQty As String    '**Geeta****Emr206**
  
  '*** 2.11 (Jen)
  Dim ws_DsgLnSvPtr As String
  Dim ws_HDTPrdPtr As String
  '*** 2.11 (Jen)
  
  Dim ws_DsgLnSvPtrOnly As String     ' ** 3.10.0
  
  ' ***** Manali 3.5.0 - 08/12/08 - DsgPrm fields added in scope
  Dim ws_DsgPrmCnd As String, ws_DsgPrmJoin As String
  Dim ws_GrpBy As String, ws_OrdBy As String, ws_DsgCatFlds As String
  ' ***** Manali 3.5.0 - 08/12/08 - DsgPrm fields added in scope
  Dim wPicNm As String, wDpCatCnd As String  ' ***** Manali 3.6.1

  Dim ws_DsgExcCnd As String      ' **** Manali 350Nxt
  
  ' ***** Manali 3.6.0 - 29/01/10
  Dim wRmMstCnd As String, wPrntCdCnd  As String

  wPrntCdCnd = IIF(adc("UoParentCdFr") <> "", " and RmParentCd >= '" + adc("UoParentCdFr") + "' ", "")
  wPrntCdCnd = wPrntCdCnd + IIF(adc("UoParentCdTo") <> "", " and RmParentCd <= '" + adc("UoParentCdTo") + "' ", "")
  wPrntCdCnd = wPrntCdCnd + IIF(adc("UoParentCdSel") <> "", " and RmParentCd In (" + adc("UoParentCdSel") + ") ", "")
  wRmMstCnd = IIF(wPrntCdCnd <> "", _
                  " and (Select count(*) from DsgRm Join RmMst On RmCd=DrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=DrPrtKey ", "") + _
                  "      Where DrDmIdNo = DmIdNo And DrTcTyp = DmTcTyp And DrCd = DmCd And DrSz = DmSz " + IIF(gs_Partition = ctCurrPrtn, " and DrPrtKey=DmPrtKey ", "") + _
                   wPrntCdCnd + ") > 0 ", "")

  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  
  Dim wGrp(3) As String, i As Integer, wDetPos As Integer
    
  Call GRP_REP.Gen3LRep
  ' **** Manali 3.5.0 - 08/12/08 - If Condition added for New Design Catalogue to avoid Grouping
  If adc("UoYN13") = "N" Then
    grpflds = GRP_REP.GrpFldLst
  Else
    grpflds = "' ' as Grp1,' ' as Grp2,' ' as Grp3,' ' as DGrp1,' ' as DGrp2,' ' as DGrp3,' ' as G1Desc,' ' as G2Desc,' ' as G3Desc "
  End If
  ' **** Manali 3.5.0 - 08/12/08 - If Condition added for New Design Catalogue to avoid Grouping
  '6.1_27 if copy picture option is selected and path is blank then raise error message
  If adc("UOYN16") = "Y" And adc("wTarPicPath") = "" Then _
    DispMsg "Picture Path Should Not Be Blank ", etError: Exit Sub
  '6.1_27 if target picture path is not valid then it will raise error
  If adc("UOYN16") = "Y" And adc("wTarPicPath") <> "" Then
    If GetAttr(adc("wTarPicPath")) <> vbDirectory Then GoTo RepErr
  End If
    
  '*** To check if Detail is selected ***
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
  
  ' **** Manali 3.5.0 - 08/12/08 - If Condition added for New Design Catalogue to avoid Grouping
  If adc("UoYN13") = "Y" And wGrp(1) <> "(Detail)" Then DispMsg "Group sort Cannot Be Applied when Print Catalogue Option Is Checked", etWarning
  ' **** Manali 3.5.0 - 08/12/08 - If Condition added for New Design Catalogue to avoid Grouping
  
  If adc("UoYN3") = "Y" Then
    ws_RmWt = "DrPrdWt"
    ws_RmQty = "DrPrdQty"
  Else
    ws_RmWt = "DrWt"
    ws_RmQty = "DrQty"
  End If
    
  ' ********* Manali 350Nxt - adc("UoYN2") changed to adc("UoDsgExc") to Accommodate new option
  '578-12 exclusive options disabled
'  If adc("UoDsgExc") = "R" Then
'    ws_DsgExcCnd = " DmRegnExc = 'Y'"
'  ElseIf adc("UoDsgExc") = "C" Then
'    ws_DsgExcCnd = " DmRegnExc = 'N'"
'  ElseIf adc("UoDsgExc") = "F" Then
'    ws_DsgExcCnd = " DmRegnExc = ''"
'  ElseIf adc("UoDsgExc") = "A" Then
'    ws_DsgExcCnd = " DmRegnExc In  ('Y', 'N', '') "
'  End If
  ' ********* Manali 350Nxt - adc("UoYN2") changed to adc("UoDsgExc") to Accommodate new option
  
  ' ***** Manali 3.6.1
  If adc("UoYN13") = "Y" Then
    adc("UoDpCatFr").CmpStr = "DpCd >= "
    adc("UoDpCatTo").CmpStr = "DpCd <="
    adc("UoDpCatSel").CmpStr = "DpCd In "
  Else
    adc("UoDpCatFr").CmpStr = ""
    adc("UoDpCatTo").CmpStr = ""
    adc("UoDpCatSel").CmpStr = ""
  End If
  ' ***** Manali 3.6.1
  
  wrepcnd = adc.RepCond
  wCnd = wrepcnd
  ' ********* Manali 350Nxt - adc("UoYN2") changed to adc("UoDsgExc") to accommodate new option
    'wCnd = IIF(wCnd <> "", wCnd + " And ", "") + " DmRegnExc = '" + adc("UoYN2") + "'"
  '578-12 exclusive options disbaled
  'wCnd = IIF(wCnd <> "", wCnd + " And ", "") + ws_DsgExcCnd
  
  ws_DsgAnaCnd = " And (Select Count(*) From UsrOptDsgAna " + _
                 "Join DsgAna on " + IIF(gs_Partition = ctCurrPrtn, " DaPrtKey='" + ctCurrPrtn + "' and ", "") + " DaAnaSr = UdAnaSr " + _
                 "where (DaAnaCd between UdAnaCdFr and UdAnaCdTo) " + _
                 "and DaTcTyp = DmTcTyp and DaCd= DmCd and DaSz= DmSz and " + _
                 "UdUsrCd = '" + adc("UoUsrCd") + "' and " + _
                 "UdMnuCd= '" + adc("UoMnuCd") + "')= " & (GrdDsgAna.Rows - 1) & " "
    
  '*** 2.11 (Jen)
  ws_HDTPrdPtr = UCase(moCn.GetFldVal("Select HDTPrdPtr from Head where HCoCd= '" + gs_CoCd + "'"))
  ' ***** Manali 3.6.0 - 31/10/09 - Detail Scope applied for Colour Stones - RmCtg Condition added in Join (UodRmCtg)
  ws_DsgLnSvPtr = " And (Select Count(*) From UsrOptDet where " + _
                  "not exists(Select 'x' From DsgRm " + _
                  "     join Param on PTyp= 'RMSCTG' and PMCd= DrRmCtg and PSCd= DrRmSCtg " + _
                  "     where " + IIF(gs_Partition = ctCurrPrtn, " DrPrtKey='" + ctCurrPrtn + "' and ", "") + " DrTcTyp= DmTcTyp and DrCd= DmCd and DrSz= DmSz " + _
                  "     and /*DrRmCtg= 'D'*/ DrRmCtg=UodRmCtg and DrRmSCtg= UodRmSCtg " + _
                  "     and (  (PValue3= 'Y' and (case when DrPrdQty> 0 and '" + ws_HDTPrdPtr + "'= 'Y' and PValue2= 'N' " + _
                  "                              then DrPrdWt/DrPrdQty else DrRmPtr end) between UodLnSvPtrFr and UodLnSvPtrTo) " + _
                  "            Or (PValue3= 'N' and DrLn1 between UodLnSvPtrFr and UodLnSvPtrTo) " + _
                  "         ) " + _
                  "     ) " + _
                  "and UodUsrCd = '" + adc("UoUsrCd") + "' and UodMnuCd= '" + adc("UoMnuCd") + "')= 0 "
  '*** 2.11 (Jen)
  
  ' ***** Manali 3.10.0 - 10/03/12 - Show Dsgs With UsrDetOpt 'D','C' SCtg Only
  If adc("UOYN") = "Y" Then
    ws_DsgLnSvPtrOnly = " and (Select count('x') From DsgRm " + _
                    " join Param on PTyp= 'RMSCTG' and PMCd= DrRmCtg and PSCd= DrRmSCtg " + _
                    " where DrDmIdNo=DmIdNo and DrTcTyp= DmTcTyp and DrCd= DmCd and DrSz= DmSz and DrRmCtg In  ('D','C') " + IIF(gs_Partition = ctCurrPrtn, " and DrPrtKey=DmPrtKey ", "") + _
                    " and not Exists(Select 'x' From UsrOptDet where DrRmCtg= UodRmCtg and DrRmSCtg= UodRmSCtg and " + _
                    " UodUsrCd = '" + adc("UoUsrCd") + "' and UodMnuCd= '" + adc("UoMnuCd") + "' and " + _
                    " ( (PValue3= 'Y' and (case when DrPrdQty> 0 and '" + ws_HDTPrdPtr + "'= 'Y' and PValue2= 'N' " + _
                    "     then DrPrdWt/DrPrdQty else DrRmPtr end) between UodLnSvPtrFr and UodLnSvPtrTo) Or " + _
                    "   (PValue3= 'N' and DrLn1 between UodLnSvPtrFr and UodLnSvPtrTo)) )) = 0 "
  End If
  ' ***** Manali 3.10.0 - 10/03/12 - Show Dsgs With UsrDetOpt 'D','C' SCtg Only
  
  
  ' ****** Manali 3.5.0 - 08/12/08 - DsgPrm fields added to report scope
  If mb_DsgCatMod = True Then
      If adc("UoDpCollFr") <> "" Or adc("UoDpCollTo") <> "" Or adc("UoDpCollSel") <> "" Then
        ws_DsgPrmCnd = " And (Select count('x') from DsgPrm where DpTyp='COLL' and DmTcTyp='DM' and DpDmCd = DmCd and DmSz='' " + _
                    IIF(adc("UoDpCollFr") <> "", " and DpCd>='" + adc("UoDpCollFr") + "'", "") + _
                    IIF(adc("UoDpCollTo") <> "", " and DpCd<='" + adc("UoDpCollTo") + "'", "") + _
                    IIF(adc("UoDpCollSel") <> "", " and DpCd in (" + adc("UoDpCollSel") + ") ", "") + ")> 0 "
      End If
      If adc("UoDpSerFr") <> 0 Or adc("UoDpSerTo") <> 0 Or adc("UoDpSerSel") <> "" Then
        ws_DsgPrmCnd = ws_DsgPrmCnd + " And (Select count('x') from DsgPrm where DpTyp='SER' and DmTcTyp='DM' and DpDmCd = DmCd and DmSz='' " + _
                    IIF(adc("UoDpSerFr") <> 0, " and Convert(Int, DpCd)>='" + CStr(adc("UoDpSerFr")) + "'", "") + _
                    IIF(adc("UoDpSerTo") <> 0, " and Convert(Int, DpCd)<='" + CStr(adc("UoDpSerTo")) + "'", "") + _
                    IIF(adc("UoDpSerSel") <> "", " and DpCd in (Select DpCd from DsgPrm where DpTyp= 'SER' and DpDmCd in (" + adc("UoDpSerSel") + ")) and DpDmCd<>''", "") + ")> 0 "
      End If
      If adc("UoDpSetFr") <> 0 Or adc("UoDpSetTo") <> 0 Or adc("UoDpSetSel") <> "" Then
        ws_DsgPrmCnd = ws_DsgPrmCnd + " And (Select count('x') from DsgPrm where DpTyp='SET' and DmTcTyp='DM' and DpDmCd = DmCd and DmSz='' " + _
                    IIF(adc("UoDpSetFr") <> 0, " and Convert(Int, DpCd)>='" + CStr(adc("UoDpSetFr")) + "'", "") + _
                    IIF(adc("UoDpSetTo") <> 0, " and Convert(Int, DpCd)<='" + CStr(adc("UoDpSetTo")) + "'", "") + _
                    IIF(adc("UoDpSetSel") <> "", " and DpCd in (Select DpCd from DsgPrm where DpTyp= 'SET' and DpDmCd in (" + adc("UoDpSetSel") + ")) and DpDmCd<>''", "") + ")> 0 "
      End If
      If adc("UoDpVerFr") <> 0 Or adc("UoDpVerTo") <> 0 Or adc("UoDpVerSel") <> "" Then
        ws_DsgPrmCnd = ws_DsgPrmCnd + " And (Select count('x') from DsgPrm where DpTyp='VER' and DmTcTyp='DM' and DpDmCd = DmCd and DmSz='' " + _
                    IIF(adc("UoDpVerFr") <> 0, " and Convert(Int, DpCd)>='" + CStr(adc("UoDpVerFr")) + "'", "") + _
                    IIF(adc("UoDpVerTo") <> 0, " and Convert(Int, DpCd)<='" + CStr(adc("UoDpVerTo")) + "'", "") + _
                    IIF(adc("UoDpVerSel") <> "", " and DpCd in (Select DpCd from DsgPrm where DpTyp= 'VER' and DpDmCd in (" + adc("UoDpVerSel") + ")) and DpDmCd<>'' ", "") + ")> 0 "
      End If
  End If

  ' ***** Manali 3.6.0 - 30/09/09 - Catalogue Entry removed from Module Check
  ''' Bef 3.6.1
  ''If (adc("UoDpCatFr") <> "" Or adc("UoDpCatTo") <> "" Or adc("UoDpCatSel") <> "") Or adc("UoYN13") = "Y" Then
  ' **** Manali 3.6.1
  If (adc("UoDpCatFr") <> "" Or adc("UoDpCatTo") <> "" Or adc("UoDpCatSel") <> "") Then
    wDpCatCnd = IIF(adc("UoDpCatFr") <> "", " and DpCd >= '" + adc("UoDpCatFr") + "' ", "")
    wDpCatCnd = wDpCatCnd + IIF(adc("UoDpCatTo") <> "", " and DpCd <= '" + adc("UoDpCatTo") + "' ", "")
    wDpCatCnd = wDpCatCnd + IIF(adc("UoDpCatSel") <> "", " and DpCd In (" + adc("UoDpCatSel") + ") ", "")
    ws_DsgPrmJoin = " Join (Select Max(DpCd) as DpCd,Max(DpNo) as DpNo,Max(DpDmNo) as DpDmNo,Max(DpPctNm) as  DpPctNm,DpDmCd from DsgPrm Where DpTyp='CAT' and DpNo<>0 " + wDpCatCnd + _
                    "       Group By DpDmCd) DsgPrm On DpDmCd=DmCd  "
  ElseIf adc("UoYN13") = "Y" Then
    ws_DsgPrmJoin = " Join DsgPrm On DpTyp='CAT' and DpDmCd=DmCd and DpNo<>0 and DmSz='' "
  End If
  
  ' ***** Manali 3.6.0 - 30/09/09 - Catalogue Entry removed from Module Check
  ''''If mb_DsgCatMod = True And adc("UoYN13") = "Y" Then
  If adc("UoYN13") = "Y" Then
    ws_DsgCatFlds = "DpCd, DpNo, DpDmNo, "
    wPicNm = ", max(DpPctNm) as DpPctNm "
    ws_GrpBy = "Group By DpCd, DpNo, DpDmNo, DmTcTyp, DmCd, DmSz  "
    ws_OrdBy = "Order By DpCd, DpDmNo, DmTcTyp, DmCd, DmSz "
  Else
    ws_DsgCatFlds = "' ' as DpCd, 0 as DpNo, 0 as DpDmNo,"
    wPicNm = ", ' ' as DpPctNm "
    ws_GrpBy = "group By DmTcTyp, DmCd, DmSz  "
    ws_OrdBy = "Order By DmTcTyp, Max(DmCtg), DmCd, DmSz  "
  End If
  ' ****** Manali 3.5.0 - 08/12/08 - DsgPrm fields added to report scope
  
  If UCase(adc("UoGrsWtFr")) > 0 Then _
    wHaveCnd = wHaveCnd + " Having Sum(Case when IsNull(DrRmCtg, '') In ('D','C') then isnull(" + ws_RmWt + ", 0) Else 0 End)>= " & adc("UoGrsWtFr")
  If UCase(adc("UoGrsWtTo")) > 0 Then _
    wHaveCnd = wHaveCnd + IIF(wHaveCnd = "", " Having ", " and ") + " Sum(Case when IsNull(DrRmCtg, '') In ('D','C') then isnull(" + ws_RmWt + ", 0) Else 0 End)<= " & adc("UoGrsWtTo")
    '*******Geeta****Emr206*****************10/04/04****************
  If UCase(adc("UoMainMetWtFr")) > 0 Then _
    wHaveCnd = wHaveCnd + IIF(wHaveCnd = "", " Having ", " and ") + " max(Case when IsNull(DrMainMet, '') ='Y' then isnull(" + ws_RmWt + ", 0) Else 0 End)>= " & adc("UoMainMetWtFr")
  If UCase(adc("UoMainMetWtTo")) > 0 Then _
    wHaveCnd = wHaveCnd + IIF(wHaveCnd = "", " Having ", " and ") + " max(Case when IsNull(DrMainMet, '') ='Y' then isnull(" + ws_RmWt + ", 0) Else 0 End)<= " & adc("UoMainMetWtTo")
  If UCase(adc("UoRmWtFr")) > 0 Then _
    wHaveCnd = wHaveCnd + IIF(wHaveCnd = "", " Having ", " and ") + " sum(Case when IsNull(DrRmCtg, '') In ('D','C') then isnull(" + ws_RmWt + "/5, 0) Else isnull(" + ws_RmWt + ", 0) End)>= " & adc("UoRmWtFr")
  If UCase(adc("UoRmWtTo")) > 0 Then _
    wHaveCnd = wHaveCnd + IIF(wHaveCnd = "", " Having ", " and ") + " Sum(Case when IsNull(DrRmCtg, '') In ('D','C') then isnull(" + ws_RmWt + "/5, 0) Else isnull(" + ws_RmWt + ", 0) End)<= " & adc("UoRmWtTo")
   
  
  'Uni.4 - SalCtg scope will check all of 3 Sal Ctg(SalCtg, SalCtg2 & SalCtg3).
  'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is
  'calling function to get where clause to DsgMst
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
                  " Where DmwDmTyp = DmTcTyp and DmwDmCd= DmCd and DmwDmSz = DmSz and DmwTyp='CUST' " + _
                  wCustCdCnd + " and DmwYN ='Y'),isnull((Select top 1 DmwYN from DmAllow " + _
                  " Where DmwDmTyp = DmTcTyp and DmwDmCd= DmCd and DmwDmSz = DmSz and DmwTyp='CUST' " + _
                  wCustCdCnd + " and DmwYN ='N'), (Select DmwYN from DmAllow " + _
                  " Where DmwDmTyp = DmTcTyp and DmwDmCd= DmCd and DmwDmSz = DmSz and DmwTyp='CUST' " + _
                  " and DmwCmCd='ZSELF' and DmwYN ='Y') ))) = 'Y' ", "")
                  
  
  '**Geeta****Emr206*************
  '****** Replaced DrWt & DrQty with ws_RmWt,ws_RmQty*******
  '****** Sachin 3.02 - Id fields in Joins
  '****** Manali 3.5.0 - 08/12/08 - ws_DsgCatFlds, ws_DsgPrmJoin, ws_DsgPrmCnd, ws_GrpBy added
  '****** Manali 3.6.0 - 29/01/10 - RmMst Cnd added
  '****** Manali 3.10.0 - 10/03/12 - Show Dsgs With UsrDetOpt 'D','C' SCtg Only  - ws_DsgLnSvPtrOnly added in Cnd
  'uni.4 max(DmSalCtg) replaced by ws_DmSalCtg
  'wDmSalCtgCnd added  in where clause
  wSqlStrg = " Select (row_Number() over(" + ws_OrdBy + " )-1)/3  as qRow,row_Number() over(" + ws_OrdBy + " )%3  as qCol,  " + _
             ws_DsgCatFlds + "DmTcTyp, DmCd, DmSz, max(DmDesc) as DmDesc, max(DmCtg) as DmCtg, " + _
             " max(DmPrdCtg) as DmPrdCtg, " + ws_DmSalCtg + " as DmSalCtg, " + _
             " max(DmCol) as DmCol, max(DmKt) as DmKt, " + _
             " Sum(IsNull(" + ws_RmWt + ", 0)/(case when IsNull(DrRmCtg, '') in ('C', 'D') then 5 else 1 end)) as qGrsWt, " + _
             "  max(PDesc) as qColDesc " + wPicNm + " Into #TmpA  " + _
             " From DsgMst " + _
             " Left Outer Join DsgRm on DrDmIdNo=DmIdNo and DrTcTyp = DmTcTyp and DrCd = DmCd and DrSz = DmSz " + IIF(gs_Partition = ctCurrPrtn, " And DrPrtKey=DmPrtKey ", "") + _
             "  Join Param On PTyp= 'DMCOL' and PMCd= DmCol and PSCd='' " + _
             ws_DsgPrmJoin + _
             " Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' And ", "") + wCnd + ws_DsgAnaCnd + ws_DsgLnSvPtr + ws_DsgLnSvPtrOnly + ws_DsgPrmCnd + _
             wRmMstCnd + wDmSalCtgCnd + _
              ws_GrpBy + wHaveCnd + ws_OrdBy
            
            moCn.CreateTmpTable "#TmpA", wSqlStrg
    
 '6.1_27 if records exists then starts picture copy
  If adc("UOYN16") = "Y" Then
    Dim wRsPicCpy As MDORowSet
    Set wRsPicCpy = moCn.OpenRes("SELECT * FROM #TmpA")
    
    If wRsPicCpy.RecCount > 0 Then
      MWLib.BeginProcess Me, "Starting Picture Copy"
      Call PicCopy(wRsPicCpy, adc("wTarPicPath"), "Y")
      MWLib.EndProcess Me
      MsgBox "Picture Copy Completed"
    End If
  End If
  
 wSqlStrg = " Select " + grpflds + ",qRow,max(DmTcTyp) as DmTcTyp," + _
              " Max((Case when qCol=1 Then DpCd Else '' End)) as DpCd1, Max((Case when qCol=1 Then DpNo Else 0 End)) as DpNo1,Max((Case when qCol=1 Then DpDmNo Else 0 End)) as DpDmNo1," + _
              " Max((Case when qCol=1 Then DmCd Else '' End)) as DmCd1, Max((Case when qCol=1 Then DmSz Else '' End)) as DmSz1," + _
              " Max((Case when qCol=1 Then DmDesc Else '' End)) as DmDesc1, Max((Case when qCol=1 Then DmCtg Else '' End)) as DmCtg1," + _
              " Max((Case when qCol=1 Then DmPrdCtg Else '' End)) as DmPrdCtg1, Max((Case when qCol=1 Then DmSalCtg Else '' End)) as DmSalCtg1," + _
              " Max((Case when qCol=1 Then DmCol Else '' End)) as DmCol1, Max((Case when qCol=1 Then DmKt Else '' End)) as DmKt1," + _
              " Max((Case when qCol=1 Then qGrsWt Else 0.0 End)) as qGrsWt1, Max((Case when qCol=1 Then qColDesc Else '' End)) as qColDesc1, Max((Case when qCol=1 Then DpPctNm Else '' End)) as DpPctNm1," + _
              " Max((Case when qCol=2 Then DpCd Else '' End)) as DpCd2, Max((Case when qCol=2 Then DpNo Else 0 End)) as DpNo2,Max((Case when qCol=2 Then DpDmNo Else 0 End)) as DpDmNo2," + _
              " Max((Case when qCol=2 Then DmCd Else '' End)) as DmCd2, Max((Case when qCol=2 Then DmSz Else '' End)) as DmSz2," + _
              " Max((Case when qCol=2 Then DmDesc Else '' End)) as DmDesc2, Max((Case when qCol=2 Then DmCtg Else '' End)) as DmCtg2," + _
              " Max((Case when qCol=2 Then DmPrdCtg Else '' End)) as DmPrdCtg2, Max((Case when qCol=2 Then DmSalCtg Else '' End)) as DmSalCtg2," + _
              " Max((Case when qCol=2 Then DmCol Else '' End)) as DmCol2, Max((Case when qCol=2 Then DmKt Else '' End)) as DmKt2," + _
              " Max((Case when qCol=2 Then qGrsWt Else 0.0 End)) as qGrsWt2, Max((Case when qCol=2 Then qColDesc Else '' End)) as qColDesc2, Max((Case when qCol=2 Then DpPctNm Else '' End)) as DpPctNm2," + _
              " Max((Case when qCol=0 Then DpCd Else '' End)) as DpCd3, Max((Case when qCol=0 Then DpNo Else 0 End)) as DpNo3,Max((Case when qCol=0 Then DpDmNo Else 0 End)) as DpDmNo3," + _
              " Max((Case when qCol=0 Then DmCd Else '' End)) as DmCd3, Max((Case when qCol=0 Then DmSz Else '' End)) as DmSz3," + _
              " Max((Case when qCol=0 Then DmDesc Else '' End)) as DmDesc3, Max((Case when qCol=0 Then DmCtg Else '' End)) as DmCtg3," + _
              " Max((Case when qCol=0 Then DmPrdCtg Else '' End)) as DmPrdCtg3, Max((Case when qCol=0 Then DmSalCtg Else '' End)) as DmSalCtg3," + _
              " Max((Case when qCol=0 Then DmCol Else '' End)) as DmCol3, Max((Case when qCol=0 Then DmKt Else '' End)) as DmKt3," + _
              " Max((Case when qCol=0 Then qGrsWt Else 0.0 End)) as qGrsWt3, Max((Case when qCol=0 Then qColDesc Else '' End)) as qColDesc3, Max((Case when qCol=0 Then DpPctNm Else '' End)) as DpPctNm3 " + _
              "  From #TmpA Group By qRow Order By qRow "
  
      
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added
  
    
  adc("UoParentCdFr").CmpStr = "": adc("UoParentCdFr").CmpStr = "": adc("UoParentCdSel").CmpStr = ""
  wrepcnd = adc.RepCond
  wCnd = wrepcnd
  '578-12 exclusive options disbaled
  'wCnd = IIF(wCnd <> "", wCnd + " And ", "") + ws_DsgExcCnd

  '****** Sachin 3.02 - Id fields in Joins
  '***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  '****** Manali 3.5.0 - 08/12/08 - ws_DsgPrmJoin added
  '***** Manali 3.6.0 - 29/01/10 - RmMst Cnd added
  'uni.4 wDmSalCtgCnd added to DsgMst join, to get SalCtg value from scope.
  'To avoid SalCtg scope condition in where clause
  wRmDet = " Select " + ws_DsgCatFlds + "DmTcTyp, DmCd, DmSz, DrRmCtg, (case DrRmCtg when 'G' then 1 when 'P' then 2 when 'S' then 3 when 'L' then 4 " + _
           " when 'D' then 5 when 'C' then 6 when 'X' then 7 when 'A' then 8 when 'M' then 9 else 0 end) " + _
           " as qTag, DrRmSCtg, DrRmCd, " + ws_RmWt + " as DrWt, " + ws_RmQty + " as DrQty " + _
           " From DsgRm " + _
           " join DsgMst on DmIdNo=DrDmIdNo and DmTcTyp = DrTcTyp and DmCd = DrCd and DmSz = DrSz " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=DrPrtKey ", "") + _
           wDmSalCtgCnd + ws_DsgPrmJoin + _
           " where " + IIF(gs_Partition = ctCurrPrtn, " DrPrtKey='" + ctCurrPrtn + "' And ", "") + wCnd + ws_DsgAnaCnd + ws_DsgPrmCnd + wRmMstCnd
    
  Call moCn.RepRes(Rep.CatSubRep.OpenSubreport, wRmDet + ctMaxDopOpt)
  Call moCn.RepRes(Rep.CatSubRep2.OpenSubreport, wRmDet + ctMaxDopOpt)
  Call moCn.RepRes(Rep.CatSubRep3.OpenSubreport, wRmDet + ctMaxDopOpt)
  
  If adc("UoPMCd") = "DM" Then
    Rep.TxtHead.SetText ("Design Master Catalog")
  ElseIf adc("UoPMCd") = "SM" Then
    Rep.TxtHead.SetText ("Design Sketch Master Catalog")
  ElseIf adc("UoPMCd") = "PM" Then
    Rep.TxtHead.SetText ("Design Part Master Catalog")
  End If
  
  Set wo_rsPic = moCn.OpenRes("Select DmTcTyp,DmCd,DmSz,'' as OmCoCd, '' as OmTc, 0 as OmYy, '' as OmChr, 0 as OmNo, 0 as OdSr,DpPctNm as PicNm from #TmpA Order By qRow,(Case when qCol=0 Then 3 Else qCol End)")
  If wo_rsPic.RecCount > 0 Then CmdChgImg.Visible = True
  
  ' ***** Manali 3.5.0 - 10/11/08 - Dsg Cat fields split into diff sections and sections suppressed or displayed as per user
  If adc("UoYN13") = "N" And wDetPos = 0 Then
      Rep.DET.Suppress = True
      Rep.DetDmCtg.Suppress = True: Rep.DetDmSz.Suppress = True: Rep.DetDmCol.Suppress = True
      Rep.DetDmWt.Suppress = True: Rep.DetPrdSal.Suppress = True
      Rep.DetCatSubRep.Suppress = True: Rep.DetLine.Suppress = True
      
  ElseIf (adc("UoYN13") = "Y" Or (adc("UoYN13") = "N" And wDetPos <> 0)) Then
      Rep.DET.Suppress = False
      If adc("UoYN4") = "Y" Then Rep.DetDmCtg.Suppress = False Else Rep.DetDmCtg.Suppress = True
      If adc("UoYN5") = "Y" Then Rep.DetDmSz.Suppress = False Else Rep.DetDmSz.Suppress = True
      If adc("UoYN6") = "Y" Then Rep.DetDmCol.Suppress = False Else Rep.DetDmCol.Suppress = True
      If adc("UoYN7") = "Y" Then Rep.DetDmWt.Suppress = False Else Rep.DetDmWt.Suppress = True
      If adc("UoYN8") = "Y" Then Rep.DetPrdSal.Suppress = False Else Rep.DetPrdSal.Suppress = True
      
      If adc("UoYN9") = "Y" Then
        Rep.CatSubRep_SecRmSCtg.Suppress = False
        Rep.CatSubRep2_SecRmSCtg.Suppress = False
        Rep.CatSubRep3_SecRmSCtg.Suppress = False
      Else
        Rep.CatSubRep_SecRmSCtg.Suppress = True
        Rep.CatSubRep2_SecRmSCtg.Suppress = True
        Rep.CatSubRep3_SecRmSCtg.Suppress = True
      End If
      
      If adc("UoYN10") = "Y" Then
        Rep.CatSubRep_SecWt.Suppress = False
        Rep.CatSubRep2_SecWt.Suppress = False
        Rep.CatSubRep3_SecWt.Suppress = False
      Else
        Rep.CatSubRep_SecWt.Suppress = True
        Rep.CatSubRep2_SecWt.Suppress = True
        Rep.CatSubRep3_SecWt.Suppress = True
      End If
      
      If adc("UoYN11") = "Y" Then
         Rep.CatSubRep_SecQty.Suppress = False
         Rep.CatSubRep2_SecQty.Suppress = False
         Rep.CatSubRep3_SecQty.Suppress = False
      Else
         Rep.CatSubRep_SecQty.Suppress = True
         Rep.CatSubRep2_SecQty.Suppress = True
         Rep.CatSubRep3_SecQty.Suppress = True
      End If
         
      If adc("UoYN12") = "Y" Then
        Rep.CatSubRep_SecRmCd.Suppress = False
        Rep.CatSubRep2_SecRmCd.Suppress = False
        Rep.CatSubRep3_SecRmCd.Suppress = False
      Else
        Rep.CatSubRep_SecRmCd.Suppress = True
        Rep.CatSubRep2_SecRmCd.Suppress = True
        Rep.CatSubRep3_SecRmCd.Suppress = True
      End If
      
      If adc("UoYN9") = "N" And adc("UoYN10") = "N" And adc("UoYN11") = "N" And adc("UoYN12") = "N" Then Rep.DetCatSubRep.Suppress = True
      If adc("UoYn12") = "Y" Then Rep.Line1.Top = 0: Rep.DetLine.Height = 45
  End If
  ' ***** Manali 3.5.0 - 10/11/08 - Dsg Cat fields split into diff sections and sections suppressed or displayed as per user
    
  If adc("UoYN13") = "N" Then Rep.GHDpCd.Suppress = True: Rep.GFDpCd.Suppress = True
    
  '3.12.0 - Show RmCtg/RmSCtg wise Summary
  If adc("UoYN14") = "Y" Then
    Rep.CatSubRep_TxtCtgStr.SetText ("Rm Ctg :")
    SetSubFormula Rep.CatSubRep, "wGrp2", "ToText ({rdo.qTag}, '#0')+Trim({rdo.DrRmCtg})"
    SetSubFormula Rep.CatSubRep, "wGrp2Desc", "Trim({rdo.DrRmCtg})"
    
    Rep.CatSubRep2_TxtCtgStr.SetText ("Rm Ctg :")
    SetSubFormula Rep.CatSubRep2, "wGrp2", "ToText ({rdo.qTag}, '#0')+Trim({rdo.DrRmCtg})"
    SetSubFormula Rep.CatSubRep2, "wGrp2Desc", "Trim({rdo.DrRmCtg})"
    
    Rep.CatSubRep3_TxtCtgStr.SetText ("Rm Ctg :")
    SetSubFormula Rep.CatSubRep3, "wGrp2", "ToText ({rdo.qTag}, '#0')+Trim({rdo.DrRmCtg})"
    SetSubFormula Rep.CatSubRep3, "wGrp2Desc", "Trim({rdo.DrRmCtg})"
  Else
    Rep.CatSubRep_TxtCtgStr.SetText ("Rm SubCtg :")
    SetSubFormula Rep.CatSubRep, "wGrp2", "ToText ({rdo.qTag}, '#0')+Trim({rdo.DrRmCtg})+{rdo.DrRmSCtg}"
    SetSubFormula Rep.CatSubRep, "wGrp2Desc", "'('+Trim({rdo.DrRmCtg})+')'+{rdo.DrRmSCtg}"
  
    Rep.CatSubRep2_TxtCtgStr.SetText ("Rm SubCtg :")
    SetSubFormula Rep.CatSubRep2, "wGrp2", "ToText ({rdo.qTag}, '#0')+Trim({rdo.DrRmCtg})+{rdo.DrRmSCtg}"
    SetSubFormula Rep.CatSubRep2, "wGrp2Desc", "'('+Trim({rdo.DrRmCtg})+')'+{rdo.DrRmSCtg}"
    
    Rep.CatSubRep3_TxtCtgStr.SetText ("Rm SubCtg :")
    SetSubFormula Rep.CatSubRep3, "wGrp2", "ToText ({rdo.qTag}, '#0')+Trim({rdo.DrRmCtg})+{rdo.DrRmSCtg}"
    SetSubFormula Rep.CatSubRep3, "wGrp2Desc", "'('+Trim({rdo.DrRmCtg})+')'+{rdo.DrRmSCtg}"
  End If
    
  CRV_REP.DisplayGroupTree = False
  CmdSC.Top = Me.ScaleHeight - CmdSC.Height - 20
  CmdDispSc.Top = Me.ScaleHeight - CmdDispSc.Height - 20      ' ******** Manali 3.5.0 - 10/11/08 - Display fields Scope Button
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******

'*** (Jen 2.12)
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
  '6.1_27 if copy pic option is selected and pic path is not valid it will show this error message
  If adc("UOYN16") = "Y" And adc("wTarPicPath") <> "" And Err.Number = 53 Then _
    DispMsg "Picture Path is Invalid. Enter Proper Picture Path ", etError
'*** (Jen 2.12)
End Sub
Private Sub ADC_Load()
    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    Set adc.FirNKeyCtl = adc("UoPMCd")
    
    adc.Child = "GrdDsgAna,GrdDsgSz"
    Set GrdDsgAna.PrevCtl = adc
    Set GrdDsgAna.NextCtl = adc
    GrdDsgAna.HotKey = "A"
    
    Set GrdDsgSz.PrevCtl = adc
    Set GrdDsgSz.NextCtl = adc
    GrdDsgSz.HotKey = "B"
    
    ' ***** Manali 3.5.0 - 08/12/08
    mb_DsgCatMod = IIF(moCn.GetFldVal("Select HDsgCatYN from head where HCoCd='" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "' ") = "Y", True, False)
    ' ***** Manali 3.5.0 - 08/12/08
    
    Call InitProp(Me)
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_PTyp1 = "DMTCTYP"
    gs_CmCtg = "C"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
      
    If mb_DsgCatMod = False Then
    ' ***** Manali 3.6.0 - 30/09/09 - Catalogue Entry removed from Module Check
    ''''      ChkBoxArr(13).Enabled = False: ChkBoxArr(13).Visible = False
    ''''      lblCat.Visible = False: adc("UoDpCatFr").Enabled = False: adc("UoDpCatFr").Visible = False
    ''''                              adc("UoDpCatTo").Enabled = False: adc("UoDpCatTo").Visible = False
    ''''                              adc("UoDpCatSel").Enabled = False: adc("UoDpCatSel").Visible = False
    ' ***** Manali 3.6.0 - 30/09/09 - Catalogue Entry removed from Module Check
      LblColl.Visible = False: adc("UoDpCollFr").Enabled = False: adc("UoDpCollFr").Visible = False
                              adc("UoDpCollTo").Enabled = False: adc("UoDpCollTo").Visible = False
                              adc("UoDpCollSel").Enabled = False: adc("UoDpCollSel").Visible = False
      LblSer.Visible = False: adc("UoDpSerFr").Enabled = False: adc("UoDpSerFr").Visible = False
                              adc("UoDpSerTo").Enabled = False: adc("UoDpSerTo").Visible = False
                              adc("UoDpSerSel").Enabled = False: adc("UoDpSerSel").Visible = False
      LblSet.Visible = False: adc("UoDpSetFr").Enabled = False: adc("UoDpSetFr").Visible = False
                              adc("UoDpSetTo").Enabled = False: adc("UoDpSetTo").Visible = False
                              adc("UoDpSetSel").Enabled = False: adc("UoDpSetSel").Visible = False
      LblVer.Visible = False: adc("UoDpVerFr").Enabled = False: adc("UoDpVerFr").Visible = False
                              adc("UoDpVerTo").Enabled = False: adc("UoDpVerTo").Visible = False
                              adc("UoDpVerSel").Enabled = False: adc("UoDpVerSel").Visible = False
    End If
    
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
  End If
  
  If adc("UoPMCd") = "" Then adc("UoPMCd") = "DM"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  ' ***** Manali 3.6.0 - Catalogue Entry removed from Module Check
  ' ***** Manali 3.5.0 -10/12/08 - DsgCat Module
'  If mb_DsgCatMod = False Then
'    adc("UoDpCatFr") = "": adc("UoDpCatTo") = "": adc("UoDpCatSel") = "": adc("UoYN13") = "N"
'  End If
  ' ***** Manali 3.5.0 -10/12/08 - DsgCat Module
  
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  If adc("UoYN3") = "Y" Then OptRmPrd(0).Value = True
  If adc("UoYN3") = "N" Then OptRmPrd(1).Value = True
  
  'Design Valid YN - 4.4.0.0
  If adc("UoYN15") = "Y" Then OptValid(0).Value = True
  If adc("UoYN15") = "N" Then OptValid(1).Value = True
  If adc("UoYN15") = "" Then OptValid(2).Value = True
    
  ' ********* Manali 350Nxt - adc("UoYN2") changed to adc("UoDsgExc") to accommodate new option
  ''  If adc("UoYN2") = "Y" Then
  ''    OptDsgExc(0).Value = True
  ''  ElseIf adc("UoYN2") = "N" Then
  ''    OptDsgExc(1).Value = True
  ''  ElseIf adc("UoYN2") = "" Then
  ''    OptDsgExc(2).Value = True
  ''  End If
  '578-12 exclusive option disabled
  'If ADC("UoDsgExc") = "" Then ADC("UoDsgExc") = "F"
  '578-12 exclusive option disabled
'  If ADC("UoDsgExc") = "R" Then
'    OptDsgExc(0).Value = True
'  ElseIf ADC("UoDsgExc") = "C" Then
'    OptDsgExc(1).Value = True
'  ElseIf ADC("UoDsgExc") = "F" Then
'    OptDsgExc(2).Value = True
'  ElseIf ADC("UoDsgExc") = "A" Then
'    OptDsgExc(3).Value = True
'  End If
  ' ********* Manali 350Nxt - adc("UoYN2") changed to adc("UoDsgExc") to accommodate new option

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  
  ' ***** Manali 3.5.0 - 22/11/08 - Check Boxes Added
  If adc("UoYN") = "" Then adc("UoYN") = "N"        ' ***** Manali 3.10.0 - 10/03/12 - Show Dsgs With UsrDetOpt 'D','C' SCtg Only
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  If adc("UoYN4") = "" Then adc("UoYN4") = "N"
  If adc("UoYN5") = "" Then adc("UoYN5") = "N"
  If adc("UoYN6") = "" Then adc("UoYN6") = "N"
  If adc("UoYN7") = "" Then adc("UoYN7") = "N"
  If adc("UoYN8") = "" Then adc("UoYN8") = "N"
  If adc("UoYN9") = "" Then adc("UoYN9") = "N"
  If adc("UoYN10") = "" Then adc("UoYN10") = "N"
  If adc("UoYN11") = "" Then adc("UoYN11") = "N"
  If adc("UoYN12") = "" Then adc("UoYN12") = "N"
  If adc("UoYN13") = "" Then adc("UoYN13") = "N"
  If adc("UoYN14") = "" Then adc("UoYN14") = "N"    '4.1.0.0 Sachin - RMCtg
  If adc("UoYN16") = "" Then adc("UoYN16") = "N"  '6.1_27
  
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked     ' ***** Manali 3.10.0 - 10/03/12 - Show Dsgs With UsrDetOpt 'D','C' SCtg Only
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If adc("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  If adc("UoYN5") = "Y" Then ChkBoxArr(5).Value = Checked Else ChkBoxArr(5).Value = Unchecked
  If adc("UoYN6") = "Y" Then ChkBoxArr(6).Value = Checked Else ChkBoxArr(6).Value = Unchecked
  If adc("UoYN7") = "Y" Then ChkBoxArr(7).Value = Checked Else ChkBoxArr(7).Value = Unchecked
  If adc("UoYN8") = "Y" Then ChkBoxArr(8).Value = Checked Else ChkBoxArr(8).Value = Unchecked
  If adc("UoYN9") = "Y" Then ChkBoxArr(9).Value = Checked Else ChkBoxArr(9).Value = Unchecked
  If adc("UoYN10") = "Y" Then ChkBoxArr(10).Value = Checked Else ChkBoxArr(10).Value = Unchecked
  If adc("UoYN11") = "Y" Then ChkBoxArr(11).Value = Checked Else ChkBoxArr(11).Value = Unchecked
  If adc("UoYN12") = "Y" Then ChkBoxArr(12).Value = Checked Else ChkBoxArr(12).Value = Unchecked
  If adc("UoYN13") = "Y" Then ChkBoxArr(13).Value = Checked Else ChkBoxArr(13).Value = Unchecked
  If adc("UoYN14") = "Y" Then ChkBoxArr(14).Value = Checked Else ChkBoxArr(14).Value = Unchecked    '4.1.0.0 Sachin - RMCtg
  If adc("UoYN16") = "Y" Then ChkBoxArr(16).Value = Checked Else ChkBoxArr(16).Value = Unchecked    '6.1_27
  
  ' ***** Manali 3.5.0 - 22/11/08 - Check Boxes Added
  'uni.4
  ChkBoxArr(2).Value = Checked
  adc("wAddAllSalCtg") = "Y"
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
'578-12 customer code disabled
'  Case Is = UCase("UoCmCdFr"), UCase("UoCmCdTo"), UCase("UoCmCdSEL")
'  ' ********* Manali 350Nxt - adc("UoYN2") changed to adc("UoDsgExc") to accommodate new option
'    If ADC("UoDsgExc") <> "F" Then
'       Cancel = False
'    ElseIf ADC("UoDsgExc") = "F" Then
'      Cancel = True
'    End If
  ''    If adc("UoYN2") <> "" Then
  ''       Cancel = False
  ''    ElseIf adc("UoYN2") = "" Then
  ''      Cancel = True
  ''    End If
  ' ********* Manali 350Nxt - adc("UoYN2") changed to adc("UoDsgExc") to accommodate new option
  End Select
  gs_PTyp1 = "DMTCTYP"
  SetProp Me, IdName, When
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

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
Select Case UCase(IdName)
'578-12 customer code disabled
'  Case Is = UCase("UoCmCdFr")
'  ' ********* Manali 350Nxt - adc("UoYN2") changed to adc("UoDsgExc") to accommodate new option
'    ''If adc("UoYN2") = "" Then
'    If ADC("UoDsgExc") = "F" Then
'  ' ********* Manali 350Nxt - adc("UoYN2") changed to adc("UoDsgExc") to accommodate new option
'      ADC("UoCmCdFr") = ""
'      ADC("UoCmCdTo") = ""
'      ADC("UoCmCdSel") = ""
'    End If
  End Select
End Sub
' ***** Manali 3.5.0 - 22/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
Select Case Index
' ***** Manali 3.10.0 - 10/03/12 - Show Dsgs With UsrDetOpt 'D','C' SCtg Only
Case Is = 0
  If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
' ***** Manali 3.10.0 - 10/03/12 - Show Dsgs With UsrDetOpt 'D','C' SCtg Only
Case Is = 1
  If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
'uni.4
Case Is = 2
  If .Value = Checked Then adc("wAddAllSalCtg") = "Y" Else adc("wAddAllSalCtg") = "N"
Case Is = 4
  If .Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"
Case Is = 5
  If .Value = Checked Then adc("UoYN5") = "Y" Else adc("UoYN5") = "N"
Case Is = 6
  If .Value = Checked Then adc("UoYN6") = "Y" Else adc("UoYN6") = "N"
Case Is = 7
  If .Value = Checked Then adc("UoYN7") = "Y" Else adc("UoYN7") = "N"
Case Is = 8
  If .Value = Checked Then adc("UoYN8") = "Y" Else adc("UoYN8") = "N"
Case Is = 9
  If .Value = Checked Then adc("UoYN9") = "Y" Else adc("UoYN9") = "N"
Case Is = 10
  If .Value = Checked Then adc("UoYN10") = "Y" Else adc("UoYN10") = "N"
Case Is = 11
  If .Value = Checked Then adc("UoYN11") = "Y" Else adc("UoYN11") = "N"
Case Is = 12
  If .Value = Checked Then adc("UoYN12") = "Y" Else adc("UoYN12") = "N"
Case Is = 13      ' Tag = "ShowDsgCat"    ' ***** Manali 3.5.0 - 08/12/08
  If .Value = Checked Then adc("UoYN13") = "Y" Else: adc("UoYN13") = "N"
Case Is = 14      '4.1.0.0 - Sachin
  If .Value = Checked Then adc("UoYN14") = "Y" Else: adc("UoYN14") = "N"
Case Is = 16      '6.1_27
  If .Value = Checked Then adc("UoYN16") = "Y" Else: adc("UoYN16") = "N"
End Select
End With
End Sub
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
Private Sub CmdChgImg_Click()
  If FraChgImg.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraChgImg.Enabled = True
    FraChgImg.Visible = True
    TAB_REP.Enabled = False
    adc.AllowSave = False
    FraChgImg.ZOrder
  Else
    TAB_REP.Enabled = True
    adc.AllowSave = True
    FraChgImg.Enabled = False
    FraChgImg.Visible = False
  End If
End Sub
Private Sub CmdChgImgGo_Click()
On Error GoTo ErrHd
 If adc("wXlFilePath") = "" Then
    DispMsg "Please Provide an Excel File to change Images", etError
    Exit Sub
 End If
 If ReplXlImg(adc("wXlFilePath"), wo_rsPic, 110, 110) Then
    TAB_REP.Enabled = True
    adc.AllowSave = True
    FraChgImg.Visible = False
    CmdChgImg.Visible = False
    DispMsg "Images Replaced Successfully", etInfo
 End If
 Exit Sub
ErrHd:
    DispMsg Err.Description, etError
End Sub
Private Sub CmdDispSc_Click()
  If FraDispSc.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraDispSc.Enabled = True
    FraDispSc.Visible = True
    TAB_REP.Enabled = False
    adc.AllowSave = False
    FraDispSc.ZOrder
    CmdSC.Enabled = False
  Else
    TAB_REP.Enabled = True
    adc.AllowSave = True
    FraDispSc.Enabled = False
    FraDispSc.Visible = False
    CmdSC.Enabled = True
  End If
End Sub
' ****** Manali 3.5.0 - 10/11/08 - Dsg Cat fields display button added
Private Sub CmdSelFile_Click()
'Sachin - 3.11.2
    On Error GoTo ErrHd
    Dim wExtn As String
    wExtn = ".xls;*.xlsx"
    With Cd1
      .FileName = ""
      .DialogTitle = "Open Report Excel"
      .Filter = "Excel Files (*" + wExtn + ") |*" + wExtn
      .CancelError = False
      .ShowOpen
      If .FileName <> "" Then adc("WXLFILEPATH") = .FileName
    End With
    Exit Sub
ErrHd:
    DispMsg Err.Description, etError
End Sub
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
Private Sub GrdDsgAna_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  With GrdDsgAna
    Select Case UCase(ColName)
    Case Is = UCase("UdAnaSr")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                                "PTyp= 'DAANAFLD' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Analysis Sr.": Exit Sub
    End Select
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
Private Sub GrdDsgAna_InitKey(ByVal RowNum As Integer)
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
Private Sub GrdDsgAna_SetRecSource()
  '*** Set the Record Source of the Grid GrdDsgAna
  GrdDsgAna.RecSource = "Select * from UsrOptDsgAna where UdUsrCd= '" + adc("UoUsrCd") + "' and " + _
                        "UdMnuCd= '" + adc("UoMnuCd") + "'"
End Sub
'*** 2.11 (Jen)
Private Sub GrdDsgSz_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Select Case ColName
  ' ***** Manali 3.6.0 - 31/10/09 - Detail Scope applied for Colour Stones
  Case Is = UCase("UodRmCtg")
    Call HlpList.PMCd("RMCTG")
  
  Case Is = UCase("UodRmSCtg")
    Call HlpList.PSCd("RMSCTG", GrdDsgSz.Value(RowNum, "UodRmCtg"))
'    Call HlpList.PSCd("RMSCTG", "D")
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
'      Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'RMSCTG' and PMCd = 'D' and PSCd= '" + NewValue + "'")
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
'*** 2.11 (Jen)

Private Sub CmdSc_Click()
  If FraSC.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    '*** Jenny (Old Code Bef Color)
    'CmdSc.BackColor = vbWhite
    'CmdSc.FontBold = True
    '*** Jenny (Old Code Bef Color)
    FraSC.Enabled = True
    FraSC.Visible = True
    TAB_REP.Enabled = False
    adc.AllowSave = False
    FraSC.ZOrder
    CmdDispSc.Enabled = False
  Else
    '*** Jenny (Old Code Bef Color)
    'CmdSc.BackColor = &H8000000F
    'CmdSc.FontBold = False
    '*** Jenny (Old Code Bef Color)
    TAB_REP.Enabled = True
    adc.AllowSave = True
    FraSC.Enabled = False
    FraSC.Visible = False
    CmdDispSc.Enabled = True
  End If
End Sub

'*** std code not to be changed
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

Private Sub MWCTL_BTN11_Click()

End Sub

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
' ****** Manali 350Nxt - adc("UoYN2") changed to adc("UoDsgExc")
Private Sub OptDsgExc_Click(Index As Integer)
  '578-12 exclusive option disabled
'  Select Case Index
'  Case Is = 0
'    'adc("UoYN2") = "Y"
'    ADC("UoDsgExc") = "R"
'  Case Is = 1
'    'adc("UoYN2") = "N"
'    ADC("UoDsgExc") = "C"
'  Case Is = 2
'    'adc("UoYN2") = ""
'    ADC("UoDsgExc") = "F"
'  Case Is = 3
'    ADC("UoDsgExc") = "A"
'  End Select
End Sub
' ****** Manali 350Nxt - adc("UoYN2") changed to adc("UoDsgExc")
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptDsgExc_GotFocus(Index As Integer)
  DispMsg FraOptDsgExc.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptRmPrd_Click(Index As Integer)
  Select Case Index
  Case Is = 0
      adc("UoYN3") = "Y"
  Case Is = 1
      adc("UoYN3") = "N"
  End Select
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptRmPrd_GotFocus(Index As Integer)
  DispMsg FraOptRm.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

Private Sub OptValid_Click(Index As Integer)
'4.4.0.0
Select Case Index
  Case Is = 0
      adc("UoYN15") = "Y"
  Case Is = 1
      adc("UoYN15") = "N"
  Case Is = 2
      adc("UoYN15") = ""
  End Select
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
Private Sub GrdDsgAna_Validate(Cancel As Boolean)
  Cancel = GrdDsgAna.Validate
End Sub

'*** (Jen 2.14)
Private Sub GrdDsgSz_Validate(Cancel As Boolean)
    Cancel = GrdDsgSz.Validate
End Sub
'*** (Jen 2.14)

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set Rep = Nothing
    '*** (09/08/05)
End Sub


