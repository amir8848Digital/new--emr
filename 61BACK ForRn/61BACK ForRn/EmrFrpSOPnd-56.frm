VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrpSOPnd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Order Pending"
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
      Left            =   11070
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   9660
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4410
      TabIndex        =   127
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
      Left            =   5175
      TabIndex        =   135
      Top             =   9660
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   136
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
         TabIndex        =   137
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
      Left            =   6390
      TabIndex        =   134
      Top             =   9675
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   131
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
      TabIndex        =   128
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
      Height          =   9735
      Left            =   0
      TabIndex        =   130
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   129
         Top             =   0
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpSOPnd.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpSOPnd.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   0
            TabIndex        =   132
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   110
               Left            =   2450
               TabIndex        =   42
               ToolTipText     =   "Enter From Stock Year"
               Top             =   4410
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOBYYFR"
               IdName          =   "UOBYYFR"
               CmpStr          =   "B.OdDmStkYy >="
            End
            Begin VB.Frame FraOptValWt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   6255
               TabIndex        =   212
               ToolTipText     =   "Select if to show Value or Weight"
               Top             =   7300
               Width           =   2940
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
                  Left            =   1605
                  TabIndex        =   214
                  Top             =   0
                  Width           =   1425
               End
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
                  Left            =   180
                  TabIndex        =   213
                  Top             =   0
                  Width           =   1425
               End
            End
            Begin VB.Frame FraRmScp 
               BorderStyle     =   0  'None
               Height          =   1695
               Left            =   0
               TabIndex        =   199
               Top             =   7230
               Width           =   15105
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   82
                  Left            =   4410
                  TabIndex        =   86
                  ToolTipText     =   "Enter Raw Material Code To Range"
                  Top             =   570
                  Width           =   1905
                  _ExtentX        =   3360
                  _ExtentY        =   503
                  MaxLength       =   12
                  DataField       =   "UORMCDTO"
                  IdName          =   "UORMCDTO"
                  CmpStr          =   "OrRmCd <="
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   83
                  Left            =   7110
                  TabIndex        =   87
                  ToolTipText     =   "Enter Raw Material Code Selection"
                  Top             =   570
                  Width           =   7650
                  _ExtentX        =   13494
                  _ExtentY        =   503
                  MaxLength       =   255
                  DataField       =   "UORMCDSEL"
                  IdName          =   "UORMCDSEL"
                  CmpStr          =   "OrRmCd In"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   84
                  Left            =   4410
                  TabIndex        =   83
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
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   85
                  Left            =   7110
                  TabIndex        =   84
                  ToolTipText     =   "Enter Raw Material Sub Category Selection"
                  Top             =   285
                  Width           =   7650
                  _ExtentX        =   13494
                  _ExtentY        =   503
                  MaxLength       =   255
                  DataField       =   "UORMSCTGSEL"
                  IdName          =   "UORMSCTGSEL"
                  CmpStr          =   "OrRmSCtg In"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   86
                  Left            =   1725
                  TabIndex        =   79
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
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   87
                  Left            =   1725
                  TabIndex        =   82
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
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   88
                  Left            =   1725
                  TabIndex        =   85
                  ToolTipText     =   "Enter Raw Material Code From Range"
                  Top             =   570
                  Width           =   1905
                  _ExtentX        =   3360
                  _ExtentY        =   503
                  MaxLength       =   12
                  DataField       =   "UORMCDFR"
                  IdName          =   "UORMCDFR"
                  CmpStr          =   "OrRmCd >="
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   89
                  Left            =   4410
                  TabIndex        =   89
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
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   90
                  Left            =   1725
                  TabIndex        =   88
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
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   91
                  Left            =   4410
                  TabIndex        =   80
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
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   92
                  Left            =   7110
                  TabIndex        =   81
                  ToolTipText     =   "Enter Raw Material Category Selection"
                  Top             =   0
                  Width           =   7650
                  _ExtentX        =   13494
                  _ExtentY        =   503
                  MaxLength       =   255
                  DataField       =   "UORMCTGSEL"
                  IdName          =   "UORMCTGSEL"
                  CmpStr          =   "OrRmCtg In "
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   93
                  Left            =   4410
                  TabIndex        =   91
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
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   94
                  Left            =   1725
                  TabIndex        =   90
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
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   95
                  Left            =   4410
                  TabIndex        =   93
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
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   96
                  Left            =   1725
                  TabIndex        =   92
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
                  TabIndex        =   205
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
                  TabIndex        =   204
                  Top             =   1425
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
                  Index           =   22
                  Left            =   90
                  TabIndex        =   203
                  Top             =   570
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
                  Index           =   6
                  Left            =   90
                  TabIndex        =   202
                  Top             =   285
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
                  Index           =   5
                  Left            =   90
                  TabIndex        =   201
                  ToolTipText     =   "Location"
                  Top             =   0
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
                  TabIndex        =   200
                  Top             =   855
                  Width           =   1425
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Picture  "
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
               Left            =   30
               TabIndex        =   99
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   8445
               Width           =   1890
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Box Det               "
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
               TabIndex        =   95
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show box Details"
               Top             =   7590
               Width           =   1890
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Only Pending      "
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
               Left            =   4410
               TabIndex        =   72
               Tag             =   "ShowPnd"
               ToolTipText     =   "Check To Show Pending Records"
               Top             =   6975
               Width           =   2250
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Spl Remark               "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Index           =   4
               Left            =   30
               TabIndex        =   97
               Tag             =   "ShowRem"
               ToolTipText     =   "Check To Show 2nd Line"
               Top             =   7875
               Width           =   1890
            End
            Begin VB.Frame FraOptDelDt 
               BorderStyle     =   0  'None
               Height          =   855
               Left            =   6420
               TabIndex        =   195
               ToolTipText     =   "Select Exp Del Date/ Prd Del Date/ Both Dates"
               Top             =   7680
               Width           =   2235
               Begin VB.OptionButton OptDelDt 
                  Caption         =   "Both Del Dates"
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
                  Left            =   0
                  TabIndex        =   78
                  ToolTipText     =   "Select Exp Del Date/ Prd Del Date/ Both Dates"
                  Top             =   570
                  Width           =   1995
               End
               Begin VB.OptionButton OptDelDt 
                  Caption         =   "Prd Del Date"
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
                  Left            =   0
                  TabIndex        =   77
                  ToolTipText     =   "Select Exp Del Date/ Prd Del Date/ Both Dates"
                  Top             =   285
                  Width           =   1995
               End
               Begin VB.OptionButton OptDelDt 
                  Caption         =   "Export Del Date"
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
                  TabIndex        =   76
                  ToolTipText     =   "Select Exp Del Date/ Prd Del Date/ Both Dates"
                  Top             =   0
                  Width           =   1995
               End
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   14
               Left            =   2145
               TabIndex        =   94
               ToolTipText     =   "Show Box Detail ? (Yes / No)"
               Top             =   7590
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   9
               Left            =   7185
               TabIndex        =   74
               ToolTipText     =   "Show Pending Orders Or All Orders? (Yes / No)"
               Top             =   6960
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   11
               Left            =   1725
               TabIndex        =   71
               ToolTipText     =   "Enter Pending As Of Date"
               Top             =   6975
               Visible         =   0   'False
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOASOFDT"
               IdName          =   "UOASOFDT"
               CmpStr          =   "BOm.OmDt <= "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   19
               Left            =   5175
               TabIndex        =   8
               ToolTipText     =   "Enter To Order Year"
               Top             =   960
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "B.OdYy <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   21
               Left            =   2505
               TabIndex        =   5
               ToolTipText     =   "Enter From Order Year"
               Top             =   960
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "B.OdYy >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   22
               Left            =   1725
               TabIndex        =   4
               ToolTipText     =   "Enter From Order Tc"
               Top             =   960
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               ReCalcOn        =   "UOOMTC"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "B.OdTc >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   23
               Left            =   2925
               TabIndex        =   6
               ToolTipText     =   "Enter From Order Character"
               Top             =   960
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "B.OdChr >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   24
               Left            =   4410
               TabIndex        =   7
               ToolTipText     =   "Enter To Order Tc"
               Top             =   960
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "B.OdTc <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   25
               Left            =   5595
               TabIndex        =   9
               ToolTipText     =   "Enter To Order Character"
               Top             =   960
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "B.OdChr <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   26
               Left            =   5295
               TabIndex        =   13
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1245
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "B.OdSr <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   27
               Left            =   2625
               TabIndex        =   11
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1245
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "B.OdSr >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   28
               Left            =   1725
               TabIndex        =   10
               ToolTipText     =   "Enter From Order Number"
               Top             =   1245
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "B.OdNo >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   29
               Left            =   4410
               TabIndex        =   12
               ToolTipText     =   "Enter To Order Number"
               Top             =   1245
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "B.OdNo <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   46
               Left            =   4410
               TabIndex        =   15
               ToolTipText     =   "Enter To Order Date"
               Top             =   1530
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "BOm.OmDt <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   48
               Left            =   1725
               TabIndex        =   14
               ToolTipText     =   "Enter From Order Date"
               Top             =   1530
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "BOm.OmDt >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   40
               Left            =   4410
               TabIndex        =   17
               ToolTipText     =   "Enter To Customer"
               Top             =   1815
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "BOm.OmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   41
               Left            =   7110
               TabIndex        =   18
               ToolTipText     =   "Enter Customer Selection"
               Top             =   1815
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "BOm.OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   42
               Left            =   1725
               TabIndex        =   16
               ToolTipText     =   "Enter From Customer"
               Top             =   1815
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "BOm.OmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   43
               Left            =   4410
               TabIndex        =   26
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   2700
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOTO"
               IdName          =   "UOPONOTO"
               CmpStr          =   "BOm.OmPoNo <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   5
               Left            =   7110
               TabIndex        =   27
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   2700
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "BOm.OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   7
               Left            =   1725
               TabIndex        =   25
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   2700
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOFR"
               IdName          =   "UOPONOFR"
               CmpStr          =   "BOm.OmPoNo >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   0
               Left            =   4410
               TabIndex        =   32
               ToolTipText     =   "Enter To Design Category"
               Top             =   3270
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGTO"
               IdName          =   "UODMCTGTO"
               CmpStr          =   "DmCtg <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   1
               Left            =   7110
               TabIndex        =   33
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   3270
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
               CmpStr          =   "DmCtg In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   30
               Left            =   1725
               TabIndex        =   31
               ToolTipText     =   "Enter From Design Category"
               Top             =   3270
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGFR"
               IdName          =   "UODMCTGFR"
               CmpStr          =   "DmCtg >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   31
               Left            =   4410
               TabIndex        =   35
               ToolTipText     =   "Enter To Design Code"
               Top             =   3555
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "B.OdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   32
               Left            =   7110
               TabIndex        =   36
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   3555
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "B.OdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   33
               Left            =   1725
               TabIndex        =   34
               ToolTipText     =   "Enter From Design Code"
               Top             =   3555
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "B.OdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   34
               Left            =   4410
               TabIndex        =   38
               ToolTipText     =   "Enter To Design Size"
               Top             =   3840
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
               CmpStr          =   "B.OdDmSz <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   36
               Left            =   1725
               TabIndex        =   37
               ToolTipText     =   "Enter From Design Size"
               Top             =   3840
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "B.OdDmSz >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   37
               Left            =   4410
               TabIndex        =   40
               ToolTipText     =   "Enter To Suffix"
               Top             =   4125
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOODSFXTO"
               IdName          =   "UOODSFXTO"
               CmpStr          =   "B.OdSfx <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   39
               Left            =   1725
               TabIndex        =   39
               ToolTipText     =   "Enter From Suffix"
               Top             =   4125
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOODSFXFR"
               IdName          =   "UOODSFXFR"
               CmpStr          =   "B.OdSfx >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   49
               Left            =   4410
               TabIndex        =   66
               ToolTipText     =   "Enter To Export Delivery Date"
               Top             =   6120
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTTO"
               IdName          =   "UODELDTTO"
               CmpStr          =   "B.OdExpDelDt <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   51
               Left            =   1725
               TabIndex        =   65
               ToolTipText     =   "Enter From Export Delivery Date"
               Top             =   6120
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTFR"
               IdName          =   "UODELDTFR"
               CmpStr          =   "B.OdExpDelDt >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   15
               Left            =   4410
               TabIndex        =   52
               ToolTipText     =   "Enter To Karat"
               Top             =   4695
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTTO"
               IdName          =   "UOKTTO"
               CmpStr          =   "B.OdKt <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   16
               Left            =   7110
               TabIndex        =   53
               ToolTipText     =   "Enter Karat Selection"
               Top             =   4695
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "B.OdKt In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   2
               Left            =   1725
               TabIndex        =   51
               ToolTipText     =   "Enter From Karat"
               Top             =   4695
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               CmpStr          =   "B.OdKt >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   3
               Left            =   4410
               TabIndex        =   58
               ToolTipText     =   "Enter To Design Sales Category"
               Top             =   5265
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGTO"
               IdName          =   "UODMSALCTGTO"
               CmpStr          =   "DmSalCtg<="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   12
               Left            =   7110
               TabIndex        =   59
               ToolTipText     =   "Enter Design Sales Category Selection"
               Top             =   5265
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMSALCTGSEL"
               IdName          =   "UODMSALCTGSEL"
               CmpStr          =   "DmSalCtg In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   4
               Left            =   1725
               TabIndex        =   57
               ToolTipText     =   "Enter From Design Sales Category"
               Top             =   5265
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGFR"
               IdName          =   "UODMSALCTGFR"
               CmpStr          =   "DmSalCtg >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   6
               Left            =   4410
               TabIndex        =   55
               ToolTipText     =   "Enter To Priority Code"
               Top             =   4980
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOPRTCDTO"
               IdName          =   "UOPRTCDTO"
               CmpStr          =   "B.OdPrtCd <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   8
               Left            =   7110
               TabIndex        =   56
               ToolTipText     =   "Enter Priority Code Selection"
               Top             =   4980
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRTCDSEL"
               IdName          =   "UOPRTCDSEL"
               CmpStr          =   "B.OdPrtCd In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   10
               Left            =   1725
               TabIndex        =   54
               ToolTipText     =   "Enter From Priority Code"
               Top             =   4980
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOPRTCDFR"
               IdName          =   "UOPRTCDFR"
               CmpStr          =   "B.OdPrtCd >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   13
               Left            =   4410
               TabIndex        =   68
               ToolTipText     =   "Enter To Balance Quantity"
               Top             =   6405
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOBALPRDTO"
               IdName          =   "UOBALPRDTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   17
               Left            =   1725
               TabIndex        =   67
               ToolTipText     =   "Enter From Balance Quantity"
               Top             =   6405
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOBALPRDFR"
               IdName          =   "UOBALPRDFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   18
               Left            =   4410
               TabIndex        =   70
               ToolTipText     =   "Enter To Balance Value"
               Top             =   6690
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTTO"
               IdName          =   "UOCSTRTTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   20
               Left            =   1725
               TabIndex        =   69
               ToolTipText     =   "Enter From Balance Value"
               Top             =   6690
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTFR"
               IdName          =   "UOCSTRTFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   35
               Left            =   1725
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   390
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "B.OdCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   38
               Left            =   2145
               TabIndex        =   98
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   8445
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   53
               Left            =   7110
               TabIndex        =   62
               ToolTipText     =   "Enter Sales Executive Selection"
               Top             =   5550
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSALPERSEL"
               IdName          =   "UOSALPERSEL"
               CmpStr          =   "OmCmSalPer In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   54
               Left            =   4410
               TabIndex        =   61
               ToolTipText     =   "Enter To Sales Executive"
               Top             =   5550
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERTO"
               IdName          =   "UOSALPERTO"
               CmpStr          =   "OmCmSalPer <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   55
               Left            =   1725
               TabIndex        =   60
               ToolTipText     =   "Enter From Sales Executive "
               Top             =   5550
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERFR"
               IdName          =   "UOSALPERFR"
               CmpStr          =   "OmCmSalPer  >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   44
               Left            =   10320
               TabIndex        =   75
               ToolTipText     =   "Show Export/Production/Both Delivery Dates? (Enter ""Y"" for Export , ""N"" for Production and ' ' for Both)"
               Top             =   6960
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   45
               Left            =   4410
               TabIndex        =   64
               ToolTipText     =   "Enter To Production Delivery Date"
               Top             =   5835
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1TO"
               IdName          =   "UODELDT1TO"
               CmpStr          =   "B.OdDelDt <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   47
               Left            =   1725
               TabIndex        =   63
               ToolTipText     =   "Enter From Production Delivery Date"
               Top             =   5835
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1FR"
               IdName          =   "UODELDT1FR"
               CmpStr          =   "B.OdDelDt >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   73
               Left            =   9285
               TabIndex        =   73
               ToolTipText     =   "Enter Currency Code Whose Orders You Wish To View Or Leave It Blank If You Wish To View All Orders In Base Currency"
               Top             =   6975
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               ReCalcOn        =   "UOINVAL"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   75
               Left            =   4410
               TabIndex        =   29
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   2985
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "B.OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   76
               Left            =   7110
               TabIndex        =   30
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   2985
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "B.OdPoNo In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   77
               Left            =   1725
               TabIndex        =   28
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   3000
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "B.OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   78
               Left            =   2145
               TabIndex        =   96
               ToolTipText     =   "Show Special Remark ? (Yes / No)"
               Top             =   7875
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN4"
               IdName          =   "UOYN4"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   80
               Left            =   1725
               TabIndex        =   3
               ToolTipText     =   "Enter Base Order Tc"
               Top             =   675
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTC"
               ReCalcParent    =   "UOOMTCFR"
               IdName          =   "UOOMTC"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   79
               Left            =   3480
               TabIndex        =   44
               ToolTipText     =   "Enter From Stock Number"
               Top             =   4410
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Mask            =   "#0"
               MaxLength       =   2
               DataType        =   1
               DataField       =   "UOODDMIDFR"
               IdName          =   "UOODDMIDFR"
               CmpStr          =   "B.OdDmStkNo>="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   81
               Left            =   6135
               TabIndex        =   48
               ToolTipText     =   "Enter To Stock Number"
               Top             =   4410
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Mask            =   "#0"
               MaxLength       =   2
               DataType        =   1
               DataField       =   "UOODDMIDTO"
               IdName          =   "UOODDMIDTO"
               CmpStr          =   "B.OdDmStkNo<="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   97
               Left            =   4410
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   390
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "B.OdCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   98
               Left            =   7110
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   390
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "B.OdCoCd In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   99
               Left            =   7110
               TabIndex        =   21
               ToolTipText     =   "Enter Customer Group Selection"
               Top             =   2115
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMGRPSEL"
               IdName          =   "UOCMGRPSEL"
               CmpStr          =   "CmGrp In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   100
               Left            =   4410
               TabIndex        =   20
               ToolTipText     =   "Enter To Customer Group"
               Top             =   2120
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMGRPTO"
               IdName          =   "UOCMGRPTO"
               CmpStr          =   "CmGrp <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   101
               Left            =   1725
               TabIndex        =   19
               ToolTipText     =   "Enter From Customer Group"
               Top             =   2120
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMGRPFR"
               IdName          =   "UOCMGRPFR"
               CmpStr          =   "CmGrp >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   102
               Left            =   7110
               TabIndex        =   24
               ToolTipText     =   "Enter Customer Type Selection"
               Top             =   2415
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMTYPSEL"
               IdName          =   "UOCMTYPSEL"
               CmpStr          =   "CmTyp In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   103
               Left            =   4410
               TabIndex        =   23
               ToolTipText     =   "Enter To Customer Type"
               Top             =   2410
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMTYPTO"
               IdName          =   "UOCMTYPTO"
               CmpStr          =   "CmTyp <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   104
               Left            =   1725
               TabIndex        =   22
               ToolTipText     =   "Enter From Customer Type"
               Top             =   2410
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMTYPFR"
               IdName          =   "UOCMTYPFR"
               CmpStr          =   "CmTyp >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   107
               Left            =   1725
               TabIndex        =   209
               ToolTipText     =   "Specify Whether Value/ Weight To Be Shown Is Actual, Invoice Or Customs "
               Top             =   7200
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataField       =   "UOINVAL"
               IdName          =   "UOINVAL"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   108
               Left            =   9360
               TabIndex        =   215
               ToolTipText     =   "Show Value Or Weight Figures ('Y' for Value / 'N' For Weight)"
               Top             =   7320
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN5"
               IdName          =   "UOYN5"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   109
               Left            =   2820
               TabIndex        =   43
               ToolTipText     =   "Enter From Stock Character"
               Top             =   4410
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOBCHRFR"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "B.OdDmStkChr >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   113
               Left            =   5025
               TabIndex        =   46
               ToolTipText     =   "Enter To Stock Year"
               Top             =   4410
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOBYYTO"
               IdName          =   "UOBYYTO"
               CmpStr          =   "B.OdDmStkYy <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   114
               Left            =   5400
               TabIndex        =   47
               ToolTipText     =   "Enter To Stock Character"
               Top             =   4410
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOBCHRTO"
               IdName          =   "UOBCHRTO"
               CmpStr          =   "B.OdDmStkChr <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   106
               Left            =   4410
               TabIndex        =   50
               ToolTipText     =   "Enter To Production Line"
               Top             =   4410
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOLINETO"
               IdName          =   "UOLINETO"
               CmpStr          =   "B.OdLine <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   111
               Left            =   1725
               TabIndex        =   41
               ToolTipText     =   "Enter From Stock Company Code"
               Top             =   4395
               Width           =   680
               _ExtentX        =   1191
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOSTKCOCDFR"
               IdName          =   "UOSTKCOCDFR"
               CmpStr          =   "B.OdDmStkCoCd>= "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   112
               Left            =   4410
               TabIndex        =   45
               ToolTipText     =   "Enter To Stock Company Code"
               Top             =   4410
               Width           =   680
               _ExtentX        =   1191
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOSTKCOCDTO"
               IdName          =   "UOSTKCOCDTO"
               CmpStr          =   "B.OdDmStkCoCd<= "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   105
               Left            =   1725
               TabIndex        =   49
               ToolTipText     =   "Enter From Production Line"
               Top             =   4410
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOLINEFR"
               IdName          =   "UOLINEFR"
               CmpStr          =   "B.OdLine >="
            End
            Begin VB.Label LblVal 
               BackStyle       =   0  'Transparent
               Caption         =   "Show "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4440
               TabIndex        =   211
               Top             =   7320
               Width           =   1665
            End
            Begin VB.Label LblValAs 
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
               TabIndex        =   210
               Top             =   7215
               Width           =   1605
            End
            Begin VB.Label LblLine 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Line"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   208
               Top             =   4410
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
               Index           =   34
               Left            =   60
               TabIndex        =   207
               Top             =   2410
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
               TabIndex        =   206
               Top             =   2120
               Width           =   1425
            End
            Begin VB.Label LblDmId 
               BackStyle       =   0  'Transparent
               Caption         =   "StkNum"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   198
               Top             =   4410
               Width           =   1665
            End
            Begin VB.Label LblBaseTc 
               BackStyle       =   0  'Transparent
               Caption         =   "Base Tc"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   197
               Top             =   675
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Spl Remark"
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
               Height          =   570
               Index           =   29
               Left            =   60
               TabIndex        =   196
               ToolTipText     =   "Location"
               Top             =   7890
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label ALBL 
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
               Index           =   28
               Left            =   60
               TabIndex        =   194
               Top             =   2985
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
               Left            =   7860
               TabIndex        =   193
               Top             =   6975
               Width           =   1005
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
               TabIndex        =   168
               Top             =   5835
               Width           =   1425
            End
            Begin VB.Label LblShowDelDt 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Delivery Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4425
               TabIndex        =   167
               ToolTipText     =   "Location"
               Top             =   7680
               Width           =   1875
            End
            Begin VB.Label LblSalPer 
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
               Left            =   60
               TabIndex        =   166
               Top             =   5550
               Width           =   1755
            End
            Begin VB.Label LblPic 
               BackStyle       =   0  'Transparent
               Caption         =   "Picture"
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
               Left            =   30
               TabIndex        =   165
               ToolTipText     =   "Location"
               Top             =   8445
               Visible         =   0   'False
               Width           =   855
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
               Index           =   25
               Left            =   60
               TabIndex        =   164
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Balance Val"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   60
               TabIndex        =   163
               Top             =   6690
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Balance Qty"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   162
               Top             =   6405
               Width           =   1425
            End
            Begin VB.Label LblPrt 
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
               Left            =   60
               TabIndex        =   161
               Top             =   4980
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Ctg"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   160
               Top             =   5265
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Karat"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   159
               Top             =   4695
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Exp Del Dt"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   158
               Top             =   6120
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               Index           =   15
               Left            =   60
               TabIndex        =   157
               Top             =   3840
               Width           =   1425
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
               TabIndex        =   156
               Top             =   4125
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
               Index           =   7
               Left            =   60
               TabIndex        =   155
               Top             =   3270
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
               Index           =   0
               Left            =   60
               TabIndex        =   154
               Top             =   3555
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
               Index           =   18
               Left            =   60
               TabIndex        =   153
               Top             =   2700
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
               TabIndex        =   152
               Top             =   1815
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
               TabIndex        =   151
               Top             =   1530
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
               Index           =   31
               Left            =   2565
               TabIndex        =   150
               Top             =   1245
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
               Left            =   5235
               TabIndex        =   149
               Top             =   1245
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
               Index           =   14
               Left            =   2445
               TabIndex        =   148
               Top             =   960
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
               Index           =   13
               Left            =   5115
               TabIndex        =   147
               Top             =   960
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
               Left            =   2865
               TabIndex        =   146
               Top             =   960
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
               Left            =   5535
               TabIndex        =   145
               Top             =   960
               Width           =   105
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
               TabIndex        =   144
               Top             =   1245
               Width           =   1485
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
               TabIndex        =   143
               Top             =   960
               Width           =   1785
            End
            Begin VB.Label LblAsOfDt 
               BackStyle       =   0  'Transparent
               Caption         =   "As Of Date"
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
               Left            =   60
               TabIndex        =   142
               Top             =   6975
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.Label LblBoxDet 
               BackStyle       =   0  'Transparent
               Caption         =   "Box Det (Y/N)"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   141
               ToolTipText     =   "Location"
               Top             =   7590
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sel"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   7110
               TabIndex        =   140
               Top             =   60
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
               Left            =   1725
               TabIndex        =   139
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
               Left            =   4410
               TabIndex        =   138
               Top             =   60
               Width           =   1275
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   133
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraGrph 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   8070
               Left            =   150
               TabIndex        =   122
               Top             =   1050
               Width           =   11160
               Begin VB.CheckBox ChkGrph 
                  Caption         =   "Show Graph"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   120
                  MaskColor       =   &H8000000F&
                  TabIndex        =   101
                  Top             =   0
                  UseMaskColor    =   -1  'True
                  Width           =   1875
               End
               Begin VB.Frame FraGrphDet 
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   7545
                  Left            =   -30
                  TabIndex        =   169
                  Top             =   450
                  Width           =   11145
                  Begin VB.ComboBox cboGrphValue 
                     Height          =   315
                     ItemData        =   "EmrFrpSOPnd.frx":0038
                     Left            =   2100
                     List            =   "EmrFrpSOPnd.frx":004E
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   106
                     ToolTipText     =   "Select Graph Value"
                     Top             =   915
                     Width           =   2865
                  End
                  Begin TabDlg.SSTab TAB_GRPH 
                     Height          =   5805
                     Left            =   0
                     TabIndex        =   110
                     Top             =   1650
                     Width           =   11115
                     _ExtentX        =   19606
                     _ExtentY        =   10239
                     _Version        =   393216
                     Tabs            =   2
                     TabsPerRow      =   2
                     TabHeight       =   520
                     TabCaption(0)   =   "T&itles"
                     TabPicture(0)   =   "EmrFrpSOPnd.frx":0081
                     Tab(0).ControlEnabled=   -1  'True
                     Tab(0).Control(0)=   "FraTabGrph(0)"
                     Tab(0).Control(0).Enabled=   0   'False
                     Tab(0).ControlCount=   1
                     TabCaption(1)   =   "Top &N"
                     TabPicture(1)   =   "EmrFrpSOPnd.frx":009D
                     Tab(1).ControlEnabled=   0   'False
                     Tab(1).Control(0)=   "FraTabGrph(1)"
                     Tab(1).ControlCount=   1
                     Begin VB.Frame FraTabGrph 
                        BorderStyle     =   0  'None
                        Height          =   5410
                        Index           =   1
                        Left            =   -74940
                        TabIndex        =   184
                        Top             =   360
                        Width           =   11030
                        Begin MwfCtl.MWCTL_MED ATXT 
                           Height          =   285
                           Index           =   64
                           Left            =   2310
                           TabIndex        =   118
                           ToolTipText     =   "Enter Sorting Type for Group1"
                           Top             =   1110
                           Width           =   2115
                           _ExtentX        =   3731
                           _ExtentY        =   503
                           MaxLength       =   8
                           DataType        =   4
                           DataField       =   "UOGRPHG1TOPN"
                           IdName          =   "UOGRPHG1TOPN"
                        End
                        Begin MwfCtl.MWCTL_MED ATXT 
                           Height          =   285
                           Index           =   65
                           Left            =   5580
                           TabIndex        =   119
                           ToolTipText     =   "Enter Top N Value For Group1"
                           Top             =   1110
                           Width           =   945
                           _ExtentX        =   1667
                           _ExtentY        =   503
                           Mask            =   "####0"
                           MaxLength       =   5
                           DataType        =   1
                           DataField       =   "UOGRPHG1NVAL"
                           IdName          =   "UOGRPHG1NVAL"
                        End
                        Begin MwfCtl.MWCTL_MED ATXT 
                           Height          =   285
                           Index           =   69
                           Left            =   1230
                           TabIndex        =   117
                           ToolTipText     =   "Enter Top N Based On Summary"
                           Top             =   480
                           Width           =   2115
                           _ExtentX        =   3731
                           _ExtentY        =   503
                           MaxLength       =   8
                           DataType        =   4
                           DataField       =   "UOGRPHSUMM"
                           IdName          =   "UOGRPHSUMM"
                        End
                        Begin MwfCtl.MWCTL_MED ATXT 
                           Height          =   285
                           Index           =   66
                           Left            =   2310
                           TabIndex        =   120
                           ToolTipText     =   "Enter Sorting Type for Group2"
                           Top             =   1740
                           Width           =   2115
                           _ExtentX        =   3731
                           _ExtentY        =   503
                           MaxLength       =   8
                           DataType        =   4
                           DataField       =   "UOGRPHG2TOPN"
                           IdName          =   "UOGRPHG2TOPN"
                        End
                        Begin MwfCtl.MWCTL_MED ATXT 
                           Height          =   285
                           Index           =   67
                           Left            =   5580
                           TabIndex        =   121
                           ToolTipText     =   "Enter Top N Value For Group2"
                           Top             =   1740
                           Width           =   945
                           _ExtentX        =   1667
                           _ExtentY        =   503
                           Mask            =   "####0"
                           MaxLength       =   5
                           DataType        =   1
                           DataField       =   "UOGRPHG2NVAL"
                           IdName          =   "UOGRPHG2NVAL"
                        End
                        Begin VB.Label LblGrphG2Val 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
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
                           Height          =   240
                           Left            =   4920
                           TabIndex        =   191
                           Top             =   1755
                           Width           =   615
                        End
                        Begin VB.Label LblGrphG2Srt 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Sort Type"
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
                           Left            =   1230
                           TabIndex        =   190
                           Top             =   1755
                           Width           =   1035
                        End
                        Begin VB.Label LblGrphGrp2 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Group 2"
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
                           Left            =   150
                           TabIndex        =   189
                           Top             =   1770
                           Width           =   825
                        End
                        Begin VB.Label LblGrphGrp1 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Group 1"
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
                           Left            =   150
                           TabIndex        =   188
                           Top             =   1125
                           Width           =   825
                        End
                        Begin VB.Label LblGrphG1Srt 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Sort Type"
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
                           Left            =   1230
                           TabIndex        =   187
                           Top             =   1125
                           Width           =   1035
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Based On"
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
                           Left            =   150
                           TabIndex        =   186
                           Top             =   495
                           Width           =   1035
                        End
                        Begin VB.Label LblGrphG1Val 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
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
                           Height          =   240
                           Left            =   4920
                           TabIndex        =   185
                           Top             =   1125
                           Width           =   615
                        End
                     End
                     Begin VB.Frame FraTabGrph 
                        BorderStyle     =   0  'None
                        Height          =   5410
                        Index           =   0
                        Left            =   60
                        TabIndex        =   125
                        Top             =   360
                        Width           =   11030
                        Begin MwfCtl.MWCTL_MED ATXT 
                           Height          =   285
                           Index           =   58
                           Left            =   1680
                           TabIndex        =   111
                           ToolTipText     =   "Enter Graph Title"
                           Top             =   435
                           Width           =   6975
                           _ExtentX        =   12303
                           _ExtentY        =   503
                           MaxLength       =   30
                           DataField       =   "UOGRPHMTITLE"
                           IdName          =   "UOGRPHMTITLE"
                        End
                        Begin MwfCtl.MWCTL_MED ATXT 
                           Height          =   285
                           Index           =   59
                           Left            =   1680
                           TabIndex        =   112
                           ToolTipText     =   "Enter Graph Sub Ttile"
                           Top             =   720
                           Width           =   6975
                           _ExtentX        =   12303
                           _ExtentY        =   503
                           MaxLength       =   30
                           DataField       =   "UOGRPHSTITLE"
                           IdName          =   "UOGRPHSTITLE"
                        End
                        Begin MwfCtl.MWCTL_MED ATXT 
                           Height          =   285
                           Index           =   61
                           Left            =   1680
                           TabIndex        =   114
                           ToolTipText     =   "Enter Group Title"
                           Top             =   1290
                           Width           =   6975
                           _ExtentX        =   12303
                           _ExtentY        =   503
                           MaxLength       =   30
                           DataField       =   "UOGRPHGTITLE"
                           IdName          =   "UOGRPHGTITLE"
                        End
                        Begin MwfCtl.MWCTL_MED ATXT 
                           Height          =   285
                           Index           =   62
                           Left            =   1680
                           TabIndex        =   115
                           ToolTipText     =   "Enter Data Title"
                           Top             =   1575
                           Width           =   6975
                           _ExtentX        =   12303
                           _ExtentY        =   503
                           MaxLength       =   30
                           DataField       =   "UOGRPHDTITLE"
                           IdName          =   "UOGRPHDTITLE"
                        End
                        Begin MwfCtl.MWCTL_MED ATXT 
                           Height          =   285
                           Index           =   63
                           Left            =   1680
                           TabIndex        =   116
                           ToolTipText     =   "Enter Data2 Title"
                           Top             =   1860
                           Width           =   6975
                           _ExtentX        =   12303
                           _ExtentY        =   503
                           MaxLength       =   30
                           DataField       =   "UOGRPHDTITLE2"
                           IdName          =   "UOGRPHDTITLE2"
                        End
                        Begin MwfCtl.MWCTL_MED ATXT 
                           Height          =   285
                           Index           =   60
                           Left            =   1680
                           TabIndex        =   113
                           ToolTipText     =   "Enter Graph Foot Note"
                           Top             =   1005
                           Width           =   6975
                           _ExtentX        =   12303
                           _ExtentY        =   503
                           MaxLength       =   30
                           DataField       =   "UOGRPHFTNOTE"
                           IdName          =   "UOGRPHFTNOTE"
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Graph Title"
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
                           Left            =   330
                           TabIndex        =   175
                           Top             =   435
                           Width           =   1170
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Sub Title"
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
                           Index           =   4
                           Left            =   330
                           TabIndex        =   174
                           Top             =   720
                           Width           =   945
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Group Title"
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
                           Index           =   6
                           Left            =   330
                           TabIndex        =   173
                           Top             =   1290
                           Width           =   1170
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Data Title"
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
                           Left            =   330
                           TabIndex        =   172
                           Top             =   1575
                           Width           =   1035
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Data2 Title"
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
                           Index           =   8
                           Left            =   330
                           TabIndex        =   171
                           Top             =   1860
                           Width           =   1155
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Foot Note"
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
                           Index           =   5
                           Left            =   330
                           TabIndex        =   170
                           Top             =   1005
                           Width           =   1035
                        End
                     End
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   56
                     Left            =   2100
                     TabIndex        =   102
                     ToolTipText     =   "Enter Graph Type"
                     Top             =   345
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "UOGRPHTYP"
                     IdName          =   "UOGRPHTYP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   57
                     Left            =   7620
                     TabIndex        =   103
                     ToolTipText     =   "Enter Graph irection"
                     Top             =   345
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "UOGRPHDIR"
                     IdName          =   "UOGRPHDIR"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   68
                     Left            =   2100
                     TabIndex        =   104
                     ToolTipText     =   "Enter Data Label"
                     Top             =   630
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "UOGRPHDATALB"
                     IdName          =   "UOGRPHDATALB"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   70
                     Left            =   7620
                     TabIndex        =   105
                     ToolTipText     =   "Enter Graph View Angle"
                     Top             =   630
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "UOGRPHVWANGL"
                     IdName          =   "UOGRPHVWANGL"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   315
                     Index           =   71
                     Left            =   4950
                     TabIndex        =   124
                     ToolTipText     =   "Enter From Customer"
                     Top             =   915
                     Visible         =   0   'False
                     Width           =   405
                     _ExtentX        =   714
                     _ExtentY        =   556
                     Mask            =   "#0"
                     MaxLength       =   2
                     DataType        =   1
                     DataField       =   "UOGRPHVALIND"
                     IdName          =   "UOGRPHVALIND"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   50
                     Left            =   7620
                     TabIndex        =   107
                     ToolTipText     =   "Enter Graph in Color (Y / N) And Blank For Default"
                     Top             =   915
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "UOGRPHCLR"
                     IdName          =   "UOGRPHCLR"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   52
                     Left            =   7620
                     TabIndex        =   109
                     ToolTipText     =   "Enter Graph View Angle"
                     Top             =   1200
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "UOGRPHORIENT"
                     IdName          =   "UOGRPHORIENT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   72
                     Left            =   2100
                     TabIndex        =   108
                     ToolTipText     =   "Enter Data Label"
                     Top             =   1210
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "UOGRPHPSZ"
                     IdName          =   "UOGRPHPSZ"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Graph Details"
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
                     TabIndex        =   192
                     Top             =   0
                     Width           =   11205
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Paper Size"
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
                     Index           =   20
                     Left            =   300
                     TabIndex        =   183
                     Top             =   1210
                     Width           =   1155
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Paper Orientation"
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
                     Index           =   19
                     Left            =   5760
                     TabIndex        =   182
                     Top             =   1200
                     Width           =   1830
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Graph in Color"
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
                     Index           =   13
                     Left            =   5760
                     TabIndex        =   181
                     Top             =   915
                     Width           =   1500
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "View Angle"
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
                     Left            =   5760
                     TabIndex        =   180
                     Top             =   630
                     Width           =   1170
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Graph Direction"
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
                     Left            =   5760
                     TabIndex        =   179
                     Top             =   345
                     Width           =   1635
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Data Label"
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
                     Left            =   300
                     TabIndex        =   178
                     Top             =   630
                     Width           =   1155
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Graph Type"
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
                     Index           =   9
                     Left            =   300
                     TabIndex        =   177
                     Top             =   345
                     Width           =   1245
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
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
                     Index           =   10
                     Left            =   300
                     TabIndex        =   176
                     Top             =   915
                     Width           =   615
                  End
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   74
                  Left            =   2010
                  TabIndex        =   123
                  ToolTipText     =   "Enter From Customer"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  DataField       =   "UOSHOWGRPH"
                  IdName          =   "UOSHOWGRPH"
               End
            End
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   100
               Top             =   360
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpSOPnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Report --------'
' This report is used to give a detail listing of all the pending orders in the system as
'   on a particular date. The report can also be used
' The user has the option of seeing/ hiding the corresponding matching invoice records.
'--------------------

Option Explicit
Dim Rep As New EmrRepSOPnd
Dim Rep1 As New EmrRepJtPndRm     ' Manali 3.7.1
Dim Rep2 As New EmrRepSalAnaDsg
Dim mRep As CRAXDRT.Report

Dim moCn As MwfLib.MDOConnection   '*** (Bef  speed) bef 24/11/06
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
Dim moGrph As EmrClsGrph
' ###########################################  Manoj  ###########################################

' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
Dim ws_MultiCoMod As Boolean

Private Sub SetGroupSort()
'VK.27 - Menu code "RepJtSoPndRm" added.
If UCase(adc.MenuCd) <> UCase("RepSOPndRM") And UCase(adc.MenuCd) <> UCase("RepJtSOPndRm") Then
   If ws_MultiCoMod = True Then _
  GRP_REP.Add "Company Code", "B.OdCoCd", "B.OdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=B.OdCoCd) "
  GRP_REP.Add "Order", "B.OdTc+B.OdYy+B.OdChr+str(B.OdNo)", "B.OdTc+'/'+B.OdYy+'/'+B.OdChr+'/'+LTrim(Str(B.OdNo))", "", "", "Order Cust,Sales Exec,Customer PO,Customer Grp,Customer Typ", "", "", "BOm.OmCmCd "
  GRP_REP.Add "Order Date", "convert(varchar(10), BOm.OmDt, 112)", "convert(varchar(8), BOm.OmDt, 3)", "OmDt", "hOmDt", "Quarter,Month,Week", "", "", ""
  GRP_REP.Add "Week", "LTrim(Str(DatePart(yyyy,BOm.OmDt)))+' - '+DateName(WW,BOm.OmDt)", "", "", "", "Quarter,Month", "", "", ""
  GRP_REP.Add "Month", "Str(DatePart(yyyy,BOm.OmDt))+Str(DatePart(mm,BOm.OmDt))", "LTrim(Str(DatePart(yyyy,BOm.OmDt)))+' '+DateName(MM,BOm.OmDt)", "", "", "Quarter", "", "", ""
  GRP_REP.Add "Quarter", "LTrim(Str(DatePart(yyyy,BOm.OmDt)))+' - '+DateName(QQ,BOm.OmDt)", "", "", "", "", "", "", ""
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  ' ***** Manali 3.7.1 - EmrFrmMDI.TXT_UCD.text Replaced by gs_UsrCd
  GRP_REP.Add "Order Cust", "BOm.OmCmCd", "", "OmCmCd", "hOmCmCd", "Sales Exec,Customer Grp,Customer Typ", "", "", "(Select ( case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= BOm.OmCmCd) "
  
  'Sachin - 4.1.0.0
  GRP_REP.Add "Customer Grp", "CmGrp", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMGRP' and PmCd = CmGrp)"
  GRP_REP.Add "Customer Typ", "CmTyp", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMTYP' and PmCd = CmTyp)"
  
  
  'GRP_REP.Add "Order Cust", "BOm.OmCmCd", "", "OmCmCd", "hOmCmCd", "Sales Exec", "", "", "(Select ( case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= BOm.OmCmCd) "
  
  ' ***** Manali 3.7.1
  End If
  
  'VK.27 - Menu code "RepJtSoPndRm" added.
  If UCase(adc.MenuCd) <> UCase("RepJtPndRm") And UCase(adc.MenuCd) <> UCase("RepSOPndRM") And UCase(adc.MenuCd) <> UCase("RepJtSOPndRm") Then
    '***** Manali 3.10.0 - 07/03/12 - Sales Ctg and DmCtg allowed below one another
    GRP_REP.Add "Design Ctg", "DmCtg", "", "DmCtg", "hDmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= DmCtg) "
    '***Geeta***Crm***212
    GRP_REP.Add "Sales Ctg", "DmSalCtg", "", "DmSalCtg", "hDmSalCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and Pmcd=DmCtg and PSCd= DmSalCtg) "
    GRP_REP.Add "Design Code", "B.OdDmCd", "", "OdDmCd", "hOdDmCd", "Design Ctg,Sales Ctg", "", "", "DmDesc "
    GRP_REP.Add "Design Size", "B.OdDmSz", "", "OdDmSz", "hOdDmSz", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= B.OdDmSz)"
    GRP_REP.Add "Suffix", "B.OdSfx", "", "OdSfx", "hOdSfx", "", "", "", ""
    GRP_REP.Add "Karat", "B.OdKt", "", "OdKt", "hOdKt", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = B.OdKt) "
    GRP_REP.Add "(DsgCd)Colour", "B.OdDmCd+B.OdDmCol", "'('+B.OdDmCd+') '+B.OdDmCol", "OdDmCd", "hOdDmCd", "Design Ctg,Design Code", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = B.OdDmCol)"
    GRP_REP.Add "Exp Del Date", "convert(varchar(10), B.OdExpDelDt, 112)", "convert(varchar(8), B.OdExpDelDt, 3)", "", "", "", "", "", ""
    GRP_REP.Add "Prd Del Date", "convert(varchar(10), B.OdDelDt, 112)", "convert(varchar(8), B.OdDelDt, 3)", "OdDelDt", "", "", "", "", ""
    GRP_REP.Add "Customer PO", "BOm.OmPoNo", "", "OmPoNo", "hTBOENo", "Order Cust,Sales Exec", "", "", "convert(varchar(8), BOm.OmPoDt, 3)"
    '*** Jay 3.2.0 [OdPoNo]
    GRP_REP.Add "Cust Sub PO", "BOm.OmPoNo+B.OdPoNo", "BOm.OmPoNo+(case when B.OdPONo<> '' then '/'+ LTrim(B.OdPONo) else '' end )", "OmPoNo", "hTBOENo", "Order Cust,Sales Exec,Order", "", "", "convert(varchar(8), BOm.OmPoDt, 3)"
    '*** Jay 3.2.0 [OdPoNo]
    'VK.27 - Menu code "RepJtSoPndRm" added.
  ElseIf UCase(adc.MenuCd) <> UCase("RepSOPndRM") And UCase(adc.MenuCd) <> UCase("RepJtSOPndRm") Then
    '***** Manali 3.10.0 - 07/03/12 - Sales Ctg and DmCtg allowed below one another
    GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= DmCtg) "
    GRP_REP.Add "Sales Ctg", "DmSalCtg", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and Pmcd=DmCtg and PSCd= DmSalCtg) "
    GRP_REP.Add "Design Code", "B.OdDmCd", "", "", "", "Design Ctg,Sales Ctg", "", "", "DmDesc "
    GRP_REP.Add "Design Size", "B.OdDmSz", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= B.OdDmSz)"
    GRP_REP.Add "Suffix", "B.OdSfx", "", "", "", "", "", "", ""
    GRP_REP.Add "Karat", "B.OdKt", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = B.OdKt) "
    GRP_REP.Add "(DsgCd)Colour", "B.OdDmCd+B.OdDmCol", "'('+B.OdDmCd+') '+B.OdDmCol", "", "", "Design Ctg,Design Code", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = B.OdDmCol)"
    GRP_REP.Add "Exp Del Date", "convert(varchar(10), B.OdExpDelDt, 112)", "convert(varchar(8), B.OdExpDelDt, 3)", "", "", "", "", "", ""
    GRP_REP.Add "Customer PO", "BOm.OmPoNo", "", "OmPoNo", "HOmPoNo", "Order Cust,Sales Exec", "", "", "convert(varchar(8), BOm.OmPoDt, 3)"
    'GRP_REP.Add "Cust Sub PO", "BOm.OmPoNo+B.OdPoNo", "BOm.OmPoNo+(case when B.OdPONo<> '' then '/'+ LTrim(B.OdPONo) else '' end )", "OmPoNo", "HOmPoNo", "Order Cust,Sales Exec,Order", "", "", "convert(varchar(8), BOm.OmPoDt, 3)"
  
    GRP_REP.Add "Rm Ctg", "OrRmCtg", "", "OrRmCtg", "hRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = OrRmCtg) "
    GRP_REP.Add "Rm SubCtg", "OrRmCtg+OrRmSCtg", "'('+ RTrim(OrRmCtg)+ ') '+ OrRmSCtg", "OrRmCtg,OrRmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = OrRmCtg and PSCd= OrRmSCtg) "
    GRP_REP.Add "Rm Code", "OrRmCtg+OrRmSCtg+OrRmCd", "'('+ RTrim(OrRmCtg)+'/'+ OrRmSCtg+ ') '+ OrRmCd", "OrRmCtg,OrRmSCtg,OrRmCd", "hRmCtg,hRmSCtg,hRmCd", "Rm Ctg,Rm SubCtg", "", "", "SubString(RmDesc, 1, 30) "
    GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "OrRmCtg,OrRmSCtg", "hOrRmCtg,hOrRmSCtg", "Rm Ctg,Rm SubCtg", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
  End If
  
  If UCase(adc.MenuCd) = UCase("RepSoPnd") Then
    GRP_REP.Add "Priority", "B.OdPrtCd", "", "OdPrtCd", "hOdPrtCd", "", "", "", "(Select PDesc from Param where PTyp= 'PRTCD' and PMCd= B.OdPrtCd) "
    GRP_REP.Add "Sales Exec", "BOm.OmCmSalPer", "", "", "", "", "", "", "(Select PDesc from Param Where PTyp = 'SP' and PMCd= BOm.OmCmSalPer) "
    GRP_REP.Add "Prd Line", "B.OdLine", "", "OdLine", "hOdLine", "", "", "", "(Select vPDesc from vParam where vPCoCd = B.OdCoCd and vPTyp = 'LINE' and vPMCd = B.OdLine)"
                                                                            
  ' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
    moGrph.AddItem "Ord Quantity", "OdOrdQty", "{rdo.OdOrdQty}", ""
    moGrph.AddItem "Exp Quantity", "OdExpQty", "{rdo.OdExpQty}", ""
    
    ' Zubin 213 (Bef 213)
    'moGrph.AddItem "Ord Sales Price", "OdOrdQty * OdSalPrc", "{rdo.OdOrdQty} * {rdo.OdSalPrc}", ""
    moGrph.AddItem "Ord Sales Price", "OdOrdQty * (case when qBaseCur= 'Y' then (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) else OdSalPrc end)", "{rdo.OdOrdQty} * {rdo.OdSalPrc}", ""
    
    moGrph.AddItem "Inv Dsg Quantity", "IdQty", "{rdo.IdQty}", ""
    ' Zubin 213 (Bef 213)
    'moGrph.AddItem "Invoice FOB", "idiFOB", "{rdo.idiFOB}", ""
    moGrph.AddItem "Invoice FOB", "(case when qBaseCur= 'Y' then (Case when InMulDiv= 'M' then idiFOB/InCnvFct else idiFOB*InCnvFct end) else idiFOB end)", "{rdo.idiFOB}", ""
    
    moGrph.AddItem "Bal Quantity", "OdOrdQty - OdExpQty", "{rdo.OdOrdQty}- {rdo.OdExpQty}", ""
    ' Zubin 213 (Bef 213)
    'moGrph.AddItem "Bal Value", "(OdOrdQty - OdExpQty) * OdSalPrc", "({rdo.OdOrdQty}- {rdo.OdExpQty}) * {rdo.OdSalPrc}", ""
    moGrph.AddItem "Bal Value", "(OdOrdQty - OdExpQty) * (case when qBaseCur= 'Y' then (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) else OdSalPrc end)", "({rdo.OdOrdQty}- {rdo.OdExpQty}) * {rdo.OdSalPrc}", ""
  ' ###########################################  Manoj  ###########################################
  End If
  
  'VK.27 - Menu code "RepJtSoPndRm" added.
  If UCase(adc.MenuCd) = UCase("RepSOPndRM") Or UCase(adc.MenuCd) = UCase("RepJtSOPndRm") Then
      
      GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "DmCtg", "hDmCtg", "Sales Ctg", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= max(DmCtg)) "
      GRP_REP.Add "Sales Ctg", "max(DmSalCtg)", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and PmCd=max(DmCtg) and PSCd= max(DmSalCtg)) "
      GRP_REP.Add "Order No", "OdCoCd+OdTc+OdYy+OdChr+str(OdNo)", "OdCoCd+'/'+OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))", "", "", "Company Code,Cust,Ord Date,Po No,Customer Grp,Customer Typ", "", "", "max(OmPoNo) "
      GRP_REP.Add "Po No", "OdCoCd+max(OmPoNo)", "'['+OdCoCd+'] '+max(OmPoNo)", "InExpNo", "hInExpNo", "Company Code,Inv Date,Inv Vch No,Invoice Cust,Customer Grp,Customer Typ", "", "", ""
      GRP_REP.Add "Order Date", "convert(varchar(10), max(OmDt), 112)", "convert(varchar(8), max(OmDt), 3)", "max(OmDt)", "hOmDt", "Quarter,Month,Week", "", "", ""
      GRP_REP.Add "Design Code", "max(B.OdDmCd)", "", "IdDmCd", "hIdDmCd", "Design Ctg,Sales Ctg", "", "", "max(DmDesc) "
      GRP_REP.Add "Design Size", "max(OdDmSz)", "", "IdDmSz", "hIdDmSz", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= max(OdDmSz))"
      GRP_REP.Add "Suffix", "max(OdSfx)", "", "IdSfx", "hIdSfx", "", "", "", ""
      GRP_REP.Add "Order Cust", "max(OmCmCd)", "", "", "", "Customer Grp,Customer Typ", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= max(OmCmCd)) "
      GRP_REP.Add "Customer Grp", "max(C.CmGrp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMGRP' and PmCd = Max(C.CmGrp))"
      GRP_REP.Add "Customer Typ", "max(C.CmTyp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMTYP' and PmCd = Max(C.CmTyp))"
      GRP_REP.Add "Cust PO", "max(OmPoNo)", "max(OmPoNo)", "", "", "", "", "", "max(OmPoNo) "
      GRP_REP.Add "Cust Sub PO", "max(OdPoNo)", "max(OdPoNo)", "", "", "", "", "", "max(OdPoNo) "
      GRP_REP.Add "Karat", "max(OdKt)", "", "IdKt", "hIdKt", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = max(OdKt)) "
      GRP_REP.Add "PCS/PRS", "max(DmUom)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMUOM' and PmCd= max(DmUom)) "
      GRP_REP.Add "VaCtg", "max(OdVaCtg)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'VACTG' and PMCd = max(OdVaCtg)) "
      GRP_REP.Add "(DsgCd)Colour", "max(OdDmCd+OdDmCol)", "'('+max(OdDmCd)+')'+max(OdDmCol)", "IdDmCd", "hIdDmCd", "Design Ctg,Sales Ctg,Design Code,Design Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(OdDmCol)) "
      GRP_REP.Add "Design Colour", "max(OdDmCol)", "", "", "", "(DsgCd)Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(OdDmCol)) "
      GRP_REP.Add "Sales Exec", "max(BOm.OmCmSalPer)", "", "", "", "", "", "", "(Select PDesc from Param Where PTyp = 'SP' and PMCd= max(BOm.OmCmSalPer)) "
      GRP_REP.Add "Prd Line", "max(OdLine)", "", "", "", "", "", "", "(Select vPDesc from vParam where vPCoCd = B.OdCoCd and vPTyp = 'LINE' and vPMCd = max(OdLine))"
      
    If ws_MultiCoMod = True Then _
      GRP_REP.Add "Company Code", "OdCoCd", "OdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=OdCoCd) "
      moGrph.AddItem "FOB", "OdSalPrc", "{rdo.IdIFOB}", ""
      moGrph.AddItem "Design Quantity", "OdOrdQty", "{rdo.IdQty}", ""
      moGrph.AddItem "Gross Weight", "QGrWt", "{rdo.QGrWt}", ""
      moGrph.AddItem "Labour Value", "QLabVal", "{rdo.QLabVal}", ""
      moGrph.AddItem "Dia CS Weight", "QDCWt", "{rdo.QDCWt}", ""
      moGrph.AddItem "(Lab + Set) Val", "QLabVal + QSetVal", "{rdo.QLabVal} + {rdo.QSetVal}", ""
  End If
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  ' ***** Manali Trading Module
  If UCase(adc.MenuCd) = UCase("RepSoPnd") Then adc("UoAsOfDt") = "01/01/80"    '*** Currently we are not giving the user the option of 'As Of Date'. Hence UoAsOfDt is set to "01/01/80"
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  
  If UCase(adc.MenuCd) = UCase("RepSoPnd") Then _
    moGrph.NonKeyValid pr_Cancel, pr_ErrMsg
    
' ###########################################  Manoj  ###########################################
'VK.27 - Menu code "RepJtSoPndRm" added.
If UCase(adc.MenuCd) = UCase("RepSOPndRM") Or UCase(adc.MenuCd) = UCase("RepJtSOPndRm") Then adc("UoAsOfDt") = "01/01/80"
If UCase(adc.MenuCd) = UCase("RepSOPndRM") Or UCase(adc.MenuCd) = UCase("RepJtSOPndRm") Then
moGrph.NonKeyValid pr_Cancel, pr_ErrMsg
End If
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
'*** Jay 3.2.0 [PicNm]
' ***** Manali 3.5.0 - 11/11/08 - [OdSalRem, qDetPos, qDispSalRemYN] added
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(45) as DGrp1, space(45) as DGrp2, space(45) as DGrp3,
'  space(45) as G1Desc, space(45) as G2Desc, space(45) as G3Desc,
'  OdTc, OdYy, OdChr, OdNo, OdSr, OmDt, OmCmCd, OmPoNo, DmCtg, OdDmCd,
'  OdSfx, OdDmSz, OdPicNm, OdOrdQty, OdSalPrc, OdKt, OdPrtCd , DmSalCtg,
'  OdExpDelDt, OdExpQty, OdDelDt as qDelDt,IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, InDt,
'  InExpNo , InCmCd, IdQty, idiFOB, OdSalRem
'  From OrdDsg, OrdMst, DsgMst, InvDsg, InvHd where 1= 2
  '*** Report Sql ***
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String
  Dim wMtchStr As String, wExpQty As String, wAllPndCnd As String, wRecSelFormula As String
  Dim wOrdDt As String
  Dim wAddlCnd As String  '****** Sachin 3.02.0
  
  ' ***** Manali Trading Module
  Dim wBaseRefIdKey As String
  
   
  Dim ws_ExpJoin As String, wBalQty As String, wOrdQty As String, wOrdAmdJoin As String, wOdQtyCnd As String    ' Manali 3.7.1
  Dim wTcDesc As String, wBalQtyRm As String
  
  Dim wOpt As String, wIdFOBOpt As String, wIrRmWtOpt As String, wIrSetValOpt As String
  Dim wIrRmValOpt As String, wiLValOpt As String, wStr1 As String, wStr2 As String
  
  
  
  ' Zubin 213
  Dim wb_InBaseCur As Boolean
  Dim ws_Join As String
  wb_InBaseCur = IIF(adc("UoCurCdFr") = "", True, False)
  ws_Join = " join (Select PMCd as qPMCd, (case when " + IIF(wb_InBaseCur, "1", "0") + " = 1 then 'Y' " + _
            "else 'N' end) as qBaseCur from Param where PTyp= 'CURNCY') Cur on qPMCd= BOm.OmCmCurCd "
  
  Set Rep = Nothing
  ' **** Manali 3.7.0
  Set Rep1 = Nothing
  
  Set Rep2 = Nothing

    'VK.27 - Menu code "RepJtSoPndRm" added.
    If UCase(adc.MenuCd) <> UCase("RepJtPndRm") And UCase(adc.MenuCd) <> UCase("RepSOPndRM") And UCase(adc.MenuCd) <> UCase("RepJtSOPndRm") Then
   Set adc.RepSource = Rep
   Set mRep = Rep
   
   Else
   Set adc.RepSource = Rep1
   Set mRep = Rep1
   End If
   
  'VK.27 - Menu code "RepJtSoPndRm" added.
  If UCase(adc.MenuCd) = UCase("RepSOPndRM") Or UCase(adc.MenuCd) = UCase("RepJtSOPndRm") Then
  Set adc.RepSource = Rep2
  Set mRep = Rep2
  End If
  
  ' **** Manali 3.7.0
  'Set adc.RepSource = rep
  Call DispCoNm
  Call SetMwName(mRep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  ' **** Crm 2.12 - 23-02-06 **** '
  Dim wCrmCnd As String
  wCrmCnd = " And BOm.OmSubmitYN <> 'N' "
  ' **** Crm 2.12 - 23-02-06 **** '
    
  '*** To check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then wDetPos = i: Exit For
  Next i
  
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then
    Rep.BoxHd.Suppress = True: Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True: Rep.DetSec2.Suppress = True
    Rep.PICSEC.Suppress = True
  End If
 
  ' ***** Manali Trading Module
  If UCase(adc.MenuCd) = UCase("RepJtPnd") Or UCase(adc.MenuCd) = UCase("RepJtPndRm") Then
    wBaseRefIdKey = "OdBaseIdKey"
  ElseIf UCase(adc.MenuCd) = UCase("RepJtOpnOrd") Then
    wBaseRefIdKey = "OdRefIdKey"
  End If
  ' ***** Manali Trading Module
 
  wExpQty = "": wAllPndCnd = "": wRecSelFormula = ""
    
  If adc("UoAsOfDt") = "01/01/80" And UCase(adc.MenuCd) <> UCase("REPJTOPNORD") Then   'And UCase(adc.MenuCd) <> UCase("RepJtPnd") And UCase(adc.MenuCd) <> UCase("RepJtPndRm")
    wExpQty = " B.OdExpQty as OdExpQty "
    ' **** Manali 3.7.1
    If UCase(adc.MenuCd) = UCase("RepJtPndRm") Then wBalQty = " (OdOrdQty- OdExpQty) "
    ' **** Manali 3.7.1
    If adc("UoYN") = "Y" Or UCase(adc.MenuCd) = UCase("RepJtPndRm") Then wAllPndCnd = " and (B.OdCls='N') and Round(B.OdOrdQty- B.OdExpQty, 1)> 0 "  'and Round(OdOrdQty- OdExpQty, 1)> 0

    adc("UoBalPrdFr").CmpStr = "B.OdOrdQty- B.OdExpQty >= "
    adc("UoBalPrdTo").CmpStr = "B.OdOrdQty- B.OdExpQty <= "
    ' Zubin 213 (Curr cnd added)
    adc("UoCstRtFr").CmpStr = "(B.OdOrdQty- B.OdExpQty)* " + IIF(wb_InBaseCur, " (Case when BOm.OmMulDiv= 'M' then (B.OdSalPrc/BOm.OmCnvFct) else (B.OdSalPrc*BOm.OmCnvFct) end) ", "B.OdSalPrc") + " >= "
    adc("UoCstRtTo").CmpStr = "(B.OdOrdQty- B.OdExpQty)* " + IIF(wb_InBaseCur, " (Case when BOm.OmMulDiv= 'M' then (B.OdSalPrc/BOm.OmCnvFct) else (B.OdSalPrc*BOm.OmCnvFct) end) ", "B.OdSalPrc") + " <= "
    ' Manali 3.7.1
    ws_ExpJoin = ""
    wOrdQty = "B.OdOrdQty"
    wOrdAmdJoin = ""
    'VK.27 - Menu code "RepJtSoPndRm" added.
    If UCase(adc.MenuCd) <> UCase("RepSoPnd") And UCase(adc.MenuCd) <> UCase("RepSOPndRM") And UCase(adc.MenuCd) <> UCase("RepJtSOPndRm") Then
      wOdQtyCnd = " and B.OdOrdQty >0 "
    End If
    ' Manali 3.7.1
    wrepcnd = adc.RepCond
    wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  Else
    If UCase(adc.MenuCd) = UCase("RepSoPnd") Then
      wExpQty = " IsNull((Select sum(IdQty) from InvDsg Id join InvHd Ih on " + _
                IIF(gs_Partition = ctCurrPrtn, " Ih.InPrtKey=Id.IdPrtKey and ", "") + " Ih.InIdNo=Id.IdInIdNo and Ih.InCoCd= Id.IdCoCd and Ih.InTc= Id.IdTc and Ih.InYy= Id.IdYy " + _
                " and Ih.InChr= Id.IdChr and Ih.InNo= Id.IdNo " + _
                " where Id.IdOdIdNo=OdIdNo and  Id.IdCoCd= OdCoCd and Id.IdExpOdTc= OdTc and Id.IdExpOdYy= OdYy " + _
                " and Id.IdExpOdChr= OdChr and Id.IdExpOdNo= OdNo and Id.IdExpOdSr= OdSr  " + IIF(gs_Partition = ctCurrPrtn, " and Id.IdPrtKey=OdPrtKey ", "") + _
               IIF(adc("UoAsOfDt") = "01/01/80", "", " and  Ih.InDt<= '" + CStr(adc("UoAsOfDt")) + "'") + "), 0) as OdExpQty "
      If adc("UoYN") = "Y" Then _
        wAllPndCnd = " and Round(OdOrdQty- IsNull((Select sum(IdQty) from InvDsg Id join InvHd Ih on " + _
                     IIF(gs_Partition = ctCurrPrtn, " Ih.InPrtKey=Id.IdPrtKey and ", "") + " Ih.InIdNo=Id.IdInIdNo and Ih.InCoCd= Id.IdCoCd and Ih.InTc= Id.IdTc and Ih.InYy= Id.IdYy " + _
                     " and Ih.InChr= Id.IdChr and Ih.InNo= Id.IdNo " + IIF(gs_Partition = ctCurrPrtn, " and Ih.InPrtKey=Id.IdPrtKey ", "") + _
                     " where " + IIF(gs_Partition = ctCurrPrtn, " Id.IdPrtKey=OdPrtKey and ", "") + " Id.IdOdIdNo=OdIdNo and Id.IdCoCd= OdCoCd and Id.IdExpOdTc= OdTc and Id.IdExpOdYy= OdYy " + _
                     " and Id.IdExpOdChr= OdChr and Id.IdExpOdNo= OdNo and Id.IdExpOdSr= OdSr " + _
                     IIF(adc("UoAsOfDt") = "01/01/80", "", " and  Ih.InDt<= '" + CStr(adc("UoAsOfDt")) + "'") + "), 0), 3)> 0 "
       
       ' ***** Manali 3.7.1 - Pending Rm Report added
       wOrdQty = "B.OdOrdQty"
       wOrdAmdJoin = ""
       wOdQtyCnd = ""
       ws_ExpJoin = ""
       wOrdAmdJoin = ""
       ' ***** Manali 3.7.1 - Pending Rm Report added
    Else
        'If UCase(adc.MenuCd) = UCase("RepJtPnd") Or UCase(adc.MenuCd) = UCase("RepJtOpnOrd") Then
      ws_ExpJoin = " left Outer Join (Select OdPrtKey as qOdPrtKey, OdCoCd as qOdCoCd, " + wBaseRefIdKey + " as qBaseRefIdKey, " + _
                   " Sum(OdOrdEnt+IsNull(qAddQty, 0)) as qExpQty " + _
                   " From OrdDsg " + _
                   " Left Outer Join (Select OaOdIdNo, OaPrtKey, OaCoCd, OaTc, OaYy, OaChr, OaNo, OaSr, Sum(Convert(Decimal(14,1),OaOrdQty)) as qAddQty " + _
                   "                  From OrdAmd  Where OaDt<='" + CStr(adc("UoAsOfDt")) + "' " + _
                   "                  Group by OaPrtKey, OaOdIdNo, OaCoCd, OaTc, OaYy, OaChr, OaNo, OaSr)OrdAmd " + _
                   "                  On OaCoCd=OdCoCd and OaTc=OdTc and OaYy=OdYy and OaChr=OdChr and OaNo=OdNo and OaSr=OdSr and OaOdIdNo=OdIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OaPrtKey=OdPrtKey ", "") + _
                 IIF(adc("UoAsOfDt") = "01/01/80", "", " Where OdOmDt<= '" + CStr(adc("UoAsOfDt")) + "' ") + _
                   "        Group By OdPrtKey, OdCoCd, " + wBaseRefIdKey + " ) OdExp " + _
                   "On (B.OdTc='JOT' Or qOdCoCd = B.OdCoCd) And qBaseRefIdKey = B.OdIdNo " + IIF(gs_Partition = ctCurrPrtn, " and qOdPrtKey=B.OdPrtKey ", "")
        
'       ws_ExpJoin = " left Outer Join (Select OdPrtKey, OdCoCd , " + wBaseRefIdKey + " as qBaseRefIdKey, Sum(OdOrdQty) as qExpQty " + _
'                   "       From OrdDsg Where OdOmDt<= '" + CStr(adc("UoAsOfDt")) + "' " + _
'                   "        Group By OdPrtKey, OdCoCd, " + wBaseRefIdKey + " )M " + _
'                   "On M.OdCoCd = B.OdCoCd And qBaseRefIdKey = B.OdIdNo " + IIF(gs_Partition = ctCurrPrtn, " and M.OdPrtKey=B.OdPrtKey ", "")
    

       wExpQty = " IsNull(qExpQty, 0) as OdExpQty "
       wOrdQty = "(B.OdOrdEnt + IsNull(qAddQty, 0))"
       wOrdAmdJoin = " Left Outer Join (Select OaOdIdNo, OaPrtKey, OaCoCd, OaTc, OaYy, OaChr, OaNo, OaSr, Sum(Convert(Decimal(14,1),OaOrdQty)) as qAddQty " + _
                      "From OrdAmd  " + IIF(adc("UoAsOfDt") = "01/01/80", "", " Where OaDt<='" + CStr(adc("UoAsOfDt")) + "' ") + _
                      "Group by OaPrtKey, OaOdIdNo, OaCoCd, OaTc, OaYy, OaChr, OaNo, OaSr)OrdAmd " + _
                      "On OaCoCd=B.OdCoCd and OaTc=B.OdTc and OaYy=B.OdYy and OaChr=B.OdChr and OaNo=B.OdNo and OaSr=B.OdSr and OaOdIdNo=B.OdIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OaPrtKey=B.OdPrtKey ", "")
       wBalQty = " (" + wOrdQty + " - IsNull(qExpQty, 0)) "
        wOdQtyCnd = " and " + wOrdQty + ">0 "
      If adc("UoYN") = "Y" Or UCase(adc.MenuCd) = UCase("RepJtPndRm") Then
        wAllPndCnd = " and Round(" + wOrdQty + "- IsNull(qExpQty, 0), 3)> 0 "
      Else
      End If
   
    End If
    adc("UoBalPrdFr").CmpStr = "": adc("UoBalPrdTo").CmpStr = ""
    adc("UoCstRtFr").CmpStr = "": adc("UoCstRtTo").CmpStr = ""
    wrepcnd = adc.RepCond
    wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    If adc("UoBalPrdFr") > 0 Then wRecSelFormula = wRecSelFormula + "{@wBalQty}>= " + CStr(adc("UoBalPrdFr")) + " "
    If adc("UoBalPrdTo") > 0 Then wRecSelFormula = wRecSelFormula + IIF(wRecSelFormula = "", "", " and ") + "{@wBalQty}<= " + CStr(adc("UoBalPrdTo")) + " "
    If adc("UoCstRtFr") > 0 Then wRecSelFormula = wRecSelFormula + IIF(wRecSelFormula = "", "", " and ") + "{@wBalVal}>= " + CStr(adc("UoCstRtFr")) + " "
    If adc("UoCstRtTo") > 0 Then wRecSelFormula = wRecSelFormula + IIF(wRecSelFormula = "", "", " and ") + "{@wBalVal}<= " + CStr(adc("UoCstRtTo")) + " "
    ' **** Manali 3.7.1 - mRep used Instead of Rep
    If wRecSelFormula <> "" Then mRep.RecordSelectionFormula = wRecSelFormula
    ' bef 3.7.1
    'If wRecSelFormula <> "" Then Rep.RecordSelectionFormula = wRecSelFormula
  End If
  
  ' Zubin 213
  wCnd = wCnd + IIF(wb_InBaseCur = False, " and CmCurCd= '" + adc("UoCurCdFr") + "'", "")
  
  If adc("UoYN1") = "Y" Then
    ' Zubin 213 (Cur Cnd and ws_Join added)
    '****** Sachin 3.02 - Id fields in Joins
    ' ***** Manali 3.7.1 - Pending Rm Report added - ws_ExpJoin, wOrdAmdJoin, wOdQtyCnd  added For Trading Reports
    If UCase(adc.MenuCd) = UCase("RepSoPnd") Then
      wMtchStr = " IsNull(IdCoCd, '') as IdCoCd, IsNull(IdTc, '') as IdTc, IsNull(IdYy, '') as IdYy, " + _
                 "IsNull(IdChr, '') as IdChr, IsNull(IdNo, 0) as IdNo, IsNull(IdSr, 0) as IdSr, " + _
                 "IsNull(InDt, '') as InDt, IsNull(InExpNo, '') as InExpNo, " + _
                 "IsNull(InCmCd, '') as InCmCd, IsNull(IdQty, 0) as IdQty, " + _
                 "IsNull(" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then (idiFOB/InCnvFct) else (idiFOB*InCnvFct) end) ", "idiFOB") + ", 0) As idiFOB " + _
                 "From OrdDsg B Join OrdMst BOm on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and " + _
                 "OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + " Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
                 "Join CustMst on CmCtg = 'C' and CmCd= OmCmCd " + _
                 "Left Outer Join InvDsg M on IdOdIdNo=OdIdNo and IdCoCd= OdCoCd and IdExpOdTc= OdTc and IdExpOdYy= OdYy and " + _
                 "IdExpOdChr= OdChr and IdExpOdNo = OdNo And IdExpOdSr = OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=IdPrtKey ", "") + _
                 IIF(adc("UoAsOfDt") = "01/01/80", " ", "and " + _
                 "IsNull((Select InDt from InvHd where InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and " + _
                 "InChr= IdChr and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IdPrtKey ", "") + "), '01/01/80')<= '" + CStr(adc("UoAsOfDt")) + "' ") + _
                 "Left Outer Join InvHd Inv on InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IdPrtKey ", "") + _
                 ws_Join + "where (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc)= 'SO' " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "")
    Else
      wMtchStr = " IsNull(M.OdCoCd, '') as IdCoCd, IsNull(M.OdTc, '') as IdTc, IsNull(M.OdYy, '') as IdYy, " + _
                 "IsNull(M.OdChr, '') as IdChr, IsNull(M.OdNo, 0) as IdNo, IsNull(M.OdSr, 0) as IdSr, " + _
                 "IsNull(MOm.OmDt, '') as InDt, '' as InExpNo, " + _
                 "IsNull(MOm.OmCmCd, '') as InCmCd, IsNull(M.OdOrdQty, 0) as IdQty, " + _
                 "IsNull(" + IIF(wb_InBaseCur, " (Case when BOm.OmMulDiv= 'M' then (B.OdSalPrc/BOm.OmCnvFct) else (B.OdSalPrc*BOm.OmCnvFct) end) ", "B.OdSalPrc") + ", 0) As idiFOB " + _
                 "From OrdDsg B Join OrdMst BOm on BOm.OmIdNo=B.OdOmIdNo and BOm.OmCoCd= B.OdCoCd and BOm.OmTc= B.OdTc and BOm.OmYy= B.OdYy and BOm.OmChr= B.OdChr and " + _
                 "BOm.OmNo= B.OdNo " + IIF(gs_Partition = ctCurrPrtn, " and BOm.OmPrtKey=B.OdPrtKey ", "") + " Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= B.OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=B.OdPrtKey ", "") + _
                 "Join CustMst on CmCtg = '" + gs_CmCtg + "' and CmCd= BOm.OmCmCd " + ws_Join + _
                 "Left Outer Join OrdDsg M on M." + wBaseRefIdKey + " =B.OdIdNo and (M.OdTc='JIN' Or M.OdCoCd = B.OdCoCd) " + IIF(gs_Partition = ctCurrPrtn, " and M.OdPrtKey=B.OdPrtKey ", "") + _
                 IIF(adc("UoAsOfDt") = "01/01/80", " ", " and M.OdOmDt <= '" + CStr(adc("UoAsOfDt")) + "' ") + _
                 "Left Outer Join OrdMst MOm on MOm.OmIdNo=M.OdOmIdNo and MOm.OmCoCd= M.OdCoCd and MOm.OmTc= M.OdTc and MOm.OmYy= M.OdYy and MOm.OmChr= M.OdChr and MOm.OmNo= M.OdNo " + IIF(gs_Partition = ctCurrPrtn, " and MOm.OmPrtKey=M.OdPrtKey ", "") + _
                 ws_ExpJoin + wOrdAmdJoin + " where (Select PValue from Param where PTyp= 'TC' and PMCd= B.OdTc)= 'JT' " + IIF(gs_Partition = ctCurrPrtn, " and B.OdPrtKey='" + ctCurrPrtn + "' ", "") + wOdQtyCnd
    End If
  Else
    ' Zubin 213 (ws_Join added)
    '****** Sachin 3.02 - Id fields in Joins
    wMtchStr = " '' as IdCoCd,'' as IdTc, '' as IdYy, '' as IdChr, 0 as IdNo, 0 as IdSr, '' as InDt, " + _
               "'' as InExpNo, '' as InCmCd, 0 as IdQty, 0 As idiFOB " + _
               "From OrdDsg B Join OrdMst BOm on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and " + _
               "OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
               "Join CustMst on CmCtg = '" + gs_CmCtg + "' and CmCd= OmCmCd  " + _
               "Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and " + _
               "DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + ws_Join + ws_ExpJoin + wOrdAmdJoin + " where " + _
               " (Select PValue from Param where PTyp= 'TC' and PMCd= B.OdTc)= '" + gs_OmTcTyp + "' " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + wOdQtyCnd
    
'    bef Trading Module
'    wMtchStr = " '' as IdTc, '' as IdYy, '' as IdChr, 0 as IdNo, 0 as IdSr, '' as InDt, " + _
'               "'' as InExpNo, '' as InCmCd, 0 as IdQty, 0 As idiFOB " + _
'               "From OrdDsg B Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and " + _
'               "OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
'               "Join CustMst on CmCtg = 'C' and CmCd= OmCmCd  " + _
'               "Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and " + _
'               "DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + ws_Join + " where " + " (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc)= 'SO' " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "")
  End If
 'Prev
'  If adc("UoYN3") = "Y" Then
'    wOrdDt = "OdExpDelDt"
'    Rep.hDelDt.SetText "Exp Del Dt"
'  ElseIf adc("UoYN3") = "N" Then
'    wOrdDt = "OdDelDt"
'    Rep.hDelDt.SetText "Prd Del Dt"
'  End If

'  OdTc, OdYy, OdChr, OdNo, OdSr, OmDt, OmCmCd, OmPoNo, DmCtg, OdDmCd,
'  OdSfx, OdDmSz, OdOrdQty, OdSalPrc, OdKt, OdPrtCd , DmSalCtg,
'  OdExpDelDt, OdExpQty, qDelDt, IdTc, IdYy, IdChr, IdNo, IdSr, InDt,
'  InExpNo, InCmCd, IdQty, idiFOB

  
  '********geeta********Emr206***********Replaced + wOrdDt + " as qDelDt with OdDelDt
  ' Zubin 213
  '*** Jay 3.2.0 [PicNm]
  ' ***** Manali 3.5.0 - 11/11/08 - [OdSalRem, qDetPos, qDispSalRemYN] added
  ' ***** Manali 3.7.1 - Pending Rm Report added - wOrdQty  Used For Order Qty
  'VK.27 - Menu code "RepJtSoPndRm" added.
  If UCase(adc.MenuCd) <> UCase("RepJtPndRm") And UCase(adc.MenuCd) <> UCase("RepSOPndRM") And UCase(adc.MenuCd) <> UCase("RepJtSOPndRm") Then
    wSqlStrg = "Set Dateformat DMY Select " + grpflds + ", " + _
             "B.OdCoCd,B.OdTc as OdTc, B.OdYy as OdYy, B.OdChr as OdChr, B.OdNo as OdNo, B.OdSr as OdSr, BOm.OmDt as OmDt, " + _
             "BOm.OmCmCd as OmCmCd, BOm.OmPoNo as OmPoNo, DmCtg, B.OdDmCd as OdDmCd, B.OdSfx as OdSfx, " + _
             "B.OdDmSz as OdDmSz, B.OdDmStkCoCd+'/'+B.OdDmStkYy+'/'+B.OdDmStkChr+'/'+ltrim(str(B.OdDmStkNo)) as OdDmStkNum, B.OdDmStkCoCd, B.OdDmStkYy, B.OdDmStkChr, B.OdDmStkNo, B.OdPicNm as OdPicNm, /*B.OdOrdQty*/ " + wOrdQty + " as OdOrdQty, " + IIF(wb_InBaseCur, " (Case when BOm.OmMulDiv= 'M' then (B.OdSalPrc/BOm.OmCnvFct) else (B.OdSalPrc*BOm.OmCnvFct) end) ", "B.OdSalPrc") + " as OdSalPrc, " + _
             "B.OdKt as OdKt, B.OdPrtCd OdPrtCd, B.OdLine as OdLine, DmSalCtg, B.OdExpDelDt as OdExpDelDt, " + _
             wExpQty + ", B.OdDelDt as qDelDt, Replace(B.OdSalRem,char(13)+ char(10),' ') as OdSalRem, '" + CStr(wDetPos) + "' as qDetPos, '" + adc("UoYN4") + "' as qDispSalRemYN, " + wMtchStr + wAllPndCnd + wCnd + _
             IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + wCrmCnd
  ElseIf UCase(adc.MenuCd) <> UCase("RepSOPndRM") And UCase(adc.MenuCd) <> UCase("RepJtSOPndRm") Then
    wSqlStrg = "Select " + grpflds + ", " + _
            "B.OdCoCd as OdCoCd, B.OdTc as OdTc, B.OdYy as OdYy, B.OdChr as OdChr, B.OdNo as OdNo, B.OdSr as OdSr, BOm.OmDt as OmDt, " + _
            "BOm.OmCmCd as OmCmCd, BOm.OmPoNo as OmPoNo, BOm.OmPoDt as OmPoDt, B.OdPicNm as OdPicNm, " + _
            " B.OdDmCd as OdDmCd, B.OdSfx as OdSfx, B.OdDmSz as OdDmSz, B.OdDmStkCoCd+'/'+B.OdDmStkYy+'/'+B.OdDmStkChr+'/'+ltrim(str(B.OdDmStkNo)) as OdDmStkNum, B.OdPoNo as OdPoNo, DmCtg, OdOrdQty, OdSalPrc, " + wBalQty + " as qBalQty, " + _
            " OrSrNo, OrRmCtg, OrRmSCtg, OrRmCd, OrLn1, OrLn2, OrRmPtr, " + _
            "Convert(int, " + wBalQty + " * OrQty) as OrQty, Convert(Decimal(14,3), " + wBalQty + " * OrWt) as OrWt, " + _
            "Convert(Decimal(14,3), " + wBalQty + " * OrWt/(case when OrRmCtg in ('D', 'C') then 5 else 1 end)) as qWtInGms,   " + _
            "Convert(Decimal(14,3), " + wBalQty + " * OrWt/(case when OrRmCtg in ('D', 'C') then 5 else 1 end) " + _
            "*(case when OrRmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityRt Else 0 end) )" + _
            "/(case when OrRmCtg in ('G', 'P', 'S', 'L', 'M') then " + _
                          "IsNull((Select RmPurityRt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= OrRmCtg), 1) " + _
            "else 1 end) as qPureWt, " + _
            "Convert(Decimal(14,3), OrSalRt) OrSalRt, Convert(Decimal(14,3), " + wBalQty + " * OrSalVal) OrSalVal " + _
            "From OrdRm " + _
            "Join OrdDsg B On B.OdCoCd=OrCoCd and B.OdTc=OrTc and B.OdYY=OrYy and B.OdChr=OrChr and B.OdNo=OrNo and OrSr=B.OdSr and OrOdIdNo=B.OdIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=B.OdPrtKey ", "") + _
            "Join OrdMst BOm on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and " + _
            "OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
            "Join RmMst On RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + "  Join CustMst on CmCtg = '" + gs_CmCtg + "' and CmCd= OmCmCd  " + _
            "Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and " + _
            "DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + ws_Join + ws_ExpJoin + wOrdAmdJoin + " where " + _
            " (Select PValue from Param where PTyp= 'TC' and PMCd= B.OdTc)= '" + gs_OmTcTyp + "' " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + _
             wAllPndCnd + wCnd + wOdQtyCnd + " Order By B.OdCoCd, B.OdTc, B.OdYy, B.OdChr, B.OdNo, B.OdSr "

  End If
' **** Bef Trading Module
'  wSqlStrg = "Set Dateformat DMY Select " + grpflds + ", " + _
'             "OdTc, OdYy, OdChr, OdNo, OdSr, OmDt, OmCmCd, OmPoNo, DmCtg, OdDmCd, OdSfx, " + _
'             "OdDmSz, OdPicNm, OdOrdQty, " + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then (OdSalPrc/OmCnvFct) else (OdSalPrc*OmCnvFct) end) ", "OdSalPrc") + " as OdSalPrc, " + _
'             "OdKt, OdPrtCd, DmSalCtg, OdExpDelDt, " + _
'             wExpQty + ", OdDelDt as qDelDt, Replace(OdSalRem,char(13)+ char(10),' ') as OdSalRem, '" + CStr(wDetPos) + "' as qDetPos, '" + ADC("UoYN4") + "' as qDispSalRemYN, " + wMtchStr + wAllPndCnd + wCnd + _
'             IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + wCrmCnd
  
  
  'VK.27 - Menu code "RepJtSoPndRm" added.
  If UCase(adc.MenuCd) = UCase("RepSOPndRM") Or UCase(adc.MenuCd) = UCase("RepJtSOPndRm") Then
  wBalQtyRm = " max(OdOrdQty-OdExpQty) "
  
  If UCase(adc("UoInVal")) = "" And OptValWt(0).Value = True Then
  adc("UoInVal") = "C"
  End If
  
     If UCase(adc("UoInVal")) = "C" And OptValWt(0).Value = True Then
        wOpt = "Cost Value"   '''"Actual"
        wIdFOBOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdCstPrc/OmCnvFct else OdCstPrc*OmCnvFct end) ", "OdCstPrc")
       ' wIrRmWtOpt = "OrWt"
        wIrSetValOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSetCstVal/OmCnvFct else OrSetCstVal*OmCnvFct end) ", "OrSetCstVal")
        wIrRmValOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrCstVal/OmCnvFct else OrCstVal*OmCnvFct end) ", "OrCstVal")
        wiLValOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OlCstVal/OmCnvFct else OlCstVal*OmCnvFct end) ", "OlCstVal")
     
     ElseIf UCase(adc("UoInVal")) = "S" And OptValWt(0).Value = True Then
        wOpt = "Sale Value"      '''"Invoice"
        wIdFOBOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc")
       ' wIrRmWtOpt = "OrWt"
        wIrSetValOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSetSalVal/OmCnvFct else OrSetSalVal*OmCnvFct end) ", "OrSetSalVal")
        wIrRmValOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal")
        wiLValOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OlSalVal/OmCnvFct else OlSalVal*OmCnvFct end) ", "OlSalVal")
     
     ElseIf UCase(adc("UoInVal")) = "" And OptValWt(1).Value = True Then
        wOpt = "Weight"   '''"Actual"
        wIdFOBOpt = 0
        wIrSetValOpt = 0
        wIrRmValOpt = 0
        wiLValOpt = 0
     
     
     End If
     
      wStr1 = "Select " + grpflds + ", " + _
          "OdCoCd as IdCoCd , OdTc as IdTc, OdYy as IdYy, OdChr as IdChr, OdNo as IdNo, OdSr as IdSr, max(OmPoNo) as InExpNo, max(DmCtg) as DmCtg, max(OdDmCd) as IdDmCd, " + _
          "max(OdSfx) as IdSfx,max(OdDmSz) as IdDmSz, max(OdKt) as IdKt, /* max(InDt) as InDt,*/ max(OmDt ) as InDt, " + _
          "max(OmCmCd) as InCmCd, max(OdLine), 0 as IdExpOdTc, max(OdExpQty) as OdExpQty, " + _
          "0 as IdExpOdYy, 0 as IdExpOdChr, 0 as IdExpOdNo, " + _
          "0 as IdExpOdSr,  max(DmSalCtg) as DmSalCtg, max(OdExpDelDt) as OdExpDelDt, max(OdDelDt) as qDelDt," + _
          "max(OdVaCtg) as IdVaCtg, max(OdOrdQty) as IdQty, (max(" + wIdFOBOpt + ") * " + wBalQtyRm + " ) as IdIFob, " + _
          "sum(case when IsNull(RmCtg, '') in ('G', 'P', 'S', 'L', 'M', 'A') then IsNull(" + wIrRmValOpt + " , 0) else 0 end)* " + wBalQtyRm + "  as qMetVal, " + _
          "sum(case when IsNull(RmCtg, '') in ('D') then IsNull(" + wIrRmValOpt + ", 0) else 0 end)* " + wBalQtyRm + "  as qDVal, " + _
          "sum(case when IsNull(RmCtg, '') in ('C') then IsNull(" + wIrRmValOpt + ", 0) else 0 end)* " + wBalQtyRm + "  as qCVal, " + _
          "sum(case when IsNull(RmCtg, '') in ('X') then IsNull(" + wIrRmValOpt + ", 0) else 0 end)* " + wBalQtyRm + "  as qXVal, " + _
          "sum(IsNull(" + wIrSetValOpt + ", 0)) * " + wBalQtyRm + " as qSetVal, " + _
          "IsNull((select sum(" + wiLValOpt + ") from OrdLab " + _
          IIF(wb_InBaseCur, " Join OrdMst on " + IIF(gs_Partition = ctCurrPrtn, " OlPrtKey=OdPrtKey and ", "") + " OmCoCd= OlCoCd and OmTc= OlTc and OmYy= OlYy  and OmChr= OlChr and OmNo= OlNo ", "") + _
          "where OLCoCd= OdCoCd and OLTc= OdTc" + _
          " and  OLYy= OdYy and OLChr= OdChr and OLNo= OdNo and OLSr= OdSr), 0)* " + wBalQtyRm + "  as qLabVal, " + _
          "sum(case when IsNull(RmCtg, '') in ('D', 'C') then IsNull(OrWt, 0)/ 5 else IsNull( OrWt, 0) end)* " + wBalQtyRm + "  as qGrWt, " + _
          "sum(case when IsNull(RmCtg, '') in ('G', 'P', 'S', 'L', 'M', 'A') then IsNull(OrWt, 0) else 0 end) * " + wBalQtyRm + "  as qMetWt, " + _
          "sum(case when IsNull(RmCtg, '') in ('D') then IsNull(OrWt, 0) else 0 end) * " + wBalQtyRm + "  as qDWt, " + _
          "sum(case when IsNull(RmCtg, '') in ('C') then IsNull(OrWt , 0) else 0 end)* " + wBalQtyRm + "  as qCWt, " + _
          "sum(case when IsNull(RmCtg, '') in ('X') then IsNull(OrWt, 0) else 0 end)* " + wBalQtyRm + "  as qXWt " + _
          IIF(adc("UoYN2") = "Y" And wDetPos > 0, ", max(OdPicNm) as OdPicNm ", ", '' As OdPicNm ")
  Dim wBSCnd As String
      wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
      wCnd = wCnd + IIF(wb_InBaseCur = False, " and C.CmCurCd= '" + adc("UoCurCdFr") + "'", "")
  
  
  wStr2 = "From OrdDsg B join OrdMst BOm on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy " + _
          " and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
          " Join CustMst C on C.CmCtg= 'C' and C.CmCd =OmCmCd " + _
          " Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and " + _
           " DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
          IIF(adc("UoAsOfDt") = "01/01/80", " ", " and  BOm.OmDt <= '" + CStr(adc("UoAsOfDt")) + "' ") + _
          " left outer join OrdRm on OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy  and " + _
          " OrChr= OdChr and OrNo= OdNo and OrSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OdPrtKey ", "") + " left outer join RmMst on RmCd= OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey  ", "") + _
           wCnd + wAllPndCnd + wBSCnd + "and (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc)= '" + gs_OmTcTyp + "'" + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + " group by OdPrtKey, OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr Order by OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr"
 
   
  wSqlStrg = wStr1 + wStr2
 
     
  End If
  
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt


'VK.27 - Menu code "RepJtSoPndRm" added.
If UCase(adc.MenuCd) = UCase("RepSOPndRM") Or UCase(adc.MenuCd) = UCase("RepJtSOPndRm") Then
    If adc("UoYN5") = "Y" Then
        Rep2.TxtHead.SetText "SO Pending with RM Detail(Value)"
        GRP_REP.SetFormula Rep2, "wRndOffDecPlaces", "2"
        GRP_REP.SetFormula Rep2, "wTotVal", "{rdo.IdIFOB}"
        GRP_REP.SetFormula Rep2, "wMetVal", "{rdo.qMetVal}"
        GRP_REP.SetFormula Rep2, "wDVal", "{rdo.qDVal}"
        GRP_REP.SetFormula Rep2, "wCVal", "{rdo.qCVal}"
        GRP_REP.SetFormula Rep2, "wXVal", "{rdo.qXVal}"
        GRP_REP.SetFormula Rep2, "wSetVal", "{rdo.qSetVal}"
        GRP_REP.SetFormula Rep2, "wLabVal", "{rdo.qLabVal}"
    
        Rep2.TxtValTyp.SetText wOpt

        Rep2.h1Tot.SetText " Val": Rep2.h2Tot.SetText "Val"
        Rep2.h1Met.SetText " Met Val": Rep2.h2Met.SetText "Met Val"
        Rep2.h1D.SetText " Dia Val": Rep2.h2D.SetText "Dia Val"
        Rep2.h1C.SetText " Cs Val": Rep2.h2C.SetText "Cs Val"
        Rep2.h1X.SetText " Acc Val": Rep2.h2X.SetText "Acc Val"
        Rep2.h1SetVal.SetText " Set Val": Rep2.h2SetVal.SetText "Set Val"
        Rep2.h1LabVal.SetText " Lab Val": Rep2.h2LabVal.SetText "Lab Val"

Rep2.Text18.SetText "Ord Sr"
        Rep2.hInExpNo.SetText "PO No"
        Rep2.hInDt.SetText "Ord Dt"


        ' Zubin 213
        If wb_InBaseCur Then
          Rep2.TxtValIn.SetText "Value In " + moCn.GetFldVal("Select HBaseCurCd from Head " + _
                                              "where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UoCoCdFr") + "'") + ""
        Else
          Rep2.TxtValIn.SetText "Value In " + UCase(adc("UoCurCdFr"))
        End If
        
    ElseIf adc("UoYN5") = "N" Then
        Rep2.TxtHead.SetText "SO Pending with RM Detail (Weights)"
        Rep2.TxtValTyp.SetText " Weights"
        GRP_REP.SetFormula Rep2, "wRndOffDecPlaces", "3"
        GRP_REP.SetFormula Rep2, "wTotVal", "{rdo.qGrWt}"
        GRP_REP.SetFormula Rep2, "wMetVal", "{rdo.qMetWt}"
        GRP_REP.SetFormula Rep2, "wDVal", "{rdo.qDWt}"
        GRP_REP.SetFormula Rep2, "wCVal", "{rdo.qCWt}"
        GRP_REP.SetFormula Rep2, "wXVal", "{rdo.qXWt}"
        GRP_REP.SetFormula Rep2, "wSetVal", "0"
        GRP_REP.SetFormula Rep2, "wLabVal", "0"
    
        Rep2.h1Tot.SetText " Grs Wt": Rep2.h2Tot.SetText " Grs Wt"
        Rep2.h1Met.SetText " Met Wt": Rep2.h2Met.SetText " Met Wt"
        Rep2.h1D.SetText " Dia Wt": Rep2.h2D.SetText " Dia Wt"
        Rep2.h1C.SetText " Cs Wt": Rep2.h2C.SetText " Cs Wt"
        Rep2.h1X.SetText " Acc Wt": Rep2.h2X.SetText " Acc Wt"
    
        Rep2.h1SetVal.SetText " ": Rep2.h2SetVal.SetText " "
        Rep2.h1LabVal.SetText " ": Rep2.h2LabVal.SetText " "
        ' Zubin 213
        Rep2.TxtValIn.SetText ""
        Rep2.Text18.SetText " Ord Sr"
        Rep2.hInExpNo.SetText " PO No"
        Rep2.hInDt.SetText " Ord Dt"
        
    End If
  
    ' ****** Manali 3.5.0 - 11/11/08 - Show Picture Option in report
    If adc("UoYN2") = "N" Or wDetPos = 0 Then
        Rep2.PICSEC.Suppress = True
        gb_ShowPic = False
    ElseIf adc("UoYN2") = "Y" And wDetPos <> 0 Then
        Rep2.PICSEC.Suppress = False
        gb_ShowPic = True
    End If
  ' ****** Manali 3.5.0 - 11/11/08 - Show Picture Option in report
  
  ' ***** Manali 3.9.0 - Fr and To Date Scope Report
    Rep2.TxtFrDt.SetText adc("UoDtFr")
    Rep2.TxtToDt.SetText adc("UoDtTo")
  ' ***** Manali 3.9.0 - Fr and To Date Scope Report
End If
  If UCase(adc.MenuCd) <> UCase("RepJtPndRm") Then
   If adc("UoYN1") = "N" Then
     Rep.BoxHd.Suppress = True
     Rep.BoxDet.Suppress = True
     Rep.BoxFt.Suppress = True
     Rep.DetSec2.Suppress = True
   End If
     
   If adc("UoYN2") = "N" Or wDetPos = 0 Then
     Rep.PICSEC.Suppress = True
     '*** urmi 2.04-3 Load Pict only is ShowPic=Y
     gb_ShowPic = False
   ElseIf adc("UoYN2") = "Y" And wDetPos <> 0 Then
     Rep.PICSEC.Suppress = False
     gb_ShowPic = True
   End If
    '*****************Geeta*************Emr206********

    If adc("UoYN3") = "Y" Then
      Rep.hExpDt.Suppress = False
      Rep.qExpDt.Suppress = False
      Rep.hDelDt.Suppress = True
      Rep.qDelDt.Suppress = True
    ElseIf adc("UoYN3") = "N" Then
       Rep.hExpDt.Suppress = True
       Rep.qExpDt.Suppress = True
       Rep.hDelDt.Suppress = False
      Rep.qDelDt.Suppress = False
    ElseIf adc("UOYN3") = "" Then
      Rep.hExpDt.Suppress = False
      Rep.qExpDt.Suppress = False
      Rep.hDelDt.Suppress = False
      Rep.qDelDt.Suppress = False
    End If

     '*****************Geeta*************Emr206********
  'VK.27 - Menu code "RepJtSoPndRm" added.
  If UCase(adc.MenuCd) = UCase("RepSOPndRM") Or UCase(adc.MenuCd) = UCase("RepJtSOPndRm") Then
    If adc("UoYN3") = "Y" Then
      Rep2.hExpDt.Suppress = False
      Rep2.qExpDt.Suppress = False
      Rep2.hDelDt.Suppress = True
      Rep2.qDelDt.Suppress = True
    ElseIf adc("UoYN3") = "N" Then
       Rep2.hExpDt.Suppress = True
       Rep2.qExpDt.Suppress = True
       Rep2.hDelDt.Suppress = False
      Rep2.qDelDt.Suppress = False
    ElseIf adc("UOYN3") = "" Then
      Rep2.hExpDt.Suppress = False
      Rep2.qExpDt.Suppress = False
      Rep2.hDelDt.Suppress = False
      Rep2.qDelDt.Suppress = False
    End If
End If
    ' Zubin 213
    Dim ws_BaseCurCd As String
    ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                    "HCoCd= '" + adc("UoCoCdFr") + "'")
    If wb_InBaseCur Then
      Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
    Else
      Rep.TxtValIn.SetText "Value In " + adc("UoCurCdFr")
    End If
    ' Zubin 213
  
  End If
    
  CRV_REP.DisplayGroupTree = False
    
  If UCase(adc.MenuCd) = UCase("RepSoPnd") Then
    ' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
    Dim wsQryStruct As String
  
    wsQryStruct = "  Select space(40) as Grp1, space(40) as Grp2, space(40) as Grp3, " + _
      "space(40) as DGrp1, space(40) as DGrp2, space(40) as DGrp3, " + _
      "space(50) as G1Desc, space(50) as G2Desc, space(50) as G3Desc, " + _
      "OdTc, OdYy, OdChr, OdNo, OdSr, OmDt, OmCmCd, OmPoNo, DmCtg, OdDmCd, " + _
      "OdSfx, OdDmSz, OdOrdQty, OdSalPrc, OdKt, OdPrtCd , OdLine, DmSalCtg, " + _
      "OdExpDelDt, OdExpQty, OdDelDt as qDelDt, IdTc, IdYy, IdChr, IdNo, IdSr, InDt, " + _
      "InExpNo , InCmCd, IdQty, idiFOB " + _
      "From OrdDsg, OrdMst, DsgMst, InvDsg, InvHd where 1= 2"
    ' Zubin 213 (Curr added)
    wSqlStrg = " Select " + grpflds + ", " + _
               "OdTc, OdYy, OdChr, OdNo, OdSr, OmDt, OmCmCd, OmPoNo, DmCtg, OdDmCd, OdSfx, " + _
               "OdDmSz, OdOrdQty, " + _
               IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then (OdSalPrc/OmCnvFct) else (OdSalPrc*OmCnvFct) end) ", "OdSalPrc") + " as OdSalPrc, " + _
               "OdKt, OdPrtCd, DmSalCtg, OdExpDelDt, " + _
               wExpQty + ", OdDelDt as qDelDt, " + wMtchStr + wAllPndCnd + wCnd + wCrmCnd
        Call moGrph.SetGraphProp(Rep, wsQryStruct, wSqlStrg + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
   
    
     
    ' ###########################################  Manoj  ###########################################
    
    Rep.hDmId.Suppress = True: Rep.FldOdDmId.Suppress = True
    If wGrp(1) <> "Prd Line" And wGrp(2) <> "Prd Line" And wGrp(3) <> "Prd Line" Then
        Rep.hOdLine.Suppress = False: Rep.FldOdLine.Suppress = False
    End If
    Rep.LblAsOfDt.Suppress = True
    
'
'     ElseIf UCase(adc.MenuCd) = UCase("RepSOPndRM") Then
'
'      Rep2.hDmId.Suppress = True: Rep.FldOdDmId.Suppress = True
'    If wGrp(1) <> "Prd Line" And wGrp(2) <> "Prd Line" And wGrp(3) <> "Prd Line" Then
'        Rep2.hOdLine.Suppress = False: Rep.FldOdLine.Suppress = False
'    End If
'    Rep2.LblAsOfDt.Suppress = True
    
  ElseIf UCase(adc.MenuCd) = UCase("RepJtPnd") Or UCase(adc.MenuCd) = UCase("RepJtOpnOrd") Then
    If UCase(adc.MenuCd) = UCase("RepJtPnd") Then
      wTcDesc = moCn.GetFldVal("Select PDesc from Param Where PTyp='TC' and PMCd='" + adc("UoOmTc") + "' and PValue='JT' ")
      Rep.TxtHead.SetText "Pending " + wTcDesc + "(Jwlry)" ' "Trading Voucher Pending (" + adc("UoOmTc") + ")"
    ElseIf UCase(adc.MenuCd) = UCase("RepJtOpnOrd") Then
      Rep.TxtHead.SetText "Open Reference Vchrs (Jwlry)"
    End If
      Rep.TxtAsOfDt.SetText adc("UoAsOfDt")
    Rep.hDmId.Suppress = False: Rep.FldOdDmId.Suppress = False
    Rep.hOdLine.Suppress = True: Rep.FldOdLine.Suppress = True
    Rep.SECGRPH.Suppress = True
    Rep.TxtInvVch.SetText "Matching Voucher"
    Rep.TxtInvDt.SetText "Date"
    Rep.TxtInvQty.SetText "Qty"
    Rep.TxtInvExpNo.SetText ""
    Rep.TxtInvCust.SetText "Cust"
    
  ElseIf UCase(adc.MenuCd) = UCase("RepJtPndRm") Then
    wTcDesc = moCn.GetFldVal("Select PDesc from Param Where PTyp='TC' and PMCd='" + adc("UoOmTc") + "' and PValue='JT' ")
    Rep1.TxtHead.SetText "Pending " + wTcDesc + " (Rm) (Jwlry)"
    Rep1.TxtAsOfDt.SetText adc("UoAsOfDt")
    
    ShowRepSummFld Rep1.Fldr1BalQty, Rep1.Fldr2BalQty, Rep1.Fldr3BalQty, Rep1.FldrBalQty, GRP_REP.GrpOpts, "Rm Ctg,Rm SubCtg,Rm Code,Parent Cd"
  
'    Rep1.FldOdOrdQty.Suppress = True: Rep1.FldOdSalPrc.Suppress = True
'
'
'    Rep1.HOdDmId.Suppress = True: Rep1.HOdDmSz.Suppress = True: Rep1.HOdDmCd.Suppress = True
'    Rep1.HOdSfx.Suppress = True: Rep1.HDmCtg.Suppress = True: Rep1.hOdSalPrc.Suppress = True
'    Rep1.hOdCstPrc.Suppress = True: Rep1.hProfit.Suppress = True: Rep1.hProfitPer.Suppress = True
'    Rep1.hOdOrdQty.Suppress = True: Rep1.hOdSalVal.Suppress = True
'    Rep1.gHOdOrdQty.Suppress = True: Rep1.ghOdSalVal.Suppress = True
'    Rep1.ghProfit.Suppress = True: Rep1.ghProfPer.Suppress = True
'
'
'    GRP_REP.SetFormula Rep1, "wVal", "iif({rdo.RmQw}='Q', {rdo.OrQty}, {rdo.orWt})*{rdo.OrSalRt}"
'    GRP_REP.SetFormula Rep1, "wProfit", "0"
'    GRP_REP.SetFormula Rep1, "wProfitPer", "0"
'    GRP_REP.SetFormula Rep1, "wOdSalVal", "0"
'    GRP_REP.SetFormula Rep1, "wOdCstVal", "0"
'
'    GRP_REP.SetFormula Rep1, "g_OdOrdQty", "0"
'    GRP_REP.SetFormula Rep1, "g1_OdOrdQty", "0"
'    GRP_REP.SetFormula Rep1, "g2_OdOrdQty", "0"
'    GRP_REP.SetFormula Rep1, "g3_OdOrdQty", "0"
'
'    Rep1.TxtHead.SetText "Pending Rm Report "
'
'
'
'    gb_ShowPic = False
'    Rep1.PicSec.Suppress = True
    
    Rep1.TxtCoCdFr.SetText adc("UoCoCdFr")
    
  End If
    'VK.27 - Menu code "RepJtSoPndRm" added.
    If UCase(adc.MenuCd) = UCase("RepSOPndRM") Or UCase(adc.MenuCd) = UCase("RepJtSOPndRm") Then
    ' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
   
    wsQryStruct = "  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3, space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3, space(50) as G1Desc, space(50) as G2Desc, " + _
      "space(50) as G3Desc, OdCoCd as IdCoCd , OdTc as IdTc, OdYy as IdYy, OdChr as IdChr, OdNo as IdNo, OdSr as IdSr, max(OmPoNo) as InExpNo, max(DmCtg) as DmCtg," + _
      "max(OdDmCd) as IdDmCd, max(OdSfx) as IdSfx,max(OdDmSz) as IdDmSz, max(OdKt) as IdKt, max(OmDt ) as InDt, max(OmCmCd) as InCmCd, 0 as IdExpOdTc, 0 as IdExpOdYy, 0 as IdExpOdChr, 0 as IdExpOdNo,0 as IdExpOdSr, " + _
      "DmSalCtg, max(OdVaCtg) as IdVaCtg, max(OdOrdQty) as IdQty,  OdSalPrc, OrSalVal as qMetVal, OrSalVal as qDCVal, OrSalVal as qXVal, OrSetSalVal as qSetVal, OlSalVal as qLabVal, " + _
      "OrWt as qGrWt, OrWt as qMetWt, OrWt As qDCWt, OrWt As qXWt from ordDsg, Ordmst, DsgMst, OrdRm, RmMst, OrdLab where 1= 2 "
      ' Zubin 213 (Curr added)
       Call moGrph.SetGraphProp(Rep2, wsQryStruct, wSqlStrg)   '****** Sachin 3.01 added ctMaxDopOpt
   End If
      

  
Call DispReport(adc) '****** Sachin 3.02 - Display Report ******

End Sub
'''' std code not to be changed
Private Sub ADC_Load()
    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    
    ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
    ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
    gb_CoCdFor = False
    If ws_MultiCoMod = True Then
      Set adc.FirNKeyCtl = adc("UoCoCdFr")
    Else
      Set adc.FirNKeyCtl = adc("UoOmTcFr")
    End If
    ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
    
    'Set adc.FirNKeyCtl = adc("UoOmTcFr") 'adc("UoCoCdFr")
    
    Call InitProp(Me)
     
    ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
    If ws_MultiCoMod = True Then
        adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
    Else
        adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
    End If
    ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
     
'    ' ***** Manali Trading Module
    If UCase(adc.MenuCd) = UCase("RepSoPnd") Then
      adc("UoOmTc").Enabled = False: adc("UoOmTc").Visible = False
      '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
      gs_DmTcTyp = "DM"
      gs_OmTcTyp = "SO"
      gs_CmCtg = "C"
      '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
            
    ' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
      Set moGrph = New EmrClsGrph
      Call moGrph.SetFormDets(Me, moCn)    '*** (Bef speed) bef 24/11/06
      'Call moGrph.SetFormDets(Me)  '*** (Jen speed) 24/11/06
    ' ###########################################  Manoj  ###########################################
      adc("UOODDMIDFR").Enabled = False: adc("UOODDMIDFR").Visible = False
      adc("UOODDMIDTO").Enabled = False: adc("UOODDMIDTO").Visible = False
      adc("UOSTKCOCDFR").Enabled = False: adc("UOSTKCOCDFR").Visible = False
        adc("UOSTKCOCDTO").Enabled = False: adc("UOSTKCOCDTO").Visible = False
        adc("UOBYYFR").Enabled = False: adc("UOBYYFR").Visible = False
        adc("UOBYYTO").Enabled = False: adc("UOBYYTO").Visible = False
        adc("UOBCHRFR").Enabled = False: adc("UOBCHRFR").Visible = False
        adc("UOBCHRTO").Enabled = False: adc("UOBCHRTO").Visible = False
      
      LblDmId.Visible = False: LblBaseTc.Visible = False
      LblLine.Visible = True
      adc("UOLINEFR").Enabled = True: adc("UOLINEFR").Visible = True
      adc("UOLINETO").Enabled = True: adc("UOLINETO").Visible = True
      FraRmScp.Enabled = False: FraRmScp.Visible = False ' **** Manali 3.7.1
      
      
      
      LblValAs.Enabled = False: LblValAs.Visible = False
      adc("UOINVAL").Enabled = False: adc("UOINVAL").Visible = False
      LblVal.Enabled = False: LblVal.Visible = False
      FraOptValWt.Enabled = False: FraOptValWt.Visible = False
      
       
            
      
      
    ElseIf UCase(adc.MenuCd) = UCase("RepJtPnd") Or UCase(adc.MenuCd) = UCase("RepJtOpnOrd") Or _
        UCase(adc.MenuCd) = UCase("RepJtPndRm") Then
      adc("UoOmTc").Enabled = True: adc("UoOmTc").Visible = True
      Set adc.FirNKeyCtl = adc("UoOmTc")
      gs_PTyp = "TC"
      gs_DmTcTyp = "DM"
      gs_OmTcTyp = ctJtTcTyp
      gs_CmCtg = "C"
      
      LblValAs.Enabled = False: LblValAs.Visible = False
      adc("UOINVAL").Enabled = False: adc("UOINVAL").Visible = False
      LblVal.Enabled = False: LblVal.Visible = False
      FraOptValWt.Enabled = False: FraOptValWt.Visible = False
      
      
      
      LblLine.Visible = False
      adc("UOLINEFR").Enabled = False: adc("UOLINEFR").Visible = False
      adc("UOLINETO").Enabled = False: adc("UOLINETO").Visible = False
      
      adc("UoOmTcFr").Enabled = False: adc("UoOmTcTo").Enabled = False
      adc("UoSalPerFr").Enabled = False: adc("UoSalPerFr").Visible = False
      adc("UoSalPerTo").Enabled = False: adc("UoSalPerTo").Visible = False
      adc("UoSalPerSel").Enabled = False: adc("UoSalPerSel").Visible = False: LblSalPer.Visible = False
      adc("UoPrtCdFr").Enabled = False: adc("UoPrtCdFr").Visible = False: LblPrt.Visible = False
      adc("UoPrtCdTo").Enabled = False: adc("UoPrtCdTo").Visible = False
      adc("UoPrtCdSel").Enabled = False: adc("UoPrtCdSel").Visible = False
      adc("UoDelDt1Fr").Enabled = False: adc("UoDelDt1Fr").Visible = False: lblPrdDelDt.Visible = False
      adc("UoDelDt1To").Enabled = False: adc("UoDelDt1To").Visible = False
      FraOptDelDt.Visible = False: FraOptDelDt.Enabled = False: LblShowDelDt.Visible = False
      FraGrph.Visible = False: FraGrph.Enabled = False
      adc("UoAsOfDt").Enabled = True: adc("UoAsOfDt").Visible = True: LblAsOfDt.Visible = True
      If UCase(adc.MenuCd) = UCase("RepJtPndRm") Then
        FraRmScp.Enabled = True: FraRmScp.Visible = True  ' **** Manali 3.7.1
        ChkBoxArr(1).Enabled = False: ChkBoxArr(1).Visible = False
        ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False
        ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False
        ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False
     
      Else
        FraRmScp.Enabled = False: FraRmScp.Visible = False ' **** Manali 3.7.1
      End If
   
   
        
      'Thirumozhi
      'VK.27 - Menu code "RepJtSoPndRm" added.
      ElseIf UCase(adc.MenuCd) = UCase("RepSOPndRM") Or UCase(adc.MenuCd) = UCase("RepJtSOPndRm") Then
        adc("UoOmTc").Enabled = False: adc("UoOmTc").Visible = False
     gs_DmTcTyp = "DM"
      gs_OmTcTyp = "SO"
      gs_CmCtg = "C"
      
        Set moGrph = New EmrClsGrph
      Call moGrph.SetFormDets(Me, moCn)    '
      '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
            
    ' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
'      Set moGrph = New EmrClsGrph
'      Call moGrph.SetFormDets(Me, moCn)
'
        FraRmScp.Enabled = False: FraRmScp.Visible = False
        ChkBoxArr(1).Enabled = False: ChkBoxArr(1).Visible = False
        ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False
        LblBoxDet(4).Enabled = False:  LblBoxDet(4).Visible = False
        
        adc("UOODDMIDFR").Enabled = False: adc("UOODDMIDFR").Visible = False
        adc("UOODDMIDTO").Enabled = False: adc("UOODDMIDTO").Visible = False
        adc("UOSTKCOCDFR").Enabled = False: adc("UOSTKCOCDFR").Visible = False
        adc("UOSTKCOCDTO").Enabled = False: adc("UOSTKCOCDTO").Visible = False
        adc("UOBYYFR").Enabled = False: adc("UOBYYFR").Visible = False
        adc("UOBYYTO").Enabled = False: adc("UOBYYTO").Visible = False
        adc("UOBCHRFR").Enabled = False: adc("UOBCHRFR").Visible = False
        adc("UOBCHRTO").Enabled = False: adc("UOBCHRTO").Visible = False
      
      
      
      LblDmId.Visible = False: LblBaseTc.Visible = False
      LblLine.Visible = True
      adc("UOLINEFR").Enabled = True: adc("UOLINEFR").Visible = True
      adc("UOLINETO").Enabled = True: adc("UOLINETO").Visible = True
     
    End If
    
   
    ' ***** Manali Trading Module
    'gb_CmCtgFor = False ' MANALI 3.8.0
    gs_CmCtg1 = ""
    
    Call SetGroupSort
    TAB_REP.Tab = 0
End Sub

Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gb_CoCdFor = False  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  gb_CmCtgFor = False ' MANALI 3.8.0
  gs_CmCtg1 = ""
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gb_CoCdFor = False  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  '''gb_CoCdFor = True         '*** (Jen 2.14 Next Patch 1)
  gb_CmCtgFor = False ' MANALI 3.8.0
  If UCase(adc.MenuCd) = UCase("RepSoPnd") Or UCase(adc.MenuCd) = UCase("RepSOPndRM") Then
    gs_DmTcTyp = "DM"
    gs_OmTcTyp = "SO"
    gs_CmCtg = "C"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  Else
    gs_PTyp = "TC"
    gs_DmTcTyp = "DM"
    gs_OmTcTyp = ctJtTcTyp
    gs_CmCtg = moCn.GetFldVal("Select TmCmCtg from TcMst Where TmCoCd='" + gs_CoCd + "' and TmTc='" + adc("UoOmTc") + "' ")
    gs_CmCtg = IIF(gs_CmCtg = "", "C", gs_CmCtg)
    
  End If
  gs_CmCtg1 = ""
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  If UCase$(IdName) = UCase$("UoGrphTyp") Then gs_GrphGrpNo = moGrph.SetGrphGrpNo
' ###########################################  Manoj  ###########################################
    
  SetProp Me, IdName, When
  
   Select Case UCase(IdName)
   Case Is = UCase("UOINVAL")      '3.11.0
    Call HlpList.PMCd("ORDVAL")
  End Select
End Sub


Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  gb_CoCdFor = False  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  Select Case UCase(IdName)
  Case Is = UCase("UoOmTcFr")
    'VK.27 - Menu code "RepJtSoPndRm" added.
    If UCase(adc.MenuCd) <> UCase("RepSoPnd") And UCase(adc.MenuCd) <> UCase("RepSOPndRM") And UCase(adc.MenuCd) <> UCase("RepJtSOPndRm") Then
      adc("UoOmTcFr") = adc("UoOmTc")
      adc("UoOmTcTo") = adc("UoOmTc")
       
    End If
  End Select
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowPnd"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1     'Tag = "ShowBoxDet"
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  Case Is = 2     'Tag = "ShowPic"
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  Case Is = 4     'Tag = "ShowRem"
    If .Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"
  End Select
End With
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptDelDt_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN3") = "Y"
Case Is = 1
  adc("UoYN3") = "N"
Case Is = 2
  adc("UoYN3") = ""
End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptDelDt_GotFocus(Index As Integer)
  DispMsg FraOptDelDt.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

'Thirumozhi
Private Sub OptValWt_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN5") = "Y"
  LblValAs.Enabled = True: LblValAs.Visible = True
  adc("UOINVAL").Enabled = True: adc("UOINVAL").Visible = True
      
Case Is = 1
  adc("UoYN5") = "N"
  LblValAs.Enabled = False: LblValAs.Visible = False
  adc("UOINVAL").Enabled = False: adc("UOINVAL").Visible = False
  adc("UoInVal") = ""
     
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
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  If ws_MultiCoMod = False Then _
      adc("UoCoCdFr") = gs_CoCd: adc("UoCoCdTo") = gs_CoCd: adc("UoCoCdSel") = "": adc("UoCoCdFr").Enabled = False
  
  'adc("UoCoCdFr") = gs_CoCd

' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  If adc("UoShowGrph") = "" Then adc("UoShowGrph") = "N"
  If UCase(adc.MenuCd) = UCase("RepSoPnd") Then
    Call moGrph.NonKeyWhen
  End If
  
  'Thirumozhi
'   If UCase(adc.MenuCd) = UCase("RepSOPndRM") Then
'    Call moGrph.NonKeyWhen
'  End If
' ###########################################  Manoj  ###########################################

  ' ***** Manali Trading Module
  'If UCase(ADC.MenuCd) = UCase("RepJtPnd") Or UCase(ADC.MenuCd) = UCase("RepJtOpnOrd") Then
  'VK.27 - Menu code "RepJtSoPndRm" added.
  If UCase(adc.MenuCd) <> UCase("RepSoPnd") And UCase(adc.MenuCd) <> UCase("RepSOPndRM") And UCase(adc.MenuCd) <> UCase("RepJtSOPndRm") Then
    adc("UoDelDt1Fr") = "01/01/80": adc("UoDelDt1To") = "01/01/80"
    adc("UoYn3") = "Y"
  End If
  ' ***** Manali Trading Module
  
  ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  If adc("UoYN3") = "Y" Then
    OptDelDt(0).Value = True
  ElseIf adc("UoYN3") = "N" Then
    OptDelDt(1).Value = True
  ElseIf adc("UoYN3") = "" Then
    OptDelDt(2).Value = True
  End If
  ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  If adc("UoYN2") = "" Then adc("UoYN2") = "N"
  If adc("UoYN4") = "" Then adc("UoYN4") = "N"
  
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If adc("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added

'Thirumozhi
  If adc("UoYN5") = "Y" Then
  OptValWt(0).Value = True
  
  End If
  
  If adc("UoYN5") = "N" Then
  OptValWt(1).Value = True
 
  End If

End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + IIF(ws_MultiCoMod = True, gs_CoCd, adc("UOCOCDFR")) + "' and " + _
                                                    "HCd = '" + ctSelfCmCd + "' ")

'  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
'                                                    "HCd = '" + ctSelfCmCd + "' ")

  mRep.wCoCd.SetText ws_HName
  'Rep.wCoCd.SetText ws_HName
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  mRep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    mRep.RHMain.UnderlaySection = True: mRep.wCoCd.Suppress = True
  Else
    mRep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  If ws_MultiCoMod = True Then
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdFr.SetText adc("UOCOCDFR"): Rep.TxtCoCdTo.SetText adc("UOCOCDTO"): Rep.TxtCoCdSel.SetText adc("UOCOCDSEL")
   
      Rep2.TxtCoCdToLbl.Suppress = False: Rep2.TxtCoCdSelLbl.Suppress = False
      Rep2.TxtCoCdFr.SetText adc("UOCOCDFR"): Rep2.TxtCoCdTo.SetText adc("UOCOCDTO"): Rep2.TxtCoCdSel.SetText adc("UOCOCDSEL")
 
  Else
      Rep.TxtCoCdTo.Suppress = True: Rep.TxtCoCdSel.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True: Rep.TxtCoCdSelLbl.Suppress = True
      Rep.TxtCoCdFr.SetText adc("UOCOCDFR")
      
       Rep2.TxtCoCdTo.Suppress = True: Rep2.TxtCoCdSel.Suppress = True
      Rep2.TxtCoCdToLbl.Suppress = True: Rep2.TxtCoCdSelLbl.Suppress = True
      Rep2.TxtCoCdFr.SetText adc("UOCOCDFR")

  End If
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  
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
  mRep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  'Rep.TxtRepUsrDtTime.SetText Trim(ADC.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  
End Sub
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  'fra_keys.Picture = Me.Picture
  'fra_keys.BorderStyle = 0
  'fra_nkeys.Picture = Me.Picture
  'fra_nkeys.BorderStyle = 0
  
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  Call SetObjColors(Me)
  '*** Jenny Colour
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
    Set moGrph = Nothing
    '*** (09/08/05)
End Sub

' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
Private Sub GRP_REP_LostFocus()
  'VK.27 - Menu code "RepJtSoPndRm" added.
  If UCase(adc.MenuCd) = UCase("RepSoPnd") Or UCase(adc.MenuCd) = UCase("RepSOPndRM") Or UCase(adc.MenuCd) = UCase("RepJtSOPndRm") Then Call moGrph.SetDefaVal
'  If UCase(adc.MenuCd) = UCase("RepSOPndRM") Then Call moGrph.SetDefaVal
End Sub
Private Sub TAB_GRPH_Click(PreviousTab As Integer)
    FraTabGrph(TAB_GRPH.Tab).Enabled = True
    FraTabGrph(PreviousTab).Enabled = False
End Sub
Private Sub ChkGrph_Click()
  If ChkGrph.Value = vbChecked Then
    adc("UoShowGrph") = "Y"
    FraGrphDet.Enabled = True
    moGrph.SetDefaVal
  ElseIf ChkGrph.Value = vbUnchecked Then
    adc("UoShowGrph") = "N"
    FraGrphDet.Enabled = False
  End If
End Sub
Private Sub cboGrphValue_Click()
  adc("UoGrphValInd") = cboGrphValue.ListIndex
End Sub
' ###########################################  Manoj  ###########################################




