VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpCustmStk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Stock"
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
      Left            =   11550
      TabIndex        =   84
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
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5655
      TabIndex        =   93
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   94
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
         TabIndex        =   95
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
      Left            =   6870
      TabIndex        =   92
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   89
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
      TabIndex        =   85
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
      TabIndex        =   88
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   87
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpCustmStk.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpCustmStk.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
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
               TabIndex        =   86
               Top             =   360
               Width           =   9375
               _ExtentX        =   16536
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Caption         =   "s"
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   90
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Rate/Value"
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
               TabIndex        =   68
               Tag             =   "ShowPnd"
               ToolTipText     =   "Check To Show Rate & Value"
               Top             =   6960
               Width           =   2070
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show 2nd Line"
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
               Left            =   30
               TabIndex        =   78
               Tag             =   "ShowLine"
               ToolTipText     =   "Check To Show 2nd Line"
               Top             =   8420
               Width           =   2070
            End
            Begin VB.Frame FraOptBestUse 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1890
               TabIndex        =   44
               ToolTipText     =   "Specify Whether To Show Records With Some Best Use Field Filled Up Yes / No / All"
               Top             =   8690
               Width           =   2685
               Begin VB.OptionButton OptBestUse 
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
                  Left            =   1680
                  TabIndex        =   82
                  ToolTipText     =   "Specify Whether To Show Records With Some Best Use Field Filled Up Yes / No / All"
                  Top             =   0
                  Width           =   735
               End
               Begin VB.OptionButton OptBestUse 
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
                  Left            =   870
                  TabIndex        =   81
                  ToolTipText     =   "Specify Whether To Show Records With Some Best Use Field Filled Up Yes / No / All"
                  Top             =   0
                  Width           =   795
               End
               Begin VB.OptionButton OptBestUse 
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
                  TabIndex        =   80
                  ToolTipText     =   "Specify Whether To Show Records With Some Best Use Field Filled Up Yes / No / All"
                  Top             =   0
                  Width           =   825
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Pending"
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
               TabIndex        =   65
               Tag             =   "ShowPnd"
               ToolTipText     =   "Check To Show Pending Records"
               Top             =   6390
               Width           =   2070
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Box Det           "
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
               TabIndex        =   67
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show box Details"
               Top             =   6660
               Width           =   2070
            End
            Begin VB.Frame FraOptDerVal 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1890
               TabIndex        =   45
               ToolTipText     =   "Select Derived Value or Actual Value"
               Top             =   8120
               Width           =   4875
               Begin VB.OptionButton OptDerVal 
                  Caption         =   "Derived Value"
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
                  TabIndex        =   73
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.OptionButton OptDerVal 
                  Caption         =   "Actual Value"
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
                  Left            =   2460
                  TabIndex        =   74
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.Frame FraOptDerRt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1890
               TabIndex        =   46
               ToolTipText     =   "Select Derived Rate or Actual Rate"
               Top             =   7835
               Width           =   4905
               Begin VB.OptionButton OptDerRt 
                  Caption         =   "Actual Rate"
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
                  Left            =   2460
                  TabIndex        =   72
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.OptionButton OptDerRt 
                  Caption         =   "Derived Rate"
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
                  TabIndex        =   71
                  ToolTipText     =   "Select Derived Rate"
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4395
               TabIndex        =   12
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   1230
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "TDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1905
               TabIndex        =   11
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   1230
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "TDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   5175
               TabIndex        =   5
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   660
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
               CmpStr          =   "B.TzYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2685
               TabIndex        =   2
               ToolTipText     =   "Enter From Transaction Year"
               Top             =   660
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
               CmpStr          =   "B.TzYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1905
               TabIndex        =   1
               ToolTipText     =   "Enter From Transaction Tc"
               Top             =   660
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
               CmpStr          =   "B.TzTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   3105
               TabIndex        =   3
               ToolTipText     =   "Enter From Transaction Character"
               Top             =   660
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
               CmpStr          =   "B.TzChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   4395
               TabIndex        =   4
               ToolTipText     =   "Enter To Transaction Tc"
               Top             =   660
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
               CmpStr          =   "B.TzTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   5595
               TabIndex        =   6
               ToolTipText     =   "Enter To Transaction Character"
               Top             =   660
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
               CmpStr          =   "B.TzChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5295
               TabIndex        =   10
               ToolTipText     =   "Enter To Transaction Serial"
               Top             =   945
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRTO"
               IdName          =   "UOTDSRTO"
               CmpStr          =   "B.TzSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   2805
               TabIndex        =   8
               ToolTipText     =   "Enter From Transaction Serial"
               Top             =   945
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRFR"
               IdName          =   "UOTDSRFR"
               CmpStr          =   "B.TzSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1905
               TabIndex        =   7
               ToolTipText     =   "Enter From Transaction Number"
               Top             =   945
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOFR"
               IdName          =   "UOTNOFR"
               CmpStr          =   "B.TzNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   4395
               TabIndex        =   9
               ToolTipText     =   "Enter To Transaction Number"
               Top             =   945
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOTO"
               IdName          =   "UOTNOTO"
               CmpStr          =   "B.TzNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   4395
               TabIndex        =   30
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   3225
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "RmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6900
               TabIndex        =   31
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   3225
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "RmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   4395
               TabIndex        =   27
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   2940
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
               Index           =   1
               Left            =   6900
               TabIndex        =   28
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   2940
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "RmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1905
               TabIndex        =   23
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   2655
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "RmCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   1905
               TabIndex        =   26
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   2940
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
               CmpStr          =   "RmSCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1905
               TabIndex        =   29
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   3225
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "RmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   4395
               TabIndex        =   24
               ToolTipText     =   "Enter To Raw Material Category"
               Top             =   2655
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
               CmpStr          =   "RmCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   6900
               TabIndex        =   25
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   2655
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCTGSEL"
               IdName          =   "UORMCTGSEL"
               CmpStr          =   "RmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   4395
               TabIndex        =   19
               ToolTipText     =   "Enter To Supplier"
               Top             =   2085
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "TSuppCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   6900
               TabIndex        =   20
               ToolTipText     =   "Enter Supplier Selection"
               Top             =   2085
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "TSuppCd In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   1905
               TabIndex        =   18
               ToolTipText     =   "Enter From Supplier"
               Top             =   2085
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "TSuppCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   4395
               TabIndex        =   14
               ToolTipText     =   "Enter To Bill Of Entry Number "
               Top             =   1515
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBOETO"
               IdName          =   "UOBOETO"
               CmpStr          =   "TBOENo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   6900
               TabIndex        =   15
               ToolTipText     =   "Enter Bill Of Entry Number Selection"
               Top             =   1515
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOBOESEL"
               IdName          =   "UOBOESEL"
               CmpStr          =   "TBOENo In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   1905
               TabIndex        =   13
               ToolTipText     =   "Enter From Bill Of Entry Number "
               Top             =   1515
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBOEFR"
               IdName          =   "UOBOEFR"
               CmpStr          =   "TBOENo >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   4395
               TabIndex        =   17
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   1800
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOBOEDTTO"
               IdName          =   "UOBOEDTTO"
               CmpStr          =   "TBOEDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1905
               TabIndex        =   16
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   1800
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOBOEDTFR"
               IdName          =   "UOBOEDTFR"
               CmpStr          =   "TBOEDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4395
               TabIndex        =   22
               ToolTipText     =   "Enter To Bill Of Entry Number "
               Top             =   2370
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBILLNOTO"
               IdName          =   "UOBILLNOTO"
               CmpStr          =   "TBillNo <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1905
               TabIndex        =   21
               ToolTipText     =   "Enter From Bill Of Entry Number "
               Top             =   2370
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBILLNOFR"
               IdName          =   "UOBILLNOFR"
               CmpStr          =   "TBillNo >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2430
               TabIndex        =   66
               ToolTipText     =   "Show Box Detail ? (Yes / No)"
               Top             =   6660
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
               Index           =   9
               Left            =   2430
               TabIndex        =   64
               ToolTipText     =   "Show Pending Records ? (Yes / No)"
               Top             =   6375
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
               Index           =   11
               Left            =   1905
               TabIndex        =   63
               ToolTipText     =   "Enter Stock As Of Date"
               Top             =   6090
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOASOFDT"
               IdName          =   "UOASOFDT"
               CmpStr          =   "TDt <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   4395
               TabIndex        =   35
               ToolTipText     =   "Enter To Purchase Rate"
               Top             =   3795
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOSTKRTTO"
               IdName          =   "UOSTKRTTO"
               CmpStr          =   "B.TzPurRt <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   1905
               TabIndex        =   34
               ToolTipText     =   "Enter From Purchase Rate"
               Top             =   3795
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOSTKRTFR"
               IdName          =   "UOSTKRTFR"
               CmpStr          =   "B.TzPurRt >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   4395
               TabIndex        =   37
               ToolTipText     =   "Enter To Average Piece Per Carat"
               Top             =   4080
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOPCPERCTTO"
               IdName          =   "UOPCPERCTTO"
               CmpStr          =   "B.TzPcPerCt <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   1905
               TabIndex        =   36
               ToolTipText     =   "Enter From Average Piece Per Carat"
               Top             =   4080
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOPCPERCTFR"
               IdName          =   "UOPCPERCTFR"
               CmpStr          =   "B.TzPcPerCt >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   1905
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   375
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "B.TzCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1905
               TabIndex        =   38
               ToolTipText     =   "Enter From Import Type Scope"
               Top             =   4365
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
               CmpStr          =   "TImpTyp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4395
               TabIndex        =   39
               ToolTipText     =   "Enter To Import Type Scope"
               Top             =   4365
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
               CmpStr          =   "TImpTyp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   6900
               TabIndex        =   40
               ToolTipText     =   "Enter Import Type Selection"
               Top             =   4365
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPMCDSEL"
               IdName          =   "UOPMCDSEL"
               CmpStr          =   "TImpTyp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1905
               TabIndex        =   70
               ToolTipText     =   "Enter Voucher Currency  US$[Y]/Rs[N]  and ' ' for ALL records in US$ "
               Top             =   7550
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1905
               TabIndex        =   69
               ToolTipText     =   "Enter Currency Code Whose Vouchers You Want To View Or Leave It Blank To View All Vouchers In Base Currency"
               Top             =   7250
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOCURCDFR"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   6855
               TabIndex        =   76
               ToolTipText     =   "Enter 'Y' for Derived Value / 'N' for Actual value"
               Top             =   8040
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
               Index           =   33
               Left            =   1905
               TabIndex        =   32
               ToolTipText     =   "Enter Raw Material Stone Class From Range"
               Top             =   3510
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSFR"
               IdName          =   "UOSTNCLSFR"
               CmpStr          =   "IsNull(SCls.PSCd, '') >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4395
               TabIndex        =   33
               ToolTipText     =   "Enter Raw Material Stone Class To Range"
               Top             =   3510
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSTO"
               IdName          =   "UOSTNCLSTO"
               CmpStr          =   "IsNull(SCls.PSCd, '') <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   5175
               TabIndex        =   51
               ToolTipText     =   "Enter To Order Year"
               Top             =   4950
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "B.TzOdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   4395
               TabIndex        =   50
               ToolTipText     =   "Enter To Order Tc"
               Top             =   4950
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "B.TzOdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   5595
               TabIndex        =   52
               ToolTipText     =   "Enter To Order Character"
               Top             =   4950
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "B.TzOdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   1905
               TabIndex        =   53
               ToolTipText     =   "Enter From Order Number"
               Top             =   5235
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "B.TzOdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   4395
               TabIndex        =   55
               ToolTipText     =   "Enter To Order Number"
               Top             =   5235
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "B.TzOdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   2685
               TabIndex        =   48
               ToolTipText     =   "Enter From Order Year"
               Top             =   4950
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "B.TzOdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   1905
               TabIndex        =   47
               ToolTipText     =   "Enter From Order Tc"
               Top             =   4950
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "B.TzOdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   3105
               TabIndex        =   49
               ToolTipText     =   "Enter From Order Character"
               Top             =   4950
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "B.TzOdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   4395
               TabIndex        =   58
               ToolTipText     =   "Enter To Order Design Code"
               Top             =   5520
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "B.TzDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   6900
               TabIndex        =   59
               ToolTipText     =   "Enter Order Design Code Selection"
               Top             =   5520
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "B.TzDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   1905
               TabIndex        =   57
               ToolTipText     =   "Enter From Order Design Code"
               Top             =   5520
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "B.TzDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   4395
               TabIndex        =   61
               ToolTipText     =   "Enter Order Raw Material Code To Range"
               Top             =   5805
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UORMCD1TO"
               IdName          =   "UORMCD1TO"
               CmpStr          =   "B.TzRmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   6900
               TabIndex        =   62
               ToolTipText     =   "Enter Order Raw Material Code Selection"
               Top             =   5805
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCD1SEL"
               IdName          =   "UORMCD1SEL"
               CmpStr          =   "B.TzRmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   1905
               TabIndex        =   60
               ToolTipText     =   "Enter Order Raw Material Code From Range"
               Top             =   5805
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UORMCD1FR"
               IdName          =   "UORMCD1FR"
               CmpStr          =   "B.TzRmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   2415
               TabIndex        =   77
               ToolTipText     =   "Specify Whether To Show Second Detail Line Or Not ('Y' Yes / 'N' No)"
               Top             =   8325
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
               Index           =   58
               Left            =   6840
               TabIndex        =   136
               ToolTipText     =   "Enter 'Y'/'N' For Best Use"
               Top             =   8640
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
               Index           =   59
               Left            =   6840
               TabIndex        =   137
               ToolTipText     =   "Enter 'Y' for Derived Rate / 'N' for Actual Rate"
               Top             =   7800
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
               Index           =   64
               Left            =   4395
               TabIndex        =   42
               ToolTipText     =   "Enter To Customer"
               Top             =   4650
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOEXPCMCDTO"
               IdName          =   "UOEXPCMCDTO"
               CmpStr          =   "TCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   65
               Left            =   6900
               TabIndex        =   43
               ToolTipText     =   "Enter Customer Selection"
               Top             =   4650
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOEXPCMCDSEL"
               IdName          =   "UOEXPCMCDSEL"
               CmpStr          =   "TCmCd In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   66
               Left            =   1905
               TabIndex        =   41
               ToolTipText     =   "Enter From Customer"
               Top             =   4650
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOEXPCMCDFR"
               IdName          =   "UOEXPCMCDFR"
               CmpStr          =   "TCmCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   67
               Left            =   2835
               TabIndex        =   54
               ToolTipText     =   "Enter From Order Serial"
               Top             =   5235
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "B.TzOdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   68
               Left            =   5325
               TabIndex        =   56
               ToolTipText     =   "Enter To Order Serial"
               Top             =   5235
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "B.TzOdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   69
               Left            =   2400
               TabIndex        =   140
               ToolTipText     =   "Show Rate & Value (Yes / No)"
               Top             =   6960
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN7"
               IdName          =   "UOYN7"
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
               Left            =   5235
               TabIndex        =   139
               Top             =   5235
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
               Left            =   2760
               TabIndex        =   138
               Top             =   5235
               Width           =   105
            End
            Begin VB.Label LblCustCd 
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
               Left            =   60
               TabIndex        =   75
               Top             =   4650
               Width           =   1425
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Rate"
               BeginProperty Font 
                  Name            =   "Arial"
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
               ToolTipText     =   "Location"
               Top             =   7835
               Width           =   1485
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Best Use"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   135
               ToolTipText     =   "Location"
               Top             =   8690
               Width           =   1485
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Show 2nd Line"
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
               TabIndex        =   134
               ToolTipText     =   "Location"
               Top             =   8420
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.Label Label4 
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
               Left            =   60
               TabIndex        =   133
               Top             =   5805
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
               TabIndex        =   132
               Top             =   5520
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
               Index           =   26
               Left            =   5115
               TabIndex        =   131
               Top             =   4950
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
               Left            =   5535
               TabIndex        =   130
               Top             =   4950
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Ord No/Sr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   129
               Top             =   5235
               Width           =   1395
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   128
               Top             =   4950
               Width           =   1515
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
               Left            =   3045
               TabIndex        =   127
               Top             =   4950
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
               Left            =   2625
               TabIndex        =   126
               Top             =   4950
               Width           =   105
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
               Index           =   18
               Left            =   60
               TabIndex        =   125
               Top             =   3510
               Width           =   1425
            End
            Begin VB.Label Label3 
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   60
               TabIndex        =   124
               ToolTipText     =   "Location"
               Top             =   8120
               Width           =   1485
            End
            Begin VB.Label Label2 
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   60
               TabIndex        =   123
               ToolTipText     =   "Location"
               Top             =   7250
               Width           =   1485
            End
            Begin VB.Label LblUSDYN 
               BackStyle       =   0  'Transparent
               Caption         =   "In US$"
               BeginProperty Font 
                  Name            =   "Arial"
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
               ToolTipText     =   "Location"
               Top             =   7550
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Import Type"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   121
               Top             =   4365
               Width           =   1545
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
               Index           =   36
               Left            =   60
               TabIndex        =   120
               Top             =   375
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Avg Pc Per Ct"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   119
               Top             =   4080
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Purchase Rate"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   118
               Top             =   3795
               Width           =   1545
            End
            Begin VB.Label ALBL 
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   6
               Left            =   60
               TabIndex        =   117
               Top             =   6090
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Pending "
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
               Index           =   5
               Left            =   60
               TabIndex        =   116
               ToolTipText     =   "Location"
               Top             =   6615
               Visible         =   0   'False
               Width           =   1485
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
               Index           =   4
               Left            =   60
               TabIndex        =   115
               ToolTipText     =   "Location"
               Top             =   6660
               Visible         =   0   'False
               Width           =   1485
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
               Left            =   6900
               TabIndex        =   114
               Top             =   60
               Width           =   1275
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Supp Bill No"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   113
               Top             =   2370
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "BOE Date"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   112
               Top             =   1800
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch Date"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   111
               Top             =   1230
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
               Left            =   2745
               TabIndex        =   110
               Top             =   945
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
               Left            =   5235
               TabIndex        =   109
               Top             =   945
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
               Left            =   2625
               TabIndex        =   108
               Top             =   660
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
               Left            =   5115
               TabIndex        =   107
               Top             =   660
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
               Index           =   29
               Left            =   3045
               TabIndex        =   106
               Top             =   660
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
               Left            =   5535
               TabIndex        =   105
               Top             =   660
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
               Index           =   12
               Left            =   60
               TabIndex        =   104
               Top             =   945
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
               TabIndex        =   103
               Top             =   660
               Width           =   1605
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
               Left            =   60
               TabIndex        =   102
               Top             =   3225
               Width           =   1425
            End
            Begin VB.Label LblRmSCtg 
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
               Left            =   60
               TabIndex        =   101
               Top             =   2940
               Width           =   1425
            End
            Begin VB.Label LblRmCtg 
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
               Left            =   60
               TabIndex        =   100
               ToolTipText     =   "Location"
               Top             =   2655
               Width           =   1425
            End
            Begin VB.Label LblSuppCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Supplier"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Top             =   2085
               Width           =   1425
            End
            Begin VB.Label LblBOENo 
               BackStyle       =   0  'Transparent
               Caption         =   "BOE No"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   98
               Top             =   1515
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "From "
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   1905
               TabIndex        =   97
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
               Left            =   4395
               TabIndex        =   96
               Top             =   60
               Width           =   1275
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpCustmStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepCustmStk
Dim Rep1 As New EmrRepCustmStkLand '*** Jay 3.1.1 (All Rates Rpt)'*** Added Again In 3.2.0
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  GRP_REP.Add "VoucherNo", "B.TzTc+B.TzYy+B.TzChr+str(B.TzNo)", "B.TzTc+'/'+B.TzYy+'/'+B.TzChr+'/'+LTrim(Str(B.TzNo))", "", "", "Supplier,Vch Date,Customer", "", "", ""
  '6.1 vch date, BOE, Boe date
  GRP_REP.Add "Vch Date", "convert(char(8),TDt,112)", "convert(varchar, TDt, " + CStr(DtFmtSlahYyyyVal()) + ")", "", "", "", "", "", ""
  GRP_REP.Add "BOE", "TBOENo", "", "TBOENo", "hTBOENo", "BOE Date", "", "", "(case when TBOEDt= '01/01/80' then '' else convert(varchar, TBOEDt, " + CStr(DtFmtSlahYyyyVal()) + ") end)"
  GRP_REP.Add "BOE Date", "convert(char(8),TBOEDt,112)", "(case when TBOEDt= '01/01/80' then '' else convert(varchar, TBOEDt, " + CStr(DtFmtSlahYyyyVal()) + ") end)", "TBOEDt", "hTBOEDt", "", "", "", ""
    '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "Supplier", "TSuppCd", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'S' and CmCd= TSuppCd) "
  '6.1
  GRP_REP.Add "Supp Bill No", "TBillNo", "", "", "", "Supplier", "", "", "(case when TBillDt='01/01/80' then '' else convert(varchar, TBillDt, " + CStr(DtFmtSlahYyyyVal()) + ") end)"
  '(Bef 3.1.1) GRP_REP.Add "Rm Code", "RmCtg+RmSCtg+B.TzRmZCd", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ B.TzRmZCd", "TzRmZCd", "hTzRmZCd", "Rm Ctg,Rm SubCtg", "", "", "(Select SubString(RmDesc, 1, 30) From RmMst Where RmCd= B.TzRmZCd) "
  '*** (Jen 3.1.1)
  GRP_REP.Add "Rm Code", "RmCtg+RmSCtg+B.TzRmZCd", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ B.TzRmZCd", "TzRmZCd", "hTzRmZCd", "Rm Ctg,Rm SubCtg", "", "", "SubString(RmDesc, 1, 30) "
  '*** (Jen 3.1.1)
  GRP_REP.Add "Rm Ctg", "RmCtg", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = RmCtg) "
  GRP_REP.Add "Rm SubCtg", "RmCtg+RmSCtg", "'('+ RTrim(RmCtg)+ ') '+ RmSCtg", "", "", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
  ' Zubin 213 (Bef 213)
  'GRP_REP.Add "Purchase Rt", "str(B.TzPurRt,9,2)", "", "TzPurRt", "hTzPurRt", "", "", "", ""
  'GRP_REP.Add "(RmCd) PurRt", "B.TzRmZCd+str(B.TzPurRt,9,2)", "'('+B.TzRmZCd+') '+LTrim(str(B.TzPurRt,9,2))", "TzRmZCd,TzPurRt", "hTzRmZCd,hTzPurRt", "Rm Code,Purchase Rt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  
  ' Zubin 213
  GRP_REP.Add "Purchase Rt", "str(case when qBaseCur= 'Y' then (case when TMulDiv= 'M' then B.TzPurRt/TCnvRt else B.TzPurRt*TCnvRt end) else B.TzPurRt End,9,2)", "", "wStkRt", "hTzPurRt", "", "", "", ""
  GRP_REP.Add "(RmCd) PurRt", "B.TzRmZCd+str(case when qBaseCur= 'Y' then (case when TMulDiv= 'M' then B.TzPurRt/TCnvRt else B.TzPurRt*TCnvRt end) else B.TzPurRt End,9,2)", "'('+B.TzRmZCd+') '+LTrim(str(case when qBaseCur= 'Y' then (case when TMulDiv= 'M' then B.TzPurRt/TCnvRt else B.TzPurRt*TCnvRt end) else B.TzPurRt End,9,2))", "TzRmZCd,wStkRt", "hTzRmZCd,hTzPurRt", "Rm Code,Purchase Rt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  
  '********* urmi 28 jun 03 ******************
  GRP_REP.Add "Import Typ", "TImpTyp", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'ImpTyp' and PmCd= TImpTyp) "
  '********* urmi 28 jun 03 ******************
  
  '********* sachin jun 05 *******************
  GRP_REP.Add "Rm Desc", "B.TzRmZCd+B.TzRmCdDesc", "B.TzRmCdDesc", "", "", "Rm Ctg,Rm SubCtg,Rm Code", "", "", "B.TzRmCdDesc"
  '********* sachin jun 05 *******************
  
  '*** (Jen 3.1.1)
  'GRP_REP.Add "Stone Class", "RmCtg+IsNull(SCls.PSCd, '')", "'('+ RTrim(RmCtg)+') '+IsNull(SCls.PSCd, '')", "RmCtg", "hRmCtg", "Rm Ctg", "", "", "SCls.PDesc"
  GRP_REP.Add "Stone Class", "RmCtg+IsNull(SCls.PSCd, '')", "'('+ RTrim(RmCtg)+') '+IsNull(SCls.PSCd, '')", "", "", "Rm Ctg", "", "", "SCls.PDesc"
  '*** (Jen 3.1.1)
  
  '*** (Jen 3.2.0)
  GRP_REP.Add "Order No", "B.TzOdTc+B.TzOdYy+B.TzOdChr+str(B.TzOdNo)", "(case when B.TzOdTc= '' then '' else B.TzOdTc+'/'+B.TzOdYy+'/'+B.TzOdChr+'/'+LTrim(Str(B.TzOdNo)) end)", "", "", "", "", "", "(case when B.TzOdTc= '' then '' else IsNull((Select '('+OmCmCd+ '/'+ OmPONo+ ')' from OrdMst where " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey=B.TzPrtKey and OmPrtKey='" + ctCurrPrtn + "' and ", "") + " OmCoCd= B.TzCoCd and OmTc= B.TzOdTc and OmYy= B.TzOdYy and OmChr= B.TzOdChr and OmNo= B.TzOdNo), '') end)"
  GRP_REP.Add "Ord Design", "B.TzDmCd", "", "", "", "", "", "", "(case when B.TzDmCd= '' then '' else IsNull((Select SubString(DmDesc, 1, 45) from DsgMst where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' and ", "") + " DmTcTyp= 'DM' and DmCd= B.TzDmCd and DmSz= ''), '') end)"
  'GRP_REP.Add "Ord Rm Code", "B.TzRmCd", "", "TzRmCd", "hTzRmCd", "Rm Ctg,Rm SubCtg", "", "", "(case when B.TzRmCd= '' then '' else IsNull((Select SubString(Orm.RmDesc, 1, 45) from RmMst Orm where Orm.RmCd= B.TzRmCd), '') end)"
  GRP_REP.Add "Ord Rm Code", "B.TzRmCd", "", "", "", "Rm Ctg", "", "", "(case when B.TzRmCd= '' then '' else IsNull((Select SubString(Orm.RmDesc, 1, 45) from RmMst Orm where " + IIF(gs_Partition = ctCurrPrtn, " Orm.RmPrtKey=B.TzPrtKey and Orm.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " Orm.RmCd= B.TzRmCd), '') end)"
  
  'Dim wsPtrYN As String
  'wsPtrYN = "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg)"
  'GRP_REP.Add "Ord Rm Fr Sv/Len", "(case when " + wsPtrYN + "= 'Y' then Str(B.TzRmSzFr, 7, 4) else Str(B.TzRmSzFr, 6, 3) end ", "", "", "", "", "", "", ""
  'GRP_REP.Add "Ord Rm To Sv/Len", "(case when " + wsPtrYN + "= 'Y' then Str(B.TzRmSzTo, 7, 4) else Str(B.TzRmSzTo, 6, 3) end ", "", "", "", "", "", "", ""
  '*** (Jen 3.2.0)
  
  ' ***** Manali 3.03 - 30/06/08
  GRP_REP.Add "Customer", "IsNull(TCmCd, '')", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= IsNull(TCmCd, '')) "

  
End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  '*** Report Sql ***
  ' **** Manali 3.03 - 30/06/08 - TCmCd added
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'TzTc as qBaseTc, TzYy as qBaseYy, TzChr as qBaseChr, TzNo as qBaseNo,
  'TzSr as qBaseSr, TDt, TBOENo, TBOEDt, TCmCd, TSuppCd, TBillNo, RmCtg, RmSCtg,
  'TzRmZCd, TzPcPerCt, TzRmWt as qBaseWt, TzRmDespPurgeWt, TzRmWt as qBalWt, TzPurRt,
  'TzPurAmt, TNetAmt, TNetAmt as qAddAmt, TzPurRt as qTzPurBaseRt, TzPurAmt as qTzPurBaseAmt,
  'TNetAmt as qTNetBaseAmt, TTotAmt as qAddBaseAmt, TzPurRt as qTzPurReRt, TzPurAmt as qTzPurReAmt,
  'TNetAmt qTNetReAmt, TTotAmt as qAddReAmt,
  'TzTc as qMatchTc, TzYy as qMatchYy, TzChr as qMatchChr,
  'TzNo as qMatchNo, TzSr as qMatchSr, TzSrNo as qMatchSrNo, InExpNo,
  'TzRmWt As qMatchWt, IzRmZRt As qMatchRt, InCmCd
  'From TxndZ, Txn, RmMst, InvZRm, InvHd where 1= 2
  '*** Report Sql ***
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String
  Dim wStr As String, wBalWt As String, wAllPndCnd As String, wImpCnd As String
  Dim wStr2 As String
  '*** Added Again In 3.2.0
  Dim wFct As String 'wPurStr As String   '*** Jay 3.1.0_1 [Derived Value]
  Dim wBaseFct As String, wReFct As String, wRepRtFlds As String '*** Jay 3.1.1 (All Rates Rpt)
  '*** Added Again In 3.2.0
  
  Dim wAddlCnd As String '****** Sachin 3.02.0 - New Fields In Where Clause
  
  Set Rep = Nothing
  '*** Added Again In 3.2.0
  Set Rep1 = Nothing    '*** Jay 3.1.1 (All Rates Rpt)
  '*** Jay 3.1.1 (All Rates Rpt)
  Select Case UCase(adc.MenuCd)
  Case UCase("RepCustmStk")
    Set adc.RepSource = Rep
    Set mRep = Rep
  Case UCase("RepCustmStkLand")
    Set adc.RepSource = Rep1
    Set mRep = Rep1
  End Select
  '*** Jay 3.1.1 (All Rates Rpt)
  '*** Added Again In 3.2.0
  Call DispCoNm
  Call SetMwName(mRep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
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
    mRep.BoxHd.Suppress = True: mRep.BoxDet.Suppress = True
    mRep.BoxFt.Suppress = True: mRep.DetSec2.Suppress = True
  End If
  
  If (adc("UoStnClsFr") <> "" Or adc("UoStnClsTo") <> "") And ((adc("UoRmCtgFr") <> adc("UoRmCtgTo")) Or (adc("UoRmCtgFr") = "")) Then
      DispMsg "Please select an RmCtg to provide a filter of Stone Class", etError: Exit Sub
  End If
  
  '*** (Jen 3.2.0)
  If wDetPos = 0 Or adc("UoYN4") = "N" Then
    mRep.SecBestUse.Suppress = True
    mRep.hqOrdNo.Suppress = True: mRep.hTzDmCd.Suppress = True: mRep.hTzRmCd.Suppress = True
    mRep.hTzRmSzFr.Suppress = True: mRep.hTzRmSzTo.Suppress = True
    mRep.FtSecBestUse.Suppress = True
  Else
    mRep.DetSec2.Suppress = True
  End If
  '*** (Jen 3.2.0)
 
  ' Zubin 213
  Dim ws_BaseCurCd As String, ws_Join As String, wMatchRt As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                "HCoCd= '" + adc("UoCoCdFr") + "'")
  wBaseFct = " (case when TMulDiv= 'M' then 1/TCnvRt else TCnvRt end) "
  wReFct = " (case when TMulDivRs= 'D' and convert (decimal (16, 4), TCnvRtRs)> 0 then 1/ TCnvRtRs else TCnvRtRs end) "
  
  If adc("UoCurCdFr") <> "" Then
    wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "TCurrCd='" + adc("UoCurCdFr") + "'"
    '*** Jay 3.1.0_1 [Derived Value]
    'wPurStr = "B.TzPurRt as TzPurRt,"
    wFct = "1"
    '*** Jay 3.1.0_1 [Derived Value]
    wMatchRt = " IsNull(IzRmZRt, 0) "
  Else
    '*** Jay 3.1.0_1 [Derived Value]
    'wPurStr = " (case when TMulDiv= 'M' then B.TzPurRt/TCnvRt else B.TzPurRt*TCnvRt end) as TzPurRt, "
    wFct = " (case when TMulDiv= 'M' then 1/TCnvRt else TCnvRt end) "
    '*** Jay 3.1.0_1 [Derived Value]
    wMatchRt = " IsNull((case when InMulDiv= 'M' then IzRmZRt/InCnvFct else IzRmZRt*InCnvFct end), 0) "
  End If
  ws_Join = " join (Select PMCd as qPMCd, (case when " + IIF(adc("UoCurCdFr") = "", "1", "0") + " = 1 then 'Y' " + _
            "else 'N' end) as qBaseCur from Param where PTyp= 'CURNCY') Cur on qPMCd= TCurrCd "
  ' Zubin 213
 
  '*** Jay 3.1.0_1 [Derived Value]
  If UCase(adc("UoYn3")) = "Y" Then
    ' ***** Bef 3.03 - Manali - 27/06/08
    'GRP_REP.SetFormula mRep, "wStkRt", "If {rdo.qBaseWt}= 0 Then {rdo.TzPurRt} " + _
                    "Else ((({rdo.TzPurAmt}/ (if {rdo.qBaseWt}> 0 then {rdo.qBaseWt} else 1))* {rdo.qBaseWt})* (1+ ({rdo.qAddAmt}/{rdo.TNetAmt})))/ {rdo.qBaseWt}"
                    
    GRP_REP.SetFormula mRep, "wBalAmt", "(({rdo.TzPurAmt}/ (if {rdo.qBaseWt}> 0 then {rdo.qBaseWt} else 1))* {rdo.qBalWt})*(1+ ({rdo.qAddAmt}/{rdo.TNetAmt}))"
    
    'GRP_REP.SetFormula Rep, "wStkRt", "If {rdo.qBaseWt}= 0 Then {rdo.TzPurRt} " + _
    '               "Else (({rdo.TzPurRt}* {rdo.qBaseWt})* (1+ ({rdo.qAddAmt}/{rdo.TNetAmt})))/ {rdo.qBaseWt}"
    'GRP_REP.SetFormula Rep, "wBalAmt", "({rdo.TzPurRt}* {rdo.qBalWt})*(1+ ({rdo.qAddAmt}/{rdo.TNetAmt}))"
    '*** Jay 3.1.1 (All Rates Rpt)
    If UCase(adc.MenuCd) = UCase("RepCustmStkLand") Then
      
      ' ***** Bef 3.03 - Manali - 27/06/08
      'GRP_REP.SetFormula mRep, "wStkBaseRt", "If {rdo.qBaseWt}= 0 Then {rdo.qTzPurBaseRt} " + _
                      "Else ((({rdo.qTzPurBaseAmt}/ (if {rdo.qBaseWt}> 0 then {rdo.qBaseWt} else 1))* {rdo.qBaseWt})* (1+ ({rdo.qAddBaseAmt}/{rdo.qTNetBaseAmt})))/ {rdo.qBaseWt}"
                      
      GRP_REP.SetFormula mRep, "wBalBaseAmt", "(({rdo.qTzPurBaseAmt}/ (if {rdo.qBaseWt}> 0 then {rdo.qBaseWt} else 1))* {rdo.qBalWt})*(1+ ({rdo.qAddBaseAmt}/{rdo.qTNetBaseAmt}))"
      
      ' ***** Bef 3.03 - Manali - 27/06/08
      'GRP_REP.SetFormula mRep, "wStkReRt", "If {rdo.qBaseWt}= 0 Then {rdo.qTzPurReRt} " + _
                      "Else ((({rdo.qTzPurReAmt}/ (if {rdo.qBaseWt}> 0 then {rdo.qBaseWt} else 1))* {rdo.qBaseWt})* (1+ ({rdo.qAddReAmt}/{rdo.qTNetReAmt})))/ {rdo.qBaseWt}"
      
      GRP_REP.SetFormula mRep, "wBalReAmt", "(({rdo.qTzPurReAmt}/ (if {rdo.qBaseWt}> 0 then {rdo.qBaseWt} else 1))* {rdo.qBalWt})*(1+ ({rdo.qAddReAmt}/{rdo.qTNetReAmt}))"
    End If
    '*** Jay 3.1.1 (All Rates Rpt)
  Else
    ' ***** Bef 3.03 - Manali - 27/06/08
    'GRP_REP.SetFormula mRep, "wStkRt", "{rdo.TzPurRt}"
    
    GRP_REP.SetFormula mRep, "wBalAmt", "({rdo.TzPurRt}* {rdo.qBalWt})"
    '*** Jay 3.1.1 (All Rates Rpt)
    If UCase(adc.MenuCd) = UCase("RepCustmStkLand") Then
      
      ' ***** Bef 3.03 - Manali - 27/06/08
      'GRP_REP.SetFormula mRep, "wStkBaseRt", "{rdo.qTzPurBaseRt}"
      
      GRP_REP.SetFormula mRep, "wBalBaseAmt", "{rdo.qTzPurBaseRt}* {rdo.qBalWt}"
      
      ' ***** Bef 3.03 - Manali - 27/06/08
      'GRP_REP.SetFormula mRep, "wStkReRt", "{rdo.qTzPurReRt}"
      
      GRP_REP.SetFormula mRep, "wBalReAmt", "{rdo.qTzPurReRt}* {rdo.qBalWt}"
    End If
    '*** Jay 3.1.1 (All Rates Rpt)
  End If
  '*** Jay 3.1.0_1 [Derived Value]
  
  ' ***** Manali 3.03 - 17/06/08 - Derived Rate Y/N
  If adc("UoYN6") = "Y" Then
      GRP_REP.SetFormula mRep, "wStkRt", "If {rdo.qBaseWt}= 0 Then {rdo.TzPurRt} " + _
                    "Else ((({rdo.TzPurAmt}/ (if {rdo.qBaseWt}> 0 then {rdo.qBaseWt} else 1))* {rdo.qBaseWt})* (1+ ({rdo.qAddAmt}/{rdo.TNetAmt})))/ {rdo.qBaseWt}"
       
      If UCase(adc.MenuCd) = UCase("RepCustmStkLand") Then
          GRP_REP.SetFormula mRep, "wStkBaseRt", "If {rdo.qBaseWt}= 0 Then {rdo.qTzPurBaseRt} " + _
                      "Else ((({rdo.qTzPurBaseAmt}/ (if {rdo.qBaseWt}> 0 then {rdo.qBaseWt} else 1))* {rdo.qBaseWt})* (1+ ({rdo.qAddBaseAmt}/{rdo.qTNetBaseAmt})))/ {rdo.qBaseWt}"
          GRP_REP.SetFormula mRep, "wStkReRt", "If {rdo.qBaseWt}= 0 Then {rdo.qTzPurReRt} " + _
                      "Else ((({rdo.qTzPurReAmt}/ (if {rdo.qBaseWt}> 0 then {rdo.qBaseWt} else 1))* {rdo.qBaseWt})* (1+ ({rdo.qAddReAmt}/{rdo.qTNetReAmt})))/ {rdo.qBaseWt}"
      End If
  Else
      GRP_REP.SetFormula mRep, "wStkRt", "{rdo.TzPurRt}"
      
      If UCase(adc.MenuCd) = UCase("RepCustmStkLand") Then
          GRP_REP.SetFormula mRep, "wStkBaseRt", "{rdo.qTzPurBaseRt}"
          GRP_REP.SetFormula mRep, "wStkReRt", "{rdo.qTzPurReRt}"
      End If
  End If
  ' ***** Manali 3.03 - 17/06/08 - Derived Rate Y/N

  wBalWt = "": wAllPndCnd = ""
  
  If adc("UoAsOfDt") = "01/01/80" Then
    wBalWt = " (B.TzRmWt- B.TzRmDespWt) as qBalWt "
    ' ******** Manali 350Nxt
    If adc("UoYN") = "Y" Then wAllPndCnd = " and (B.TzCls='N' and Round(B.TzRmWt- B.TzRmDespWt, 3)> 0) "
    
'    If adc("UoYN") = "Y" Then wAllPndCnd = " and /*(B.TzCls='N' and Round(B.TzRmWt- B.TzRmDespWt, 3)> 0)*/ " + _
'        "(B.TzCls='N' or Round(B.TzRmWt- B.TzRmDespWt, 3)> 0) "
        '****** Manali 3.2.2 only open voucher allowed ****** Reverted (/*(B.TzCls='N' and Round(B.TzRmWt- B.TzRmDespWt, 3)> 0)*/ )
        ' ***** Manali 3.2.2 - 08/10/08 - the above said change is reverted back to its original status, change in sql condition is commented
  Else
   ' wBalWt = " B.TzRmWt- IsNull((Select sum(Tz.TzRmWt) from TxndZ Tz " + _
             " join InvHd on InCoCd= TzCoCd and InTc= TzTc and " + _
             " InYy= TzYy and InChr= TzChr and InNo= TzNo where " + _
             " Tz.TzCoCd= B.TzCoCd and Tz.TzBM= 'M' and Tz.TzCoCd= B.TzCoCd and " + _
             " Tz.TzRefYy= B.TzYy and Tz.TzRefKey= B.TzKey and " + _
             " InDt<= '" + CStr(adc("UoAsOfDt")) + "'), 0) as qBalWt "
    'Geeta***Emr210
    '****** Sachin 3.02 - Id fields in Joins
    wBalWt = " B.TzRmWt- " + IIF(gs_Partition = ctCurrPrtn, "B.TzRmDespPurgeWt", "0") + " - IsNull((Select sum(Tz.TzRmWt) from TxndZ Tz " + _
             " join InvHd on InIdNo=TzInIdNo and InCoCd=TzCoCd and InTc= TzTc and " + _
             " InYy= TzYy and InChr= TzChr and InNo= TzNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=TzPrtKey ", "") + " where " + _
             " Tz.TzCoCd= B.TzCoCd and Tz.TzBM= 'M' and Tz.TzCoCd= B.TzCoCd and " + _
             " Tz.TzRefYy= B.TzYy and Tz.TzRefKey= B.TzKey and " + _
             " InDt<= '" + CStr(adc("UoAsOfDt")) + "'" + IIF(gs_Partition = ctCurrPrtn, " And Tz.TzPrtKey='" + ctCurrPrtn + "'", "") + "), 0) as qBalWt "
             
             
    If adc("UoYN") = "Y" Then
      '*** (Bef 2.14 Next)
      'wAllPndCnd = " and Round(B.TzRmWt- IsNull((Select sum(Tz.TzRmWt) from TxndZ Tz " + _
                   " join InvHd on InCoCd= TzCoCd and InTc= TzTc and InYy= TzYy and " + _
                   " InChr= TzChr and InNo= TzNo where " + _
                   " Tz.TzBM= 'M' and Tz.TzCoCd= B.TzCoCd and Tz.TzRefYy= B.TzYy and " + _
                   " Tz.TzRefKey= B.TzKey and InDt<= '" + CStr(adc("UoAsOfDt")) + "'), 0), 3)> 0 "
      '*** (Bef 2.14 Next)
      
      '*** (Jen 2.14 Next)
      '****** Sachin 3.02 - Id fields in Joins
      '****** Manali 3.2.2 - only open voucher allowed
      '****** Manali 3.2.2 - 08/10/08 - above said change Reverted back (/*and B.TzCls= 'N'*/)
      ' ***** Manali 350Nxt
      ' ***** Manali 3.6.2 - TzClsDt clause changed - (Sangeeta)
      wAllPndCnd = " and Round(B.TzRmWt- " + IIF(gs_Partition = ctCurrPrtn, "B.TzRmDespPurgeWt", "0") + "- IsNull((Select sum(Tz.TzRmWt) from TxndZ Tz " + _
                   " join InvHd on InIdNo=TzInIdNo and InCoCd= TzCoCd and InTc= TzTc and InYy= TzYy and " + _
                   " InChr= TzChr and InNo= TzNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=TzPrtKey ", "") + " where " + _
                   " Tz.TzBM= 'M' and Tz.TzCoCd= B.TzCoCd and Tz.TzRefYy= B.TzYy and " + _
                   " Tz.TzRefKey= B.TzKey and InDt<= '" + CStr(adc("UoAsOfDt")) + "'" + IIF(gs_Partition = ctCurrPrtn, " And Tz.TzPrtKey='" + ctCurrPrtn + "'", "") + "), 0), 3)> 0 " + _
                  " and (B.TzClsDt='01/01/80' Or B.TzClsDt>'" + CStr(adc("UoAsOfDt")) + "' ) "
                  
                  ' Bef 3.6.2
                  ' " and (B.TzCls= 'N' Or B.TzClsDt>'" + CStr(adc("UoAsOfDt")) + "' ) "
      '*** (Jen 2.14 Next)
      
    End If
    
  End If
  
'''  ' Zubin 213
'''  Dim ws_BaseCurCd As String, ws_Join As String, wMatchRt As String
'''  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
'''                "HCoCd= '" + adc("UoCoCdFr") + "'")
'''
'''  If adc("UoCurCdFr") <> "" Then
'''    wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "TCurrCd='" + adc("UoCurCdFr") + "'"
'''    '*** Jay 3.1.0_1 [Derived Value]
'''    'wPurStr = "B.TzPurRt as TzPurRt,"
'''    wFct = "1"
'''    '*** Jay 3.1.0_1 [Derived Value]
'''    wMatchRt = " IsNull(IzRmZRt, 0) "
'''  Else
'''    '*** Jay 3.1.0_1 [Derived Value]
'''    'wPurStr = " (case when TMulDiv= 'M' then B.TzPurRt/TCnvRt else B.TzPurRt*TCnvRt end) as TzPurRt, "
'''    wFct = " (case when TMulDiv= 'M' then 1/TCnvRt else TCnvRt end) "
'''    '*** Jay 3.1.0_1 [Derived Value]
'''    wMatchRt = " IsNull((case when InMulDiv= 'M' then IzRmZRt/InCnvFct else IzRmZRt*InCnvFct end), 0) "
'''  End If
'''  ws_Join = " join (Select PMCd as qPMCd, (case when " + IIF(adc("UoCurCdFr") = "", "1", "0") + " = 1 then 'Y' " + _
'''            "else 'N' end) as qBaseCur from Param where PTyp= 'CURNCY') Cur on qPMCd= TCurrCd "
'''  ' Zubin 213
  
  '*** (Jen 3.2.0)
  '''*** and not ( (RepFr> 0 and FldTo> 0 and FldTo< RepFr)
  '''         or (RepTo> 0 and FldFr> RepTo)
  '''         )
  'wCnd = wCnd + " and not (   (convert (decimal (7, 4), " + CStr(adc("UoLn1Fr")) + ")> 0 " + _
  '              "             and convert (decimal (7, 4), B.TzRmSzTo)> 0 " + _
  '              "             and convert (decimal (7, 4), B.TzRmSzTo)< " + CStr(adc("UoLn1Fr")) + _
  '              "             ) " + _
  '              "         or  (convert (decimal (7, 4), " + CStr(adc("UoLn1To")) + ")> 0 " + _
  '              "             and convert (decimal (7, 4), B.TzRmSzFr)> " + CStr(adc("UoLn1To")) + _
  '              "             ) " + _
  '              "         ) "
  '*** (Jen 3.2.0)
    
  '****** Sachin 3.02.0 - Additional Fields in Where Clause
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " B.TzRmCtg >= '" + Trim(adc("UORMCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGTO")) <> "", " B.TzRmCtg <= '" + Trim(adc("UORMCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGSEL")) <> "", " B.TzRmCtg In (" + Trim(adc("UORMCTGSEL")) + ")", "")

  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " B.TzRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " B.TzRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " B.TzRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  '****** Sachin 3.02.0 - Additional Fields in Where Clause
  
  If adc("UoYN1") = "Y" Then
    '*** Added Again In 3.2.0
    '****** (Jen 3.1.1 added jon for STNCLS) ******
    ' Zubin 213 (wMatchRt and ws_Join added)
    'IsNull(IzRmZRt, 0) (Zubin - Bef 213)
    '****** Sachin 3.02 - Id fields added in joins
    'Jm.1- Inv date added
    'JewelSmith.2 ',' + SCls.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SCls.PSCd = RmStnCls
    wStr = "IsNull(M.TzTc, '') as qMatchTc, " + _
           "IsNull(M.TzYy, '') as qMatchYy, IsNull(M.TzChr, '') as qMatchChr, IsNull(M.TzNo, 0) as qMatchNo, " + _
           "IsNull(M.TzSr, 0) as qMatchSr, IsNull(M.TzSrNo, 0) as qMatchSrNo, IsNull(InExpNo, '') as InExpNo, InDt, " + _
           "IsNull(M.TzRmWt, 0) as qMatchWt, " + wMatchRt + " as qMatchRt, IsNull(InCmCd, '') as InCmCd " + _
           "From TxndZ B join Txn On TIdNo=B.TzTIdNo and TCoCd= B.TzCoCd and TTc= B.TzTc and TYy= B.TzYy and " + _
           "TChr= B.TzChr and TNo= B.TzNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey=B.TzPrtKey ", "") + ws_Join + " join RmMst on RmCd= B.TzRmZCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=B.TzPrtKey ", "") + _
           "left outer join TxndZ M on " + IIF(gs_Partition = ctCurrPrtn, " M.TzPrtKey=B.TzPrtKey and ", "") + " M.TzCoCd= B.TzCoCd " + _
           "and M.TzRefYy= B.TzYy and M.TzRefKey= B.TzKey and M.TzBM= 'M' " + _
           IIF(adc("UoAsOfDt") = "01/01/80", " ", " and " + _
           "IsNull((Select InDt from InvHd where InIdNo=M.TzInIdNo and InCoCd= M.TzCoCd and InTc= M.TzTc and InYy= M.TzYy and " + _
           "InChr= M.TzChr and InNo= M.TzNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=M.TzPrtKey ", "") + " ), '01/01/80')<= '" + CStr(adc("UoAsOfDt")) + "' ") + _
           IIF(gs_Partition = ctCurrPrtn, " And M.TzPrtKey= '" + ctCurrPrtn + "' ", " ") + _
           "left outer join InvZRm on IzCoCd= M.TzCoCd and IzTc= M.TzTc " + _
           "and IzYy= M.TzYy and IzChr= M.TzChr and IzNo= M.TzNo and IzSr= M.TzSr " + IIF(gs_Partition = ctCurrPrtn, " And IzPrtKey=M.TzPrtKey ", "") + _
           "left outer join InvHd  on  InCoCd= IzCoCd and InTc= IzTc and InYy= IzYy " + _
           "and InChr= IzChr and InNo=IzNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IzPrtKey ", "") + _
           "Left Outer Join Param SCls on SCls.PTyp= 'STNCLS' and SCls.PmCd= RmCtg and SCls.PSCd = RmStnCls " + _
           "where B.TzBM= 'B' " + IIF(gs_Partition = ctCurrPrtn, " And B.TzPrtKey='" + ctCurrPrtn + "' ", "")
           
    wStr2 = "'' as qMatchTc, '' as qMatchYy, '' as qMatchChr, 0 as qMatchNo, " + _
           "0 as qMatchSr, 0 as qMatchSrNo, '' as InExpNo, '' as InDt, 0 as qMatchWt, " + _
           "0 as qMatchRt, '' as InCmCd " + _
           "From TxndZ B join Txn On TIdNo=B.TzTIdNo and TCoCd= B.TzCoCd and TTc= B.TzTc and   TYy= B.TzYy and " + _
           "TChr= B.TzChr and TNo= B.TzNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey= B.TzPrtKey ", "") + _
           ws_Join + "join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=B.TzPrtKey and ", "") + " RmCd= B.TzRmZCd " + _
           "Left Outer Join Param SCls on SCls.PTyp= 'STNCLS' and SCls.PmCd= RmCtg and SCls.PSCd = RmStnCls " + _
           "where B.TzBM= 'B' "
'    wStr = "IsNull(M.TzTc, ''), " + _
'           "IsNull(M.TzYy, ''), IsNull(M.TzChr, ''), IsNull(M.TzNo, 0), " + _
'           "IsNull(M.TzSr, 0), IsNull(M.TzSrNo, 0), IsNull(InExpNo, ''), " + _
'           "IsNull(M.TzRmWt, 0), IsNull(IzRmZRt, 0) as qMatchRt, IsNull(InCmCd, '') " + _
'           "From TxndZ B join Txn On TTc= B.TzTc and TYy= B.TzYy and TChr= B.TzChr and " + _
'           "TNo= B.TzNo join RmMst on RmCd= B.TzRmZCd " + _
'           "left outer join TxndZ M (Index= xTzRefKey) on M.TzRefYy= B.TzYy " + _
'           "and M.TzRefKey= B.TzKey and M.TzBM= 'M' " + _
'           "and IsNull((Select InDt from InvHd where InTc= M.TzTc and InYy= M.TzYy " + _
'           "and InChr= M.TzChr and InNo= M.TzNo ), '01/01/80')<= '" + CStr(adc("UoAsOfDt")) + "' " + _
'           "left outer join InvZRm on IzTc= M.TzTc and IzYy= M.TzYy and IzChr= M.TzChr " + _
'           "and IzNo= M.TzNo and IzSr= M.TzSr " + _
'           "left outer join InvHd on InTc= IzTc and InYy= IzYy and InChr= IzChr " + _
'           "and InNo= IzNo where B.TzBM= 'B' "
  Else
    '*** Added Again In 3.2.0
    '****** (Jen 3.1.1 added jon for STNCLS) ******
    ' Zubin 213 (ws_Join added)
    '****** Sachin 3.02 - Id fields in Joins
    'Jm.1- Inv date added
    'JewelSmith.2 ',' + SCls.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SCls.PSCd = RmStnCls
    wStr = "'' as qMatchTc, '' as qMatchYy, TzChr as qMatchChr, 0 as qMatchNo, " + _
           "0 as qMatchSr, 0 as qMatchSrNo, '' as InExpNo, '' as InDt, 0 as qMatchWt, " + _
           "0 as qMatchRt, '' as InCmCd " + _
           "From TxndZ B join Txn On TIdNo=B.TzTIdNo and TCoCd= B.TzCoCd and TTc= B.TzTc and   TYy= B.TzYy and " + _
           "TChr= B.TzChr and TNo= B.TzNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey= B.TzPrtKey ", "") + _
           ws_Join + "join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=B.TzPrtKey and ", "") + " RmCd= B.TzRmZCd " + _
           "Left Outer Join Param SCls on SCls.PTyp= 'STNCLS' and SCls.PmCd= RmCtg and SCls.PSCd = RmStnCls " + _
           "where B.TzBM= 'B' "
  End If
  '******************Geeta*********Emr207*********
    ' Zubin 213 (Before 213)
'    If ADC("UoYN2") <> "" Then
'      wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "TUSDYn='" + ADC("UoYN2") + "'"
'      wPurStr = "B.TzPurRt as TzPurRt,"
'    Else
'       wPurStr = " (case when TUSDYn='N' then B.TzPurRt/TCnvRt else B.TzPurRt end ) as TzPurRt, "
'    End If

  '*** Jay 3.1.1 (All Rates Rpt)
  wRepRtFlds = "B.TzPurRt * " + wFct + " as TzPurRt, B.TzPurAmt * " + wFct + " as TzPurAmt, " + _
             "TNetAmt * " + wFct + " as TNetAmt, ((TTotAmt- TNetAmt)* " + wFct + ") as qAddAmt, "
  If UCase(adc.MenuCd) = UCase("RepCustmStkLand") Then
    wRepRtFlds = wRepRtFlds + " B.TzPurRt * " + wBaseFct + " as qTzPurBaseRt, B.TzPurAmt * " + wBaseFct + " as qTzPurBaseAmt, " + _
             "TNetAmt * " + wBaseFct + " as qTNetBaseAmt, ((TTotAmt- TNetAmt)* " + wBaseFct + ") as qAddBaseAmt, " + _
             "B.TzPurRt * " + wReFct + " as qTzPurReRt, B.TzPurAmt * " + wReFct + " as qTzPurReAmt, " + _
             "TNetAmt * " + wReFct + " as qTNetReAmt, ((TTotAmt- TNetAmt)* " + wReFct + ") as qAddReAmt, "
  End If
  '*** Jay 3.1.1 (All Rates Rpt)
  
  '************Geeta****************Emr207*******************
  '*********************Replaced  B.TzPurRt with  wPurStr*********
  '****** (Jen 3.2.0) added B.TzOdTc+'/'+B.TzOdYy+'/'+B.TzOdChr+'/'+LTrim(Str(B.TzOdNo)) as qOrdNo, B.TzDmCd as TzDmCd, B.TzRmCd as TzRmCd, B.TzRmSzFr as TzRmSzFr, B.TzRmSzTo as TzRmSzTo
  '****** (Jen 3.2.0) also added condition for Best Use (UoYN5)
  '*** Jay 3.1.0_1 (added: TzPurAmt,TNetAmt, (TTotAmt- TNetAmt) as qAddAmt, )
  '6.1
  wSqlStrg = "Set Dateformat " + DtFmtStr() + " Select " + grpflds + ", " + _
           "B.TzTc as qBaseTc, B.TzYy as qBaseYy, B.TzChr as qBaseChr, B.TzNo as qBaseNo, B.TzSr as qBaseSr, " + _
           "(case when B.TzOdTc= '' then '' else B.TzOdTc+'/'+B.TzOdYy+'/'+B.TzOdChr+'/'+LTrim(Str(B.TzOdNo))+'/'+LTrim(Str(B.TzOdSr)) end) as qOrdNo, " + _
           "B.TzDmCd as TzDmCd, B.TzRmCd as TzRmCd, B.TzRmSzFr as TzRmSzFr, B.TzRmSzTo as TzRmSzTo, TDt, TBOENo, TBOEDt, " + _
           "TCmCd, TSuppCd, TBillNo, RmCtg, RmSCtg, B.TzRmZCd, B.TzPcPerCt, B.TzRmWt as qBaseWt, " + IIF(gs_Partition = ctCurrPrtn, "B.TzRmDespPurgeWt", "0") + " as TzRmDespPurgeWt, " + _
           wBalWt + ", " + wRepRtFlds + wStr + wAllPndCnd + wCnd + wImpCnd + _
           IIF(adc("UoYN5") = "", " ", _
              IIF(adc("UoYN5") = "Y", _
                  " and not (TCmCd= '' and B.TzOdNo= 0 and B.TzOdSr= 0 and B.TzDmCd= '' and B.TzRmCd= '' and B.TzRmSzFr= 0 and B.TzRmSzTo= 0) ", _
                  " and (TCmCd= '' and B.TzOdNo= 0 and B.TzOdSr= 0 and B.TzDmCd= '' and B.TzRmCd= '' and B.TzRmSzFr= 0 and B.TzRmSzTo= 0) "))
           
  If UCase(adc.MenuCd) = UCase("RepCustmStk") Then
     If adc("UoYN1") = "Y" Then
     '6.1
      wSqlStrg = "Set Dateformat " + DtFmtStr() + " Select " + grpflds + ",'B' as qTag, " + _
           "B.TzTc as qBaseTc, B.TzYy as qBaseYy, B.TzChr as qBaseChr,B.TzNo as qBaseNo, B.TzSr as qBaseSr, " + _
           "(case when B.TzOdTc= '' then '' else B.TzOdTc+'/'+B.TzOdYy+'/'+B.TzOdChr+'/'+LTrim(Str(B.TzOdNo))+'/'+LTrim(Str(B.TzOdSr)) end) as qOrdNo, " + _
           "B.TzDmCd as TzDmCd, B.TzRmCd as TzRmCd, B.TzRmSzFr as TzRmSzFr, B.TzRmSzTo as TzRmSzTo, TDt, TBOENo, TBOEDt, " + _
           "TCmCd, TSuppCd, TBillNo, RmCtg, RmSCtg, B.TzRmZCd,B.TzPcPerCt, B.TzRmWt as qBaseWt, " + IIF(gs_Partition = ctCurrPrtn, "B.TzRmDespPurgeWt", "0") + " as TzRmDespPurgeWt, " + _
           wBalWt + ", " + wRepRtFlds + wStr2 + wAllPndCnd + wCnd + wImpCnd + _
           IIF(adc("UoYN5") = "", " ", _
              IIF(adc("UoYN5") = "Y", _
                  " and not (TCmCd= '' and B.TzOdNo= 0 and B.TzOdSr= 0 and B.TzDmCd= '' and B.TzRmCd= '' and B.TzRmSzFr= 0 and B.TzRmSzTo= 0) ", _
                  " and (TCmCd= '' and B.TzOdNo= 0 and B.TzOdSr= 0 and B.TzDmCd= '' and B.TzRmCd= '' and B.TzRmSzFr= 0 and B.TzRmSzTo= 0) "))
      
      wSqlStrg = wSqlStrg + " union all Select " + grpflds + ",'M' as qTag, " + _
           "B.TzTc as qBaseTc, B.TzYy as qBaseYy, B.TzChr as qBaseChr,B.TzNo as qBaseNo, B.TzSr as qBaseSr, " + _
           "'' as qOrdNo, '' as TzDmCd, '' as TzRmCd, 0.0 as TzRmSzFr, 0.0 as TzRmSzTo, TDt, TBOENo, TBOEDt, " + _
           "TCmCd, TSuppCd, TBillNo, RmCtg, RmSCtg, '' as TzRmZCd, 0 as TzPcPerCt, 0 as qBaseWt, " + IIF(gs_Partition = ctCurrPrtn, "B.TzRmDespPurgeWt", "0") + " as TzRmDespPurgeWt, " + _
           "0 as qBalWt , 0 as TzPurRt, 0 as TzPurAmt, 0 as TNetAmt, 0 as qAddAmt, " + wStr + wAllPndCnd + wCnd + wImpCnd + _
           IIF(adc("UoYN5") = "", " ", _
              IIF(adc("UoYN5") = "Y", _
                  " and not (TCmCd= '' and B.TzOdNo= 0 and B.TzOdSr= 0 and B.TzDmCd= '' and B.TzRmCd= '' and B.TzRmSzFr= 0 and B.TzRmSzTo= 0) ", _
                  " and (TCmCd= '' and B.TzOdNo= 0 and B.TzOdSr= 0 and B.TzDmCd= '' and B.TzRmCd= '' and B.TzRmSzFr= 0 and B.TzRmSzTo= 0) ")) + _
           " order by qTag "
     End If
   End If
   
  'wSqlStrg = "Set Dateformat DMY Select " + grpflds + ", " + _
           "B.TzTc as qBaseTc, B.TzYy as qBaseYy, B.TzChr as qBaseChr, B.TzNo as qBaseNo, B.TzSr as qBaseSr, " + _
           "TDt, TBOENo, TBOEDt, TSuppCd, " + _
           "TBillNo, RmCtg, RmSCtg, B.TzRmZCd, B.TzPcPerCt, B.TzRmWt as qBaseWt, B.TzRmDespPurgeWt, " + _
           wBalWt + ", " + wPurStr + " B.TzPurAmt, TNetAmt, (TTotAmt- TNetAmt) as qAddAmt, " + wStr + wAllPndCnd + wCnd + wImpCnd '**** urmi ****'
  '*** Jay 3.1.0_1 (added: TzPurAmt,TNetAmt, (TTotAmt- TNetAmt) as qAddAmt, )
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  'Sachin 3.01 ctMaxDopOpt added
  
  If adc("UOASOFDT") <> MWLib.EmptyDate Then _
    mRep.TxtAsOfDt.SetText "As of Date " + CStr(adc("UOASOFDT"))
    ' Zubin 213 (Before 213)
'    If ADC("UoYN2") = "" Or ADC("UoYN2") = "Y" Then
'      mRep.TxtPRCurCd.SetText "Purchase Currency (US$)"
'    ElseIf ADC("UoYN2") = "N" Then
'      mRep.TxtPRCurCd.SetText "Purchase Currency (RS)"
'    End If
    ' Zubin 213
  If adc("UoCurCdFr") = "" Then
    mRep.TxtPRCurCd.SetText "Purchase Currency (" + ws_BaseCurCd + ")"
  Else
    mRep.TxtPRCurCd.SetText "Purchase Currency (" + adc("UoCurCdFr") + ")"
  End If
  '*** Jay 3.1.1 (All Rates Rpt)
  If UCase(adc.MenuCd) = UCase("RepCustmStkLand") Then
    mRep.txthBaseCurr.SetText "in " + ws_BaseCurCd
    mRep.txthReCurr.SetText "in " + ctIndCurncy
  End If
  '*** Jay 3.1.1 (All Rates Rpt)
  
  If adc("UoYN1") = "N" Then
    mRep.BoxHd.Suppress = True
    mRep.BoxDet.Suppress = True
    mRep.BoxFt.Suppress = True
    mRep.DetSec2.Suppress = True
  End If
  'Bombay.1 -if Show Rate&Value=N then will not show rate and value fields
  If adc("UoYN7") = "N" Then
    Rep.qDerivedStkRt1.Suppress = True: Rep.hTzPurRt.Suppress = True
    Rep.FldBalAmt.Suppress = True
    Rep.g3BalAmt.Suppress = True: Rep.g2BalAmt.Suppress = True
    Rep.g1BalAmt.Suppress = True: Rep.gBalAmt.Suppress = True
    Rep.hBalAmt.Suppress = True: Rep.hGrpBalAmt.Suppress = True
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula mRep, "wRmCdFontSz", GetRmcdFontSz("{rdo.TzRmCd}", 7)
  GRP_REP.SetFormula mRep, "wRmCdFontSz", GetRmcdFontSz("{rdo.TzRmZCd}", 8)
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
    
Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
Select Case Index
Case Is = 0         ' Tag = ShowPnd
  If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
Case Is = 1         ' Tag = ShowBoxDet
  If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
Case Is = 4         ' Tag = ShowLine
  If .Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"
'Bombay.1 - check to show rate & value
Case Is = 2         ' Tag = ShowRateVal
  If .Value = Checked Then adc("UoYN7") = "Y" Else adc("UoYN7") = "N"
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
    Set adc.FirNKeyCtl = adc("UoTTcFr") 'adc("UoCoCdFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_TxnTcTyp = "TZ"
    ' **** Bef 3.03 - Manali
    'gs_CmCtg = "S"
    
    'URMI 2.04-3 hELP OF RmCd should show only ZRmCd
    gb_RmZ = True
    gb_RmZ1 = False '*** (Jen 3.2.0)
    gb_RmCtgFor = False
    
    gs_PTyp = "ImpTyp"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
    'Bombay.1 -Show Rate&Value check box not shown for Custom stock Landscape report
    If UCase(adc.MenuCd) = UCase("RepCustmStkLand") Then ChkBoxArr(2).Visible = False
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_PTyp = "ImpTyp"
  
  'URMI 2.04-3 hELP OF RmCd should show only ZRmCd
  gb_RmZ = True
  gb_RmZ1 = False '*** (Jen 3.2.0)
  gb_RmCtgFor = False
  
  ' **** Manali 3.03 - 30/06/08
  Select Case UCase(IdName)
    Case Is = UCase("UoCmCdFr"), UCase("UoCmCdTo"), UCase("UoCmCdSel")
      gs_CmCtg = "S"
    Case Is = UCase("UoExpCmCdFr"), UCase("UoExpCmCdTo"), UCase("UoExpCmCdSel")
      gs_CmCtg = "C"
  End Select
  ' **** Manali 3.03 - 30/06/08
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_PTyp = "ImpTyp"
  
  'URMI 2.04-3 hELP OF RmCd should show only ZRmCd
  gb_RmZ = True
  gb_RmZ1 = False '*** (Jen 3.2.0)
  gb_RmCtgFor = False
  
  ' **** Manali 3.03 - 30/06/08
  Select Case UCase(IdName)
    Case Is = UCase("UoCmCdFr"), UCase("UoCmCdTo"), UCase("UoCmCdSel")
      gs_CmCtg = "S"
    Case Is = UCase("UoExpCmCdFr"), UCase("UoExpCmCdTo"), UCase("UoExpCmCdSel")
      gs_CmCtg = "C"
  End Select
  ' **** Manali 3.03 - 30/06/08
  
  SetProp Me, IdName, When
End Sub

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptBestUse_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UOYN5") = "Y"
Case Is = 1
 adc("UOYN5") = "N"
Case Is = 2
 adc("UOYN5") = ""
End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added
Private Sub OptBestUse_GotFocus(Index As Integer)
DispMsg FraOptBestUse.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptDerRt_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UOYN6") = "Y"
Case Is = 1
 adc("UOYN6") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added
Private Sub OptDerRt_GotFocus(Index As Integer)
DispMsg FraOptDerRt.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptDerVal_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UOYN3") = "Y"
Case Is = 1
 adc("UOYN3") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptDerVal_GotFocus(Index As Integer)
DispMsg FraOptDerVal.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

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
  If adc("UoYN3") = "Y" Then OptDerVal(0).Value = True
  If adc("UoYN3") = "N" Then OptDerVal(1).Value = True
  If adc("UoYN6") = "Y" Then OptDerRt(0).Value = True
  If adc("UoYN6") = "N" Then OptDerRt(1).Value = True
  If adc("UoYN5") = "Y" Then
    OptBestUse(0).Value = True
  ElseIf adc("UoYN5") = "N" Then
    OptBestUse(1).Value = True
  ElseIf adc("UoYN5") = "" Then
    OptBestUse(2).Value = True
  End If
  ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  
  ' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  If adc("UoYN4") = "" Then adc("UoYN4") = "N"
  
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If adc("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  ' ****** Manali 3.5.0 - 24/11/08 - Check Boxes Added
  'Bombay.1 - setting default value as Y to show rate & value
  adc("UoYN7") = "Y"
  If adc("UoYN7") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
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
  
' ********* Manali 3.2.2 - Before 3.2.2
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  mRep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  mRep.wSrvrTime.SetText ws_DtTm

' ******** Manali 3.2.2 - User date and time in reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  Rep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  Rep1.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
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
