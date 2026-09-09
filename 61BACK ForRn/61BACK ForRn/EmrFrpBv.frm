VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpBv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bag Movement"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   13440
   ScaleWidth      =   25890
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   10830
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4365
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4965
      TabIndex        =   73
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   74
         Top             =   105
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
         TabIndex        =   75
         Top             =   120
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
      Left            =   9030
      TabIndex        =   72
      Top             =   9630
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   69
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
      Left            =   30
      TabIndex        =   65
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
      Height          =   9645
      Left            =   0
      TabIndex        =   68
      Top             =   60
      Width           =   15225
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   67
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpBv.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpBv.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   71
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   66
               Top             =   360
               Width           =   9825
               _ExtentX        =   17330
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   70
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraLoc 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1695
               TabIndex        =   117
               Top             =   590
               Width           =   4425
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
                  Left            =   50
                  TabIndex        =   119
                  ToolTipText     =   "Select Production or Process"
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   2085
               End
               Begin VB.OptionButton OptLoc 
                  Caption         =   "Process Loc"
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
                  Left            =   2560
                  TabIndex        =   118
                  ToolTipText     =   "Select Production or Process"
                  Top             =   0
                  Width           =   1545
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Addl Dets"
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
               Left            =   80
               TabIndex        =   62
               Tag             =   "ShowFgMvm"
               ToolTipText     =   "Click To Show Stock Rate Field"
               Top             =   6080
               Width           =   1860
            End
            Begin VB.Frame FraOptBagAck 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1740
               TabIndex        =   115
               ToolTipText     =   "Select if to show Acknowledged Bags / Unacknowledged bags/ all the Bags"
               Top             =   5160
               Width           =   5865
               Begin VB.OptionButton OptBagAck 
                  Caption         =   "Acknowledged"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Index           =   0
                  Left            =   0
                  TabIndex        =   57
                  ToolTipText     =   "Select if to show Acknowledged Bags / Unacknowledged bags/ all the Bags"
                  Top             =   0
                  Width           =   1995
               End
               Begin VB.OptionButton OptBagAck 
                  Caption         =   "UnAcknowledged"
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
                  Left            =   2130
                  TabIndex        =   58
                  ToolTipText     =   "Select if to show Acknowledged Bags / Unacknowledged bags/ all the Bags"
                  Top             =   0
                  Width           =   2055
               End
               Begin VB.OptionButton OptBagAck 
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
                  Left            =   4740
                  TabIndex        =   59
                  ToolTipText     =   "Select if to show Acknowledged Bags / Unacknowledged bags/ all the Bags"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4260
               TabIndex        =   5
               ToolTipText     =   "Enter Receiving Location To Range"
               Top             =   1170
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UORECLOCTO"
               IdName          =   "UORECLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6780
               TabIndex        =   6
               ToolTipText     =   "Enter Receiving Location Selection"
               Top             =   1170
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UORECLOCSEL"
               IdName          =   "UORECLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4260
               TabIndex        =   2
               ToolTipText     =   "Enter Issue Location To Range"
               Top             =   885
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6780
               TabIndex        =   3
               ToolTipText     =   "Enter Issue Location Selection"
               Top             =   885
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1740
               TabIndex        =   1
               ToolTipText     =   "Enter Issue Location From Range"
               Top             =   885
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1740
               TabIndex        =   4
               ToolTipText     =   "Enter Receiving Location From Range"
               Top             =   1170
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UORECLOCFR"
               IdName          =   "UORECLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   4260
               TabIndex        =   8
               ToolTipText     =   "Enter To Transaction Date"
               Top             =   1455
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
               Index           =   5
               Left            =   1740
               TabIndex        =   7
               ToolTipText     =   "Enter From Transaction Date"
               Top             =   1455
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
               Index           =   6
               Left            =   4260
               TabIndex        =   12
               ToolTipText     =   "Enter To Bag Year"
               Top             =   1740
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYTO"
               IdName          =   "UOBYYTO"
               CmpStr          =   "TdBYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1740
               TabIndex        =   9
               ToolTipText     =   "Enter From Bag Year"
               Top             =   1740
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYFR"
               IdName          =   "UOBYYFR"
               CmpStr          =   "TdBYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4260
               TabIndex        =   16
               ToolTipText     =   "Enter To Design Code "
               Top             =   2025
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
               Index           =   12
               Left            =   6780
               TabIndex        =   17
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2025
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "BOdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   1740
               TabIndex        =   15
               ToolTipText     =   "Enter From Design Code "
               Top             =   2025
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
               Left            =   4260
               TabIndex        =   43
               ToolTipText     =   "Enter To Design Category"
               Top             =   3735
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
               Index           =   31
               Left            =   6780
               TabIndex        =   44
               ToolTipText     =   "Enter Design Category Code Selection"
               Top             =   3735
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
               Index           =   32
               Left            =   1740
               TabIndex        =   42
               ToolTipText     =   "Enter From Design Category"
               Top             =   3735
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
               Index           =   33
               Left            =   4260
               TabIndex        =   46
               ToolTipText     =   "Enter To Design Production Category"
               Top             =   4020
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
               Index           =   34
               Left            =   6780
               TabIndex        =   47
               ToolTipText     =   "Enter Design Production Category Selection"
               Top             =   4020
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
               Index           =   35
               Left            =   1740
               TabIndex        =   45
               ToolTipText     =   "Enter From Design Production Category"
               Top             =   4020
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
               Left            =   4260
               TabIndex        =   49
               ToolTipText     =   "Enter To Karat"
               Top             =   4305
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOKTTO"
               IdName          =   "UOKTTO"
               CmpStr          =   "OdKt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   6780
               TabIndex        =   50
               ToolTipText     =   "Enter Karat Selection"
               Top             =   4305
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "OdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   1740
               TabIndex        =   48
               ToolTipText     =   "Enter From Karat"
               Top             =   4305
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               CmpStr          =   "OdKt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   4260
               TabIndex        =   52
               ToolTipText     =   "Enter To Customer"
               Top             =   4590
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
               Index           =   40
               Left            =   6780
               TabIndex        =   53
               ToolTipText     =   "Enter Customer Selection"
               Top             =   4590
               Width           =   7950
               _ExtentX        =   14023
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
               Left            =   1740
               TabIndex        =   51
               ToolTipText     =   "Enter From Customer"
               Top             =   4590
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
               Index           =   42
               Left            =   4260
               TabIndex        =   55
               ToolTipText     =   "Enter To Time"
               Top             =   4875
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               Mask            =   "#0.00"
               MaxLength       =   5
               DataType        =   5
               DataField       =   "UOMODTIMETO"
               IdName          =   "UOMODTIMETO"
               CmpStr          =   "Txnd.ModTime <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   1740
               TabIndex        =   54
               ToolTipText     =   "Enter From Time"
               Top             =   4875
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               Mask            =   "#0.00"
               MaxLength       =   5
               DataType        =   5
               DataField       =   "UOMODTIMEFR"
               IdName          =   "UOMODTIMEFR"
               CmpStr          =   "Txnd.ModTime >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   2160
               TabIndex        =   10
               ToolTipText     =   "Enter From Bag Character"
               Top             =   1740
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRFR"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "TdBChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   2940
               TabIndex        =   11
               ToolTipText     =   "Enter From Bag Number"
               Top             =   1740
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOFR"
               IdName          =   "UOBNOFR"
               CmpStr          =   "TdBNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   4680
               TabIndex        =   13
               ToolTipText     =   "Enter To Bag Character"
               Top             =   1740
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRTO"
               IdName          =   "UOBCHRTO"
               CmpStr          =   "TdBChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   5460
               TabIndex        =   14
               ToolTipText     =   "Enter To Bag Number"
               Top             =   1740
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOTO"
               IdName          =   "UOBNOTO"
               CmpStr          =   "TdBNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   4260
               TabIndex        =   20
               ToolTipText     =   "Enter To Design Size"
               Top             =   2310
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
               CmpStr          =   "BOdDmSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1740
               TabIndex        =   18
               ToolTipText     =   "Enter From Design Size"
               Top             =   2310
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "BOdDmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   2400
               TabIndex        =   19
               ToolTipText     =   "Enter From Design Suffix"
               Top             =   2310
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOODSFXFR"
               IdName          =   "UOODSFXFR"
               CmpStr          =   "BOdSfx >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4920
               TabIndex        =   21
               ToolTipText     =   "Enter To Design Suffix"
               Top             =   2310
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOODSFXTO"
               IdName          =   "UOODSFXTO"
               CmpStr          =   "BOdSfx <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5040
               TabIndex        =   26
               ToolTipText     =   "Enter To Order Year"
               Top             =   2595
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "BOdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2520
               TabIndex        =   23
               ToolTipText     =   "Enter From Order Year"
               Top             =   2595
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "BOdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1740
               TabIndex        =   22
               ToolTipText     =   "Enter From Order Tc"
               Top             =   2595
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "BOdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   2940
               TabIndex        =   24
               ToolTipText     =   "Enter From Order Character"
               Top             =   2595
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "BOdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4260
               TabIndex        =   25
               ToolTipText     =   "Enter To Order Tc"
               Top             =   2595
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "BOdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   5460
               TabIndex        =   27
               ToolTipText     =   "Enter To Order Character"
               Top             =   2595
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "BOdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5160
               TabIndex        =   31
               ToolTipText     =   "Enter To Order Serial"
               Top             =   2880
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "BOdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   2640
               TabIndex        =   29
               ToolTipText     =   "Enter From Order Serial"
               Top             =   2880
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "BOdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1740
               TabIndex        =   28
               ToolTipText     =   "Enter From Order Number"
               Top             =   2880
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "BOdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   4260
               TabIndex        =   30
               ToolTipText     =   "Enter To Order Number"
               Top             =   2880
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "BOdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   5040
               TabIndex        =   36
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   3165
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
               CmpStr          =   "TdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2520
               TabIndex        =   33
               ToolTipText     =   "Enter From Transaction Year"
               Top             =   3165
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
               CmpStr          =   "TdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1740
               TabIndex        =   32
               ToolTipText     =   "Enter From Transaction Tc"
               Top             =   3165
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
               CmpStr          =   "TdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2940
               TabIndex        =   34
               ToolTipText     =   "Enter From Transaction Character"
               Top             =   3165
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
               CmpStr          =   "TdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   4260
               TabIndex        =   35
               ToolTipText     =   "Enter To Transaction Tc"
               Top             =   3165
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
               CmpStr          =   "TdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   5460
               TabIndex        =   37
               ToolTipText     =   "Enter To Transaction Character"
               Top             =   3165
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
               CmpStr          =   "TdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   5160
               TabIndex        =   41
               ToolTipText     =   "Enter To Transaction Serial"
               Top             =   3450
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRTO"
               IdName          =   "UOTDSRTO"
               CmpStr          =   "TdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   2640
               TabIndex        =   39
               ToolTipText     =   "Enter From Transaction Serial"
               Top             =   3450
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRFR"
               IdName          =   "UOTDSRFR"
               CmpStr          =   "TdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   1740
               TabIndex        =   38
               ToolTipText     =   "Enter From Transaction Number"
               Top             =   3450
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOFR"
               IdName          =   "UOTNOFR"
               CmpStr          =   "TdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   4260
               TabIndex        =   40
               ToolTipText     =   "Enter To Transaction Number"
               Top             =   3450
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOTO"
               IdName          =   "UOTNOTO"
               CmpStr          =   "TdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   1740
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   300
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "TdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   1740
               TabIndex        =   56
               ToolTipText     =   "Enter Y for Acknowledged Bags / N for all the Unacknowledged bags/ leave Blank to see all the Bags "
               Top             =   5160
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
               Index           =   57
               Left            =   1740
               TabIndex        =   60
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   5445
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOBAGTYP"
               IdName          =   "UOBAGTYP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   2115
               TabIndex        =   116
               ToolTipText     =   "Show Open Bags (Yes / No / All)"
               Top             =   6080
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
               Index           =   60
               Left            =   6840
               TabIndex        =   120
               ToolTipText     =   "Show Production or Process Location"
               Top             =   600
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
               Index           =   71
               Left            =   1740
               TabIndex        =   61
               ToolTipText     =   "Show Production Bags (Yes / No)"
               Top             =   5750
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               DataType        =   4
               IdName          =   "WPRDBAGSYN"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Note: Selecting ‘Y’es may slow down the report heavily"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   38
               Left            =   2880
               TabIndex        =   122
               Top             =   5760
               Width           =   5535
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Only Prd Bags"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   105
               TabIndex        =   121
               Top             =   5770
               Width           =   1605
            End
            Begin VB.Label LblBagTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "BagType"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   105
               TabIndex        =   114
               ToolTipText     =   "Location"
               Top             =   5460
               Width           =   1035
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag AckwOpt"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   105
               TabIndex        =   113
               Top             =   5160
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
               Index           =   36
               Left            =   105
               TabIndex        =   112
               Top             =   300
               Width           =   1335
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
               Left            =   2580
               TabIndex        =   111
               Top             =   3450
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
               Left            =   5100
               TabIndex        =   110
               Top             =   3450
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
               Left            =   2460
               TabIndex        =   109
               Top             =   3165
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
               Left            =   4980
               TabIndex        =   108
               Top             =   3165
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
               Left            =   2880
               TabIndex        =   107
               Top             =   3165
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
               Left            =   5400
               TabIndex        =   106
               Top             =   3165
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
               Left            =   2580
               TabIndex        =   105
               Top             =   2880
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
               Left            =   5100
               TabIndex        =   104
               Top             =   2880
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
               Left            =   2460
               TabIndex        =   103
               Top             =   2595
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
               Left            =   4980
               TabIndex        =   102
               Top             =   2595
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
               Left            =   2880
               TabIndex        =   101
               Top             =   2595
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
               Left            =   5400
               TabIndex        =   100
               Top             =   2595
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
               Index           =   23
               Left            =   2340
               TabIndex        =   99
               Top             =   2310
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
               Index           =   22
               Left            =   4860
               TabIndex        =   98
               Top             =   2310
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
               Index           =   21
               Left            =   5400
               TabIndex        =   97
               Top             =   1740
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
               Index           =   20
               Left            =   2880
               TabIndex        =   96
               Top             =   1740
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
               Index           =   19
               Left            =   4620
               TabIndex        =   95
               Top             =   1740
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
               Index           =   18
               Left            =   2100
               TabIndex        =   94
               Top             =   1740
               Width           =   105
            End
            Begin VB.Label LblScpTime 
               BackStyle       =   0  'Transparent
               Caption         =   "Time"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   105
               TabIndex        =   93
               Top             =   4875
               Width           =   1605
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
               Index           =   16
               Left            =   105
               TabIndex        =   92
               Top             =   4590
               Width           =   1605
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
               Index           =   15
               Left            =   105
               TabIndex        =   91
               Top             =   4305
               Width           =   1605
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
               Index           =   14
               Left            =   105
               TabIndex        =   90
               Top             =   4020
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Dsg Ctg"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   105
               TabIndex        =   89
               Top             =   3735
               Width           =   1605
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
               Left            =   105
               TabIndex        =   88
               Top             =   3450
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
               Left            =   105
               TabIndex        =   87
               Top             =   3165
               Width           =   1605
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
               Index           =   10
               Left            =   105
               TabIndex        =   86
               Top             =   2880
               Width           =   1605
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
               Index           =   9
               Left            =   105
               TabIndex        =   85
               Top             =   2595
               Width           =   1725
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Sz/Sfx"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   105
               TabIndex        =   84
               Top             =   2310
               Width           =   1605
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
               Index           =   6
               Left            =   105
               TabIndex        =   83
               Top             =   2025
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Yy/Chr/No"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   105
               TabIndex        =   82
               Top             =   1740
               Width           =   1605
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
               Left            =   105
               TabIndex        =   81
               Top             =   1455
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Issuing Loc"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   105
               TabIndex        =   80
               Top             =   885
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Receiving Loc"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   105
               TabIndex        =   79
               Top             =   1170
               Width           =   1605
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
               Left            =   6780
               TabIndex        =   78
               Top             =   30
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
               Left            =   4260
               TabIndex        =   77
               Top             =   30
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
               Left            =   1740
               TabIndex        =   76
               Top             =   30
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpBv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepBv
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
Dim ms_BvDets As String         '4.1.2.0
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
    
Private Sub SetGroupSort()
  GRP_REP.Add "Issuing Loc", "TdFrBLoc", "", "TdFrBLoc", "hTdFrBLoc", "", "", "", "(Select LocDesc From Loc Where LocCoCd= TdCoCd and LocTyp= 'P' and LocCd = TdFrBLoc) "
  GRP_REP.Add "Issuing Prcs", "LocForTdFrBLoc.LocPrcs", "", "TdFrBLoc", "hTdFrBLoc", "", "", "", "(Select LocDesc From Loc Where LocCoCd= TdCoCd and LocTyp='R' and LocCd =LocForTdFrBLoc.LocPrcs ) "
  GRP_REP.Add "Receiving Loc", "TdToBLoc", "", "TdToBLoc", "hTdToBLoc", "", "", "", "(Select LocDesc From Loc Where LocCoCd= TdCoCd and LocTyp= 'P' and LocCd = TdToBLoc) "
  GRP_REP.Add "Receiving Prcs", "LocForTdToBLoc.LocPrcs", "", "TdToBLoc", "hTdToBLoc", "", "", "", "(Select LocDesc From Loc Where LocCoCd= TdCoCd and LocTyp='R' and LocCd = LocForTdToBLoc.LocPrcs) "
  '6.1
  GRP_REP.Add "Date", "convert(char(10),TDt,112)", "convert(varchar, TDt, " + CStr(DtFmtSlahVal()) + ")", "TDt", "hTDt", "", "", "", ""
  GRP_REP.Add "Dsg Cd", "BOdDmCd", "", "BOdDmCd", "hBOdDmCd", "DsgCtg,PrdCtg,(DsgCd)Colour", "", "", ""
  'Repair - repair bag fields included
  If UCase(ADC.MenuCd) = UCase("RprWIPBv") Then
    GRP_REP.Add "OrderNo", "BRepOdTc+'/'+BRepOdYy+'/'+BRepOdChr+'/'+Str(BRepOdNo)", "BRepOdTc+'/'+BRepOdYy+'/'+BRepOdChr+'/'+LTrim(Str(BRepOdNo))", "", "", "Customer", "", "", ""
    GRP_REP.Add "Order Sr", "BRepOdTc+'/'+BRepOdYy+'/'+BRepOdChr+'/'+Str(BRepOdNo)+'/'+Str(BRepOdSr)", "BRepOdTc+'/'+BRepOdYy+'/'+BRepOdChr+'/'+LTrim(Str(BRepOdNo))+'/'+LTrim(Str(BRepOdSr))", "", "", "OrderNo,Dsg Cd,DsgCtg,PrdCtg,Karat,(DsgCd)Colour,Customer", "", "", ""
  Else
    GRP_REP.Add "OrderNo", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+Str(BOdNo)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "", "", "Customer", "", "", ""
    GRP_REP.Add "Order Sr", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+Str(BOdNo)+'/'+Str(BOdSr)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))", "", "", "OrderNo,Dsg Cd,DsgCtg,PrdCtg,Karat,(DsgCd)Colour,Customer", "", "", ""
  End If
  
  GRP_REP.Add "Vch Sr", "TdTc+'/'+TdYy+'/'+TdChr+'/'+Str(TdNo)", "TdTc+'/'+TdYy+'/'+TdChr+'/'+LTrim(Str(TdNo))", "wVchNo", "hVchNo", "Date", "", "", ""
  GRP_REP.Add "Karat", "OdKt", "", "OdKt", "hOdKt", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = OdKt) "
  GRP_REP.Add "(DsgCd)Colour", "BOdDmCd+OdDmCol", "'('+OdDmCd+') '+OdDmCol", "BOdDmCd", "hBOdDmCd", "DsgCtg,PrdCtg,Dsg Cd", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
  
  GRP_REP.Add "DsgCtg", "DmCtg", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PMCd = DmCtg) "
  GRP_REP.Add "PrdCtg", "DmPrdCtg", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'PRDCTG' and PMCd = DmPrdCtg) "
  '******Geeta*****************Emr206****06/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "Customer", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd = OmCmCd) "
  
  'GRP_REP.Value(gltGroup1) = "(Detail)"

'*** Jay 2.13(CT) ***
  GRP_REP.Add "Bag Type", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", _
    "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", "", "", "Bag Typ+Chr", "", "", "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')"
  GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
  GRP_REP.Add "Bag Typ+Chr", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+BChr", _
    "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+'/ '+ BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
' **** Zubin 211 **** '
'  GRP_REP.Add "Bag Type", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end )", "", "", "Bag Typ+Chr", "", "", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)"
'  GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'  GRP_REP.Add "Bag Typ+Chr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+BChr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+ '/ ' + BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
' **** Zubin 211 **** '
'*** Jay 2.13(CT) ***
End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr

  '*** Report Sql ***
  '  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  '  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  '  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  '  TdFrBLoc, TdToBLoc, TdTc, TdYy, TdChr, TdNo, TdSr, TdBYy, TdBChr,
  '  TdBNo, TdBQty, TdPtQty, TdRjQty, TdRjPtQty, BOdDmCd, BOdSfx,
  '  BOdDmSz, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, TdBGrWt, ModTime, TDt,
  '  OmCmCd, OdKt
  '  From Txnd, Bag, OrdMst, OrdDsg, Txn, DsgMst where 1= 2
  '*** Report Sql ***
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String
  Dim wBAckwOpt As String
  
  '*** Jay 2.13(CT) ***
  Dim ws_BagJoin As String
  'Dim ws_FluteBagJoin As String   ' **** Zubin 211 **** '
  '*** Jay 2.13(CT) ***
  
  Dim wAddlCnd As String '****** Sachin 3.02.0 - New Fields
  
  Set Rep = Nothing
  Set ADC.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  'repair - repair field selection
  If UCase(ADC.MenuCd) = UCase("RprWIPBv") Then
    ADC("UOOMTCFR").CmpStr = "BRepOdTc >="
    ADC("UOOMYYFR").CmpStr = "BRepOdYy >="
    ADC("UOOMCHRFR").CmpStr = "BRepOdChr >="
    ADC("UOOMTCTO").CmpStr = "BRepOdTc <="
    ADC("UOOMYYTO").CmpStr = "BRepOdYy <="
    ADC("UOOMCHRTO").CmpStr = "BRepOdChr <="
    ADC("UOOMNOFR").CmpStr = "BRepOdNo >="
    ADC("UOODSRFR").CmpStr = "BRepOdSr >="
    ADC("UOOMNOTO").CmpStr = "BRepOdNo <="
    ADC("UOODSRTO").CmpStr = "BRepOdSr <="
  End If
   
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    
    If OptLoc(0).Value = True Then
      If ADC("UoIssLocFr") <> "" Then wCnd = wCnd + " and TdFrBLoc >= '" + ADC("UoIssLocFr") + "'"
      If ADC("UoIssLocTo") <> "" Then wCnd = wCnd + " and TdFrBLoc <= '" + ADC("UoIssLocTo") + "'"
      If ADC("UoIssLocSel") <> "" Then wCnd = wCnd + " and TdFrBLoc In (" + ADC("UoIssLocSel") + ")"
      If ADC("UoRecLocFr") <> "" Then wCnd = wCnd + " and TdToBLoc >= '" + ADC("UoRecLocFr") + "'"
      If ADC("UoRecLocTo") <> "" Then wCnd = wCnd + " and TdToBLoc <= '" + ADC("UoRecLocTo") + "'"
      If ADC("UoRecLocSel") <> "" Then wCnd = wCnd + " and TdToBLoc In (" + ADC("UoRecLocSel") + ")"
    ElseIf OptLoc(1).Value = True Then
      If ADC("UoIssLocFr") <> "" And ADC("UoIssLocTo") <> "" Then wCnd = wCnd + " and TdFrBLoc in (select LocCd from Loc where LocPrcs >= '" + ADC("UoIssLocFr") + "'"
      If ADC("UoIssLocFr") <> "" And ADC("UoIssLocTo") = "" Then wCnd = wCnd + " and TdFrBLoc in (select LocCd from Loc where LocPrcs >= '" + ADC("UoIssLocFr") + "')"
      If ADC("UoIssLocTo") <> "" And ADC("UoIssLocFr") <> "" Then wCnd = wCnd + " and LocPrcs <= '" + ADC("UoIssLocTo") + "')"
      If ADC("UoIssLocTo") <> "" And ADC("UoIssLocFr") = "" Then wCnd = wCnd + " and TdFrBLoc in (select LocCd from Loc where LocPrcs <= '" + ADC("UoIssLocTo") + "')"
      If ADC("UoIssLocSel") <> "" Then wCnd = wCnd + " and TdFrBLoc in (select LocCd from Loc where LocPrcs In (" + ADC("UoIssLocSel") + "))"
      If ADC("UoRecLocFr") <> "" And ADC("UoRecLocTo") <> "" Then wCnd = wCnd + " and TdToBLoc in (select LocCd from Loc where LocPrcs >= '" + ADC("UoRecLocFr") + "'"
      If ADC("UoRecLocFr") <> "" And ADC("UoRecLocTo") = "" Then wCnd = wCnd + " and TdToBLoc in (select LocCd from Loc where LocPrcs >= '" + ADC("UoRecLocFr") + "')"
      If ADC("UoRecLocTo") <> "" And ADC("UoRecLocFr") <> "" Then wCnd = wCnd + " and LocPrcs <= '" + ADC("UoRecLocTo") + "')"
      If ADC("UoRecLocTo") <> "" And ADC("UoRecLocFr") = "" Then wCnd = wCnd + " and TdToBLoc in (select LocCd from Loc where LocPrcs <= '" + ADC("UoRecLocTo") + "')"
      If ADC("UoRecLocSel") <> "" Then wCnd = wCnd + " and TdToBLoc in (select LocCd from Loc where LocPrcs In (" + ADC("UoRecLocSel") + "))"
    End If


  gs_ValidPrdLocSeq = Trim(moCn.GetFldVal("Select PDesc225 from param where PTyp= 'USR' and " + _
                                     "PMCd= '" + EmrFrmMDI.TXT_UCD + "'"))
  If ADC("UoYN") = "Y" Then
    wBAckwOpt = " And TdAck = 'Y' "
  Else
    If ADC("UoYN") = "N" Then
      wBAckwOpt = " And TdAck = 'N' "
    ElseIf ADC("UoYN") = "" Then
      wBAckwOpt = ""
    End If
  End If
  
  'If adc("UoYN") = "" Then wBAckwOpt = " "

    '*** Jay 2.13(CT) ***
    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                IIF(ADC("UoBagTyp") <> "", " and PValue= '" + ADC("UoBagTyp") + "' ", " ")
                        
    '  ' **** Zubin 211 (ws_FluteBagJoin added in wSqlStrg) **** '
    '  If UCase(adc("UOYN2")) = "N" Then
    '    ws_FluteBagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'N' "
    '  ElseIf UCase(adc("UOYN2")) = "Y" Then
    '    ws_FluteBagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'Y' "
    '  Else
    '    ws_FluteBagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' "
    '  End If
    '*** Jay 2.13(CT) ***
      
  '****** Sachin 3.02.0 - Additional Fields in Where Clause
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(ADC("UOCMCDFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(ADC("UOCMCDTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(ADC("UOCMCDSEL")) + ")", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(ADC("UODTFR")) <> "01/01/80", " TdDt >= '" + Trim(ADC("UODTFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(ADC("UODTTO")) <> "01/01/80", " TdDt <= '" + Trim(ADC("UODTTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Fields in Where Clause
 
  'Repair - Join for normal bag/repair bag
  Dim ws_OrdBagJoin As String, ws_PrdBag As String
  
  If UCase(ADC.MenuCd) = UCase("RprWIPBv") Then
    ws_OrdBagJoin = "Join OrdMst on OmCoCd= BCoCd and OmTc= BRepOdTc and OmYy= BRepOdYy " + _
               "and OmChr= BRepOdChr and OmNo= BRepOdNo" + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey= BPrtKey ", "") + _
               "Join OrdDsg on OdCoCd= BCoCd and OdTc= BRepOdTc and OdYy= BRepOdYy " + _
               "and OdChr= BRepOdChr and OdNo= BRepOdNo and OdSr= BRepOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "")
  Else
    ws_OrdBagJoin = "Join OrdMst on BOmIdNo=OmIdNo and BCoCd= OmCoCd and BOdTc= OmTc and BOdYy= OmYy " + _
               "and BOdChr= OmChr and BOdNo = OmNo " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=OmPrtKey ", "") + _
               "Join OrdDsg on BOdIdNo=OdIdNo and BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy " + _
               "and BOdChr= OdChr and BOdNo = OdNo And BOdSr = OdSr " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=OdPrtKey ", "")
  End If
  'kbs.136-selects only Production=Y records from txnd, it will execute only if PrdBag is Yes
  If ADC("wPrdBagsYN") = "Y" Then
    ws_PrdBag = "and exists (Select 'a' from Txnd ForPrd where ForPrd.TdByy = Txnd.TdByy and ForPrd.TdBChr = Txnd.TdBChr " + _
              "and ForPrd.TdBNo = Txnd.TdBNo and ForPrd.TdPrdYn = 'Y' and ForPrd.TdByLoc = Txnd.TdFrBLoc and " + _
              "ForPrd.TdYyKey < Txnd.TdYyKey and ForPrd.TdYyKey > isnull((Select max(TdYyKey) from Txnd BagIssueRecTxnd where " + _
              "BagIssueRecTxnd.TdByy = Txnd.TdByy And BagIssueRecTxnd.TdBChr = Txnd.TdBChr And BagIssueRecTxnd.TdBNo = Txnd.TdBNo and " + _
              "BagIssueRecTxnd.TdToBLoc = Txnd.TdFrBLoc and BagIssueRecTxnd.TdTc = 'BV' and BagIssueRecTxnd.TdKey < Txnd.TdKey),0)) "
  End If
  
  '****** Sachin 3.02 - Id fields in Joins
  'Repair - vParam join added to find repair CHR
  'kbs.136- ws_PrdBag added
  'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp=BOdDmTcTyp in 2 DsgMst join
  wSqlStrg = "Select " + grpflds + ", " + _
             IIF(OptLoc(0).Value = True, "TdFrBLoc", "(select LocPrcs from Loc where LocCd =TdFrBLoc and LocCoCd =TdCoCd)") + " as TdFrBLoc, " + _
             IIF(OptLoc(0).Value = True, "TdToBLoc", "(select LocPrcs from Loc where LocCd =TdToBLoc and LocCoCd =TdCoCd)") + " as TdToBLoc, " + _
             "TdTc, TdYy, TdChr, TdNo, TdSr, TdBYy, TdBChr, TdBNo, " + _
             "TdBQty, TdPtQty, TdRjQty, TdRjPtQty, BOdDmCd, BOdSfx, BOdDmSz, " + _
             IIF(UCase(ADC.MenuCd) = UCase("RprWIPBv"), "BRepOdTc", "BOdTc") + " as BOdTc," + IIF(UCase(ADC.MenuCd) = UCase("RprWIPBv"), "BRepOdYy", "BOdYy") + " as BOdYy, " + _
             IIF(UCase(ADC.MenuCd) = UCase("RprWIPBv"), "BRepOdChr", "BOdChr") + " as BOdChr," + IIF(UCase(ADC.MenuCd) = UCase("RprWIPBv"), "BRepOdNo", "BOdNo") + " as BOdNo, " + _
             IIF(UCase(ADC.MenuCd) = UCase("RprWIPBv"), "BRepOdSr", "BOdSr") + " as BOdSr," + IIF(UCase(ADC.MenuCd) = UCase("RprWIPBv"), "BRepGrWt", "TdBGrWt") + " as TdBGrWt, " + _
             " IsNull(Txnd.ModTime,0) as ModTime, TDt, OmCmCd, " + _
             "OdKt From Txnd " + _
             "Join Bag on BIdNo=TdBIdNo and BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo = TdBNo " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=TdPrtKey ", "") + ws_BagJoin + _
             ws_OrdBagJoin + _
             "Join Txn on " + IIF(UCase(ADC.MenuCd) = UCase("RprWIPBv"), "TDt >=OmDt And ", "") + " TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr " + _
             "and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey=TdPrtKey ", "") + _
             "Join DsgMst on DmTcTyp=BOdDmTcTyp and DmCd= BOdDmCd " + _
             "and DmSz = isnull((Select DmSz from DsgMst where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey=BPrtKey And ", "") + " DmTcTyp=BOdDmTcTyp and " + _
             "DmCd= BOdDmCd and DmSz= BOdDmSz), '') " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=BPrtKey ", "") + _
             "join Loc LocForTdFrBLoc on LocForTdFrBLoc.LocCoCd = TdCoCd and LocForTdFrBLoc.LocCd = TdFrBLoc " + _
             "join Loc LocForTdToBLoc on LocForTdToBLoc.LocCoCd = TdCoCd and LocForTdToBLoc.LocCd = TdToBLoc " + _
             "join vParam on vPCoCd =TdCoCd and vPTyp ='CHR' and vPMCd =TdTc and vPSCd =TdChr and vPValue3='" + IIF(UCase(ADC.MenuCd) = UCase("RprWIPBv"), "Y", "N") + "' " + _
             "where " + IIF(gs_Partition = ctCurrPrtn, " TdPrtKey='" + ctCurrPrtn + "' And ", "") + " TdTc in (Select PMCd from Param where PTyp= 'TC' and PValue= 'BV') " + wBAckwOpt + _
             wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + ws_PrdBag
             '(Index = Bag) '(Index=OrdMst) '(Index=OrdDsg) '(Index=Txn)  '(index=DsgMst)
             '" left outer join Loc on LocCoCd =TdCoCd and LocCd=TdToBLoc " + _
             'uma added " left outer join Loc on LocCoCd =TdCoCd and LocCd=TdFrBLoc "
    
  ADC.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added
  
  Rep.TxtFrIssLoc.SetText ADC("UoIssLocFr")
  Rep.TxtToIssLoc.SetText ADC("UoIssLocTo")
  Rep.TxtFrRecLoc.SetText ADC("UoRecLocFr")
  Rep.TxtToRecLoc.SetText ADC("UoRecLocTo")
  Rep.TxtFrDt.SetText ADC("UoDtFr")
  Rep.TxtToDt.SetText ADC("UoDtTo")
  
  '*** Jay 2.13(CT) ***
  If ADC("UoBagTyp") <> "" Then
     If ADC("UoBagTyp") = "F" Then
       Rep.TxtBagTypAndChr.SetText "(Flute Bags)"
     ElseIf ADC("UoBagTyp") = "P" Then
       Rep.TxtBagTypAndChr.SetText "(Primary Bags)"
     ElseIf ADC("UoBagTyp") = "C" Then
       Rep.TxtBagTypAndChr.SetText "(Component Bags)"
     End If
  End If
  
  'Sachin - 4.1.2.0
  If ADC("UOYN1") = "N" Then
    Rep.FldModTime.Suppress = True
    Rep.hdModTime.Suppress = True
  Else
    Rep.FldModTime.Suppress = False
    Rep.hdModTime.Suppress = False
  End If

  '  '***(Jen 2.11)
  '  If adc("UoYn2") <> "" Then Rep.TxtBagTypAndChr.SetText IIF(adc("UoYn2") = "Y", "(Flute Bags", "(Primary Bags") & ")"
  '  '***(Jen 2.11)
  '*** Jay 2.13(CT) ***

  CRV_REP.DisplayGroupTree = False
  Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******
  
Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
'Sachin - 4.1.2.0
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 1
    If .Value = Checked Then ADC("UoYN1") = "Y" Else ADC("UoYN1") = "N"
  End Select
End With
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
  gb_CoCdFor = True
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
    Set ADC.FirNKeyCtl = ADC("UoCoCdFr")
  Else
    Set ADC.FirNKeyCtl = ADC("UoIssLocFr")
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******

  Call InitProp(Me)
    
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"
  gs_TxnTcTyp = "BV"
   
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
   ADC("UOCOCDFR").Enabled = ws_MultiCoMod
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    
  '4.1.2.0
  ms_BvDets = moCn.GetFldVal("Select hShowBgHistDetsYN From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'")
  If Trim(ms_BvDets) = "Y" Or Trim(ms_BvDets) = "" Then
     ChkBoxArr(1).Visible = IIF(Trim(ms_BvDets) = "", False, True)
    LblScpTime.Visible = True: ADC("UOMODTIMEFR").Visible = True: ADC("UOMODTIMETO").Visible = True
  Else
    ChkBoxArr(1).Visible = False:
    ADC("UOMODTIMEFR") = 0#: ADC("UOMODTIMETO") = 0#
    LblScpTime.Visible = False: ADC("UOMODTIMEFR").Visible = False: ADC("UOMODTIMETO").Visible = False
  End If
 
  Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)

  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  gb_CoCdFor = True
  Select Case UCase(IdName)
    Case Is = UCase("UoCoCdFr")
    If Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'") Then _
          Cancel = True: ErrMsg = "Invalid Company Code": Exit Sub
     Me.Caption = GetFrmCaption(pv_NewValue, "Bag Movement")
    'kbs.136- checking valid option
    Case Is = UCase("wPrdBagsYN")
      Cancel = Not moCn.RecSeek(" Select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' ")
      If Cancel = True Then ErrMsg = "Invalid Production Bags Option. Enter(Y/N)": Exit Sub
        
  End Select
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = True '****** Sachin 2.14.0 - [Multi-Company]
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"

  Select Case UCase(IdName)
    Case Is = UCase("UOOMTCFR")
      If UCase(ADC.MenuCd) = UCase("RprWIPBv") Then HlpList.PMCd "TC", "'JT'":  Exit Sub
    Case Is = UCase("UOOMTCTO")
      If UCase(ADC.MenuCd) = UCase("RprWIPBv") Then HlpList.PMCd "TC", "'JT'":  Exit Sub
    
    Case Is = UCase("UOISSLOCFR"), UCase("UOISSLOCTO"), UCase("UOISSLOCSEL"), UCase("UORECLOCFR"), UCase("UORECLOCTO"), UCase("UORECLOCSEL")
      If ADC("UOLOCTYPFR") = "R" Then gs_LocTyp = "'R'"
      If ADC("UOLOCTYPFR") = "P" Then gs_LocTyp = "'P'"
    'kbs.136-calling help
    Case Is = UCase("wPrdBagsYN")
      Call HlpList.PMCd("YN")
  End Select
   
  SetProp Me, IdName, When
  
  
End Sub



' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptBagAck_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN") = "Y"
Case Is = 1
  ADC("UoYN") = "N"
Case Is = 2
  ADC("UoYN") = ""
End Select
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptBagAck_GotFocus(Index As Integer)
  DispMsg FraOptBagAck.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  
Private Sub OptLoc_GotFocus(Index As Integer)
  DispMsg FraLoc.ToolTipText, etInfo
Select Case Index
Case Is = 0
  ADC("UOISSLOCFR") = "":   ADC("UOISSLOCTO") = "":   ADC("UOISSLOCsel") = ""
  ADC("UORECLOCFR") = "":   ADC("UORECLOCTO") = "":   ADC("UORECLOCSEL") = ""
  ADC("UoLocTypFr") = "P"
Case Is = 1
  ADC("UOISSLOCFR") = "":   ADC("UOISSLOCTO") = "":   ADC("UOISSLOCsel") = ""
  ADC("UORECLOCFR") = "":   ADC("UORECLOCTO") = "":   ADC("UORECLOCSEL") = ""
  ADC("UoLocTypFr") = "R"
End Select

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
  ADC("UoCoCdFr") = gs_CoCd
 
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(ADC("UOCOCDFR"), "Bag Movement")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  '*** Jay 2.13(CT) ***
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + ctSelfCoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    ADC("UoBagTyp").Visible = False
    ADC("UoBagTyp") = "P"
  End If
  Set wHeadChk = Nothing
  '*** Jay 2.13(CT) ***
    
'*** Jay 2.13(CT) ***
'  '*** (Jen 2.11)
'  If moCn.GetFldVal("Select hFluteBagYN from Head where HCoCd= '" + adc("UoCoCdFr") + "' and " + _
'                                                    "HCd= '" + ctSelfCmCd + "' ") = "N" Then
'    LblPrimFluteBags.Visible = False
'    adc("UoYN2").Visible = False
'    adc("UOYN2") = "N"
'  End If
'  '*** (Jen 2.11)
'*** Jay 2.13(CT) ***

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
If ADC("UoYN") = "Y" Then
  OptBagAck(0).Value = True
ElseIf ADC("UoYN") = "N" Then
  OptBagAck(1).Value = True
ElseIf ADC("UoYN") = "" Then
  OptBagAck(2).Value = True
End If
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

'Sachin - 4.1.2.0
ADC("UOYN1") = IIF(Trim(ms_BvDets) = "N", "N", "Y")
If ADC("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked

  If ADC("UOLOCTYPFR") = "P" Then
    OptLoc(0).Value = True
  ElseIf ADC("UOLOCTYPFR") = "R" Then
    OptLoc(1).Value = True
  End If
  'kbs.136- default value for Only PrdBag is No
  ADC("wPrdBagsYN") = "N"

End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + ADC("UoCoCdFr") + "' and " + _
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
